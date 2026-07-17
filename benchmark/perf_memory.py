"""
Reproducible wall-time and peak-RSS benchmark for pyfastexcel.

Each sample runs in a fresh subprocess.  This keeps peak RSS independent
between samples and deliberately avoids ``tracemalloc``, whose allocation
hooks distort the hot Python loop.

Examples
--------
    uv run python benchmark/perf_memory.py --rows 50000 --cols 30
    uv run python benchmark/perf_memory.py --wire json --output before.json
    uv run python benchmark/perf_memory.py --compare before.json

"""

from __future__ import annotations

import argparse
import hashlib
import importlib.metadata
import json
import os
import platform
import shutil
import statistics
import subprocess  # nosec B404
import sys
import tempfile
import time
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
HARNESS = Path(__file__).resolve()
WORKLOAD_SCHEMA_VERSION = 1
WORKLOAD_DESCRIPTION = {
    'writer_engine': 'StreamWriter',
    'sheet_count': 1,
    'no_style': False,
    'value_pattern': 'row-major sequential integers',
    'formula_cells': 0,
    'style_pattern': 'column modulo four fixed styles',
}
METADATA_TOOLS = frozenset({'git', 'go'})


def _package_version(distribution: str) -> str | None:
    try:
        return importlib.metadata.version(distribution)
    except importlib.metadata.PackageNotFoundError:
        return None


def _cpu_name() -> str | None:
    if processor := platform.processor():
        return processor
    try:
        for line in Path('/proc/cpuinfo').read_text(encoding='utf-8').splitlines():
            if line.lower().startswith('model name'):
                return line.partition(':')[2].strip()
    except OSError:  # pragma: no cover - platform dependent
        pass
    return None


def _file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open('rb') as file:
        for chunk in iter(lambda: file.read(1024 * 1024), b''):
            digest.update(chunk)
    return digest.hexdigest()


def _peak_rss_bytes() -> int | None:
    if sys.platform == 'win32':
        return _windows_peak_rss_bytes()

    try:
        import resource
    except ImportError:  # pragma: no cover - platform dependent
        return None

    rss = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    # macOS reports bytes; Linux and the BSDs report KiB.
    return int(rss if sys.platform == 'darwin' else rss * 1024)


def _windows_peak_rss_bytes() -> int | None:  # pragma: no cover - Windows only
    import ctypes
    from ctypes import wintypes

    class ProcessMemoryCounters(ctypes.Structure):
        _fields_ = [
            ('cb', wintypes.DWORD),
            ('PageFaultCount', wintypes.DWORD),
            ('PeakWorkingSetSize', ctypes.c_size_t),
            ('WorkingSetSize', ctypes.c_size_t),
            ('QuotaPeakPagedPoolUsage', ctypes.c_size_t),
            ('QuotaPagedPoolUsage', ctypes.c_size_t),
            ('QuotaPeakNonPagedPoolUsage', ctypes.c_size_t),
            ('QuotaNonPagedPoolUsage', ctypes.c_size_t),
            ('PagefileUsage', ctypes.c_size_t),
            ('PeakPagefileUsage', ctypes.c_size_t),
        ]

    counters = ProcessMemoryCounters()
    counters.cb = ctypes.sizeof(counters)
    process = ctypes.windll.kernel32.GetCurrentProcess()
    ok = ctypes.windll.psapi.GetProcessMemoryInfo(
        process,
        ctypes.byref(counters),
        counters.cb,
    )
    return int(counters.PeakWorkingSetSize) if ok else None


def _worker(
    rows: int,
    cols: int,
    wire: str,
    destination: str,
) -> dict[str, int | float | str | None]:
    sys.path.insert(0, str(ROOT))
    os.environ['PYFASTEXCEL_WIRE'] = wire

    from pyfastexcel import CustomStyle, StreamWriter
    from pyfastexcel.driver import NativeExcelClient
    from pyfastexcel.wire import WIRE_MAGIC, encode_payload

    class BenchmarkWriter(StreamWriter):
        style_0 = CustomStyle(font_color='111111')
        style_1 = CustomStyle(font_color='222222', fill_color='DDDDDD')
        style_2 = CustomStyle(font_bold=True, font_color='333333')
        style_3 = CustomStyle(font_color='444444', number_format='0.00')

    writer = BenchmarkWriter()
    style_names = ('style_0', 'style_1', 'style_2', 'style_3')
    native = NativeExcelClient(writer._read_lib(None))
    if wire == 'msgpack' and not native.supports_v2_export:
        raise RuntimeError('MessagePack benchmark requires a freshly built ABI-v2 library')
    if destination == 'file' and not native.supports_direct_file_export:
        raise RuntimeError('File benchmark requires ExportToFileV2 support')
    native_path = Path(native.library._name).resolve()

    started = time.perf_counter()
    for row in range(rows):
        for col in range(cols):
            writer.row_append(
                row * cols + col,
                style=style_names[col % len(style_names)],
            )
        writer.create_row()
    built = time.perf_counter()

    export_data = writer._build_export_data()
    payload = encode_payload(export_data, force_json=not native.supports_v2_export)
    effective_wire = 'msgpack' if payload.startswith(WIRE_MAGIC) else 'json'
    if wire == 'msgpack' and effective_wire != 'msgpack':
        raise RuntimeError('benchmark workload unexpectedly fell back to the JSON wire')

    if destination == 'file':
        with tempfile.TemporaryDirectory() as tmp_dir:
            output = Path(tmp_dir, 'benchmark-output.bin')
            native.export_to_file(payload, str(output), 1)
            xlsx_bytes = output.stat().st_size
    else:
        xlsx_bytes = len(native.export_bytes(payload, 1))
    exported = time.perf_counter()

    return {
        'rows': rows,
        'cols': cols,
        'cells': rows * cols,
        'wire': wire,
        'effective_wire': effective_wire,
        'destination': destination,
        'build_seconds': built - started,
        'export_seconds': exported - built,
        'total_seconds': exported - started,
        'peak_rss_bytes': _peak_rss_bytes(),
        'xlsx_bytes': xlsx_bytes,
        'wire_bytes': len(payload),
        'native_abi': native.abi_version,
        'native_library': str(native_path),
        'native_library_sha256': _file_sha256(native_path),
    }


def _run_sample(rows: int, cols: int, wire: str, destination: str) -> dict[str, Any]:
    # Preserve virtual-environment launcher symlinks while making PATH lookup
    # impossible for the worker executable.
    python_executable = str(Path(sys.executable).absolute())
    command = [
        python_executable,
        str(HARNESS),
        '--worker',
        '--rows',
        str(rows),
        '--cols',
        str(cols),
        '--wire',
        wire,
        '--destination',
        destination,
    ]
    # The executable and harness are absolute, validated values are separate
    # arguments, and no shell interprets the command.
    completed = subprocess.run(  # nosec B603
        command,
        cwd=ROOT,
        check=True,
        capture_output=True,
        text=True,
    )
    return json.loads(completed.stdout)


def _require_current_native_build() -> None:
    """Refuse release evidence when the checked-in sources are newer than the library."""
    make_executable = shutil.which('make')
    if make_executable is None:
        raise RuntimeError('make is required to verify the native benchmark build')
    try:
        resolved_make = str(Path(make_executable).resolve(strict=True))
        # ``make`` is resolved to an absolute executable and receives only
        # module-owned arguments without shell parsing.
        completed = subprocess.run(  # nosec B603
            [resolved_make, '-q', 'build'],
            cwd=ROOT,
            capture_output=True,
            text=True,
        )
    except OSError as exc:  # pragma: no cover - depends on the host toolchain
        raise RuntimeError('make is required to verify the native benchmark build') from exc
    if completed.returncode == 1:
        raise RuntimeError('native library is stale; run `make build` before benchmarking')
    if completed.returncode != 0:
        detail = completed.stderr.strip() or completed.stdout.strip()
        raise RuntimeError(f'could not verify native build freshness: {detail}')


def _command_output(tool: str, *arguments: str) -> str | None:
    if tool not in METADATA_TOOLS:
        raise ValueError(f'unsupported metadata tool: {tool}')
    executable = shutil.which(tool)
    if executable is None:
        return None
    try:
        resolved_executable = str(Path(executable).resolve(strict=True))
        # The executable is allowlisted and absolute. All arguments come from
        # fixed call sites in this benchmark and are never parsed by a shell.
        return subprocess.check_output(  # nosec B603
            [resolved_executable, *arguments],
            cwd=ROOT,
            text=True,
        ).strip()
    except (OSError, subprocess.CalledProcessError):
        return None


def _summarize(samples: list[dict[str, Any]]) -> dict[str, Any]:
    timing_keys = ('build_seconds', 'export_seconds', 'total_seconds')
    summary: dict[str, Any] = {}
    for key in timing_keys:
        values = [float(sample[key]) for sample in samples]
        summary[key] = {
            'mean': statistics.mean(values),
            'min': min(values),
            'max': max(values),
            'stdev': statistics.stdev(values) if len(values) > 1 else 0.0,
        }

    rss_values = [
        int(sample['peak_rss_bytes']) for sample in samples if sample['peak_rss_bytes'] is not None
    ]
    summary['peak_rss_bytes'] = (
        {
            'mean': statistics.mean(rss_values),
            'min': min(rss_values),
            'max': max(rss_values),
        }
        if rss_values
        else None
    )
    return summary


def _comparison(current: dict[str, Any], baseline: dict[str, Any]) -> dict[str, float]:
    for key in (
        'schema_version',
        'rows',
        'cols',
        'cells',
        'styles',
        *WORKLOAD_DESCRIPTION,
    ):
        if current['workload'][key] != baseline['workload'][key]:
            raise ValueError(f'cannot compare different {key} workloads')
    current_destination = current['workload'].get('destination', 'bytes')
    baseline_destination = baseline['workload'].get('destination', 'bytes')
    if current_destination != baseline_destination:
        raise ValueError('cannot compare byte-return and direct-file destinations')

    changes: dict[str, float] = {}
    for key in ('build_seconds', 'export_seconds', 'total_seconds'):
        old = baseline['summary'][key]['mean']
        new = current['summary'][key]['mean']
        changes[f'{key}_percent'] = (new - old) / old * 100

    old_rss = baseline['summary'].get('peak_rss_bytes')
    new_rss = current['summary'].get('peak_rss_bytes')
    if old_rss and new_rss:
        changes['peak_rss_percent'] = (new_rss['mean'] - old_rss['mean']) / old_rss['mean'] * 100
    return changes


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--rows', type=int, default=50_000)
    parser.add_argument('--cols', type=int, default=30)
    parser.add_argument('--repeat', type=int, default=3)
    parser.add_argument('--wire', choices=('msgpack', 'json'), default='msgpack')
    parser.add_argument('--destination', choices=('bytes', 'file'), default='bytes')
    parser.add_argument('--output', type=Path)
    parser.add_argument('--compare', type=Path)
    parser.add_argument('--worker', action='store_true', help=argparse.SUPPRESS)
    return parser.parse_args()


def main() -> None:
    args = _parse_args()
    if args.rows < 1 or args.cols < 1 or args.repeat < 1:
        raise SystemExit('rows, cols, and repeat must all be positive')

    if args.worker:
        print(json.dumps(_worker(args.rows, args.cols, args.wire, args.destination)))
        return

    _require_current_native_build()
    samples = [
        _run_sample(args.rows, args.cols, args.wire, args.destination) for _ in range(args.repeat)
    ]
    native_fingerprints = {
        (
            sample['native_abi'],
            sample['native_library'],
            sample['native_library_sha256'],
        )
        for sample in samples
    }
    if len(native_fingerprints) != 1:
        raise RuntimeError('native library changed between benchmark subprocesses')
    effective_wires = {sample['effective_wire'] for sample in samples}
    if len(effective_wires) != 1:
        raise RuntimeError('effective wire format changed between benchmark subprocesses')
    report: dict[str, Any] = {
        'environment': {
            'platform': platform.platform(),
            'cpu': _cpu_name(),
            'python': platform.python_version(),
            'python_dependencies': {
                'msgspec': _package_version('msgspec'),
                'pydantic': _package_version('pydantic'),
            },
            'go': _command_output('go', 'version'),
            'go_dependencies': {
                'excelize': _command_output(
                    'go', 'list', '-m', '-f', '{{.Version}}', 'github.com/xuri/excelize/v2'
                ),
                'msgpack': _command_output(
                    'go', 'list', '-m', '-f', '{{.Version}}', 'github.com/vmihailenco/msgpack/v5'
                ),
            },
            'git_commit': _command_output('git', 'rev-parse', 'HEAD'),
            'git_dirty': bool(_command_output('git', 'status', '--porcelain')),
            'harness_sha256': _file_sha256(HARNESS),
            'native_abi': samples[0]['native_abi'],
            'native_library': samples[0]['native_library'],
            'native_library_sha256': samples[0]['native_library_sha256'],
            'native_build_info': _command_output(
                'go', 'version', '-m', samples[0]['native_library']
            ),
        },
        'workload': {
            'schema_version': WORKLOAD_SCHEMA_VERSION,
            'rows': args.rows,
            'cols': args.cols,
            'cells': args.rows * args.cols,
            'styles': 4,
            'wire': args.wire,
            'wire_requested': args.wire,
            'wire_effective': samples[0]['effective_wire'],
            'destination': args.destination,
            'repeat': args.repeat,
            **WORKLOAD_DESCRIPTION,
        },
        'samples': samples,
        'summary': _summarize(samples),
    }

    if args.compare:
        with args.compare.open(encoding='utf-8') as baseline_file:
            report['change_from_baseline'] = _comparison(report, json.load(baseline_file))

    rendered = json.dumps(report, indent=2)
    print(rendered)
    if args.output:
        args.output.write_text(f'{rendered}\n', encoding='utf-8')


if __name__ == '__main__':
    main()
