"""Reproduce the pre-perf-rework baseline on the OLD pyfastexcel API.

``perf_memory.py`` cannot run against code from before the perf rework: the wire
module, ``NativeExcelClient`` and ABI-v2 it depends on were all introduced by
that rework.  This probe fills the gap.  It reproduces perf_memory.py's
50000x30 four-style ``StreamWriter`` workload using only the *old* public API
(``row_append`` / ``create_row`` / ``read_lib_and_create_excel``) and emits the
**same JSON schema** ``perf_memory.py`` produces, so ``plot_perf.py`` can plot a
baseline report alongside current ones.

Each sample runs in a fresh subprocess so peak RSS is independent between
samples, exactly like ``perf_memory.py``.

Reproduction workflow
---------------------
Run this against a checkout of a pre-rework commit in an isolated worktree::

    # 1. Worktree at the commit just before the perf rework.
    git worktree add --detach /tmp/pfx-baseline <pre-rework-commit>

    # 2. Build that commit's native library.  If the toolchain reports a
    #    GOROOT/version mismatch, unset the stale GOROOT so the pinned Go
    #    toolchain uses its own bundled one:
    (cd /tmp/pfx-baseline && env -u GOROOT make build)

    # 3. Copy this probe in and run it there (its directory must be the old
    #    package root so ``import pyfastexcel`` resolves to the old code).
    cp benchmark/baseline_probe.py /tmp/pfx-baseline/
    (cd /tmp/pfx-baseline && uv run python baseline_probe.py \
        --rows 50000 --cols 30 --repeat 3 --output baseline.json)

    # 4. Clean up.
    git worktree remove --force /tmp/pfx-baseline

Then compare with a current run::

    uv run python benchmark/perf_memory.py --output new.json
    uv run python benchmark/plot_perf.py baseline.json new.json \
        --label "v1 baseline" "v2 pfx2"
"""

from __future__ import annotations

import argparse
import json
import platform
import statistics
import subprocess
import sys
import time
from pathlib import Path

# The probe must live in the old package root (see the module docstring), so its
# own directory is the import root for the old ``pyfastexcel`` package.
ROOT = Path(__file__).resolve().parent


def _peak_rss_bytes() -> int:
    import resource

    rss = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    # macOS reports bytes; Linux and the BSDs report KiB.
    return int(rss if sys.platform == 'darwin' else rss * 1024)


def _worker(rows: int, cols: int) -> dict:
    from pyfastexcel import CustomStyle, StreamWriter

    class BenchmarkWriter(StreamWriter):
        style_0 = CustomStyle(font_color='111111')
        style_1 = CustomStyle(font_color='222222', fill_color='DDDDDD')
        style_2 = CustomStyle(font_bold=True, font_color='333333')
        style_3 = CustomStyle(font_color='444444', number_format='0.00')

    writer = BenchmarkWriter()
    style_names = ('style_0', 'style_1', 'style_2', 'style_3')

    started = time.perf_counter()
    for row in range(rows):
        for col in range(cols):
            writer.row_append(row * cols + col, style=style_names[col % len(style_names)])
        writer.create_row()
    built = time.perf_counter()

    xlsx = writer.read_lib_and_create_excel()
    exported = time.perf_counter()

    return {
        'rows': rows,
        'cols': cols,
        'cells': rows * cols,
        'build_seconds': built - started,
        'export_seconds': exported - built,
        'total_seconds': exported - started,
        'peak_rss_bytes': _peak_rss_bytes(),
        'xlsx_bytes': len(xlsx),
    }


def _run_sample(rows: int, cols: int) -> dict:
    completed = subprocess.run(
        [
            sys.executable,
            str(Path(__file__).resolve()),
            '--worker',
            '--rows',
            str(rows),
            '--cols',
            str(cols),
        ],
        cwd=str(ROOT),
        check=True,
        capture_output=True,
        text=True,
    )
    return json.loads(completed.stdout)


def _command_output(*command: str) -> str | None:
    try:
        return subprocess.check_output(command, cwd=str(ROOT), text=True).strip()
    except (OSError, subprocess.CalledProcessError):
        return None


def _summary(samples: list[dict], key: str) -> dict:
    values = [float(sample[key]) for sample in samples]
    return {
        'mean': statistics.mean(values),
        'min': min(values),
        'max': max(values),
        'stdev': statistics.stdev(values) if len(values) > 1 else 0.0,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--rows', type=int, default=50_000)
    parser.add_argument('--cols', type=int, default=30)
    parser.add_argument('--repeat', type=int, default=3)
    parser.add_argument('--output', type=Path)
    parser.add_argument('--worker', action='store_true', help=argparse.SUPPRESS)
    args = parser.parse_args()

    if args.worker:
        print(json.dumps(_worker(args.rows, args.cols)))
        return

    samples = [_run_sample(args.rows, args.cols) for _ in range(args.repeat)]
    rss = [int(sample['peak_rss_bytes']) for sample in samples]
    report = {
        'environment': {
            'platform': platform.platform(),
            'python': platform.python_version(),
            'git_commit': _command_output('git', 'rev-parse', 'HEAD'),
            'note': (
                'Baseline probe (old pre-rework API, ABI-v1 export via '
                'read_lib_and_create_excel). Not produced by perf_memory.py.'
            ),
        },
        'workload': {
            'schema_version': 1,
            'rows': args.rows,
            'cols': args.cols,
            'cells': args.rows * args.cols,
            'styles': 4,
            'wire': 'v1',
            'destination': 'bytes',
            'repeat': args.repeat,
            'writer_engine': 'StreamWriter',
            'sheet_count': 1,
            'no_style': False,
            'value_pattern': 'row-major sequential integers',
            'formula_cells': 0,
            'style_pattern': 'column modulo four fixed styles',
        },
        'samples': samples,
        'summary': {
            'build_seconds': _summary(samples, 'build_seconds'),
            'export_seconds': _summary(samples, 'export_seconds'),
            'total_seconds': _summary(samples, 'total_seconds'),
            'peak_rss_bytes': {'mean': statistics.mean(rss), 'min': min(rss), 'max': max(rss)},
        },
    }

    rendered = json.dumps(report, indent=2)
    print(rendered)
    if args.output:
        args.output.write_text(rendered + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
