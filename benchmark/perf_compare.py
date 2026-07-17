"""
Legacy broad-case throughput benchmark.

This script is retained for its Workbook/StreamWriter size matrix, but its
``perf_baseline.json`` files are not release evidence because they lack native
binary provenance and fresh-process peak RSS. Use ``perf_memory.py`` for
before/after acceptance measurements.

benchmark/perf_compare.py

Run once before changes (saves results to perf_baseline.json),
then run again after changes to print the comparison.

Usage:
    python benchmark/perf_compare.py          # first run → saves baseline
    python benchmark/perf_compare.py          # subsequent runs → shows diff
    python benchmark/perf_compare.py --reset  # clears saved baseline
"""

from __future__ import annotations

import argparse
import json
import os
import statistics
import sys
import time
import timeit

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from example import prepare_example_data  # noqa: E402
from pyfastexcel import StreamWriter, Workbook  # noqa: E402
from pyfastexcel.driver import NativeExcelClient  # noqa: E402
from pyfastexcel.wire import encode_payload  # noqa: E402

REPEAT = 5
CASES = [
    (500, 10),
    (2000, 20),
    (10000, 30),
    (50000, 30),
]
BASELINE_FILE = os.path.join(os.path.dirname(__file__), 'perf_baseline.json')


# ---------------------------------------------------------------------------
# Benchmark functions
# ---------------------------------------------------------------------------


def bench_workbook(data: list[dict]) -> None:
    wb = Workbook()
    ws = wb['Sheet1']
    for i, record in enumerate(data):
        ws[i] = list(record.values())
    wb.read_lib_and_create_excel()


def bench_stream(data: list[dict]) -> None:
    class _W(StreamWriter):
        def run(self) -> None:
            for row in self.data:
                for v in row.values():
                    self.row_append(v)
                self.create_row()
            self.read_lib_and_create_excel()

    _W(data).run()


def measure_components(
    data: list[dict], repeat: int = 3
) -> tuple[dict[str, dict[str, float]], str, int]:
    """
    Break down the time inside read_lib_and_create_excel into components.

    Components:
      prepare        – Python-side data and style metadata assembly
      wire_enc       – the production JSON or PFX2 encoder
      native_export  – the production ABI call, result copy and pointer cleanup

    ABI detection deliberately goes through ``NativeExcelClient``. Calling an
    unknown C function with a guessed signature can corrupt the process before
    Python has an opportunity to catch an exception.
    """
    timings: dict[str, list[float]] = {
        'prepare': [],
        'wire_enc': [],
        'native_export': [],
    }

    for _ in range(repeat):
        wb = Workbook()
        ws = wb['Sheet1']
        for i, record in enumerate(data):
            ws[i] = list(record.values())
        native = NativeExcelClient(wb._read_lib(None))

        # ── prepare ──────────────────────────────────────────────────────────
        t0 = time.perf_counter()
        export_data = wb._build_export_data()
        t1 = time.perf_counter()

        # ── Production wire encode ────────────────────────────────────────────
        wire_data = encode_payload(export_data, force_json=not native.supports_v2_export)
        t2 = time.perf_counter()

        # ── Production native export, copy and free ───────────────────────────
        native.export_bytes(wire_data, 1)
        t3 = time.perf_counter()
        transport = 'raw bytes' if native.supports_v2_export else 'base64'
        api_label = f'v{native.abi_version} ({transport})'

        timings['prepare'].append(t1 - t0)
        timings['wire_enc'].append(t2 - t1)
        timings['native_export'].append(t3 - t2)

    return (
        {k: {'mean': statistics.mean(v), 'min': min(v)} for k, v in timings.items()},
        api_label,
        len(wire_data),
    )


# ---------------------------------------------------------------------------
# Timing helper
# ---------------------------------------------------------------------------


def time_fn(fn, data: list[dict], repeat: int = REPEAT) -> dict:
    times = timeit.repeat(lambda: fn(data), number=1, repeat=repeat)
    return {
        'mean': statistics.mean(times),
        'min': min(times),
        'max': max(times),
        'stdev': statistics.stdev(times) if repeat > 1 else 0.0,
    }


# ---------------------------------------------------------------------------
# Reporting
# ---------------------------------------------------------------------------


def _ms(s: float) -> str:
    return f'{s * 1000:8.1f} ms'


def print_row(label: str, r: dict, baseline: dict | None = None) -> None:
    mean = _ms(r['mean'])
    min_ = _ms(r['min'])
    max_ = _ms(r['max'])
    line = f'  {label:<18}  mean={mean}  min={min_}  max={max_}'
    if baseline:
        pct = (r['mean'] - baseline['mean']) / baseline['mean'] * 100
        arrow = '↓' if pct < 0 else '↑'
        line += f'  {arrow} {abs(pct):5.1f}% vs baseline'
    print(line)


def print_component_row(label: str, mean_s: float, total_s: float) -> None:
    pct = mean_s / total_s * 100 if total_s > 0 else 0
    print(f'    {label:<14} {_ms(mean_s)}  ({pct:4.1f}% of total)')


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument('--reset', action='store_true', help='Clear saved baseline')
    parser.add_argument('--repeat', type=int, default=REPEAT)
    args = parser.parse_args()

    if args.reset and os.path.exists(BASELINE_FILE):
        os.remove(BASELINE_FILE)
        print('Baseline cleared.')
        return

    baseline: dict = {}
    if os.path.exists(BASELINE_FILE):
        with open(BASELINE_FILE) as f:
            baseline = json.load(f)
        print(f'Loaded baseline from {BASELINE_FILE}')

    current_results: dict = {}

    separator = '=' * 72
    print(f'\n{separator}')
    print('pyfastexcel performance benchmark')
    print(f'repeat={args.repeat} per case')
    print(separator)

    for rows, cols in CASES:
        data = prepare_example_data(rows=rows, cols=cols)
        key = f'{rows}r_{cols}c'
        print(f'\n[ {rows} rows × {cols} cols ]')

        wb_r = time_fn(bench_workbook, data, args.repeat)
        sw_r = time_fn(bench_stream, data, args.repeat)

        print_row('Workbook', wb_r, baseline.get(key, {}).get('workbook'))
        print_row('StreamWriter', sw_r, baseline.get(key, {}).get('stream'))

        current_results[key] = {'workbook': wb_r, 'stream': sw_r}

    # ── component breakdown for medium-sized dataset ──────────────────────
    separator = '─' * 72
    print(f'\n{separator}')
    print('Component breakdown  (10 000 rows × 30 cols, 3 repeats)')
    print(separator)
    data_med = prepare_example_data(rows=10_000, cols=30)
    comp, api_ver, wire_bytes = measure_components(data_med, repeat=3)
    total = sum(v['mean'] for v in comp.values())
    print(f'  API version detected : {api_ver}')
    print(f'  Wire payload size    : {wire_bytes / 1024:.1f} KB')
    for label, v in comp.items():
        print_component_row(label, v['mean'], total)
    total_label = 'total'
    print(f'    {total_label:<14} {_ms(total)}')

    # ── save baseline if none exists ──────────────────────────────────────
    if not baseline:
        with open(BASELINE_FILE, 'w') as f:
            json.dump(current_results, f, indent=2)
        print(f'\nBaseline saved to {BASELINE_FILE}')
        print('Re-run after optimizations to see improvements.')
    else:
        print('\nOptimization summary (Workbook, mean time):')
        for key in current_results:
            if key in baseline:
                old = baseline[key]['workbook']['mean']
                new = current_results[key]['workbook']['mean']
                pct = (old - new) / old * 100
                label = key.replace('_', ' × ')
                print(f'  {label:<14}  {_ms(old)} → {_ms(new)}  ({pct:+.1f}%)')


if __name__ == '__main__':
    main()
