"""
Measure pyfastexcel performance before and after optimization.

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
import base64
import ctypes
import json
import os
import statistics
import sys
import time
import timeit

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import msgspec  # noqa: E402

from example import prepare_example_data  # noqa: E402
from pyfastexcel import StreamWriter, Workbook  # noqa: E402
from pyfastexcel.manager import StyleManager  # noqa: E402
from pyfastexcel.validators import TableFinalValidation  # noqa: E402

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


def measure_components(data: list[dict], repeat: int = 3) -> dict[str, float]:
    """
    Break down the time inside read_lib_and_create_excel into components.

    Components:
      prepare   – Python-side data assembly (_create_style, _transfer_to_dict)
      json_enc  – msgspec.json.encode
      ffi       – actual Go shared-library call
      post      – bytes retrieval + freeing the C pointer

    Works with BOTH the old API (Export returns base64 *C.char)
    and the new API (Export returns raw bytes with a length out-param).
    """
    timings: dict[str, list[float]] = {
        'prepare': [],
        'json_enc': [],
        'ffi': [],
        'post': [],
    }

    for _ in range(repeat):
        wb = Workbook()
        ws = wb['Sheet1']
        for i, record in enumerate(data):
            ws[i] = list(record.values())

        # ── prepare ──────────────────────────────────────────────────────────
        t0 = time.perf_counter()
        wb._create_style()
        for sheet in wb._sheet_list:
            wb._dict_wb[sheet] = wb.workbook[sheet]._transfer_to_dict()
            if wb.workbook[sheet]._table_list:
                TableFinalValidation(
                    data=wb.workbook[sheet]._data,
                    table_list=wb.workbook[sheet]._table_list,
                )
        payload = {
            'content': wb._dict_wb,
            'file_props': wb.file_props,
            'style': wb.style._style_map,
            'protection': wb.protection,
            'sheet_order': wb._sheet_list,
        }
        t1 = time.perf_counter()

        # ── JSON encode ───────────────────────────────────────────────────────
        json_data = msgspec.json.encode(payload)
        t2 = time.perf_counter()

        # ── FFI call ──────────────────────────────────────────────────────────
        lib = wb._read_lib(None)
        ignore_go_panic = ctypes.c_int64(1)

        # Detect API version: new API has 3 args (data, *outLen, catchPanic)
        try:
            out_len = ctypes.c_int64(0)
            lib.Export.argtypes = [
                ctypes.c_char_p,
                ctypes.POINTER(ctypes.c_int64),
                ctypes.c_int64,
            ]
            lib.Export.restype = ctypes.c_void_p
            lib.FreeCPointer.argtypes = [ctypes.c_void_p, ctypes.c_int64]

            t3 = time.perf_counter()
            ptr = lib.Export(json_data, ctypes.byref(out_len), ignore_go_panic)
            t4 = time.perf_counter()

            bytes(ctypes.string_at(ptr, out_len.value))
            lib.FreeCPointer(ptr, ctypes.c_int64(0))
            t5 = time.perf_counter()
            api_label = 'new (raw bytes)'

        except Exception:
            # Fall back to old API
            lib.Export.argtypes = [ctypes.c_char_p, ctypes.c_int64]
            lib.Export.restype = ctypes.c_void_p

            t3 = time.perf_counter()
            ptr = lib.Export(json_data, ignore_go_panic)
            t4 = time.perf_counter()

            base64.b64decode(ctypes.cast(ptr, ctypes.c_char_p).value)
            lib.FreeCPointer.argtypes = [ctypes.c_void_p, ctypes.c_int64]
            lib.FreeCPointer(ptr, ctypes.c_int64(0))
            t5 = time.perf_counter()
            api_label = 'old (base64)'

        StyleManager.reset_style_configs()

        timings['prepare'].append(t1 - t0)
        timings['json_enc'].append(t2 - t1)
        timings['ffi'].append(t4 - t3)
        timings['post'].append(t5 - t4)

    return (
        {k: {'mean': statistics.mean(v), 'min': min(v)} for k, v in timings.items()},
        api_label,
        len(json_data),
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
    comp, api_ver, json_bytes = measure_components(data_med, repeat=3)
    total = sum(v['mean'] for v in comp.values())
    print(f'  API version detected : {api_ver}')
    print(f'  JSON payload size    : {json_bytes / 1024:.1f} KB')
    print(
        f'  base64-encoded size  : {json_bytes * 4 // 3 / 1024:.1f} KB  '
        f'(savings if removed: {json_bytes // 3 / 1024:.1f} KB)'
    )
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
