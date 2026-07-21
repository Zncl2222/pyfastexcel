"""Plot old-vs-new comparison charts from ``perf_memory.py`` JSON reports.

Measurement lives in ``perf_memory.py``; this script only *reads* the committed
reports under ``benchmark/results`` and never mutates them, so historical data is
preserved by construction.

The figure has two panels:

  * left  – grouped bars for build / export / total wall time (mean, with a
            min-max error bar), one bar per report;
  * right – grouped bars for peak RSS in MiB, one bar per report.

The first report is treated as the baseline; every other report is annotated
with its percentage change against that baseline.

Examples::

    # Auto-discover and plot every report in benchmark/results.
    uv run python benchmark/plot_perf.py

    # Pick specific reports and give them short legend labels.
    uv run python benchmark/plot_perf.py \
        results/2026-07-16-stage-a-baseline.json \
        results/2026-07-16-stage-a-pfx2.json \
        --label "v1 baseline" "v2 pfx2" \
        --output results/perf-compare.png
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any

import matplotlib

matplotlib.use('Agg')
import matplotlib.pyplot as plt  # noqa: E402
import numpy as np  # noqa: E402

RESULTS_DIR = Path(__file__).resolve().parent / 'results'
TIME_METRICS = ('build_seconds', 'export_seconds', 'total_seconds')
TIME_LABELS = ('build', 'export', 'total')
_DATE_PREFIX = re.compile(r'^\d{4}-\d{2}-\d{2}-')


class Report:
    """A single loaded perf_memory report with the fields the plot needs."""

    def __init__(self, path: Path, label: str | None = None):
        with path.open(encoding='utf-8') as report_file:
            raw: dict[str, Any] = json.load(report_file)
        self.path = path
        self.workload = raw['workload']
        self.summary = raw['summary']
        self.label = label or _default_label(path)

    def time_stat(self, metric: str) -> dict[str, float]:
        return self.summary[metric]

    def peak_rss_mib(self) -> tuple[float, float, float] | None:
        rss = self.summary.get('peak_rss_bytes')
        if not rss:
            return None
        scale = 1024 * 1024
        return rss['mean'] / scale, rss['min'] / scale, rss['max'] / scale


def _default_label(path: Path) -> str:
    """Filename stem with a leading ``YYYY-MM-DD-`` date prefix stripped."""
    return _DATE_PREFIX.sub('', path.stem)


def _err_bounds(mean: float, low: float, high: float) -> tuple[float, float]:
    """Return (lower, upper) error-bar magnitudes, clamped non-negative."""
    return max(mean - low, 0.0), max(high - mean, 0.0)


def _pct(new: float, base: float) -> str:
    if base == 0:
        return 'n/a'
    change = (new - base) / base * 100
    return f'{change:+.0f}%'


def _grouped_positions(num_groups: int, num_series: int) -> tuple[np.ndarray, float]:
    """x-centres for each group and the per-series bar width."""
    indices = np.arange(num_groups)
    width = 0.8 / num_series
    return indices, width


def _plot_time(ax: Any, reports: list[Report]) -> None:
    indices, width = _grouped_positions(len(TIME_METRICS), len(reports))
    baseline = reports[0]

    for series_idx, report in enumerate(reports):
        offset = (series_idx - (len(reports) - 1) / 2) * width
        means, lower, upper = [], [], []
        for metric in TIME_METRICS:
            stat = report.time_stat(metric)
            means.append(stat['mean'])
            low, high = _err_bounds(stat['mean'], stat['min'], stat['max'])
            lower.append(low)
            upper.append(high)

        bars = ax.bar(
            indices + offset,
            means,
            width,
            label=report.label,
            edgecolor='black',
            linewidth=0.5,
            yerr=[lower, upper],
            capsize=3,
        )
        if series_idx == 0:
            continue
        for bar, metric in zip(bars, TIME_METRICS):
            base_mean = baseline.time_stat(metric)['mean']
            ax.annotate(
                _pct(bar.get_height(), base_mean),
                xy=(bar.get_x() + bar.get_width() / 2, bar.get_height()),
                xytext=(0, 3),
                textcoords='offset points',
                ha='center',
                va='bottom',
                fontsize=8,
            )

    ax.set_title('Wall time (mean, min-max bars)')
    ax.set_ylabel('Time (s)')
    ax.set_xticks(indices)
    ax.set_xticklabels(TIME_LABELS)
    ax.legend()


def _plot_rss(ax: Any, reports: list[Report]) -> None:
    rss_values = [report.peak_rss_mib() for report in reports]
    if any(value is None for value in rss_values):
        ax.set_visible(False)
        return

    indices = np.arange(len(reports))
    baseline_mean = rss_values[0][0]
    means = [value[0] for value in rss_values]
    lower = [_err_bounds(*value)[0] for value in rss_values]
    upper = [_err_bounds(*value)[1] for value in rss_values]

    bars = ax.bar(
        indices,
        means,
        0.6,
        color=[f'C{i}' for i in range(len(reports))],
        edgecolor='black',
        linewidth=0.5,
        yerr=[lower, upper],
        capsize=3,
    )
    for series_idx, bar in enumerate(bars):
        text = f'{bar.get_height():.0f} MiB'
        if series_idx > 0:
            text += f'\n{_pct(bar.get_height(), baseline_mean)}'
        ax.annotate(
            text,
            xy=(bar.get_x() + bar.get_width() / 2, bar.get_height()),
            xytext=(0, 3),
            textcoords='offset points',
            ha='center',
            va='bottom',
            fontsize=8,
        )

    ax.set_title('Peak RSS')
    ax.set_ylabel('Peak RSS (MiB)')
    ax.set_xticks(indices)
    ax.set_xticklabels([report.label for report in reports], rotation=15, ha='right')


def _figure_title(reports: list[Report]) -> str:
    workloads = {(r.workload['rows'], r.workload['cols']) for r in reports}
    if len(workloads) == 1:
        rows, cols = next(iter(workloads))
        return f'pyfastexcel perf comparison — {rows} rows × {cols} cols ({rows * cols:,} cells)'
    return 'pyfastexcel perf comparison (mixed workloads)'


def build_figure(reports: list[Report]) -> Any:
    fig, (ax_time, ax_rss) = plt.subplots(1, 2, figsize=(13, 6))
    _plot_time(ax_time, reports)
    _plot_rss(ax_rss, reports)
    fig.suptitle(_figure_title(reports), fontsize=14)
    fig.tight_layout()
    return fig


def _discover_reports() -> list[Path]:
    return sorted(RESULTS_DIR.glob('*.json'))


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        'reports',
        nargs='*',
        type=Path,
        help='perf_memory JSON reports, oldest/baseline first (default: benchmark/results/*.json)',
    )
    parser.add_argument(
        '--label',
        nargs='+',
        help='legend labels, one per report (default: filename stem without date prefix)',
    )
    parser.add_argument(
        '--output',
        type=Path,
        default=RESULTS_DIR / 'perf-compare.png',
        help='output PNG path (default: benchmark/results/perf-compare.png)',
    )
    return parser.parse_args()


def main() -> None:
    args = _parse_args()
    paths = args.reports or _discover_reports()
    if not paths:
        raise SystemExit(f'no reports given and none found in {RESULTS_DIR}')
    if args.label and len(args.label) != len(paths):
        raise SystemExit(f'--label needs {len(paths)} labels, got {len(args.label)}')

    labels = args.label or [None] * len(paths)
    reports = [Report(path, label) for path, label in zip(paths, labels)]

    workloads = {
        (r.workload['rows'], r.workload['cols'], r.workload.get('destination')) for r in reports
    }
    if len(workloads) != 1:
        print(
            'warning: reports have different workloads; comparison may be misleading',
            file=sys.stderr,
        )

    figure = build_figure(reports)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    figure.savefig(args.output, dpi=120)
    print(f'wrote {args.output} comparing {len(reports)} report(s):')
    for report in reports:
        print(f'  {report.label:<20} {report.path}')


if __name__ == '__main__':
    main()
