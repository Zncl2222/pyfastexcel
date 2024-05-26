import os
import statistics
import sys
import timeit

import matplotlib.pyplot as plt
import numpy as np
from openpyxl import Workbook as OpenpyxlWorkbook

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))  # noqa

from example import prepare_example_data  # noqa
from pyfastexcel import StreamWriter  # noqa
from pyfastexcel import Workbook as PyFastExcelWorkbook  # noqa

data = None
result_dict = {}


setup = """
from __main__ import write_excel_with_pyfastexcel_with_double_for_loop
from __main__ import write_excel_with_pyfastexcel_with_row
from __main__ import write_excel_with_stream_writer
from __main__ import write_excel_with_openpyxl_normal_wb
from __main__ import write_excel_with_openpyxl_write_only_wb
"""


def write_excel_with_pyfastexcel_with_double_for_loop() -> None:
    from pyfastexcel.utils import index_to_column

    wb = PyFastExcelWorkbook()
    ws = wb['Sheet1']
    ws[0] = list(data[0].keys())
    for i, record in enumerate(data):
        for j, (_, value) in enumerate(record.items()):
            col = index_to_column(j + 2)
            ws[f'{col}{i}'] = value
    wb.save('pyfastexcel_double_for_loop.xlsx')


def write_excel_with_pyfastexcel_with_row() -> None:
    wb = PyFastExcelWorkbook()
    ws = wb['Sheet1']
    for i, record in enumerate(data):
        ws[i] = list(record.values())
    wb.save('pyfastexcel_by_row.xlsx')


def write_excel_with_stream_writer() -> None:
    class CustomWriter(StreamWriter):
        def create_excel(self) -> bytes:
            self._set_header()
            self._create_style()
            self._create_body()
            self.save('pyfastexcel_stream_writer.xlsx')

        def _set_header(self):
            self.headers = list(self.data[0].keys())
            for h in self.headers:
                self.row_append(h)
            self.create_row()

        def _create_body(self) -> None:
            for row in self.data:
                for h in self.headers:
                    self.row_append(row[h])
                self.create_row()

    CustomWriter(data).create_excel()


def write_excel_with_openpyxl_normal_wb() -> None:
    wb = OpenpyxlWorkbook()
    ws = wb.create_sheet()
    ws.append(list(data[0].keys()))
    for record in data:
        ws.append(list(record.values()))
    wb.save('openpyxl_wb.xlsx')


def write_excel_with_openpyxl_write_only_wb() -> None:
    wb = OpenpyxlWorkbook(write_only=True)
    ws = wb.create_sheet()
    ws.append(list(data[0].keys()))
    for record in data:
        ws.append(list(record.values()))
    wb.save('openpyxl_write_only_wb.xlsx')


def run_test_case(test_case_name, test_case_func, repeat=5, number=1):
    benchmark = f'\nExecution time ({test_case_name}): \n'
    results = timeit.repeat(
        f'{test_case_func}()',
        setup=setup,
        repeat=repeat,
        number=number,
    )
    result_dict[test_case_name] = {'results': results}
    for i, time in enumerate(results):
        benchmark += f'Test {i + 1}: {time} s\n'

    mean = statistics.mean(results)
    std_dev = statistics.stdev(results)
    max_val = max(results)
    min_val = min(results)
    benchmark += f'Mean: {mean} s\nMax: {max_val} s\nMin: {min_val} s\nStd: {std_dev} s\n'
    result_dict[test_case_name]['mean'] = mean
    result_dict[test_case_name]['max_val'] = max_val
    result_dict[test_case_name]['min_vl'] = min_val
    result_dict[test_case_name]['std_dev'] = std_dev
    return benchmark


def extract_plot_data(result_dict):
    labels = list(result_dict.keys())
    means = [entry['mean'] for entry in result_dict.values()]
    max_vals = [entry['max_val'] for entry in result_dict.values()]
    min_vals = [entry['min_vl'] for entry in result_dict.values()]
    std_devs = [entry['std_dev'] for entry in result_dict.values()]
    return labels, means, max_vals, min_vals, std_devs


def add_labels(bars, ax, orientation='v'):
    for bar in bars:
        if orientation == 'v':
            height = bar.get_height()
            ax.annotate(
                '{}'.format(round(height, 3)),
                xy=(bar.get_x() + bar.get_width() / 2, height),
                xytext=(0, 3),  # 3 points vertical offset
                textcoords='offset points',
                ha='center',
                va='bottom',
            )
        else:
            width = bar.get_width()
            ax.annotate(
                '{}'.format(round(width, 3)),
                xy=(width, bar.get_y() + bar.get_height() / 2),
                xytext=(3, 0),  # 3 points horizontal offset
                textcoords='offset points',
                ha='left',
                va='center',
            )


def plot_bars(orientation='v', title='Method', fig_name='benchmark.png'):
    labels, means, max_vals, min_vals, std_devs = extract_plot_data(result_dict)
    num_bars = len(labels)
    indices = np.arange(num_bars)
    bar_width = 0.2

    fig, ax = plt.subplots(figsize=(10, 6))

    if orientation == 'v':
        bars1 = ax.bar(
            indices - bar_width,
            min_vals,
            bar_width,
            label='Min',
            color='orange',
            edgecolor='black',
        )
        bars2 = ax.bar(indices, means, bar_width, label='Mean', color='cyan', edgecolor='black')
        bars3 = ax.bar(
            indices + bar_width,
            max_vals,
            bar_width,
            label='Max',
            color='gray',
            edgecolor='black',
        )
        ax2 = ax.twinx()
        ax2.plot(indices, std_devs, '--o', color='red', label='Std Dev')
        ax.set_xlabel(title)
        ax.set_ylabel('Time (s)')
        ax.set_xticks(indices)
        ax.set_xticklabels(labels, rotation=45, ha='right')
        ax2.set_ylabel('Standard Deviation')
    else:
        bars1 = ax.barh(
            indices - bar_width,
            min_vals,
            bar_width,
            label='Min',
            color='orange',
            edgecolor='black',
        )
        bars2 = ax.barh(indices, means, bar_width, label='Mean', color='cyan', edgecolor='black')
        bars3 = ax.barh(
            indices + bar_width,
            max_vals,
            bar_width,
            label='Max',
            color='gray',
            edgecolor='black',
        )
        bars4 = ax.barh(
            indices + 2 * bar_width,
            std_devs,
            bar_width,
            label='Std',
            color='pink',
            edgecolor='black',
        )
        ax.set_ylabel(title)
        ax.set_xlabel('Time (s)')
        ax.set_yticks(indices)
        ax.set_yticklabels(labels)

    handles1, labels1 = ax.get_legend_handles_labels()
    handles2, labels2 = ax2.get_legend_handles_labels() if orientation == 'v' else ([], [])
    ax.legend(handles1 + handles2, labels1 + labels2)

    add_labels(bars1, ax, orientation)
    add_labels(bars2, ax, orientation)
    add_labels(bars3, ax, orientation)
    if orientation == 'h':
        add_labels(bars4, ax, orientation)

    if orientation == 'v':
        for i, std in enumerate(std_devs):
            ax2.annotate(
                f'{std:.3f}',
                xy=(indices[i], std),
                xytext=(0, 5),
                textcoords='offset points',
                ha='center',
                va='bottom',
            )

    plt.tight_layout()
    plt.savefig(fig_name)


def plot_vertical_bar(title='Method', fig_name='vbars.png'):
    plot_bars(orientation='v', title=title, fig_name=fig_name)


def plot_horizontal_bar(title='Method', fig_name='hbars.png'):
    plot_bars(orientation='h', title=title, fig_name=fig_name)


if __name__ == '__main__':
    cases = [(500, 30), (5000, 30), (50000, 30)]
    for row, col in cases:
        data = prepare_example_data(rows=row, cols=col)
        benchmark = run_test_case(
            'WorkBook double loop',
            'write_excel_with_pyfastexcel_with_double_for_loop',
        )
        benchmark += run_test_case('WorkBook by row', 'write_excel_with_pyfastexcel_with_row')
        benchmark += run_test_case('StreamWriter', 'write_excel_with_stream_writer')
        benchmark += run_test_case('Openpyxl Workbook', 'write_excel_with_openpyxl_normal_wb')
        benchmark += run_test_case(
            'Openpyxl Write Only Workbook',
            'write_excel_with_openpyxl_write_only_wb',
        )
        print(benchmark)
        plot_vertical_bar(f'Method (rows={row}, columns={col})', f'{row}+{col}_vertical.png')
        plot_horizontal_bar(f'Method (rows={row}, columns={col})', f'{row}+{col}_horizontal.png')
