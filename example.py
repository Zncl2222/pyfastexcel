import time

from openpyxl_style_writer import CustomStyle
from openpyxl.styles import Side
from pyfastexcel.driver import FastWriter, NormalWriter


def prepare_example_data(rows: int = 1000, cols: int = 10) -> list[dict[str, str]]:
    import random

    random.seed(42)
    headers = [f'Column_{i}' for i in range(cols)]
    data = [[random.random() for _ in range(cols)] for _ in range(rows)]
    records = []
    for row in data:
        record = {}
        for header, value in zip(headers, row):
            record[header] = str(round(value * 100, 2))
        records.append(record)
    return records


class StyleCollections:
    black_fill_style = CustomStyle(
        font_size='11',
        font_bold=True,
        font_color='F62B00',
        fill_color='000000',
    )
    green_fill_style = CustomStyle(
        font_size='29',
        font_bold=False,
        font_color='000000',
        fill_color='375623',
    )
    test_fill_style = CustomStyle(
        font_params={
            'size': 20,
            'bold': True,
            'italic': True,
            'color': '5e03fc',
        },
        fill_params={
            'patternType': 'solid',
            'fgColor': '375623',
        },
        border_params={
            'left': Side(style='thin', color='e12aeb'),
            'right': Side(style='thick', color='e12aeb'),
            'top': Side(style=None, color='e12aeb'),
            'bottom': Side(style='dashDot', color='e12aeb'),
        },
        ali_params={
            'wrapText': True,
            'shrinkToFit': True,
        },
        number_format='0.00%',
    )


class PyExcelizeFastExample(FastWriter, StyleCollections):
    def create_excel(self) -> bytes:
        self._set_header()
        self._create_style()
        self._create_single_header()
        self._create_body()
        return self._read_lib_and_create_excel()

    def _set_header(self):
        self.headers = list(self.data[0].keys())

    def _create_single_header(self):
        for i, h in enumerate(self.headers):
            self.row_append(h, style='green_fill_style', row_idx=0, col_idx=i)
        self.apply_to_header()

    def _create_body(self) -> None:
        for i, row in enumerate(self.data):
            for j, h in enumerate(self.headers):
                if h[-1] in ('1', '3', '5', '7', '9'):
                    self.row_append(row[h], style='black_fill_style', row_idx=i, col_idx=j)
                else:
                    self.row_append(row[h], style='test_fill_style', row_idx=i, col_idx=j)
            self.create_row(i)

        self.switch_sheet('Sheet2')
        for i, row in enumerate(self.data):
            for j, h in enumerate(self.headers):
                if h[-1] in ('1', '3', '5', '7', '9'):
                    self.row_append(row[h], style='test_fill_style', row_idx=i, col_idx=j)
                else:
                    self.row_append(row[h], style='black_fill_style', row_idx=i, col_idx=j)
            self.create_row(i)


class PyExcelizeNormalExample(NormalWriter, StyleCollections):
    def create_excel(self) -> bytes:
        self._set_header()
        self._create_style()
        self._create_single_header()
        self._create_body()
        return self._read_lib_and_create_excel()

    def _set_header(self):
        self.headers = list(self.data[0].keys())

    def _create_single_header(self):
        for h in self.headers:
            self.row_append(h, style='green_fill_style')
        self.create_row()

    def _create_body(self) -> None:
        for row in self.data:
            for h in self.headers:
                if h[-1] in ('1', '3', '5', '7', '9'):
                    self.row_append(row[h], style='black_fill_style')
                else:
                    self.row_append(row[h], style='test_fill_style')
            self.create_row()

        self.switch_sheet('Sheet2')
        for row in self.data:
            for h in self.headers:
                if h[-1] in ('1', '3', '5', '7', '9'):
                    self.row_append(row[h], style=self.green_fill_style)
                else:
                    self.row_append(row[h], style='black_fill_style')
            self.create_row()


if __name__ == '__main__':
    data = prepare_example_data(653, 90)
    fast_start_time = time.perf_counter()
    excel_fast = PyExcelizeFastExample(data).create_excel()
    fast_end_time = time.perf_counter()
    print('PYExcelizeFastDriver time: ', fast_end_time - fast_start_time)
    normal_start_time = time.perf_counter()
    excel_normal = PyExcelizeNormalExample(data).create_excel()
    notmal_end_time = time.perf_counter()
    print('PYExcelizeNormalDriver time: ', notmal_end_time - normal_start_time)

    file_path = 'pyexample_fast.xlsx'
    file_path2 = 'pyexample_normal.xlsx'

    with open(file_path, 'wb') as file:
        file.write(excel_fast)

    with open(file_path2, 'wb') as file:
        file.write(excel_normal)
