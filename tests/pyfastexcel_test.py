from __future__ import annotations

import pytest

from openpyxl_style_writer import CustomStyle
from openpyxl.styles import Side
from pyfastexcel import FastWriter, NormalWriter


font_params = {
    'size': 11,
    'bold': True,
    'italic': True,
    'color': 'FF000000',
    'vertAlign': 'baseline',
    'strike': True,
    'name': 'Calibri',
    'family': 1,
    'underline': 'doubleAccounting',
}

fill_params = {
    'fill_type': 'solid',
    'start_color': 'FFFFFFFF',
    'end_color': 'FF000000',
}

border_params = {
    'left': Side(style='thin', color='FF000000'),
    'right': Side(style='thick', color='FF000000'),
    'top': Side(style='dotted', color='FF000000'),
    'bottom': Side(style='dashDot', color='FF000000'),
    'diagonal': Side(style='hair', color='FF000000'),
    'diagonal_direction': 1,
    'outline': Side(style='medium', color='FF000000'),
    'vertical': Side(style='mediumDashed', color='FF000000'),
    'horizontal': Side(style='slantDashDot', color='FF000000'),
}

ali_params = {
    'horizontal': 'general',
    'vertical': 'bottom',
    'text_rotation': 12,
    'wrap_text': True,
    'shrink_to_fit': True,
    'indent': 1,
    'justifyLastLine': True,
    'readingOrder': 1,
    'relativeIndent': 1,
}


def prepare_example_data(rows: int = 1000, cols: int = 10) -> list[dict[str, str]]:
    headers = [f'Column_{i}' for i in range(cols)]
    data = [[i for i in range(cols)] for j in range(rows)]
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
    test_style = CustomStyle(
        font_params=font_params,
        fill_params=fill_params,
        border_params=border_params,
        ali_params=ali_params,
        number_format='0.00%',
        protect=True,
    )
    test_style.protection.hidden = True


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
        self.set_file_props('Creator', 'Hello')
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


def test_pyexcelize_fast_example():
    data = prepare_example_data(rows=25, cols=9)
    excel_example = PyExcelizeFastExample(data)
    excel_bytes = excel_example.create_excel()
    assert isinstance(excel_bytes, bytes)


def test_set_file_props():
    excel_example = PyExcelizeFastExample([])
    with pytest.raises(ValueError):
        excel_example.set_file_props('Test', 'Test')


def test_pyexcelize_normal_example():
    data = prepare_example_data(rows=3, cols=3)
    excel_example = PyExcelizeNormalExample(data)
    excel_example.create_sheet('Test')
    excel_example.remove_sheet('Test')
    excel_bytes = excel_example.create_excel()
    assert isinstance(excel_bytes, bytes)