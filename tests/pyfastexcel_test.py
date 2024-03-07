from __future__ import annotations

import pytest
from openpyxl.styles import Side
from openpyxl_style_writer import CustomStyle

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


style_for_set_custom_style = CustomStyle(font_color='fcfcfc')


class StyleCollections:
    black_fill_style = CustomStyle(
        font_name='Time News Roman',
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
        self.set_cell_height(self.sheet, 5, 12)
        self.set_cell_width(self.sheet, 'A', 12)
        self.set_cell_width(self.sheet, 3, 12)

    def _create_single_header(self):
        for i, h in enumerate(self.headers):
            self.row_append(h, style='green_fill_style', col_idx=i)
        self.create_row()

    def _create_body(self) -> None:
        for row in self.data:
            for j, h in enumerate(self.headers):
                if h[-1] in ('1', '3', '5', '7', '9'):
                    self.row_append(row[h], style='black_fill_style', col_idx=j)
                else:
                    self.row_append(row[h], style='test_fill_style', col_idx=j)
            self.create_row()

        self.create_sheet('Sheet2')
        self.switch_sheet('Sheet2')
        for row in self.data:
            for j, h in enumerate(self.headers):
                if h[-1] in ('1', '3', '5', '7', '9'):
                    self.row_append(row[h], style='test_fill_style', col_idx=j)
                else:
                    self.row_append(row[h], style='black_fill_style', col_idx=j)
            self.create_row()


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
        self.set_cell_height(self.sheet, 5, 12)
        self.set_cell_width(self.sheet, 'A', 12)
        self.set_cell_width(self.sheet, 3, 12)

    def _create_single_header(self):
        for h in self.headers:
            self.row_append(h, style='green_fill_style')
        self.create_row()

    def _create_body(self) -> None:
        for row in self.data:
            for h in self.headers:
                if h[-1] in ('1', '3', '5', '7', '9'):
                    self.row_append(row[h], style=self.black_fill_style)
                else:
                    self.row_append(row[h], style='test_fill_style')
            self.create_row()

        self.create_sheet('Sheet2')
        for row in self.data:
            for h in self.headers:
                if h[-1] in ('1', '3', '5', '7', '9'):
                    self.row_append(row[h], style=self.green_fill_style)
                else:
                    self.row_append(row[h], style='black_fill_style')
            self.create_row()


def test_pyexcelize_fast_example():
    from pyfastexcel.utils import set_custom_style

    set_custom_style('test', style_for_set_custom_style)
    data = prepare_example_data(rows=25, cols=9)
    excel_example = PyExcelizeFastExample(data)
    excel_bytes = excel_example.create_excel()
    assert isinstance(excel_bytes, bytes)


def test_set_file_props():
    excel_example = PyExcelizeFastExample([[None] * 1000 for _ in range(1000)])
    with pytest.raises(ValueError):
        excel_example.set_file_props('Test', 'Test')


@pytest.mark.parametrize(
    'sheet, column, width, expected_exception',
    [
        ('Sheet1', 16385, 12, ValueError),  # Invalid case
        ('qwe', '', '', KeyError),  # Invalid: Single cell is not a merge cell
    ],
)
def test_set_cell_width(sheet, column, width, expected_exception):
    excel_example = PyExcelizeNormalExample([])
    with pytest.raises(expected_exception):
        excel_example.set_cell_width(sheet, column, width)


@pytest.mark.parametrize(
    'sheet, row, height, expected_exception',
    [
        ('Sheet1', 1048577, 12, ValueError),  # Invalid case
        ('qwe', '', '', KeyError),  # Invalid: Single cell is not a merge cell
    ],
)
def test_set_cell_height(sheet, row, height, expected_exception):
    excel_example = PyExcelizeFastExample([[None] * 1000 for _ in range(1000)])
    with pytest.raises(expected_exception):
        excel_example.set_cell_height(sheet, row, height)


@pytest.mark.parametrize(
    'sheet, top_left_cell, bottom_right_cell, expected_exception',
    [
        ('Sheet1', 'A1', 'C2', None),  # Valid case
        ('Sheet1', 'A1', 'A1', None),  # Invalid: Single cell is not a merge cell
        ('Sheet1', 'A1048577', 'C2', ValueError),  # Invalid: Row number exceeds limit
        ('Sheet1', 'A1', 'C1048577', ValueError),  # Invalid: Row number exceeds limit
        ('Sheet1', 'XFD1', 'XFD1048576', None),  # Valid: Maximum row and column numbers
        ('Sheet1', 'A1', 'XFE1048576', ValueError),  # Invalid: Column number exceeds limit
        ('Sheet1', 'A2', 'A1', ValueError),  # Invalid: Top number less than bottom number
        ('Sheet1', 'C1', 'A1', ValueError),  # Invalid: Top column less than bottom column
        ('Sheet1', 'A0', 'A1', ValueError),  # Invalid: Row number too small
        ('Sheet1', 'A0', 'C0', ValueError),  # Invalid: Row number too small
        ('abcd', '', '', KeyError),  # Invalid: Sheet name not found
    ],
)
def test_set_merge_cell(sheet, top_left_cell, bottom_right_cell, expected_exception):
    excel = PyExcelizeFastExample([[None] * 1000 for _ in range(1000)])
    if expected_exception is not None:
        with pytest.raises(expected_exception):
            excel.set_merge_cell(sheet, top_left_cell, bottom_right_cell)
    else:
        excel.set_merge_cell(sheet, top_left_cell, bottom_right_cell)
        assert (top_left_cell, bottom_right_cell) in excel.workbook[sheet]['MergeCells']


def test_pyexcelize_normal_example():
    data = prepare_example_data(rows=3, cols=3)
    excel_example = PyExcelizeNormalExample(data)
    excel_example.create_sheet('Test')
    excel_example.remove_sheet('Test')
    excel_example.switch_sheet('Sheet1')
    excel_bytes = excel_example.create_excel()
    assert isinstance(excel_bytes, bytes)
