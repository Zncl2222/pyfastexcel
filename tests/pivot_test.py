from __future__ import annotations

import pytest
import random

from pyfastexcel import Workbook

from pyfastexcel.pivot import PivotTable, PivotTableField
from pyfastexcel.enums import PivotSubTotal


def get_wb():
    wb = Workbook()
    ws = wb['Sheet1']
    ws[0] = ['Month', 'Year', 'Types', 'Sales', 'Mart']
    month = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    year = [2023, 2024, 2025]
    types = ['Coke', 'Sprite', 'Dr.Peppr', 'Juice']
    mart = ['A', 'B', 'C', 'D']
    for row in range(1, 32):
        ws.cell(row=row, column=1, value=random.choice(month))
        ws.cell(row=row, column=2, value=random.choice(year))
        ws.cell(row=row, column=3, value=random.choice(types))
        ws.cell(row=row, column=4, value=random.randint(0, 5000))
        ws.cell(row=row, column=5, value=random.choice(mart))

    return wb, ws


@pytest.mark.parametrize(
    'case',
    [
        (1),
        (2),
        (3),
        (4),
        (5),
        (6),
    ],
)
def test_pivot_table(case):
    wb, ws = get_wb()

    pivot_table = PivotTable(
        data_range='Sheet1!A1:E31',
        pivot_table_range='Sheet1!G2:M34',
        rows=[
            PivotTableField(data='Month', default_subtotal=True),
            PivotTableField(data='Year'),
        ],
        filter=[PivotTableField(data='mart')],
        columns=[PivotTableField(data='Type', default_subtotal=True)],
        data=[PivotTableField(data='Sales', name='Summarize', subtotal='sum')],
        row_grand_totals=True,
        column_grand_totals=True,
        show_drill=True,
        show_row_headers=True,
        show_last_column=True,
    )

    if case == 1:
        ws.add_pivot_table(pivot_table)
    elif case == 2:
        ws.add_pivot_table([pivot_table])
    elif case == 3:
        wb.add_pivot_table('Sheet1', pivot_table)
    elif case == 4:
        wb.add_pivot_table('Sheet1', [pivot_table])
    elif case == 5:
        ws.add_pivot_table(
            data_range='Sheet1!A1:E31',
            pivot_table_range='Sheet1!G2:M34',
            rows=[
                PivotTableField(data='Month', default_subtotal=True),
                PivotTableField(data='Year'),
            ],
            filter=[PivotTableField(data='mart')],
            columns=[PivotTableField(data='Type', default_subtotal=True)],
            data=[PivotTableField(data='Sales', name='Summarize', subtotal='sum')],
            row_grand_totals=True,
            column_grand_totals=True,
            show_drill=True,
            show_row_headers=True,
            show_last_column=True,
        )
    elif case == 6:
        wb.add_pivot_table(
            'Sheet1',
            data_range='Sheet1!A1:E31',
            pivot_table_range='Sheet1!G2:M34',
            rows=[
                PivotTableField(data='Month', default_subtotal=True),
                PivotTableField(data='Year'),
            ],
            filter=[PivotTableField(data='mart')],
            columns=[PivotTableField(data='Type', default_subtotal=True)],
            data=[PivotTableField(data='Sales', name='Summarize', subtotal='sum')],
            row_grand_totals=True,
            column_grand_totals=True,
            show_drill=True,
            show_row_headers=True,
            show_last_column=True,
        )

    wb.read_lib_and_create_excel()


@pytest.mark.parametrize(
    'data',
    [('A1:E31'), ('Sheet1A1:E31')],
)
def test_pivot_table_data_range_failed(data):
    with pytest.raises(ValueError):
        PivotTable(
            data_range=data,
        )


@pytest.mark.parametrize(
    'data',
    [('A1:E31'), ('Sheet1A1:E31')],
)
def test_pivot_table_pivot_data_range_failed(data):
    with pytest.raises(ValueError):
        PivotTable(
            pivot_table_range=data,
        )


@pytest.mark.parametrize(
    'name',
    [('123'), ('Sqwer')],
)
def test_pivot_table_style_name_failed(name):
    with pytest.raises(ValueError):
        PivotTable(
            data_range='Sheet1!A1:E31',
            pivot_table_range='Sheet1!G2:M34',
            pivot_table_style_name=name,
        )


@pytest.mark.parametrize(
    'value, expected_result',
    [('sum', 'Sum'), ('Sum', 'Sum'), ('sUm', 'Sum'), ('SUM', 'Sum'), (PivotSubTotal.Max, 'Max')],
)
def test_pivot_table_field_subtotal_serialize(value, expected_result):
    field = PivotTableField(data='Sales', subtotal=value)
    f_dump = field.model_dump()
    assert f_dump['subtotal'] == expected_result
