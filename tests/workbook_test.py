from __future__ import annotations

import pytest
from openpyxl_style_writer import CustomStyle

from pyfastexcel import Workbook

style_for_set_custom_style = CustomStyle(font_color='fcfcfc')


@pytest.mark.parametrize(
    'range_slice, values, expected_output',
    [
        (
            slice('A1', 'G1'),
            [1, 2, 3, 9, 8, 45, 11],
            [
                (1, 'DEFAULT_STYLE'),
                (2, 'DEFAULT_STYLE'),
                (3, 'DEFAULT_STYLE'),
                (9, 'DEFAULT_STYLE'),
                (8, 'DEFAULT_STYLE'),
                (45, 'DEFAULT_STYLE'),
                (11, 'DEFAULT_STYLE'),
            ],
        ),
        (slice('A1', 'B1'), [1, 2], [(1, 'DEFAULT_STYLE'), (2, 'DEFAULT_STYLE')]),
        # Ensure the expand_row_and_cols do not induce value reference issue
        (
            'B3',
            1,
            [
                [('', 'DEFAULT_STYLE')],
                [('', 'DEFAULT_STYLE'), ('', 'DEFAULT_STYLE')],
                [('', 'DEFAULT_STYLE'), (1, 'DEFAULT_STYLE')],
            ],
        ),
    ],
)
def test_workbook(range_slice, values, expected_output):
    wb = Workbook()
    ws = wb['Sheet1']
    ws[range_slice] = values

    if not isinstance(range_slice, slice):
        actual_output = ws.data
    else:
        actual_output = [tuple([cell for cell in row]) for row in ws[range_slice]]
    assert actual_output == expected_output


@pytest.mark.parametrize(
    'input_data, expected_exception',
    [
        ([slice('A1', 'G3'), [1, 2, 3]], ValueError),  # Invalid slice assignment
        ((1.62, [1]), TypeError),  # Invalid row assignment
    ],
)
def test_invalid_assignment(input_data, expected_exception):
    wb = Workbook()
    ws = wb['Sheet1']

    with pytest.raises(expected_exception):
        ws[input_data[0]] = input_data[1]


@pytest.mark.parametrize(
    'row_slice, value_list, expected_output',
    [
        (
            slice('A1', 'E1'),
            [2, 6, 7, 8, 9],
            [
                (2, 'DEFAULT_STYLE'),
                (6, 'DEFAULT_STYLE'),
                (7, 'DEFAULT_STYLE'),
                (8, 'DEFAULT_STYLE'),
                (9, 'DEFAULT_STYLE'),
            ],
        ),
        (
            slice('B6', 'F6'),
            ['qwe', 6, 7, -8, 'hello'],
            [
                ('', 'DEFAULT_STYLE'),
                ('qwe', 'DEFAULT_STYLE'),
                (6, 'DEFAULT_STYLE'),
                (7, 'DEFAULT_STYLE'),
                (-8, 'DEFAULT_STYLE'),
                ('hello', 'DEFAULT_STYLE'),
            ],
        ),
        (
            slice('E100', 'I100'),
            [('qwe', 'bold_font_style'), 6, 7, -8, 'hello'],
            [
                ('', 'DEFAULT_STYLE'),
                ('', 'DEFAULT_STYLE'),
                ('', 'DEFAULT_STYLE'),
                ('', 'DEFAULT_STYLE'),
                ('qwe', 'bold_font_style'),
                (6, 'DEFAULT_STYLE'),
                (7, 'DEFAULT_STYLE'),
                (-8, 'DEFAULT_STYLE'),
                ('hello', 'DEFAULT_STYLE'),
            ],
        ),
        (slice('A1', 'G100'), [1, 2, 3], ValueError),
    ],
)
def test_workbook_slice(row_slice, value_list, expected_output):
    wb = Workbook()
    ws = wb['Sheet1']

    if not isinstance(expected_output, list):
        with pytest.raises(expected_output):
            ws[row_slice] = value_list
        with pytest.raises(expected_output):
            print(ws[row_slice])
    else:
        ws[row_slice] = value_list
        print(ws[row_slice])
        assert ws[row_slice] == expected_output


@pytest.mark.parametrize(
    'index, value_list, expected_output',
    [
        (
            0,
            [2, 6, 7, 8, 9],
            [
                (2, 'DEFAULT_STYLE'),
                (6, 'DEFAULT_STYLE'),
                (7, 'DEFAULT_STYLE'),
                (8, 'DEFAULT_STYLE'),
                (9, 'DEFAULT_STYLE'),
            ],
        ),
        (
            1,
            ['qwe', 6, 7, -8, 'hello'],
            [
                ('qwe', 'DEFAULT_STYLE'),
                (6, 'DEFAULT_STYLE'),
                (7, 'DEFAULT_STYLE'),
                (-8, 'DEFAULT_STYLE'),
                ('hello', 'DEFAULT_STYLE'),
            ],
        ),
        (
            36669,
            [('qwe', 'bold_font_style'), [6, 7, 78], 7, -8.5435, 'hello'],
            [
                ('qwe', 'bold_font_style'),
                ('[6, 7, 78]', 'DEFAULT_STYLE'),
                (7, 'DEFAULT_STYLE'),
                (-8.5435, 'DEFAULT_STYLE'),
                ('hello', 'DEFAULT_STYLE'),
            ],
        ),
        (-1, [1, 2, 3], ValueError),
        (1048576, [1], ValueError),
        (2, 99, ValueError),
        (6, 'STRING', ValueError),
    ],
)
def test_worksheet_row_get_and_set(index, value_list, expected_output):
    from pyfastexcel.utils import set_custom_style

    style = CustomStyle(font_size=12, font_bold=True)
    set_custom_style('bold_font_style', style)
    wb = Workbook()
    ws = wb['Sheet1']

    if not isinstance(expected_output, list):
        print('qweqweqw')
        with pytest.raises(expected_output):
            ws[index] = value_list
    else:
        ws[index] = value_list
        print(ws[index])
        assert ws[index] == expected_output


@pytest.mark.parametrize(
    'cell_value',
    [
        ([('1', '2', '3')]),
        (('1', 3, 2)),
        (('1')),
        ((1)),
    ],
)
def test_set_worksheet_with_wrong_format(cell_value):
    wb = Workbook()
    ws = wb['Sheet1']
    with pytest.raises(ValueError):
        ws[0] = cell_value


def test_save_workbook():
    wb = Workbook()
    ws = wb['Sheet1']
    ws['A1':'G1'] = [1, 2, 3, 9, 8, 45, 11]

    # Save without calling read_lib_and_create_excel()
    wb.save('test1.xlsx')

    wb.read_lib_and_create_excel()
    wb.save('test2.xlsx')


def test_if_style_is_reset():
    from pyfastexcel.style import StyleManager

    wb = Workbook()
    style = CustomStyle(font_size=11, font_color='000000')

    ws = wb['Sheet1']
    ws['A1'] = ('test', style)
    wb._create_style()
    assert len(StyleManager._style_map) != 0
    assert len(StyleManager._STYLE_NAME_MAP) != 0
    assert StyleManager._STYLE_ID == 1
    assert StyleManager.REGISTERED_STYLES == {
        'DEFAULT_STYLE': StyleManager.DEFAULT_STYLE,
        'Custom Style 0': style,
    }
    wb.read_lib_and_create_excel()
    assert len(StyleManager._style_map) == 0
    assert len(StyleManager._STYLE_NAME_MAP) == 0
    assert StyleManager._STYLE_ID == 0
    assert StyleManager.REGISTERED_STYLES == {
        'DEFAULT_STYLE': StyleManager.DEFAULT_STYLE,
    }

    # Create another Workbook in one process to ensure that after style configs
    # reset, everythings is still working.
    wb2 = Workbook()
    style2 = CustomStyle(font_size=99, font_color='fcfcfc')

    ws2 = wb2['Sheet1']
    ws2['A1'] = ('test', style2)
    wb2._create_style()
    assert len(StyleManager._style_map) != 0
    assert len(StyleManager._STYLE_NAME_MAP) != 0
    assert StyleManager._STYLE_ID == 1
    assert StyleManager.REGISTERED_STYLES == {
        'DEFAULT_STYLE': StyleManager.DEFAULT_STYLE,
        'Custom Style 0': style2,
    }
    wb2.read_lib_and_create_excel()
    assert len(StyleManager._style_map) == 0
    assert len(StyleManager._STYLE_NAME_MAP) == 0
    assert StyleManager._STYLE_ID == 0
    assert StyleManager.REGISTERED_STYLES == {
        'DEFAULT_STYLE': StyleManager.DEFAULT_STYLE,
    }


@pytest.mark.parametrize(
    'target, expected_output1',
    [
        ('A1', ('test', 'bold_font_style')),
        ('XD1', ('test', 'bold_font_style')),
    ],
)
def test_set_style_with_str(target, expected_output1):
    from pyfastexcel.style import StyleManager
    from pyfastexcel.utils import set_custom_style

    wb = Workbook()
    ws = wb['Sheet1']

    bold_style = CustomStyle(font_bold=True)
    set_custom_style('bold_font_style', bold_style)
    color_style = CustomStyle(font_color='d33513')

    ws[target] = 'test'

    ws.set_style(target, 'bold_font_style')
    assert ws[target] == expected_output1

    ws.set_style(target, color_style)
    assert ws[target][1] == f'Custom Style {StyleManager._STYLE_ID - 1}'

    with pytest.raises(ValueError):
        ws.set_style(target, 'wrong_style')


@pytest.mark.parametrize(
    'target, expected_output1',
    [
        ('A1:B1', [('test', 'bold_font_style'), ('q', 'bold_font_style')]),
        (slice('A1', 'B1'), [('test', 'bold_font_style'), ('q', 'bold_font_style')]),
    ],
)
def test_set_style_with_silce(target, expected_output1):
    from pyfastexcel.style import StyleManager
    from pyfastexcel.utils import set_custom_style

    wb = Workbook()
    ws = wb['Sheet1']

    bold_style = CustomStyle(font_bold=True)
    set_custom_style('bold_font_style', bold_style)
    color_style = CustomStyle(font_color='d33513')

    # Index assignment is not supported for slice currently.
    if isinstance(target, str) and ':' in target:
        t = target.split(':')
        t = slice(t[0], t[1])
    else:
        t = target

    ws[t] = ['test', 'q']

    ws.set_style(target, 'bold_font_style')
    assert ws[t] == expected_output1

    ws.set_style(target, color_style)
    assert ws[t][1][1] == f'Custom Style {StyleManager._STYLE_ID - 1}'

    with pytest.raises(ValueError):
        ws.set_style(target, 'wrong_style')


@pytest.mark.parametrize(
    'row, target, expected_output1',
    [
        (0, [0, 1], [('test', 'DEFAULT_STYLE'), ('q', 'bold_font_style'), ('1', 'DEFAULT_STYLE')]),
        (1, [1, 1], [('test', 'DEFAULT_STYLE'), ('q', 'bold_font_style'), ('1', 'DEFAULT_STYLE')]),
    ],
)
def test_set_style_with_list(row, target, expected_output1):
    from pyfastexcel.style import StyleManager
    from pyfastexcel.utils import set_custom_style

    wb = Workbook()
    ws = wb['Sheet1']

    bold_style = CustomStyle(font_bold=True)
    set_custom_style('bold_font_style', bold_style)
    color_style = CustomStyle(font_color='d33513')

    ws[row] = ['test', 'q', '1']

    ws.set_style(target, 'bold_font_style')
    assert ws[row] == expected_output1

    ws.set_style(target, color_style)
    assert ws[row][target[1]][1] == f'Custom Style {StyleManager._STYLE_ID - 1}'


@pytest.mark.parametrize(
    'target, expected_output',
    [
        ([1048577, 1], ValueError),
        ([1, 16385], ValueError),
        (['awer', 1], TypeError),
        ({}, TypeError),
    ],
)
def test_set_style_error(target, expected_output):
    from pyfastexcel.utils import set_custom_style

    wb = Workbook()
    ws = wb['Sheet1']

    bold_style = CustomStyle(font_bold=True)
    set_custom_style('bold_font_style', bold_style)

    with pytest.raises(expected_output):
        ws.set_style(target, 'bold_font_style')


@pytest.mark.parametrize(
    'data_range, expected_result',
    [
        ('A1:J1', set(['A1:J1'])),
        ('B1:J5', set(['B1:J5'])),
        ('C2:D5', set(['C2:D5'])),
    ],
)
def test_auto_filter(data_range, expected_result):
    wb = Workbook()
    ws = wb['Sheet1']
    ws[0] = [f'col{i}' for i in range(10)]
    row = 1
    for i in range(1, 9):
        ws[i] = [row * (j + 1) for j in range(10)]
        row += 1

    ws.auto_filter(data_range)
    assert ws.auto_filter_set == expected_result


@pytest.mark.parametrize(
    'data_range, expected_result',
    [
        ('A1J1', ValueError),
        ('5', ValueError),
        ('AAAAA', ValueError),
    ],
)
def test_auto_filter_failed(data_range, expected_result):
    wb = Workbook()
    ws = wb['Sheet1']
    with pytest.raises(expected_result):
        ws.auto_filter(data_range)


@pytest.mark.parametrize(
    'algorithm, expected_result',
    [
        ('qwe', ValueError),
        (123, ValueError),
        ('XOR', None),
        ('MD4', None),
        ('MD5', None),
        ('SHA-1', None),
        ('SHA-256', None),
        ('SHA-384', None),
        ('SHA-512', None),
    ],
)
def test_protect_workbook(algorithm, expected_result):
    wb = Workbook()
    ws = wb['Sheet1']
    ws['A1'] = 'test'
    if expected_result is None:
        wb.protect_workbook(algorithm, '12345', True, False)
    else:
        with pytest.raises(expected_result):
            wb.protect_workbook(algorithm, '12345', True, False)


def test_workbook_plain_data():
    plain_data = [[1, 2, 3, 4, 5], [6, 2, 9, 10]]
    wb = Workbook(plain_data=plain_data)
    wb.read_lib_and_create_excel()

    with pytest.raises(TypeError):
        Workbook(plain_data='failed')

    with pytest.raises(TypeError):
        Workbook(plain_data=['failed'])


def test_worksheet_plain_data():
    plain_data = [[1, 2, 3, 4, 5], [6, 2, 9, 10]]
    wb = Workbook()
    wb.create_sheet('New Sheet', plain_data=plain_data)
    wb.read_lib_and_create_excel()

    plain_data = 12
    with pytest.raises(TypeError):
        wb.create_sheet('New Sheet2', plain_data=plain_data)

    plain_data = '123'
    with pytest.raises(TypeError):
        wb.create_sheet('New Sheet3', plain_data=plain_data)

    with pytest.raises(ValueError):
        wb.create_sheet('neew', plain_data=plain_data, pre_allocate='123')


def test_pre_allocate():
    pre_allocate = {'n_rows': 1000, 'n_cols': 20}
    wb = Workbook(pre_allocate=pre_allocate)

    wb.create_sheet('new', pre_allocate=pre_allocate)
    wb.read_lib_and_create_excel()
