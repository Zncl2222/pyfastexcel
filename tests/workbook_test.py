from __future__ import annotations

import pytest
from pydantic import ValidationError

from pyfastexcel import CustomStyle, Workbook
from pyfastexcel._typing import SelectionDict
from pyfastexcel.utils import CommentText, Selection

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
                [()],
                [(), ()],
                [(), (1, 'DEFAULT_STYLE')],
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
        # Test ws.sheet work as expected
        sheet_info_actual_output = ws.sheet['Data']
        assert sheet_info_actual_output == expected_output
    else:
        actual_output = ws[range_slice]
    assert actual_output == expected_output


@pytest.mark.parametrize(
    'sheet_name',
    [
        ('Sheet2'),
        ('New Sheet'),
        ('Gogoasd'),
    ],
)
def test_rename_sheet(sheet_name):
    wb = Workbook()
    wb.rename_sheet('Sheet1', sheet_name)

    assert sheet_name in wb.sheet_list


@pytest.mark.parametrize(
    'sheet_name',
    [
        ('Sheet2'),
        ('New Sheet'),
        ('Gogoasd'),
    ],
)
def test_rename_sheet_old_sheet_failed(sheet_name):
    wb = Workbook()
    with pytest.raises(IndexError):
        wb.rename_sheet('gogo', sheet_name)


@pytest.mark.parametrize(
    'sheet_name',
    [
        ('Sheet2'),
        ('New Sheet'),
        ('Gogoasd'),
    ],
)
def test_rename_sheet_new_sheet_failed(sheet_name):
    wb = Workbook()
    wb.create_sheet(sheet_name)
    with pytest.raises(ValueError):
        wb.rename_sheet(sheet_name, sheet_name)
    with pytest.raises(ValueError):
        wb.rename_sheet('Sheet1', sheet_name)


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
                (),
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
                (),
                (),
                (),
                (),
                ('qwe', 'bold_font_style'),
                (6, 'DEFAULT_STYLE'),
                (7, 'DEFAULT_STYLE'),
                (-8, 'DEFAULT_STYLE'),
                ('hello', 'DEFAULT_STYLE'),
            ],
        ),
        (slice('A1', 'G100'), [1, 2, 3], ValueError),
        (
            'A1:E1',
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
            'B6:F6',
            ['qwe', 6, 7, -8, 'hello'],
            [
                (),
                ('qwe', 'DEFAULT_STYLE'),
                (6, 'DEFAULT_STYLE'),
                (7, 'DEFAULT_STYLE'),
                (-8, 'DEFAULT_STYLE'),
                ('hello', 'DEFAULT_STYLE'),
            ],
        ),
        (
            'E100:I100',
            [('qwe', 'bold_font_style'), 6, 7, -8, 'hello'],
            [
                (),
                (),
                (),
                (),
                ('qwe', 'bold_font_style'),
                (6, 'DEFAULT_STYLE'),
                (7, 'DEFAULT_STYLE'),
                (-8, 'DEFAULT_STYLE'),
                ('hello', 'DEFAULT_STYLE'),
            ],
        ),
        ('A1:G100', [1, 2, 3], ValueError),
        ('A1:C4', [[1, 2, 3], [1, 2, 3]], ValueError),
        ('A1:C4', [[1, 2, 3], [1, 2, 3], [1, 2, 3], [1, 2, 3, 4]], ValueError),
    ],
)
def test_workbook_slice(row_slice, value_list, expected_output):
    wb = Workbook()
    ws = wb['Sheet1']

    if not isinstance(expected_output, list):
        with pytest.raises(expected_output):
            ws[row_slice] = value_list
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
        (tuple(['1'])),
        (tuple([1])),
    ],
)
def test_set_worksheet_with_wrong_format(cell_value):
    wb = Workbook()
    ws = wb['Sheet1']
    with pytest.raises(ValueError):
        ws[0] = cell_value


def test_save_workbook():
    import io

    wb = Workbook()
    ws = wb['Sheet1']
    ws['A1':'G1'] = [1, 2, 3, 9, 8, 45, 11]

    # Save without calling read_lib_and_create_excel()
    wb.save('test1.xlsx')

    wb.read_lib_and_create_excel()
    wb.save('test2.xlsx')

    # Save with Writable object
    buffer = io.BytesIO()
    wb.save(buffer)


def test_if_style_is_reset():
    from pyfastexcel.manager import StyleManager

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
    from pyfastexcel.manager import StyleManager
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
    from pyfastexcel.manager import StyleManager
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
    from pyfastexcel.manager import StyleManager
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
    assert ws._auto_filter_set == expected_result


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

    with pytest.raises(TypeError):
        wb.create_sheet('new1', pre_allocate={})
    with pytest.raises(TypeError):
        wb.create_sheet('new1', pre_allocate={'n_rows': [], 'n_cols': 20})
    with pytest.raises(TypeError):
        wb.create_sheet('new1', pre_allocate={'n_rows': 10, 'n_cols': 'asdf'})


def test_freeze_set_panes():
    wb = Workbook()
    ws = wb['Sheet1']
    ws[0] = [1, 2, 3]
    ws[1] = [4, 5, 6]

    ws.set_panes(
        freeze=True,
        x_split=6,
        top_left_cell='H1',
        active_pane='topRight',
    )

    wb.read_lib_and_create_excel()


@pytest.mark.parametrize(
    'selection,',
    [
        ([{'sq_ref': 'G36', 'active_cell': 'G36', 'pane': 'topRight'}]),
        ([SelectionDict(sq_ref='G36', active_cell='G36', pane='topRight')]),
        ([Selection(sq_ref='G36', active_cell='G36', pane='topRight')]),
        (Selection(sq_ref='G36', active_cell='G36', pane='topRight')),
    ],
)
def test_split_set_panes(selection):
    wb = Workbook()
    ws = wb['Sheet1']
    ws[0] = [1, 2, 3]
    ws[1] = [4, 5, 6]

    wb.set_panes(
        'Sheet1',
        split=True,
        x_split=6200,
        y_split=9999,
        top_left_cell='N11',
        active_pane='bottomLeft',
        selection=selection,
    )

    ws.set_panes(
        split=True,
        x_split=6200,
        y_split=9999,
        top_left_cell='N11',
        active_pane='bottomLeft',
        selection=selection,
    )

    wb.read_lib_and_create_excel()


@pytest.mark.parametrize(
    'top_left_cell, expected_result',
    [
        ('qwe', ValueError),
        (123, ValidationError),
        ('XFDDDD1', ValueError),
        ('X99999999999999', ValueError),
    ],
)
def test_set_panes_failed_top_left_cell(top_left_cell, expected_result):
    wb = Workbook()
    ws = wb['Sheet1']

    with pytest.raises(expected_result):
        ws.set_panes(
            top_left_cell=top_left_cell,
        )


@pytest.mark.parametrize(
    'x_split, expected_result',
    [
        (-15, ValueError),
        (-999, ValueError),
    ],
)
def test_set_panes_failed_x_split_and_y_split(x_split, expected_result):
    wb = Workbook()
    ws = wb['Sheet1']

    with pytest.raises(expected_result):
        ws.set_panes(
            x_split=x_split,
        )

    with pytest.raises(expected_result):
        ws.set_panes(
            y_split=x_split,
        )


@pytest.mark.parametrize(
    'active_pane, expected_result',
    [
        ('qwe', ValueError),
        (787, ValueError),
        ('hqllw', ValueError),
    ],
)
def test_set_panes_failed_active_pane(active_pane, expected_result):
    wb = Workbook()
    ws = wb['Sheet1']

    with pytest.raises(expected_result):
        ws.set_panes(
            active_pane=active_pane,
        )


@pytest.mark.parametrize(
    'sq_ref, set_range, drop_list, input_msg, error_msg',
    [
        (
            'A1',
            [2, 9],
            ['2', 3, 4, 5],
            ['input_title', 'input_body'],
            ['error_title', 'error_body'],
        ),
        ('A1', [2, 9], ['2', 3, 4, 5], ['input_title', 'input_body'], ''),
        ('A1:A5', [2, 9], 'B1:B5', '', ['error_title', 'error_body']),
    ],
)
def test_set_data_validation(sq_ref, set_range, drop_list, input_msg, error_msg):
    wb = Workbook()
    ws = wb['Sheet1']
    ws['B1'] = 'B1'
    ws['B2'] = 'B2'
    ws['B3'] = 'B3'
    ws['B4'] = 'B4'
    ws['B5'] = 'B5'
    kwargs = {}
    if set_range:
        kwargs['set_range'] = set_range
    if drop_list:
        kwargs['drop_list'] = drop_list
    if input_msg:
        kwargs['input_msg'] = input_msg
    if error_msg:
        kwargs['error_msg'] = error_msg

    ws.set_data_validation(
        sq_ref=sq_ref,
        **kwargs,
    )
    wb.set_data_validation(
        'Sheet1',
        sq_ref=sq_ref,
        **kwargs,
    )


@pytest.mark.parametrize(
    'drop_list, expected_resp',
    [
        ('A2', ValueError),
        ('XFD11', ValueError),
        (123, ValueError),
    ],
)
def test_set_data_validation_drop_list_error(drop_list, expected_resp):
    wb = Workbook()
    ws = wb['Sheet1']

    with pytest.raises(expected_resp):
        ws.set_data_validation(
            sq_ref='A1',
            drop_list=drop_list,
        )


@pytest.mark.parametrize(
    'set_range, expected_resp',
    [
        ('A2', ValueError),
        ([], ValueError),
        (['1', '2', '3'], ValueError),
        ([1, 2, 3], ValueError),
        (12, ValueError),
    ],
)
def test_set_data_validation_set_range_error(set_range, expected_resp):
    wb = Workbook()
    ws = wb['Sheet1']

    with pytest.raises(expected_resp):
        ws.set_data_validation(
            sq_ref='A1',
            set_range=set_range,
        )


@pytest.mark.parametrize(
    'input_msg, error_msg, expected_resp',
    [
        ('title', 'title', ValueError),
        (1, 'title', ValueError),
        (['1', '1', '1'], 'title', ValueError),
        ('title', 2, ValueError),
    ],
)
def test_set_data_validation_msg_error(input_msg, error_msg, expected_resp):
    wb = Workbook()
    ws = wb['Sheet1']
    with pytest.raises(expected_resp):
        ws.set_data_validation(
            sq_ref='A1',
            error_msg=error_msg,
        )

    with pytest.raises(expected_resp):
        ws.set_data_validation(
            sq_ref='A1',
            input_msg=input_msg,
        )


@pytest.mark.parametrize(
    'cell, author, text',
    [
        ('A1', 'Author', 'qqwwee'),
        ('B4', 'qwe', [{'text': 'tqer', 'bold': True, 'color': 'FF0000'}]),
        (
            'B9',
            'aaa',
            [{'text': 'tqer', 'bold': True, 'color': 'FF0000'}, {'text': 'hello', 'italic': True}],
        ),
        ('B9', 'aaa', [CommentText(text='tqer', bold=True, color='FF0000')]),
        ('B9', 'aaa', CommentText(text='tqer', italic=True, vert_align='bottom')),
        ('A1', 'Author', CommentText(text='qwer12')),
    ],
)
def test_add_comment(cell, author, text):
    wb = Workbook()
    ws = wb['Sheet1']

    wb.add_comment('Sheet1', cell, author, text)
    ws.add_comment(cell, author, text)
    wb.read_lib_and_create_excel()


@pytest.mark.parametrize(
    'text',
    [(12321), ([['qwd']]), ([1102]), ({'qwe': 123}), ([{'qq': 'ad'}])],
)
def test_add_comment_failed_text(text):
    wb = Workbook()
    ws = wb['Sheet1']

    with pytest.raises(ValueError):
        ws.add_comment(
            'A1',
            'author',
            text,
        )


@pytest.mark.parametrize(
    'start, end, level, hidden, engine',
    [
        # ('A', 'A', 1, True, 'openpyxl'),
        # ('A', 'C', 12, True, 'openpyxl'),
        # ('A', 'D', 3, True, 'openpyxl'),
        # ('A', 'XD', 1, True, 'openpyxl'),
        # ('A', 'A', 1, False, 'openpyxl'),
        # ('A', 'C', 12, False, 'openpyxl'),
        # ('A', 'D', 3, False, 'openpyxl'),
        # ('A', 'XD', 1, False, 'openpyxl'),
        ('A', 'A', 1, True, 'pyfastexcel'),
        ('A', 'C', 12, True, 'pyfastexcel'),
        ('A', 'D', 3, True, 'pyfastexcel'),
        ('A', 'XD', 1, True, 'pyfastexcel'),
        ('A', 'A', 1, False, 'pyfastexcel'),
        ('A', 'C', 12, False, 'pyfastexcel'),
        ('A', 'D', 3, False, 'pyfastexcel'),
        ('A', 'XD', 1, False, 'pyfastexcel'),
    ],
)
def test_group_column_openpyxl(start, end, level, hidden, engine):
    wb = Workbook()
    ws = wb['Sheet1']
    ws.group_columns(start, end, level, hidden, engine=engine)
    wb.read_lib_and_create_excel()

    wb.create_sheet('Sheet2')
    ws2 = wb['Sheet2']
    ws2.group_columns(start, end, level, hidden, engine=engine)
    wb.read_lib_and_create_excel()

    wb.create_sheet('Sheet3')
    wb.group_columns('Sheet3', start, end, level, hidden, engine=engine)
    wb.read_lib_and_create_excel()


@pytest.mark.parametrize(
    'start, end, level, hidden, engine',
    [
        # (1, 1, 1, True, 'openpyxl'),
        # (1, 3, 12, True, 'openpyxl'),
        # (1, 4, 3, True, 'openpyxl'),
        # (1, 2445, 1, True, 'openpyxl'),
        # (1, 1, 1, False, 'openpyxl'),
        # (1, 3, 12, False, 'openpyxl'),
        # (1, 4, 3, False, 'openpyxl'),
        # (1, 2445, 1, False, 'openpyxl'),
        # (1, None, 1, False, 'openpyxl'),
        # (1, None, 1, False, 'openpyxl'),
        (1, 1, 1, True, 'pyfastexcel'),
        (1, 3, 12, True, 'pyfastexcel'),
        (1, 4, 3, True, 'pyfastexcel'),
        (1, 2445, 1, True, 'pyfastexcel'),
        (1, 1, 1, False, 'pyfastexcel'),
        (1, 3, 12, False, 'pyfastexcel'),
        (1, 4, 3, False, 'pyfastexcel'),
        (1, 2445, 1, False, 'pyfastexcel'),
        (1, None, 1, False, 'pyfastexcel'),
        (1, None, 1, False, 'pyfastexcel'),
    ],
)
def test_group_row_openpyxl(start, end, level, hidden, engine):
    wb = Workbook()
    ws = wb['Sheet1']
    ws.group_rows(start, end, level, hidden, engine=engine)
    wb.read_lib_and_create_excel()

    wb.create_sheet('Sheet2')
    ws2 = wb['Sheet2']
    ws2.group_rows(start, end, level, hidden, engine=engine)
    wb.read_lib_and_create_excel()

    wb.create_sheet('Sheet3')
    wb.group_rows('Sheet3', start, end, level, hidden, engine=engine)
    wb.read_lib_and_create_excel()


@pytest.mark.parametrize(
    'ws1, ws2, cell_range',
    [
        ([1, 2, 3, 4], [5, 6, 7, 8], 'A1:B2'),
        (['q', 'w', 'e', 'r', 'd', 'f'], [1, 2, 3, 4, 5, 7], 'A1:D2'),
        (['q', 'w', 'e', 'r', 'd', 'f'], [1, 2, 3, 4, 5, 7], 'A1:F2'),
    ],
)
def test_create_table(ws1, ws2, cell_range):
    wb = Workbook()
    ws = wb['Sheet1']

    ws[0] = ws1
    ws[1] = ws2

    ws.create_table(cell_range, 'test')

    ws[2] = ['q', 'w', 'e', 'r']
    ws[3] = [2, 2, 3, 4]

    wb.create_table('Sheet1', 'A3:D4', 'test', validate_table=False)

    ws[4] = ['q', 'w', 'e', 'r']
    ws[5] = [2, 2, 3, 4]
    ws.create_table('A5:D6', 'test', show_first_column=True, show_last_column=False)

    ws[6] = ['q', 'w', 'e', 'r']
    ws[7] = [2, 2, 3, 4]
    ws.create_table(cell_range='A7:D8', name='test', show_first_column=True, show_last_column=False)

    ws[8] = ['q', 'w', 'e', 'r']
    ws[9] = [2, 2, 3, 4]
    ws.create_table('A9:D10', name='test', show_first_column=True, show_last_column=False)
    wb.read_lib_and_create_excel()


@pytest.mark.parametrize(
    'ws1, ws2, cell_range',
    [
        ([1, 2, 3, 4], [5, 6, 7, 8], 'A1:B2'),
        (['q', 'w', 'e', 'r', 'd', 'f'], [1, 2, 3, 4, 5, 7], 'A1:D2'),
        (['q', 'w', 'e', 'r', 'd', 'f'], [1, 2, 3, 4, 5, 7], 'A1:F2'),
    ],
)
def test_create_table_with_normal_writer(ws1, ws2, cell_range):
    wb = Workbook()
    ws = wb['Sheet1']

    ws[0] = ws1
    ws[1] = ws2

    ws.create_table(cell_range, 'test')
    # Make pyfastexcel use normal wirter to write content
    ws.group_columns('F1')
    wb.read_lib_and_create_excel()


def test_create_table_failed_args():
    wb = Workbook()
    ws = wb['Sheet1']

    ws[0] = ['t1', 't2', 't3', 't4']
    ws[1] = [2, 2, 3, 4]

    with pytest.raises(ValueError):
        ws.create_table('A1B2', 'test')

    with pytest.raises(ValueError):
        ws.create_table('A1:B2:1', 'test')

    with pytest.raises(ValueError):
        ws.create_table('A1:B2', 'test', style_name='wrong_style')


@pytest.mark.parametrize(
    'ws1, ws2, cell_range',
    [
        ([1, 2, 3, 4], [5, 6, 7, 8], 'A1:G2'),
        ([1, 1, 1, 1], [5, 6, 7, 8], 'A1:D2'),
        ([1, 2, 3, 4], [1, 2, 3, 4], 'A1:D2'),
    ],
)
def test_create_table_failed_final_validation(ws1, ws2, cell_range):
    wb = Workbook()
    ws = wb['Sheet1']

    ws[0] = ws1
    ws[1] = ws2

    with pytest.raises(ValueError):
        ws.create_table(cell_range, 'test')
        wb.read_lib_and_create_excel()


def test_sheet_visible():
    wb = Workbook()
    wb.create_sheet('Sheet2')
    wb.create_sheet('Sheet3')

    wb['Sheet2'].sheet_visible = False
    assert not wb['Sheet2'].sheet_visible
    print(wb['Sheet1'].sheet_visible)

    with pytest.raises(ValueError):
        wb['Sheet1'].sheet_visible = 'qwe'
