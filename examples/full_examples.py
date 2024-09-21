import random

from pyfastexcel import CustomStyle, Workbook
from pyfastexcel.worksheet import WorkSheet

from pyfastexcel.utils import CommentText, Selection, set_custom_style
from pyfastexcel.chart import (
    Chart,
    ChartSeries,
    RichTextRun,
    Font,
    ChartAxis,
    ChartLegend,
    Fill,
    Marker,
)
from pyfastexcel.pivot import PivotTable, PivotTableField


def setup(wb: Workbook, sheet_name: str) -> WorkSheet:
    wb.create_sheet(sheet_name)
    ws = wb[sheet_name]

    columns = ['Year', 'Month', 'Market', 'Location', 'Sales']
    month = [
        'Jan',
        'Feb',
        'Mar',
        'Apr',
        'May',
        'Jun',
        'Jul',
        'Aug',
        'Sep',
        'Oct',
        'Nov',
        'Dec',
    ]
    year = [2024, 2025, 2026, 2027, 2028]
    location = ['North', 'South', 'East', 'West']
    market = ['A', 'B', 'C', 'D']

    ws[0] = columns
    for i in range(1, 60):
        ws[f'A{i+1}'] = random.choice(year)
        ws[f'B{i+1}'] = random.choice(month)
        ws[f'C{i+1}'] = random.choice(market)
        ws[f'D{i+1}'] = random.choice(location)
        ws[f'E{i+1}'] = random.randint(-1000, 10000)

    return ws


def set_style(wb: Workbook):
    ws = setup(wb, 'SetStyle')
    red_font = CustomStyle(font_color='FF0000')
    green_font = CustomStyle(font_color='00FF00')
    black_fill = CustomStyle(fill_color='000000')

    ws.set_style('A1', red_font)
    ws.set_style('B1', green_font)

    set_custom_style('black_fill', black_fill)
    ws.set_style('C1', 'black_fill')


def merge_cell_examples(wb: Workbook):
    ws = setup(wb, 'MergeCells')
    ws.merge_cell('A1', 'C1')
    ws.merge_cell('B1', 'C2')
    ws.merge_cell('D1:G5')


def set_panes_split_examples(wb: Workbook):
    ws = setup(wb, 'SetPanes')

    # Set panes with Selection instance
    ws.set_panes(
        split=True,
        x_split=3500,
        y_split=3500,
        top_left_cell='L30',
        active_pane='bottomLeft',
        selection=[Selection(sq_ref='A1', active_cell='A1', pane='topRight')],
    )


def set_panes_freeze_examples(wb: Workbook):
    ws = setup(wb, 'FreezePanes')

    # Freeze 1 to 6 rows
    ws.set_panes(
        freeze=True,
        y_split=6,
        top_left_cell='A34',
        active_pane='bottomLeft',
        selection=[Selection(sq_ref='A7', active_cell='A7', pane='bottomLeft')],
    )


def set_data_validation_examples(wb: Workbook):
    ws = setup(wb, 'DataValidation')
    # Example 1: Setting data validation with a specified range, input message, drop-down list, and error message
    ws.set_data_validation(
        sq_ref='A1:B2',
        set_range=[1, 10],
        input_msg=['Input Title', 'Input Body'],
        drop_list=['Option1', 'Option2', 'Option3'],
        error_msg=['Error Title', 'Error Body'],
    )

    # Example 2: Setting data validation with a drop-down list based on cell values
    ws.set_data_validation(
        sq_ref='A1:B2',
        drop_list='C1:C5',
    )


def add_comment_examples(wb: Workbook):
    ws = setup(wb, 'Comment')
    # Add a comment to cell A1 with CommentText Instance
    comment_text = CommentText(text='Comment', bold=True)
    ws.add_comment('A1', 'pyfastexcel', comment_text)

    # Add a comment to cell A1 with list of CommentText Instance
    comment_text = CommentText(text='Comment', bold=True)
    comment_text2 = CommentText(text=' Comment two', color='00ff00')
    ws.add_comment('A2', 'pyfastexcel', [comment_text, comment_text2])

    # Add a comment to cell A1, and use string as the comment text
    ws.add_comment('B1', 'pyfastexcel', 'This is a comment.')

    # Add a comment to cell A1, and use dictionary as the comment text and set the font style
    ws.add_comment(
        'C1', 'pyfastexcel', {'text': 'This is a comment.', 'bold': True, 'italic': True}
    )

    # Add a comment to cell A1, and use list of dictionary as the comment text and set the font style
    # This will create "This is a comment" with bold and italic font style, and "This is another comment" with bold and red color font style.
    ws.add_comment(
        'D1',
        'pyfastexcel',
        [
            {'text': 'This is a comment.', 'bold': True, 'italic': True},
            {'text': 'This is another comment.', 'bold': True, 'color': 'FF0000'},
        ],
    )


def group_columns_and_rows_examples(wb: Workbook):
    ws = setup(wb, 'GroupColumnAndRows')
    ws.group_columns('A', 'C', 1, False)
    ws.group_columns('D', 'E', 1, True)
    ws.group_rows(1, 3, 1, False)
    ws.group_rows(5, 8, 1, True)


def create_table_examples(wb: Workbook):
    ws = setup(wb, 'Table')
    ws.create_table(
        'A1:B3',
        'table_name',
        'TableStyleLight1',
        True,
        True,
        False,
        True,
    )


def add_chart_examples(wb: Workbook):
    wb.create_sheet('Chart')
    ws = wb['Chart']

    ws[0] = ['Category', '2024/01', '2024/02', '2024/03']
    ws[1] = ['Food', 123, 125, 645]
    ws[2] = ['Book', 456, 789, 321]
    ws[3] = ['Phone', 777, 66, 214]

    column_chart = Chart(
        chart_type='col',
        series=[
            ChartSeries(
                name='Chart!A2',
                categories='Chart!B1:D1',
                values='Chart!B2:D2',
                fill=Fill(ftype='pattern', pattern=1, color='ebce42'),
                marker=Marker(symbol='none'),
            ),
            ChartSeries(
                name='Chart!A3',
                categories='Chart!B1:D1',
                values='Chart!B3:D3',
                fill=Fill(ftype='pattern', pattern=1, color='29a64b'),
                marker=Marker(symbol='none'),
            ),
        ],
        legend=ChartLegend(position='top', show_legend_key=True),
    )

    line_chart = Chart(
        chart_type='line',
        series=[
            ChartSeries(
                name='Chart!A4',
                categories='Chart!B1:D1',
                values='Chart!B4:D4',
                fill=Fill(ftype='pattern', pattern=1, color='0000FF'),
                marker=Marker(
                    symbol='circle',
                    fill=Fill(ftype='pattern', pattern=1, color='FFFF00'),
                ),
            ),
        ],
        title=[RichTextRun(text='Example Chart', font=Font(color='FF0000', bold=True))],
        x_axis=ChartAxis(major_grid_lines=True, font=Font(color='000000')),
        y_axis=ChartAxis(major_grid_lines=True, font=Font(color='000000')),
        legend=ChartLegend(position='top', show_legend_key=True),
    )
    ws.add_chart('E1', [column_chart, line_chart])


def add_pivot_table_examples(wb: Workbook):
    ws = setup(wb, 'PivotTable')

    pivot = PivotTable(
        data_range='PivotTable!A1:E60',
        pivot_table_range='PivotTable!H2:N60',
        rows=[PivotTableField(data='Year', default_subtotal=True), PivotTableField(data='Month')],
        pivot_filter=[PivotTableField(data='Market')],
        columns=[PivotTableField(data='Location')],
        data=[PivotTableField(data='Sales', name='Summation', subtotal='sum')],
        show_drill=True,
        row_grand_totals=True,
        column_grand_totals=True,
        show_row_headers=True,
        show_column_headers=True,
        show_last_column=True,
        pivot_table_style_name='PivotStyleLight16',
    )
    ws.add_pivot_table(pivot)


if __name__ == '__main__':
    wb = Workbook()
    set_style(wb)
    merge_cell_examples(wb)
    wb.remove_sheet('Sheet1')
    set_panes_split_examples(wb)
    set_panes_freeze_examples(wb)
    set_data_validation_examples(wb)
    add_comment_examples(wb)
    group_columns_and_rows_examples(wb)
    create_table_examples(wb)
    add_chart_examples(wb)
    add_pivot_table_examples(wb)
    wb.save('FullExamples.xlsx')
