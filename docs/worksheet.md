# WorkSheet

The Worksheet is the core component of the application. It stores cell values, formulas, and styles, and provides methods to manipulate the cells.

## Create a WorkSheet

The default `WorkSheet` 'Sheet1' is created when the `Workbook` is created. You can access the `WorkSheet` through an index, as shown in the code snippet below:

| Parameter           | Data Type    | Description                                     |
|---------------------|--------------|-------------------------------------------------|
| `sheet_name`        | `str`        | Sheet Name                                      |
| `pre_allocate`      | `dict[str, int]` | Pre-allocate the memory space of given row and column numbers |
| `plain_data`        | `list[list]` (Optional) | Row and Column to write the excel without style |

```python title="Access the default WorkSheet"
from pyfastexcel import Workbook

wb = Workbook()
ws = wb['Sheet1']
```

You can also create a new `WorkSheet` by calling the `create_sheet(sheet_name: str)` function:


```python title="Create a new WorkSheet"
wb.create_sheet('New Sheet')
ws_new = wb['New Sheet']
```

If you know the dimension of the data you want to write. You can use `pre_allocate`
to pre_allocate the memory space of the pyfastexcel to improve the performance.

```python
from pyfastexcel import Workbook


pre_allocate = {'n_rows': 1000, 'n_cols': 10}
# This will pre_allocate the memory space of the pyfastexcel
wb = Workbook(pre_allocate=pre_allocate)
wb.save('pre_allocate.xlsx')
```

If you don't need any style for the excel. You could also write the excel without any style with the following code **(This is the fastest way to write the excel)**:

```python
from pyfastexcel import Workbook


data = [[1, 2, 3], [4, 5, 6, 7, 8]]
# This will write the data into the default sheet 'Sheet1'
wb = Workbook()
wb.create_sheet('New Sheet', plain_data=data)
wb.save('plain_data.xlsx')
```

!!! note="Note"
    "You can only specify either `pre_allocate` or `plain_data` at a time, not both.

## Assign a value to a cell

There are multiple methods to assign a value and style to a cell. If you would like to adopt
a style to a cell, you should create a `CustomStyle` first and register it with the `set_custom_style` function. The first args in the `set_custom_style` function is the name of the style, and the second args is the `CustomStyle` instance. After you register the style, you can assign the style to a cell by passing the style name as the second element in the tuple. Besides, you can also assign the style by directly passing the `CustomStyle` instance to the cell.

```python title="Assign a value via excel index"
from pyfastexcel import Workbook
from pyfastexcel.utils import set_custom_style

from openpyxl_style_writer import CustomStyle

wb = Workbook()
ws = wb['Sheet1']

# Create and register the style
bold_style = CustomStyle(font_size=19, font_bold=True)
set_custom_style('bold_style', bold_style)

# Assign a value to a cell with default_style
ws['A1'] = 1
# Assign a value to a cell with custom_style
ws['A2'] = (2, 'bold_style')

# Assign a value to a cell with custom_style instance (This method don't need to call set_custom_style)
ws['A3'] = (9, bold_style)
```

!!! note "Style Setting"
    `pyfastexcel` accept the style as a tuple `(value, style)`. If the style is not provided, the default style will be used. The default style is defined in the `Workbook` class.

!!! note "Style assignment without `set_custom_style`"
    `pyfastexcel` support the style assignment without calling `set_custom_style`. You can directly pass the `CustomStyle` instance to the cell. `pyfastexcel` will register the style name with the auto increment id. This may lead to the style being duplicated in the workbook, and potentially decrease the performance.

!!! info "Why not use a class instance for 'Cells'?"
    Although creating the cell value as a Cell class can make the cell
    assignment more readable, such as `#!python ws['A1'].value = 10`
    or `#!python ws['A1'].style = a_style`, the tuple method is more
    efficient in terms of performance.


```python title="Assign a value via slicing"
from pyfastexcel import CustomStyle, Workbook
from pyfastexcel.utils import set_custom_style


wb = Workbook()
ws = wb['Sheet1']

# Create and register the style
bold_style = CustomStyle(font_size=19, font_bold=True)
set_custom_style('bold_style', bold_style)

# Assign a value to a cell with default_style
ws['A1':'D1'] = [1, 2, 3, 4]
# Assign a value to a cell with string slice
ws['E1:G1'] = [1, 2, 3]
# Assign a value to a cell with custom_style
ws['A2':'D2'] = [
    (2, 'bold_style'),
    ('12', 'bold_style'),
    6,  # (1)
    (8, 'bold_style')
]
```

1. The default style will be used.

```python title="Assign a value via row index"
from pyfastexcel import CustomStyle, Workbook
from pyfastexcel.utils import set_custom_style


wb = Workbook()
ws = wb['Sheet1']

# Create and register the style
bold_style = CustomStyle(font_size=19, font_bold=True)
set_custom_style('bold_style', bold_style)

# Assign a value to a cell with default_style
ws[0] = [1, 2, 3, 4]
# Assign a value to a cell with custom_style
ws[1] = [
    (2, 'bold_style'),
    ('12', 'bold_style'),
    6
    (8, 'bold_style')
]
```

## Set Cell by row and columns index

Set the cell value by row and column index.

| Parameter |     Data Type      | Description                   |
|-----------|------------------- |-------------------------------|
| `row`     | int                | Target row.                   |
| `column`  | int                | Target column.                |
| `value`   | Any                | The value to set in the cell. |
| `style`   | CustomStyle or str | Style to apply to the cells.  |

```python title="Set Cell by row and column index"
from pyfastexcel import Workbook
from pyfastexcel.utils import set_custom_style

wb = Workbook()
ws = wb['Sheet1']

# Create and register the style
bold_style = CustomStyle(font_size=19, font_bold=True)
set_custom_style('bold_style', bold_style)

# Set style with the register name
ws.cell(0, 0, 'Hello', style='bold_style')
# Set style with the CustomStyle instance
ws.cell(0, 1, '123', style=bold_style)
```

!!! note "Note"
    The row and column index are 0-based. So if you want to set the value in the first row and the first column like `A1` in excel, you should use `ws.cell(0, 0, 'Hello')`.

## Set Style

Set style with input coordinate.

| Parameter | Data Type                      | Description                    |
|-----------|--------------------------------|--------------------------------|
| `target`  | str, slice, or list[int, int]  | Target cells to apply style.   |
| `style`   | CustomStyle or str             | Style to apply to the cells.   |

```python
from pyfastexcel import Workbook
from pyfastexcel.utils import set_custom_style

wb = Workbook()
ws = wb['Sheet1']

# Create and register the style
bold_style = CustomStyle(font_size=19, font_bold=True)
set_custom_style('bold_style', bold_style)

ws[0] = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']

# Set style with CustomStyle instance
ws.set_style('A1', bold_style)
# Set style with the register name
ws.set_style('B1', 'bold_style')

# Set style with the string slice
ws.set_style('C1:E1', 'bold_style')
ws.set_style(slice('F1', 'H1'), 'bold_style')

# Set style with the row and column index ([0, 8] = 'I1')
ws.set_style([0, 8], 'bold_style')
```

## Set cell width and height

Set column width

| Parameter           | Data Type | Description                    |
|---------------------|-----------|--------------------------------|
| `col`               | str       | The column number             |
| `value`             | str       | The value of the width        |

```python title="Set Width"
# Set through alphabet
ws.set_cell_width('A', 20)
# Set through number
ws.set_cell_width(1, 23)
```

Set row height

| Parameter           | Data Type | Description                    |
|---------------------|-----------|--------------------------------|
| `row`               | str or int| The row number                |
| `value`             | int       | The value of the height       |

```python title="Set height"
# Set row 15 to height = 20
ws.set_cell_height(15, 20)
```

## Merge Cell

The cell can be merged through the function. You can choose to use either two
parameters or one parameter to merge.

!!! note "Note"
    1. The function supports merging cells using either individual
    cell references or a cell range.
    2. Ensure that the `top-left cell` is specified before the
    `bottom-right cell` when using two parameters.
    3. Cell references should be valid within the sheet's
    boundaries (rows: 1 to 1,048,576, columns: A to XFD).

| Parameter           | Data Type | Description                    |
|---------------------|-----------|--------------------------------|
| `top_left_cell`     | str       | The index of the top left cell|
| `bottom_right_cell` | str       | The index of the bottom right cell|

```python title='Merge Cells'
# Merge cells using individual cell references
ws.merge_cell('A1', 'B2')
```

| Parameter           | Data Type | Description                    |
|---------------------|-----------|--------------------------------|
| `cell_range`        | str       | The cell range to merge|

```python title='Merge Cells'
# Merge cells using a cell range
ws.merge_cell('A1:B2')
```

## AutoFilter

Create an auto filter in a worksheet.

### Parameters

| Parameter      | Data Type | Description                       |
|----------------|-----------|-----------------------------------|
| `target_range` | str       | The range where the auto filter will be applied. |

### Example

```python title='Auto Filter'
ws.auto_filter("A1:C1")
```

## Set Panes

Configure the pane settings for a specific sheet in an Excel file using the provided Excelize file.

### Parameters

| Parameter      | Data Type                | Description                                     |
|----------------|--------------------------|-------------------------------------------------|
| `freeze`       | bool                     | Determines if the panes are frozen.              |
| `split`        | bool                     | Determines if the panes are split.               |
| `x_split`      | int                      | The horizontal position where the panes are split, or the column index that should be frozen. |
| `y_split`      | int                      | The vertical position where the panes are split, or the row index that should be frozen.  |
| `top_left_cell`| str                      | The cell at the top left of the visible window.   |
| `active_pane`  | str                      | The active pane.                                 |
| `selection`    | list[dict[str, str]]     | The selection settings for panes.                |

!!! note "Key of selection"
    The key of the selection is `sq_ref`, `active_cell`, and `pane`. `sq_ref` and `active_cell`
    should be a cell reference, and `pane` should be one of the `topLeft`, `topRight`, `bottomLeft`, and `bottomRight`.

!!! note "Options of active_pane"
    The options for `active_pane` and are `topLeft`, `topRight`, `bottomLeft`, and `bottomRight`.

!!! note "Note"
    When `freeze` is set to `true`, `x_split` and `y_split` represent the column or row index where the panes are
    frozen. When `split` is set to `true`, `x_split` and `y_split` represent the pixel position where the panes are split.

### Example

```python title='Panes Configuration'
from pyfastexcel.utils import Selection

# Set panes with Selection instance
ws.set_panes(
    freeze=True,
    split=False,
    x_split=0,
    y_split=0,
    top_left_cell="A1",
    active_pane="topRight",
    selection=[
        Selection(sq_ref="A1", active_cell="A1", pane="topRight")
    ],
)

# Set panes's selection with dict
ws.set_panes(
    freeze=True,
    split=False,
    x_split=0,
    y_split=0,
    top_left_cell="A1",
    active_pane="topRight",
    selection=[
        {
            "sq_ref": "A1",
            "active_cell": "A1",
            "pane": "topRight",
        }
    ],
)
```

## Set Data Validation

Set data validation for a specified range in a worksheet.

### Parameters

| Parameter      | Data Type                | Description                                                             |
|----------------|--------------------------|-------------------------------------------------------------------------|
| `sq_ref`       | str                      | The range to set the data validation.                                   |
| `set_range`    | list[int or float]       | The range of values to set the data validation.                         |
| `input_msg`    | list[str]                | The input message for the data validation. Must be a list with two elements: [Title, Body]. |
| `drop_list`    | list[str] or str         | The drop list for the data validation. Can be a list of strings or a range in the format "A1:B2".  |
| `error_msg`    | list[str]                | The error message for the data validation. Must be a list with two elements: [Title, Body]. |

### Example

```python title='Set Data Validation'
# Example 1: Setting data validation with a specified range, input message, drop-down list, and error message
ws.set_data_validation(
    sq_ref="A1:B2",
    set_range=[1, 10],
    input_msg=["Input Title", "Input Body"],
    drop_list=["Option1", "Option2", "Option3"],
    error_msg=["Error Title", "Error Body"]
)

# Example 2: Setting data validation with a drop-down list based on cell values
ws.set_data_validation(
    sq_ref="A1:B2",
    drop_list="C1:C5",
)
```

## Add Comment

Adds a comment to the specified cell.

### Parameters

| Parameter  | Data Type                                 | Description                                           |
|------------|-------------------------------------------|-------------------------------------------------------|
| `cell`     | `str`                                       | The cell location to add the comment.                 |
| `author`   | `str`                                       | The author of the comment.                            |
| `text`     | `str` or `dict[str, str]` or `list[str or dict[str, str]]` | The text of the comment, and it's font style|

### Example

```python title='Add Comment'
from pyfastexcel.utils import CommentText

# Add a comment to cell A1 with CommentText Instance
comment_text = CommentText(text='Comment', bold=True)
ws.add_comment("A1", "pyfastexcel", comment_text)

# Add a comment to cell A1 with list of CommentText Instance
comment_text = CommentText(text='Comment', bold=True)
comment_text2 = CommentText(text=' Comment two', color='00ff00')
ws.add_comment("A1", "pyfastexcel", [comment_text, comment_text2])

# Add a comment to cell A1, and use string as the comment text
ws.add_comment("A1", "pyfastexcel", "This is a comment.")

# Add a comment to cell A1, and use dictionary as the comment text and set the font style
ws.add_comment("A1", "pyfastexcel", {"text": "This is a comment.", 'bold': True, 'italic': True})

# Add a comment to cell A1, and use list of dictionary as the comment text and set the font style
# This will create "This is a comment" with bold and italic font style, and "This is another comment" with bold and red color font style.
ws.add_comment(
    "A1",
    "pyfastexcel",
    [
        {
            "text": "This is a comment.",
            'bold': True,
            'italic': True
        },
        {
            "text": "This is another comment.",
            'bold': True,
            'color': 'FF0000'
        }
    ]
)
```

Here is the `key words of the comment when using the dictionary`:

| Key            | Data Type | Description                       |
|----------------|-----------|-----------------------------------|
| `text`         | str       | The text of the comment.          |
| `size`         | int       | The font size of the comment text. |
| `name`         | str       | The font name of the comment text. |
| `bold`         | bool      | Sets the comment text to bold. |
| `italic`       | bool      | Sets the comment text to italic. |
| `underline`    | str       | Sets the underline style of the comment text. |
| `strike`       | bool      | Sets whether the comment text is strike through. |
| `vertAlign`    | str       | Sets the vertical alignment of the comment text. |
| `color`        | str       | Sets the font color of the comment text. |


!!! note "Note"
    The `text` parameter can be a string, a dictionary, or a list of dictionaries. If it is a string, it will be treated as the comment text. If it is a dictionary, it should contain the key `text` with the comment's text as the corresponding value. If it is a list of dictionaries, each dictionary should contain the key `text` with the comment's text as the corresponding value.

## Group Columns

Group columns in a worksheet. This function is currently implemented using `openpyxl`.
It is not recommended to use this function when dealing with large files.

!!! note "Note"
    - **Excelize** does not currently support column grouping in Streaming mode.
    - By default, if you call the `group_columns` function, **pyfastexcel** will
    write the Excel file using the normal API of **Excelize**, which is slower
    than Streaming mode.
    - As an alternative method, you can set the `engine` parameter to
    `openpyxl`. This will allow **pyfastexcel** to first write other Excel
    content using Streaming mode, and then load_workbook with `openpyxl` and
    use the `openpyxl` API to group columns and save the file.

### Parameters

| Parameter      | Data Type        | Description                                      |
|----------------|------------------|------------------------------------------------- |
| `start_col`    | str              | The cell reference where grouping starts.        |
| `end_col`      | Optional[str]    | The cell reference where grouping ends.          |
| `outline_level`| int              | The level of grouping.                           |
| `hidden`       | bool             | Whether to hide the group or not.                |
| `engine`       | Literal['pyfastexcel', 'openpyxl'] | The engine to group columns    |

```python
ws.group_columns('A', 'C', 1, False)
```

## Group Rows

Group Rows in a worksheet. This function is currently implemented using `openpyxl`.
It is not recommended to use this function when dealing with large files.

!!! note "Note"
    - **Excelize** does not currently support column grouping in Streaming mode.
    - By default, if you call the `group_rows` function, **pyfastexcel** will
    write the Excel file using the normal API of **Excelize**, which is slower
    than Streaming mode.
    - As an alternative method, you can set the `engine` parameter to
    `openpyxl`. This will allow **pyfastexcel** to first write other Excel
    content using Streaming mode, and then load_workbook with `openpyxl` and
    use the `openpyxl` API to group rows and save the file.

### Parameters

| Parameter      | Data Type        | Description                                      |
|----------------|------------------|------------------------------------------------- |
| `start_row`    | int              | The row reference where grouping starts.         |
| `end_row`      | Optional[int]    | The row reference where grouping ends.           |
| `outline_level`| int              | The level of grouping.                           |
| `hidden`       | bool             | Whether to hide the group or not.                |
| `engine`       | Literal['pyfastexcel', 'openpyxl'] | The engine to group rows       |

```python
ws.group_rows(1, 3, 1, False)
```

## Create Table

Create a table in a sheet.

!!! note "Note"
    There are some limitations when creating a table:
    1. A table must always be created with at least one row of data.
    For example, if you want to create a table in the range 'A1:B3',
    you should first ensure that there is data in the range 'A1:A3';
    otherwise, the table will not be created correctly.
    2. Tables should not overlap with one another.

### Parameters

| Parameter            | Data Type | Description                                      |
|----------------------|-----------|--------------------------------------------------|
| `cell_range`         | str       | The cell reference range, e.g., 'A1:B3'.         |
| `name`               | str       | The name of the table.                           |
| `style_name`         | str       | The built-in style name for the table in Excel.  |
| `show_first_column`  | bool      | Whether to display the first column.             |
| `show_last_column`   | bool      | Whether to display the last column.              |
| `show_row_stripes`   | bool      | Whether to display row stripes.                  |
| `show_column_stripes`| bool      | Whether to display column stripes.               |

```python
ws.create_table(
    'A1:B3',
    'table_name',
    'TableStyleLight1,
    True,
    True,
    False,
    True,
)
```
