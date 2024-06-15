# WorkSheet

The Worksheet is the core component of the application. It stores cell values, formulas, and styles, and provides methods to manipulate the cells.

## Create a WorkSheet

The default `WorkSheet` 'Sheet1' is created when the `Workbook` is created. You can access the `WorkSheet` through an index, as shown in the code snippet below:

```python title="Access the default WorkSheet"
from pyfastexcel import Workbook

wb = Workbook()
ws = wb['Sheet1']
```

You can also create a new `WorkSheet` by calling the `create_sheet(sheet_name: str)` function:

| Parameter           | Data Type    | Description                                     |
|---------------------|--------------|-------------------------------------------------|
| `sheet_name`        | `str`        | Sheet Name                                      |
| `plain_data`        | `list[list]` (Optional) | Row and Column to write the excel without style |

```python title="Create a new WorkSheet"
wb.create_sheet('New Sheet')
ws_new = wb['New Sheet']
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

The cell can be merge through the function

| Parameter           | Data Type | Description                    |
|---------------------|-----------|--------------------------------|
| `top_left_cell`     | str       | The index of the top left cell|
| `bottom_right_cell` | str       | The index of the bottom right cell|

```python title='Merge Cells'
ws.merge_cell('A1', 'B2')
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
