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

```python title="Create a new WorkSheet"
wb.create_sheet('New Sheet')
ws_new = wb['New Sheet']
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
from pyfastexcel import Workbook
from pyfastexcel.utils import set_custom_style

from openpyxl_style_writer import CustomStyle

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
from pyfastexcel import Workbook
from pyfastexcel.utils import set_custom_style

from openpyxl_style_writer import CustomStyle

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

## Set cell width and height

The cell widht can be set with the function

| Parameter           | Data Type | Description                    |
|---------------------|-----------|--------------------------------|
| `col`               | str       | The column number             |
| `value`             | str       | The value of the width        |

```python title="Set Width"
# Set through alphabet
ws.set_cell_width('A', 20)
# Set through number
ws.set_cell_width(, 23)
```

The cell height can be set with the function

| Parameter           | Data Type | Description                    |
|---------------------|-----------|--------------------------------|
| `row`               | str or int| The row number                |
| `value`             | int       | The value of the height       |

```python title="Set height"
# Set row 15 to height = 20
ws.set_cell_height( 15, 20)
```

## Merge Cell

The cell can be merge through the function

| Parameter           | Data Type | Description                    |
|---------------------|-----------|--------------------------------|
| `top_left_cell`     | str       | The index of the top left cell|
| `bottom_right_cell` | str       | The index of the bottom right cell|

```python title='Merge Cells'
ws.set_merge_cell('A1', 'B2')
```
