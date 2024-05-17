# Common Functions

This page list the utility functions of pyfastexcel

## set_custom_style

Register the custom_style

| Parameter    | Data Type   | Description              |
|--------------|-------------|--------------------------|
| `style_name` | str         | The name of the style    |
| `style`      | CustomStyle | CustomStyle instance     |

```python title="set_custom_style"
from pyfastexcel import CustomStyle
from pyfastexcel.utils import set_custom_style

bold_style = CustomStyle(font_bold=True)

set_custom_style('bold_style', bold_style)
```

!!! note "Note"
    After registering the style using set_custom_style, you can choose to set
    the style either by its `name` or by the `CustomStyle` instance. For example,
    you can use `#!python ws['A1'] = (123, 'bold_style')` or
    `#!python ws['A1'] = (123, bold_style)`.

## `column_to_index`

Converts an Excel column name to an index, e.g., 'A' -> 1.

| Parameter | Data Type | Description          |
|-----------|-----------|----------------------|
| `column`  | str       | The Excel column name |

```python title="column_to_index"
from pyfastexcel.utils import column_to_index

index = column_to_index('A') # index = 1
```

## index_to_column

Converts an index to an Excel column name.

| Parameter    | Data Type   | Description              |
|--------------|-------------|--------------------------|
| `index`      | int         | The index                |

```python title="column_to_index"
from pyfastexcel.utils import index_to_column

column = column_to_index(1) # column = 'A'
```
