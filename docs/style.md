## CustomStyle

### Overview

The `CustomStyle` class is a wrapper for defining Excel styles using the [openpyxl](https://openpyxl.readthedocs.io/en/stable/) library. It provides a convenient way to create custom styles for formatting cells in Excel spreadsheets. `pyfastexcel` use `CustomStyle` as an interface to get and set the style for each cell.

!!! info "CustomStyle"

    Currently, Customstyle is depends on the [openpyxl](https://openpyxl.readthedocs.io/en/stable/)
    and [openpyxl_style_writer](https://github.com/Zncl2222/openpyxl_style_writer).
    We plan to implement the Style object making it no longer depends on
    openpyxl and openpyxl_style_writer in the future.


### Create CustomStyle

Users can set the `CustomStyle` via the alias defined by `openpyxl_style_writer`.

For example we can create a `CustomStyle` where the font is bold and font color is yellow like this.

```python title="CustomStyle"
from openpyxl_style_writer import CustomStyle


yellow_bold_style = CustomStyle(font_bold=True, font_color='ffff00')
```

Also, you can globally change the default style like

```python title="Change Default Style Globally"
from openpyxl_style_writer import DefautStyle


DefaultStyle.set_default(font_size=16)
```

Additionally, it is possible to set the style with the original style name from `opnepyxl`

```python title="Set style by params"
from opnpyxl.styles import Side

from openpyxl_style_writer import CustomStyle

# Create the dict that the key is the style name from openpyxl
blue_title_font = {
    'color': '0000ff',
    'bold': True,
    'size': 15,
}
cyan_title_pattern = {
    'patternType': 'solid',
    'fgColor': '00ffff'
}
border = {
    'left': Side(style='medium', color='cccccc'),
    'right': Side(style='thin', color='cccccc'),
    'top': Side(style='double', color='cccccc'),
    'bottom': Side(style='dashed', color='cccccc'),
}

custom_title_style = CustomStyle(
    font_params=blue_title_font,
    fill_params=cyan_titl_patter,
    border_params=border,
)
```

### Alias in CustomStyle

The following code snippet demonstrates the available aliases in CustomStyle.

```python title="CustomStyle with all arguments"
from opnepyxl_style_writer import CustomStyle


# This is the defualt settings, you can pass the arguments you want only.
default_style = CustomStyle(
    # Font
    font_size=11,
    font_name='Calibri',
    font_bold=False,
    font_italic=False,
    font_underline='none',
    font_strike=False,
    font_vertAlign=None,
    font_color='000000',
    fill_pattern='solid',
    fill_color='ffffff',

    # alignment
    ali_horizontal=None,
    ali_vertical='bottom',
    ali_text_rotation=0,
    ali_wrap_text=False,
    ali_shrink_to_fit=False,
    ali_indent=0,

    # border
    border_style_top='thin',
    border_style_right='thin',
    border_style_left='thin',
    border_style_bottom='thin',
    border_color_top='C0C0C0',
    border_color_right='C0C0C0',
    border_color_left='C0C0C0',
    border_color_bottom='C0C0C0',

    # protect
    protect=False,
)
```
