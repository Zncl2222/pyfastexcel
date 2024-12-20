## CustomStyle

### Overview

The `CustomStyle` provides a convenient way to create custom styles for formatting cells in Excel spreadsheets. `pyfastexcel` use `CustomStyle` as an interface to get and set the style for each cell.


### Create CustomStyle

Create a `CustomStyle` where the font is bold and font color is yellow

```python title="CustomStyle"
from pyfastexcel import CustomStyle


yellow_bold_style = CustomStyle(font_bold=True, font_color='ffff00')
```

Then you can use the `yellow_bold_style` to set the style for the cell. See [worksheet]() for more details that how to set the style for the cell.

```python title="Set Style"
# Set the style for the cell A1
ws['A1'] = ('Hello', yellow_bold_style)

# Set the style for the cell B1
ws.cell(row=1, column=2, value='World', style=yellow_bold_style)
```

### Set Default Style

You can change the default style globally by using the following code:

```python title="Change Default Style Globally"
from pyfastexcel import DefautStyle


DefaultStyle.set_default(font_size=16)
```

Once you set the default style, any new `CustomStyle` will have a `font_size` of `16` unless you specify a different `font_siz`e for that `CustomStyle`.

### Set Style by Params

Additionally, it is possible to set the style with the original style name with
dictionary.

```python title="Set style by params"
from pyfastexcel.style import BorderStyle
from pyfastexcel import CustomStyle

# Create the dict that the key is the style name and the value is the style value
blue_title_font = {
    'color': '0000ff',
    'bold': True,
    'size': 15,
}
cyan_title_pattern = {
    'pattern': 'solid',
    'color': '0000ff',
}
border = {
    'left': BorderStyle(style='medium', color='cccccc'),
    'right': BorderStyle(style='thin', color='cccccc'),
    'top': BorderStyle(style='double', color='cccccc'),
    'bottom': BorderStyle(style='dashed', color='cccccc'),
}

custom_title_style = CustomStyle(
    font_params=blue_title_font,
    fill_params=cyan_title_patter,
    border_params=border,
)
```

### CustomStyle & DefaultStyle

The `CustomStyle` and `DefaultStyle` have the same arguments. The `DefaultStyle` is used to set the default style for all the `CustomStyle` instances. The `CustomStyle` is used to set the style for a specific cell.

| Parameter             | Data Type | Description                                                                                     |
|-----------------------|-----------|-------------------------------------------------------------------------------------------------|
| `font_params`         | `dict`     | The dictionary that contain the key value pair of `Font` style                                 |
| `fill_params`         | `dict`     | The dictionary that contain the key value pair of `Fill` style                                 |
| `ali_params`          | `dict`     | The dictionary that contain the key value pair of `Alignment` style                            |
| `border_params`       | `dict`     | The dictionary that contain the key value pair of `Border` style                               |
| `font_size`           | `int`     | Size of the font.                                                                               |
| `font_name`           | `str`     | Name of the font.                                                                               |
| `font_bold`           | `bool`    | Whether the font is bold.                                                                       |
| `font_italic`         | `bool`    | Whether the font is italic.                                                                     |
| `font_underline`      | `str`     | Underline style of the font. Options include 'none', 'single', 'double', etc.                   |
| `font_strike`         | `bool`    | Whether the font is struck through.                                                             |
| `font_vertAlign`      | `str or None` | Vertical alignment of the font. Options include 'subscript', 'superscript', or `None`.      |
| `font_color`          | `str`     | Color of the font in hex format (e.g., '000000' for black).                                     |
| `fill_pattern`        | `str or int` | Specifies the fill pattern style. String options include 'solid', 'gray75', 'gray50', etc. Integer options range from 0 to 18.                           |
| `fill_type`           | `str`     | Defines the fill type, such as 'pattern' or 'gradient'.                                                      |
| `fill_color`          | `str`     | Fill color in hex format (e.g., 'ffffff' for white).                                            |
| `fill_shading`        | `int`     | Determines the shading intensity for the fill pattern.                                            |
| `ali_horizontal`      | `str or None` | Horizontal alignment. Options include 'left', 'center', 'right', 'justify', or `None`.      |
| `ali_vertical`        | `str`     | Vertical alignment. Options include 'top', 'middle', 'bottom'.                                  |
| `ali_text_rotation`   | `int`     | Degree of text rotation, from 0 to 180.                                                         |
| `ali_wrap_text`       | `bool`    | Whether to wrap the text within the cell.                                                       |
| `ali_shrink_to_fit`   | `bool`    | Whether to shrink text to fit within the cell.                                                  |
| `ali_indent`          | `int`     | Indentation level for the cell.                                                                 |
| `border_style_top`    | `str`     | Style of the top border. Options include 'thin', 'medium', 'thick', etc.                        |
| `border_style_right`  | `str`     | Style of the right border.                                                                      |
| `border_style_left`   | `str`     | Style of the left border.                                                                       |
| `border_style_bottom` | `str`     | Style of the bottom border.                                                                     |
| `border_color_top`    | `str`     | Color of the top border in hex format (e.g., 'C0C0C0' for silver).                              |
| `border_color_right`  | `str`     | Color of the right border.                                                                      |
| `border_color_left`   | `str`     | Color of the left border.                                                                       |
| `border_color_bottom` | `str`     | Color of the bottom border.                                                                     |
| `protect`             | `bool`    | Whether the cell is protected from editing.                                                     |
| `hidden`              | `bool`    | Whether the cell is hidden.                                                                     |
| `number_format`       | `str`     | Format for displaying numbers (e.g., 'General', '0.00', '#,##0', 'mm/dd/yyyy').                 |

```python title="CustomStyle with all arguments"
from pyfastexcel import DefaultStyle, CustomStyle


DefaultStyle.set_default(
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
    fill_type=None,
    fill_shading=None,

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
    hidden=False,

    # number_format
    number_format="General",
)

custom_style = CustomStyle(
    # Font
    font_size=11,
    font_name='Calibri',
    font_bold=False,
    font_italic=False,
    font_underline='none',
    font_strike=False,
    font_vertAlign=None,
    font_color='000000',
    fill_pattern=1,
    fill_color='ffffff',
    fill_type='pattern,
    fill_shading=None,

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
    hidden=False,

    # number_format
    number_format="General",
)
```

### Font

Set the font style for the text in the cell.

| Parameter    | Data Type    | Description                                                                                         |
|--------------|--------------|-----------------------------------------------------------------------------------------------------|
| `bold`       | `Optional[bool]` | Specifies if the text is bold.                                                                    |
| `color`      | `Optional[str]`  | The color of the text.                                                                             |
| `fgColor`    | `Optional[str]`  | Backward compatibility for openpyxl_style_writer; same as 'color'.                                 |
| `family`     | `Optional[str]`  | The font family for the text.                                                                     |
| `name`       | `Optional[str]`  | Equivalent to family; backward compatibility for openpyxl_style_writer.                           |
| `italic`     | `Optional[bool]` | Specifies if the text is italic.                                                                  |
| `size`       | `Optional[float]`| The font size for the text.                                                                       |
| `strike`     | `Optional[bool]` | Specifies if the text has a strikethrough.                                                        |
| `underline`  | `Optional[str]`  | The style of underline for the text.                                                              |
| `vert_align` | `Optional[str]`  | Vertical alignment for the text, such as "baseline", "superscript", or "subscript".               |


```python title="Font Style"
from pyfastexcel import CustomStyle


font_params = {
    'bold': True,
    'color': '0000ff',
    'size': 15,
}
custom_font_style = CustomStyle(font_params=font_params)
```

### Fill

Set the fill style for the cell.

| Parameter | Data Type                               | Description                                                                                                                |
|-----------|-----------------------------------------|----------------------------------------------------------------------------------------------------------------------------|
| `color`   | `Optional[str]`                         | The fill color (only supports hex color value).                                                                            |
| `fgColor` | `Optional[str]`                         | Backward compatibility for openpyxl_style_writer; same as 'color'.                                                          |
| `pattern` | `Optional[str | int]`                   | Specifies the fill pattern style. String options include 'solid', 'gray75', 'gray50', etc. Integer options range from 0 to 18. |
| `type` | `Optional[Literal['pattern', 'gradient']]`  | The type of the fill, it can be pattern or gradient |
| `shading` | `Optional[int]`                         | The shading index for the fill (between 0 and 5). (Not yet supported)                                                      |

```python title="Fill Style"
from pyfastexcel import CustomStyle


fill_params = {
    'color': '0000ff',
    'pattern': 'solid',
}
custom_fill_style = CustomStyle(fill_params=fill_params)
```

### Alignment

Set the alignment style for the cell.

| Parameter         | Data Type          | Description                                      |
|-------------------|--------------------|--------------------------------------------------|
| `horizontal`      | `Optional[str]`    | Horizontal alignment.                            |
| `vertical`        | `Optional[str]`    | Vertical alignment. Defaults to 'bottom'.        |
| `text_rotation`   | `Optional[int]`    | Degree of text rotation, from 0 to 180. Defaults to 0.     |
| `wrap_text`       | `Optional[bool]`   | Whether to wrap the text within the cell. Defaults to False.            |
| `shrink_to_fit`   | `Optional[bool]`   | Whether to shrink text to fit within the cell. Defaults to False.       |
| `indent`          | `Optional[int]`    | Indentation level for the cell. Defaults to 0.                          |
| `reading_order`   | `Optional[int]`    | Reading order for the text.                            |
| `justify_last_line` | `Optional[bool]` | Whether to justify the last line of text.                        |
| `relative_indent`  | `Optional[int]`   | Relative indentation level for the cell.                             |

```python title="Alignment Style"
from pyfastexcel import CustomStyle


ali_params = {
    'horizontal': 'center',
    'vertical': 'center',
    'text_rotation': 45,
    'wrap_text': True,
    'shrink_to_fit': True,
    'indent': 1,
}
custom_ali_style = CustomStyle(ali_params=ali_params)
```

### BorderStyle

The implementation of border style.

| Parameter | Data Type       | Description                                  |
|-----------|-----------------|----------------------------------------------|
| `style`   | `Optional[str]` | Style of the border.                         |
| `color`   | `Optional[str]` | Color of the border.                         |


### Border

Set the border style for the cell.

| Parameter | Data Type            | Description                            |
|-----------|----------------------|----------------------------------------|
| `left`    | `Optional[BorderStyle]` | Left border style. Defaults to 'thin' with color 'C0C0C0'.          |
| `right`   | `Optional[BorderStyle]` | Right border style. Defaults to 'thin' with color 'C0C0C0'.         |
| `top`     | `Optional[BorderStyle]` | Top border style. Defaults to 'thin' with color 'C0C0C0'.           |
| `bottom`  | `Optional[BorderStyle]` | Bottom border style. Defaults to 'thin' with color 'C0C0C0'.        |

```python title="Border Style"
from pyfastexcel import CustomStyle
from pyfastexcel.style import BorderStyle


border_params = {
    'left': BorderStyle(style='medium', color='cccccc'),
    'right': BorderStyle(style='thin', color='cccccc'),
    'top': BorderStyle(style='double', color='cccccc'),
    'bottom': BorderStyle(style='dashed', color='cccccc'),
}
custom_border_style = CustomStyle(border_params=border_params)
```

### Protection

Set the protection style for the cell.

| Parameter | Data Type     | Description                      |
|-----------|---------------|----------------------------------|
| `locked`  | `Optional[bool]` | Whether the cell is locked. Defaults to False. |
| `hidden`  | `Optional[bool]` | Whether the cell is hidden. Defaults to False. |

```python title="Protection Style"
from pyfastexcel import CustomStyle


protect_params = {
    'locked': True,
    'hidden': False,
}
custom_protect_style = CustomStyle(protect_params=protect_params)
```
