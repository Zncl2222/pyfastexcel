from __future__ import annotations

import re
import string
import warnings
from dataclasses import dataclass
from typing import Any, Literal

from openpyxl_style_writer import CustomStyle

warnings.simplefilter('always', DeprecationWarning)


@dataclass
class CommentText:
    text: str
    size: int | None = None
    name: str | None = None
    bold: bool | None = None
    italic: bool | None = None
    underline: Literal['single', 'double'] | None = None
    strike: bool | None = None
    vertAlign: str | None = None
    color: str | None = None

    def to_dict(self):
        result = {
            k[0].upper() + k[1:]: v
            for k, v in self.__dict__.items()
            if v is not None and k != 'text'
        }
        result['text'] = self.text
        return result


def deprecated_warning(msg: str):
    warnings.warn(
        msg,
        DeprecationWarning,
        stacklevel=2,
    )


def set_custom_style(style_name: str, style: CustomStyle) -> None:
    from .style import StyleManager

    StyleManager.set_custom_style(style_name, style)


def validate_and_register_style(style: CustomStyle) -> None:
    from .style import StyleManager

    if not isinstance(style, CustomStyle):
        raise TypeError(
            f'Invalid type ({type(style)}). Style should be a CustomStyle object.',
        )
    set_custom_style(f'Custom Style {StyleManager._STYLE_ID}', style)
    StyleManager._STYLE_ID += 1


def validate_and_format_value(
    value: Any,
    set_default_style: bool = True,
) -> tuple[Any, Literal['DEFAULT_STYLE']] | Any:
    # Convert non-numeric value to string
    value = f'{value}' if not isinstance(value, (int, float, str)) else value
    # msgpec does not support np.float64, so we should convert
    # it to python float.
    value = float(value) if isinstance(value, float) else value

    return (value, 'DEFAULT_STYLE') if set_default_style else value


def transfer_string_slice_to_slice(string_slice: str) -> slice:
    '''
    Transfer a string slice to a slice object.
    For example, 'A1:B2' will be transfered to slice(1, 2, 1).
    '''
    start, end = string_slice.split(':')
    alpha_start, row_start = _separate_alpha_numeric(start)
    alpha_end, row_end = _separate_alpha_numeric(end)
    return slice(f'{alpha_start}{row_start}', f'{alpha_end}{row_end}')


def _separate_alpha_numeric(input_string: str) -> tuple[str, str]:
    '''
    Separate the alpha and numeric part of a string.
    Return alpha_part at first index and num_part at second index.
    '''
    alpha_part = re.findall(r'[a-zA-Z]+', input_string)
    num_part = re.findall(r'[0-9]+', input_string)
    if len(alpha_part) == 0 or len(num_part) == 0:
        raise ValueError(f'Invalid input string {input_string}.')
    return alpha_part[0], num_part[0]


def _is_valid_column(column: str) -> bool:
    """
    Validate the alphabet part of the column.
    """
    column = column.upper()
    index = 0
    for c in column:
        index = index * 26 + (ord(c) - ord('A')) + 1
    return 1 <= index <= 16384


def column_to_index(column: str) -> int:
    """
    Translate the column name to the column index.
    """
    if not isinstance(column, str):
        raise TypeError(f'Invalid type ({type(column)}). Column should be a string.')
    if len(column) > 3:
        raise ValueError(f"Invalid column ({column}). Maximum Column is 'XFD'.")
    if not all(c in string.ascii_uppercase for c in column):
        raise ValueError(f'Invalid column ({column}). Column should be in uppercase.')
    if not _is_valid_column(column):
        raise ValueError(f"Invalid column ({column}). Maximum Column is 'XFD'.")
    column = column.upper()
    index = 0
    for c in column:
        index = index * 26 + (ord(c) - ord('A')) + 1
    return index


def index_to_column(index: int) -> str:
    """
    Translate the index to the column name.
    """
    if not isinstance(index, int):
        raise TypeError(f'Invalid type ({type(index)}). Index should be a string.')
    if index < 1 or index > 16384:
        raise ValueError(f'Invalid index ({index}). Index should less and equal to 16384.')
    name = ''
    while index > 0:
        index, r = divmod(index - 1, 26)
        name = chr(r + ord('A')) + name
    return name


def cell_reference_to_index(index: str) -> tuple[int, int]:
    """
    Return the row and column index of the given Excel cell reference.
    """
    alpha, num = _separate_alpha_numeric(index)
    column = column_to_index(alpha)
    row = int(num)
    return row - 1, column - 1


def _validate_cell_reference(index: str) -> bool:
    alpha, num = _separate_alpha_numeric(index)
    num = int(num)
    if not all(c in string.ascii_uppercase for c in alpha):
        raise ValueError(f'Invalid column ({alpha}). Column should be in uppercase.')
    if not _is_valid_column(alpha):
        raise ValueError(f"Invalid column ({alpha}). Maximum Column is 'XFD'.")
    if num < 1 or num > 16384:
        raise ValueError(f'Invalid index ({num}). Index should less and equal to 16384.')
    return True
