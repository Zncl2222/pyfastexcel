from __future__ import annotations

from typing import Any, Optional

from openpyxl_style_writer import CustomStyle

from .utils import validate_and_format_value, validate_and_register_style
from .workbook import Workbook
from .worksheet import WorkSheet


class StreamWriter(Workbook):
    """
    A class for writing data to Excel files with or without custom styles.

    Attributes:
        _row_list (list[Tuple[str, str | CustomStyle]]): A list of tuples
            representing rows with values and styles.
        data (list[dict[str, str]]): The data to be written to the Excel file.

    Methods:
        __init__(data: list[dict[str, str]]): Initializes the StreamWriter.
        row_append(value: str, style: str | CustomStyle): Appends a value to
            the row list.
        create_row(is_header: bool = False): Creates a row in the Excel data.
    """

    def __init__(self, data: Optional[list[dict[str, str]]] = None):
        """
        Initializes the NormalWriter.

        Args:
            data (list[dict[str, str]]): The data to be written to the
            Excel file.
        """
        super().__init__()
        self._row_list = []
        self.data = data

    @property
    def wb(self) -> StreamWriter:
        return self

    @property
    def ws(self) -> WorkSheet:
        return self.workbook[self.sheet]

    def row_append(self, value: Any, style: str | CustomStyle = 'DEFAULT_STYLE'):
        """
        Appends a value to the row list.

        Args:
            value (Any): The value to be appended.
            style (str | CustomStyle): The style of the value, can be either
                a style name or a CustomStyle object.
        """
        if isinstance(style, CustomStyle):
            if self.style._STYLE_NAME_MAP.get(style) is None:
                validate_and_register_style(style)
            style = self.style._STYLE_NAME_MAP[style]
        value = validate_and_format_value(value, set_default_style=False)
        self._row_list.append((value, style))

    def create_row(self):
        """
        Creates a row in the Excel data, and clean the current _row_list.
        """
        self.workbook[self.sheet].data.append(self._row_list)
        self._row_list = []
