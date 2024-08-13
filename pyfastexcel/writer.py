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
        super().__init__()
        self._row_list = []
        self.data = data
        self._collections = self._get_style_collections()
        self._collections_list = list(self._collections)
        self._cache = {}

    @property
    def wb(self) -> StreamWriter:
        return self

    @property
    def ws(self) -> WorkSheet:
        return self.workbook[self.sheet]

    def _handle_custom_style(self, style_instance: CustomStyle, kwargs) -> None:
        """
        Handle the case when style is a CustomStyle instance.
        """
        if self.style._STYLE_NAME_MAP.get(style_instance) is None:
            validate_and_register_style(style_instance)
        style_name = self.style._STYLE_NAME_MAP[style_instance]

        if not kwargs:
            return style_name

        if self.style_key in self._cache:
            return self._cache[self.style_key]

        new_style = style_instance.clone_and_modify(**kwargs)
        validate_and_register_style(new_style)
        style_name = self.style._STYLE_NAME_MAP[new_style]
        self._cache[self.style_key] = style_name
        return style_name

    def _handle_string_style(self, style: str, kwargs) -> None:
        """
        Handle the case when style is a string.
        """
        if style == 'DEFAULT_STYLE':
            return style

        if style not in (self._collections_list + list(self.style.REGISTERED_STYLES)):
            raise ValueError(f'Style {style} not found !')

        if not kwargs:
            return style

        if self.style_key in self._cache:
            return self._cache[self.style_key]

        self._collections.update(self.style.REGISTERED_STYLES)
        new_style = self._collections[style].clone_and_modify(**kwargs)
        validate_and_register_style(new_style)
        style_name = self.style._STYLE_NAME_MAP[new_style]
        self._cache[self.style_key] = style_name
        return style_name

    def row_append(
        self,
        value: Any,
        style: str | CustomStyle = 'DEFAULT_STYLE',
        **kwargs,
    ) -> None:
        """
        Appends a value to the row list.

        Args:
            value (Any): The value to be appended.
            style (str | CustomStyle): The style of the value, can be either
                a style name or a CustomStyle object.
            **kwargs: Additional keyword arguments to modify the style.
        """
        self.style_key = f'{style}{kwargs}'

        if isinstance(style, CustomStyle):
            style = self._handle_custom_style(style, kwargs)
        elif isinstance(style, str):
            style = self._handle_string_style(style, kwargs)

        value = validate_and_format_value(value, set_default_style=False)
        self._row_list.append((value, style))

    def row_append_list(
        self,
        value: list[Any],
        style: str | CustomStyle = 'DEFAULT_STYLE',
        create_row: bool = False,
        **kwargs,
    ) -> None:
        """
        Appends a value to the row list.

        Args:
            value (list[Any]): The value to be appended.
            style (str | CustomStyle): The style of the value, can be either
                a style name or a CustomStyle object.
            create_row (bool): Whether to create row.
            **kwargs: Additional keyword arguments to modify the style.
        """
        self.style_key = f'{style}{kwargs}'

        if isinstance(style, CustomStyle):
            style = self._handle_custom_style(style, kwargs)
        elif isinstance(style, str):
            style = self._handle_string_style(style, kwargs)

        value = tuple((validate_and_format_value(x, set_default_style=False), style) for x in value)

        if create_row:
            self.workbook[self.sheet].data.append(value)
        else:
            self._row_list.extend(value)

    def create_row(self):
        """
        Creates a row in the Excel data, and clean the current _row_list.
        """
        self.workbook[self.sheet].data.append(self._row_list)
        self._row_list = []
