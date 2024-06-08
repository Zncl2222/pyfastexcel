from __future__ import annotations

from typing import Any

from openpyxl_style_writer import CustomStyle

from .style import StyleManager
from .utils import (
    _separate_alpha_numeric,
    _validate_excel_index,
    column_to_index,
    deprecated_warning,
    excel_index_to_list_index,
    validate_and_format_value,
    validate_and_register_style,
)


class WorkSheet:
    """
    A class representing a worksheet in a spreadsheet. Remember to call
    _transfer_to_dict before turning the worksheet to JSON.

    Attributes:
        sheet (dict): A dictionary representing the default sheet structure.
        data (list): A list of rows containing cell data.
        header (list): A list containing the header row.
        merge_cells (list): A list of merged cell ranges.
        width (dict): A dictionary mapping column indices to column widths.
        height (dict): A dictionary mapping row indices to row heights.

    Methods:
        _transfer_to_dict():
            Transfers the worksheet data to a dictionary representation.

        _get_default_sheet():
            Returns a dictionary representing the default sheet structure.

        cell(row: int, column: int, value: any, style: str | CustomStyle = 'DEFAULT_STYLE') -> None:
            Sets the value and style of a cell in the worksheet.

        _expand_row_and_cols(target_row: int, target_col: int):
            Expands the rows and columns of the worksheet to accommodate
            the given row and column indices.

        _validate_value_and_set_default(value: Any) -> Tuple[str, Union[str, CustomStyle]]:
            Validates the input value and ensures it is a tuple with the correct
            format.

        set_style(target: str | slice | list[int, int], style: CustomStyle | str) -> None:
            Applies a style to specified cells.

        __getitem__(key: str | slice | int) -> tuple | list[tuple]:
            If index_supported is True, retrieves the cell value at the
            specified index. Raises TypeError if index_supported is False.

        __setitem__(key: str | slice | int, value: Any) -> None:
            If index_supported is True, sets the cell value at the specified
            index. Raises TypeError if index_supported is False.
    """

    def __init__(self):
        """
        Initializes a WorkSheet instance.

        Args:
            index_supported (bool, optional): A flag indicating whether
                index-based access is supported. Defaults to False.
        """
        self.sheet = self._get_default_sheet()
        self.data = [[('', 'DEFAULT_STYLE')]]
        self.header = []
        self.merge_cells = []
        self.width = {}
        self.height = {}
        self.auto_filter_set = set()

    def cell(
        self,
        row: int,
        column: int,
        value: any,
        style: str | CustomStyle = 'DEFAULT_STYLE',
    ) -> None:
        """
        Sets the value and style of a cell in the worksheet.

        Args:
            row (int): The row index of the cell.
            col (int): The column index of the cell.
            value (any): The value to set in the cell.
            style (str | CustomStyle, optional): The style to apply to the cell.
                Defaults to 'DEFAULT_STYLE'.
        """
        if not isinstance(value, tuple):
            value = (f'{value}', style)
        elif not isinstance(value[1], (str, CustomStyle)):
            raise TypeError('Style should be a string or CustomStyle object.')
        if row < 1 or row > 1048576:
            raise ValueError(f'Invalid row index: {row}')
        if column < 1 or column > 16384:
            raise ValueError(f'Invalid column index: {column}')
        try:
            self.data[row][column] = value
        except IndexError:
            self._expand_row_and_cols(row, column)
            self.data[row][column] = value

    def set_style(
        self,
        target: str | slice | list[int, int],
        style: CustomStyle | str,
    ) -> None:
        """
        Applies a specified style to a target range of cells.

        Args:
            target (str | slice | list[int, int]): Target cells to apply style.
            style (CustomStyle | str): Style to apply to the cells.

        Raises:
            TypeError: If target type is invalid.
            ValueError: If style is not registered.
        """
        if isinstance(style, str):
            if StyleManager.REGISTERED_STYLES.get(style) is None:
                raise ValueError(
                    f'Style not found: {style}. Style should be register by '
                    'set_custom_style function when you set a style with '
                    'string.',
                )
        elif isinstance(style, CustomStyle):
            if StyleManager._STYLE_NAME_MAP.get(style) is None:
                validate_and_register_style(style)
            style = StyleManager._STYLE_NAME_MAP[style]

        if isinstance(target, str):
            self._apply_style_to_string_target(target, style)
        elif isinstance(target, slice):
            self._apply_style_to_slice_target(target, style)
        elif isinstance(target, list) and len(target) == 2:
            self._apply_style_to_list_target(target, style)
        else:
            raise TypeError('Target should be a string, slice, or list[row, index].')

    def _apply_style_to_string_target(self, target: str, style: str):
        if ':' not in target:
            row, col = excel_index_to_list_index(target)
            self.data[row][col] = (self.data[row][col][0], style)
        else:
            target_slice = target.split(':')
            target = slice(target_slice[0], target_slice[1])
            self._apply_style_to_slice_target(target, style)

    def _apply_style_to_slice_target(self, target: slice, style: str):
        start_row, start_col, col_stop = self._extract_slice_indices(target)
        for col in range(start_col, col_stop + 1):
            self.data[start_row][col] = (self.data[start_row][col][0], style)

    def _apply_style_to_list_target(self, target: list[int, int], style: str):
        row = target[0]
        col = target[1]
        if not isinstance(row, int) or not isinstance(col, int):
            raise TypeError('Target should be a list of integers.')
        if row < 0 or row > 1048576:
            raise ValueError(f'Invalid row index: {row}')
        if col < 0 or col > 16384:
            raise ValueError(f'Invalid column index: {col}')
        self.data[row][col] = (self.data[row][col][0], style)

    def set_cell_width(self, col: str | int, value: int) -> None:
        if isinstance(col, str):
            col = column_to_index(col)
        if col < 1 or col > 16384:
            raise ValueError(f'Invalid column index: {col}')
        self.width[col] = value

    def set_cell_height(self, row: int, value: int) -> None:
        if row < 1 or row > 1048576:
            raise ValueError(f'Invalid row index: {row}')
        self.height[row] = value

    def set_merge_cell(self, top_left_cell: str, bottom_right_cell: str) -> None:
        deprecated_warning(
            "This function is going to deprecated in v1.0.0. Please use 'ws.merge_cell' instead",
        )
        self.merge_cell(top_left_cell, bottom_right_cell)

    def merge_cell(self, top_left_cell: str, bottom_right_cell: str) -> None:
        """
        Sets a merge cell range in the specified sheet.

        Args:
            top_left_cell (str): The cell location of the top-left corner of the
                merge cell range (e.g., 'A1').
            bottom_right_cell (str): The cell location of the bottom-right corner
                of the merge cell range (e.g., 'C3').

        Raises:
            ValueError: If any of the following conditions are met:
                - Either the top_left_cell or bottom_right_cell has an invalid
                    row number (not between 1 and 1048576).
                - The top_left_cell number is larger than the bottom_right_cell number.
                - The top_left_cell column index is larger than the bottom_right_cell
                    column index.
            IndexError: If sheet does not exist.

        Returns:
            None
        """
        top_alpha, top_number = _separate_alpha_numeric(top_left_cell)
        bottom_alpha, bottom_number = _separate_alpha_numeric(bottom_right_cell)
        top_idx = column_to_index(top_alpha)
        bottom_idx = column_to_index(bottom_alpha)

        if (
            int(top_number) > 1048576
            or int(bottom_number) > 1048576
            or int(top_number) < 1
            or int(bottom_number) < 1
        ):
            raise ValueError('Invalid row number. Row number should be between 1 and 1048576.')

        if int(top_number) > int(bottom_number):
            raise ValueError(
                'Invalid cell range. The top-left cell number should be'
                + 'smaller than or equal to the bottom-right cell number.',
            )

        if top_idx > bottom_idx:
            raise ValueError(
                'Invalid cell range. The top-left cell column should be'
                + 'smaller than or equal to the bottom-right cell column.',
            )

        self.merge_cells.append((top_left_cell, bottom_right_cell))

    def auto_filter(self, target_range: str) -> None:
        """
        Sets the auto filter for the specified range.

        Args:
            target_range (str): The target range to set the auto filter.

        Raises:
            ValueError: If the target range is invalid.

        Returns:
            None
        """
        if ':' not in target_range:
            raise ValueError('Invalid target range. Target range should be in the format "A1:B2".')
        target_list = target_range.split(':')
        _validate_excel_index(target_list[0])
        _validate_excel_index(target_list[1])
        self.auto_filter_set.add(target_range)

    def _expand_row_and_cols(self, target_row: int, target_col: int):
        data_row_len = len(self.data)
        data_col_len = len(self.data[0])
        # Case when the memory space of self.data row is enough
        # but the memory space of the target_col is not enough
        if data_row_len > target_row:
            if data_col_len <= target_col:
                self.data[target_row].extend(
                    [('', 'DEFAULT_STYLE') for _ in range(target_col + 1 - data_col_len)],
                )
        else:
            current_row = max(data_row_len, target_row + 1)
            current_col = max(data_col_len, target_col + 1)
            default_value = ('', 'DEFAULT_STYLE')
            self.data.extend(
                [
                    [default_value for _ in range(current_col)]
                    for _ in range(current_row - data_row_len)
                ],
            )

    def _transfer_to_dict(self) -> None:
        self.sheet = {
            'Header': self.header,
            'Data': self.data,
            'MergeCells': self.merge_cells,
            'Width': self.width,
            'Height': self.height,
            'AutoFilter': self.auto_filter_set,
        }
        return self.sheet

    def _get_default_sheet(self) -> dict[str, dict[str, list]]:
        return {
            'Header': [],
            'Data': [],
            'MergeCells': [],
            'Width': {},
            'Height': {},
            'AutoFilter': set(),
        }

    def _validate_value_and_set_default(self, value: Any):
        """
        Validates the input value and ensures it is a tuple with the correct
        format.

        If the input value is not a tuple, it is converted to a tuple with
        the value as the first element and the string 'DEFAULT_STYLE' as the
        second element.

        If the input value is a tuple, it checks if the second element is a
        string or a CustomStyle object. If not, it raises a TypeError.

        Args:
            value (Any): The value to be validated and formatted.

        Returns:
            Tuple[str, Union[str, CustomStyle]]: A tuple with the value as the
            first element and the style as the second element. The style can be
            either a string or a CustomStyle object.

        Raises:
            TypeError: If the second element of the input tuple is not a string
            or a CustomStyle object.
        """
        if not isinstance(value, tuple):
            value = validate_and_format_value(value)
        else:
            if len(value) != 2:
                raise ValueError(
                    'Cell value should be a tuple with two element like (value, style).',
                )
            if not isinstance(value[1], (str, CustomStyle)):
                raise TypeError(
                    'Style should be a string or CustomStyle object.',
                )
            # The case that user do not register the Custom Style by 'Class attributes'
            # or set_cumston_style function.
            if (
                isinstance(value[1], CustomStyle)
                and StyleManager._STYLE_NAME_MAP.get(value[1]) is None
            ):
                validate_and_register_style(value[1])
                style = StyleManager._STYLE_NAME_MAP[value[1]]
                value = (value[0], style)
        return value

    def __getitem__(self, key: str | slice) -> tuple | list[tuple]:
        if isinstance(key, slice):
            return self._get_cell_by_slice(key)
        elif isinstance(key, int):
            return self.data[key]
        elif isinstance(key, str):
            return self._get_cell_by_location(key)

    def __setitem__(self, key: str | slice | int, value: Any) -> None:
        if isinstance(key, slice):
            self._set_cell_by_slice(key, value)
        elif isinstance(key, int):
            self._set_row_by_index(key, value)
        elif isinstance(key, str):
            self._set_cell_by_location(key, value)
        else:
            raise TypeError('Key should be a string or slice.')

    def _get_cell_by_slice(self, cell_slice: slice) -> list[tuple]:
        _, start_row = _separate_alpha_numeric(cell_slice.start)
        _, stop_row = _separate_alpha_numeric(cell_slice.stop)
        if start_row != stop_row:
            raise ValueError('Only support row-wise slicing.')
        return self.data[int(start_row) - 1]

    def _get_cell_by_location(self, key: str) -> tuple:
        row, col = excel_index_to_list_index(key)
        return self.data[row][col]

    def _extract_slice_indices(self, cell_slice: slice) -> tuple[int, int, int]:
        _, start_row = _separate_alpha_numeric(cell_slice.start)
        _, stop_row = _separate_alpha_numeric(cell_slice.stop)
        if start_row != stop_row:
            raise ValueError('Only support row-wise slicing.')
        start_row, start_col = excel_index_to_list_index(cell_slice.start)
        _, col_stop = excel_index_to_list_index(cell_slice.stop)
        self._expand_row_and_cols(start_row, col_stop)
        return start_row, start_col, col_stop

    def _set_cell_by_slice(self, cell_slice: slice, value: Any) -> None:
        start_row, start_col, col_stop = self._extract_slice_indices(cell_slice)
        for idx, col in enumerate(range(start_col, col_stop + 1)):
            val = self._validate_value_and_set_default(value[idx])
            self.data[start_row][col] = val

    def _set_row_by_index(self, row: int, value: Any) -> None:
        if row < 0 or row > 1048575:
            raise ValueError(f'Invalid row index: {row}')
        if not isinstance(value, list):
            raise ValueError('Value should be a list.')
        value = [self._validate_value_and_set_default(v) for v in value]
        self._expand_row_and_cols(row, len(value) - 1)
        self.data[row] = value

    def _set_cell_by_location(self, key: str, value: Any) -> None:
        row, col = excel_index_to_list_index(key)
        value = self._validate_value_and_set_default(value)
        try:
            self.data[row][col] = value
        except IndexError:
            self._expand_row_and_cols(row, col)
            self.data[row][col] = value
