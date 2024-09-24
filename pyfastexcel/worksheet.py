from __future__ import annotations

from typing import Any, Literal, Optional, List, overload

from openpyxl_style_writer import CustomStyle
from pydantic import validate_call as pydantic_validate_call

from .chart import (
    Chart,
    ChartSeries,
    GraphicOptions,
    RichTextRun,
    ChartLegend,
    ChartAxis,
    ChartPlotArea,
    Fill,
    Line,
    ChartDimension,
)
from .pivot import PivotTable, PivotTableField
from .style import StyleManager
from ._typing import CommentTextStructure, SetPanesSelection
from .utils import (
    CommentText,
    Selection,
    _separate_alpha_numeric,
    column_to_index,
    deprecated_warning,
    cell_reference_to_index,
    validate_and_format_value,
    validate_and_register_style,
    transfer_string_slice_to_slice,
)
from .validators import validate_call
from .serializers import CommentSerializer, PanesSerializer, DataValidationSerializer


class WorkSheetBase:
    """
    The base worksheet class for private functions and utilities.
    """

    def __init__(
        self,
        pre_allocate: Optional[dict[str, int]] = None,
        plain_data: Optional[list[list[str]]] = None,
    ):
        """
        Initializes a WorkSheet instance with optional pre-allocation of data or initialization
        with plain data.

        Args:
            pre_allocate (dict[str, int], optional): A dictionary containing 'n_rows' and 'n_cols'
                keys specifying the dimensions for pre-allocating data.
                This can enhancement the performance when you need to write a large excel
            plain_data (list[list[str]], optional): A 2D list of strings representing the
                initial data to populate the worksheet.

        Notes:
            If both `pre_allocate` and `plain_data` are provided, `plain_data` takes precedence.

        Attributes:
            _sheet (dict): Default sheet settings.
            _data (list[list]): The main data structure holding worksheet contents.
            _merged_cells_list (list): List of merged cell coordinates.
            _width_dict (dict): Column widths.
            _height_dict (dict): Row heights.
            _auto_filter_set (set): Set of auto-filter settings.
            _data_validation_list (list): list of dv settings.
            _grouped_columns_list (list): list of settings to group columns.
            _grouped_rows_list (list): list of settings to group rows.
            _engine (str): choice to use excelize normalWriter or openpyxl

        Raises:
            TypeError: If `plain_data` is provided but is not a valid 2D list of strings.

        """
        self._sheet = self._get_default_sheet()
        self._data = []
        self._merged_cells_list = []
        self._width_dict = {}
        self._height_dict = {}
        self._panes_dict = {}
        self._comment_list = []
        self._auto_filter_set = set()
        self._data_validation_list = []
        self._grouped_columns_list = []
        self._grouped_rows_list = []
        self._table_list = []
        self._chart_list = []
        self._pivot_table_list = []
        self._sheet_visible = True
        # Using pyfastexcel to write as default
        self._engine: Literal['pyfastexcel', 'openpyxl'] = 'pyfastexcel'

        if plain_data is not None and pre_allocate is not None:
            raise ValueError(
                "You can only specify either 'pre_allocate' or 'plain_data' at a time, not both.",
            )

        if pre_allocate is not None:
            if (
                not isinstance(pre_allocate, dict)
                or 'n_rows' not in pre_allocate
                or 'n_cols' not in pre_allocate
            ):
                raise TypeError('Invalid pre_allocate dictionary format.')
            if not isinstance(pre_allocate['n_rows'], int) or not isinstance(
                pre_allocate['n_cols'],
                int,
            ):
                raise TypeError('n_rows and n_cols must be integers.')
            self._data = [[None] * pre_allocate['n_cols'] for _ in range(pre_allocate['n_rows'])]

        if plain_data is not None:
            if not isinstance(plain_data, list) or any(
                not isinstance(row, list) for row in plain_data
            ):
                raise TypeError('plain_data should be a valid 2D list of strings.')
            self._data = plain_data
            self._sheet['NoStyle'] = True

    @property
    def data(self):
        return self._data

    @property
    def sheet(self):
        return self._transfer_to_dict()

    @property
    def sheet_visible(self):
        return self._sheet_visible

    @sheet_visible.setter
    def sheet_visible(self, value: bool):
        if not isinstance(value, bool):
            raise ValueError('Sheet visible should be a boolean.')
        self._sheet_visible = value

    def _apply_style_to_string_target(self, target: str, style: str) -> None:
        row, col = cell_reference_to_index(target)
        self._data[row][col] = (self._data[row][col][0], style)

    def _apply_style_to_slice_target(self, target: slice, style: str) -> None:
        start_row, start_col, col_stop = self._extract_slice_indices(target)
        for col in range(start_col, col_stop + 1):
            self._data[start_row][col] = (self._data[start_row][col][0], style)

    def _apply_style_to_list_target(self, target: list[int, int], style: str) -> None:
        row = target[0]
        col = target[1]
        if not isinstance(row, int) or not isinstance(col, int):
            raise TypeError('Target should be a list of integers.')
        if row < 0 or row > 1048576:
            raise ValueError(f'Invalid row index: {row}')
        if col < 0 or col > 16384:
            raise ValueError(f'Invalid column index: {col}')
        self._data[row][col] = (self._data[row][col][0], style)

    def _expand_row_and_cols(self, target_row: int, target_col: int) -> None:
        data_row_len = len(self._data)
        d = ()
        if data_row_len == 0:
            self._data.append([d])
            data_row_len = 1
        data_col_len = len(self._data[0])

        if data_row_len > target_row and len(self._data[target_row]) > target_col:
            return

        # Case when the memory space of self._data row is enough
        # but the memory space of the target_col is not enough
        if data_row_len > target_row:
            if data_col_len <= target_col:
                self._data[target_row].extend(
                    [d for _ in range(target_col + 1 - data_col_len)],
                )
        else:
            current_row = max(data_row_len, target_row + 1)
            current_col = max(data_col_len, target_col + 1)
            self._data.extend(
                [[d for _ in range(current_col)] for _ in range(current_row - data_row_len)],
            )

    def _transfer_to_dict(self) -> dict[str, Any]:
        self._sheet = {
            'Data': self._data,
            'MergeCells': self._merged_cells_list,
            'Width': self._width_dict,
            'Height': self._height_dict,
            'AutoFilter': self._auto_filter_set,
            'Panes': self._panes_dict,
            'DataValidation': self._data_validation_list,
            'NoStyle': self._sheet['NoStyle'],
            'Comment': self._comment_list,
            'GroupedRow': self._grouped_rows_list,
            'GroupedCol': self._grouped_columns_list,
            'Table': self._table_list,
            'Chart': self._chart_list,
            'PivotTable': self._pivot_table_list,
            'SheetVisible': self._sheet_visible,
        }
        return self._sheet

    def _get_default_sheet(self) -> dict[str, dict[str, list]]:
        return {
            'Data': [],
            'MergeCells': [],
            'Width': {},
            'Height': {},
            'AutoFilter': set(),
            'Panes': {},
            'DataValidation': [],
            'NoStyle': False,
            'Comment': [],
            'GroupedRow': [],
            'GroupedCol': [],
            'Table': [],
            'Chart': [],
            'PivotTable': [],
            'SheetVisible': True,
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
            # or set_custom_style function.
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
            return self._data[key]
        elif isinstance(key, str):
            if ':' in key:
                target = transfer_string_slice_to_slice(key)
                return self._get_cell_by_slice(target)
            return self._get_cell_by_location(key)

    def __setitem__(self, key: str | slice | int, value: Any) -> None:
        if isinstance(key, slice):
            self._set_cell_by_slice(key, value)
        elif isinstance(key, int):
            self._set_row_by_index(key, value)
        elif isinstance(key, str):
            if ':' in key:
                target = transfer_string_slice_to_slice(key)
                self._set_cell_by_slice(target, value)
            else:
                self._set_cell_by_location(key, value)
        else:
            raise TypeError('Key should be a string or slice.')

    def _get_cell_by_slice(self, cell_slice: slice) -> list[tuple]:
        _, start_row = _separate_alpha_numeric(cell_slice.start)
        _, stop_row = _separate_alpha_numeric(cell_slice.stop)
        if start_row != stop_row:
            raise ValueError('Only support row-wise slicing.')
        return self._data[int(start_row) - 1]

    def _get_cell_by_location(self, key: str) -> tuple:
        row, col = cell_reference_to_index(key)
        return self._data[row][col]

    def _extract_slice_indices(self, cell_slice: slice) -> tuple[int, int, int]:
        _, start_row = _separate_alpha_numeric(cell_slice.start)
        _, stop_row = _separate_alpha_numeric(cell_slice.stop)
        if start_row != stop_row:
            raise ValueError('Only support row-wise slicing.')
        start_row, start_col = cell_reference_to_index(cell_slice.start)
        _, col_stop = cell_reference_to_index(cell_slice.stop)
        self._expand_row_and_cols(start_row, col_stop)
        return start_row, start_col, col_stop

    def _set_cell_by_slice(self, cell_slice: slice, value: Any) -> None:
        start_row, start_col, col_stop = self._extract_slice_indices(cell_slice)
        for idx, col in enumerate(range(start_col, col_stop + 1)):
            val = self._validate_value_and_set_default(value[idx])
            self._data[start_row][col] = val

    def _set_row_by_index(self, row: int, value: Any) -> None:
        if row < 0 or row > 1048575:
            raise ValueError(f'Invalid row index: {row}')
        if not isinstance(value, list):
            raise ValueError('Value should be a list.')
        value = [self._validate_value_and_set_default(v) for v in value]
        self._expand_row_and_cols(row, len(value) - 1)
        self._data[row] = value

    def _set_cell_by_location(self, key: str, value: Any) -> None:
        row, col = cell_reference_to_index(key)
        value = self._validate_value_and_set_default(value)
        try:
            self._data[row][col] = value
        except IndexError:
            self._expand_row_and_cols(row, col)
            self._data[row][col] = value


class WorkSheet(WorkSheetBase):
    """
    A class representing a worksheet in a spreadsheet. Remember to call
    _transfer_to_dict before turning the worksheet to JSON.

    Attributes:
        sheet (dict): A dictionary representing the default sheet structure.
        data (list): A list of rows containing cell data.
        header (list): A list containing the header row.
        self._merged_cells_list (list): A list of merged cell ranges.
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

    def cell(
        self,
        row: int,
        column: int,
        value: Any,
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
            self._data[row][column] = value
        except IndexError:
            self._expand_row_and_cols(row, column)
            self._data[row][column] = value

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
            if ':' in target:
                target = transfer_string_slice_to_slice(target)
                self._apply_style_to_slice_target(target, style)
            else:
                self._apply_style_to_string_target(target, style)
        elif isinstance(target, slice):
            self._apply_style_to_slice_target(target, style)
        elif isinstance(target, list) and len(target) == 2:
            self._apply_style_to_list_target(target, style)
        else:
            raise TypeError('Target should be a string, slice, or list[row, index].')

    @pydantic_validate_call
    def set_cell_width(self, col: str | int, value: int) -> None:
        if isinstance(col, str):
            col = column_to_index(col)
        if col < 1 or col > 16384:
            raise ValueError(f'Invalid column index: {col}')
        self._width_dict[col] = value

    @pydantic_validate_call
    def set_cell_height(self, row: int, value: int) -> None:
        if row < 1 or row > 1048576:
            raise ValueError(f'Invalid row index: {row}')
        self._height_dict[row] = value

    @overload
    def set_merge_cell(
        self,
        top_lef_cell: Optional[str],
        bottom_right_cell: Optional[str],
    ) -> None:
        '''
        This function is going to deprecated in v1.0.0. Please use 'ws.merge_cell' instead
        Sets a merge cell range in the specified sheet.

        Args:
            top_left_cell (str): The cell location of the top-left corner of the
                merge cell range (e.g., 'A1').
            bottom_right_cell (str): The cell location of the bottom-right corner
                of the merge cell range (e.g., 'C3').
        '''
        ...

    @overload
    def set_merge_cell(self, cell_range: Optional[str]) -> None:
        '''
        "This function is going to deprecated in v1.0.0. Please use 'ws.merge_cell' instead"
        Sets a merge cell range in the specified sheet.

        Args:
            cell_range: The cell range to merge cell.
        '''
        ...

    def set_merge_cell(self, *args) -> None:
        deprecated_warning(
            "ws.set_merge_cell is going to deprecated in v1.0.0. Please use 'ws.merge_cell' instead",
        )
        self.merge_cell(*args)

    @overload
    def merge_cell(self, top_lef_cell: Optional[str], bottom_right_cell: Optional[str]) -> None:
        '''
        Sets a merge cell range in the specified sheet.

        Args:
            top_left_cell (str): The cell location of the top-left corner of the
                merge cell range (e.g., 'A1').
            bottom_right_cell (str): The cell location of the bottom-right corner
                of the merge cell range (e.g., 'C3').
        '''
        ...

    @overload
    def merge_cell(self, cell_range: Optional[str]) -> None:
        '''
        Sets a merge cell range in the specified sheet.

        Args:
            cell_range: The cell range to merge cell.
        '''
        ...

    def merge_cell(self, *args) -> None:
        if len(args) == 1:
            cell_range = args[0]
            top_left_cell, bottom_right_cell = cell_range.split(':')
        elif len(args) == 2:
            top_left_cell, bottom_right_cell = args
        else:
            raise ValueError(
                'Invalid arguments. Use either ws.merge_cell(cell_range) or'
                ' ws.merge_cell(top_left_cell, bottom_right_cell).'
            )

        if top_left_cell == bottom_right_cell:
            raise ValueError('Invalid arguments. Single cell is not a merge cell.')

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
                ' smaller than or equal to the bottom-right cell number.'
            )

        if top_idx > bottom_idx:
            raise ValueError(
                'Invalid cell range. The top-left cell column should be'
                ' smaller than or equal to the bottom-right cell column.'
            )

        self._merged_cells_list.append((top_left_cell, bottom_right_cell))

    @validate_call
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
        self._auto_filter_set.add(target_range)

    @validate_call
    def set_panes(
        self,
        freeze: bool = False,
        split: bool = False,
        x_split: int = 0,
        y_split: int = 0,
        top_left_cell: str = '',
        active_pane: Literal['bottomLeft', 'bottomRight', 'topLeft', 'topRight', ''] = '',
        selection: Optional[SetPanesSelection | list[Selection] | Selection] = None,
    ) -> None:
        """
        Sets the panes for the worksheet with options for freezing, splitting, and selection.

        Args:
            freeze (bool): Whether to freeze the panes.
            split (bool): Whether to split the panes.
            x_split (int): The column position to split or freeze.
            y_split (int): The row position to split or freeze.
            top_left_cell (str): The top-left cell in the split or frozen panes.
            active_pane (Literal['bottomLeft', 'bottomRight', 'topLeft', 'topRight', '']):
            The active pane.
            selection (Optional[SetPanesSelection | list[Selection]]): The selection
            details for the panes.

        Raises:
            ValueError: If x_split or y_split is negative, or if active_pane is
                invalid.

        Returns:
            None
        """
        selection = PanesSerializer.serialize_selection(selection)

        self._panes_dict = {
            'freeze': freeze,
            'split': split,
            'x_split': x_split,
            'y_split': y_split,
            'top_left_cell': top_left_cell,
            'active_pane': active_pane,
            'selection': selection,
        }

    @validate_call
    def set_data_validation(
        self,
        sq_ref: str = '',
        set_range: Optional[list[int | float]] = None,
        input_msg: Optional[list[str]] = None,
        drop_list: Optional[list[str | int | float] | str] = None,
        error_msg: Optional[list[str]] = None,
    ):
        """
        Set data validation for the specified range.

        Args:
            sq_ref (str): The range to set the data validation.
            set_range (list[int | float]): The range of values to set the data validation.
            input (list[str]): The input message for the data validation.
            drop_list (list[str] | str): The drop list for the data validation.
            error (list[str]): The error message for the data validation.

        Raises:
            ValueError: If the range is invalid.

        Returns:
            None
        """
        dv = DataValidationSerializer(
            set_range=set_range,
            input_msg=input_msg,
            drop_list=drop_list,
            error_msg=error_msg,
        ).model_dump()
        dv['sq_ref'] = sq_ref
        self._data_validation_list.append(dv)

    @validate_call
    def add_comment(
        self,
        cell: str,
        author: str,
        text: CommentTextStructure | CommentText | List[CommentText],
    ) -> None:
        """
        Adds a comment to the specified cell.
        Args:
            cell (str): The cell location to add the comment.
            author (str): The author of the comment.
            text (str | dict[str, str] | list[str | dict[str, str]]): The text of the comment.
        Raises:
            ValueError: If the cell location is invalid.
        Returns:
            None
        """
        text = CommentSerializer.serialize_text(text)

        self._comment_list.append({'cell': cell, 'author': author, 'paragraph': text})

    @pydantic_validate_call
    def group_columns(
        self,
        start_col: str,
        end_col: Optional[str] = None,
        outline_level: int = 1,
        hidden: bool = False,
        engine: Literal['pyfastexcel', 'openpyxl'] = 'pyfastexcel',
    ):
        """
        Groups columns between start_col and end_col with specified outline
        level and visibility.

        Args:
            start_col (str): The starting column to group.
            end_col (Optional[str]): The ending column to group. If None, only
            start_col will be grouped.
            outline_level (int): The outline level of the group.
            hidden (bool): Whether the grouped columns should be hidden.
            engine (Literal['pyfastexcel', 'openpyxl']): The engine to use for
            grouping.

        Returns:
            None
        """
        self._grouped_columns_list.append(
            {
                'start_col': start_col,
                'end_col': end_col,
                'outline_level': outline_level,
                'hidden': hidden,
            }
        )
        self._engine = engine

    @pydantic_validate_call
    def group_rows(
        self,
        start_row: int,
        end_row: Optional[int] = None,
        outline_level: int = 1,
        hidden: bool = False,
        engine: Literal['pyfastexcel', 'openpyxl'] = 'pyfastexcel',
    ):
        """
        Groups rows between start_row and end_row with specified outline level
        and visibility.

        Args:
            start_row (int): The starting row to group.
            end_row (Optional[int]): The ending row to group. If None,
            only start_row will be grouped.
            outline_level (int): The outline level of the group.
            hidden (bool): Whether the grouped rows should be hidden.
            engine (Literal['pyfastexcel', 'openpyxl']): The engine to use for
            grouping.

        Returns:
            None
        """
        self._grouped_rows_list.append(
            {
                'start_row': start_row,
                'end_row': end_row,
                'outline_level': outline_level,
                'hidden': hidden,
            }
        )
        self._engine = engine

    @validate_call
    def create_table(
        self,
        cell_range: str,
        name: str,
        style_name: str = '',
        show_first_column: bool = True,
        show_last_column: bool = True,
        show_row_stripes: bool = False,
        show_column_stripes: bool = True,
        validate_table: bool = True,
    ):
        """
        Creates a table within the specified cell range with given style and display options.

        Args:
            cell_range (str): The cell range where the table should be created (e.g., 'A1:D10').
            name (str): The name of the table.
            style_name (str): The style to apply to the table. Defaults to an empty string, which
            applies the default style.
            show_first_column (bool): Whether to emphasize the first column.
            show_last_column (bool): Whether to emphasize the last column.
            show_row_stripes (bool): Whether to show row stripes for alternate row shading.
            show_column_stripes (bool): Whether to show column stripes for alternate column shading.
            validate_table (bool): Whether to validate the table through TableFinalValidation.

        Returns:
            None
        """
        table = {
            'range': cell_range,
            'name': name,
            'style_name': style_name,
            'show_first_column': show_first_column,
            'show_last_column': show_last_column,
            'show_row_stripes': show_row_stripes,
            'show_column_stripes': show_column_stripes,
            # validate_table is a flag to decide whether
            # to validate the table by TableFinalValidation
            'validate_table': validate_table,
        }

        self._table_list.append(table)

    @overload
    def add_chart(self, cell: str, chart_model: Chart | List[Chart]): ...

    @overload
    def add_chart(
        self,
        cell: str,
        chart_type: str,
        series: List[ChartSeries] | ChartSeries,
        graph_format: Optional[GraphicOptions] = None,
        title: Optional[List[RichTextRun]] = None,
        legend: Optional[ChartLegend] = None,
        dimension: Optional[ChartDimension] = None,
        vary_colors: Optional[bool] = None,
        x_axis: Optional[ChartAxis] = None,
        y_axis: Optional[ChartAxis] = None,
        plot_area: Optional[ChartPlotArea] = None,
        fill: Optional[Fill] = None,
        border: Optional[Line] = None,
        show_blanks_as: Optional[str] = None,
        bubble_size: Optional[int] = None,
        hole_size: Optional[int] = None,
        order: Optional[int] = None,
    ): ...

    def add_chart(
        self,
        cell: str,
        chart_model: Optional[List[Chart] | Chart] = None,
        chart_type: Optional[str] = None,
        series: Optional[List[ChartSeries] | ChartSeries] = None,
        graph_format: Optional[GraphicOptions] = None,
        title: Optional[List[RichTextRun]] = None,
        legend: Optional[ChartLegend] = None,
        dimension: Optional[ChartDimension] = None,
        vary_colors: Optional[bool] = None,
        x_axis: Optional[ChartAxis] = None,
        y_axis: Optional[ChartAxis] = None,
        plot_area: Optional[ChartPlotArea] = None,
        fill: Optional[Fill] = None,
        border: Optional[Line] = None,
        show_blanks_as: Optional[str] = None,
        bubble_size: Optional[int] = None,
        hole_size: Optional[int] = None,
        order: Optional[int] = None,
    ):
        if chart_model is not None:
            if isinstance(chart_model, list):
                self._chart_list.append(
                    {
                        'cell': cell,
                        'chart': [chart.model_dump(by_alias=True) for chart in chart_model],
                    }
                )
            else:
                self._chart_list.append(
                    {'cell': cell, 'chart': [chart_model.model_dump(by_alias=True)]}
                )
        elif chart_type is not None and series is not None:
            chart = Chart(
                chart_type=chart_type,
                series=series,
                graph_format=graph_format,
                title=title,
                legend=legend,
                dimension=dimension,
                vary_colors=vary_colors,
                x_axis=x_axis,
                y_axis=y_axis,
                plot_area=plot_area,
                fill=fill,
                border=border,
                show_blanks_as=show_blanks_as,
                bubble_size=bubble_size,
                hole_size=hole_size,
                order=order,
            )
            self._chart_list.append({'cell': cell, 'chart': [chart.model_dump(by_alias=True)]})
        else:
            raise ValueError('Invalid arguments provided to add_chart function')

    @overload
    def add_pivot_table(self, pivot_table: PivotTable | list[PivotTable]) -> None:
        """
        Adds a pivot table to the worksheet.

        Args:
            pivot_table (PivotTable | list[PivotTable]): The pivot table to add to the worksheet.
        """
        ...

    @overload
    def add_pivot_table(
        self,
        data_range: str,
        pivot_table_range: str,
        rows: list[PivotTableField] = None,
        pivot_filter: list[PivotTableField] = None,
        columns: list[PivotTableField] = None,
        data: list[PivotTableField] = None,
        row_grand_totals: Optional[bool] = None,
        column_grand_totals: Optional[bool] = None,
        show_drill: Optional[bool] = None,
        show_row_headers: Optional[bool] = None,
        show_column_headers: Optional[bool] = None,
        show_row_stripes: Optional[bool] = None,
        show_col_stripes: Optional[bool] = None,
        show_last_column: Optional[bool] = None,
        use_auto_formatting: Optional[bool] = None,
        page_over_then_down: Optional[bool] = None,
        merge_item: Optional[bool] = None,
        compact_data: Optional[bool] = None,
        show_error: Optional[bool] = None,
        pivot_table_style_name: Optional[str] = None,
    ) -> None:
        """
        Add Pivot table.

        Args:
            data_range (str): The range of data to be used in the pivot table, e.g., "Sheet1!A1:B2".
            pivot_table_range (str): The range where the pivot table will be positioned, e.g., "Sheet1!C3:D4".
            rows (list[PivotTableField]): List of fields used as rows in the pivot table.
            filter ([PivotTableField]): List of fields used as filters in the pivot table.
            columns (list[PivotTableField]): List of fields used as columns in the pivot table.
            data (list[PivotTableField]): List of fields used as data fields in the pivot table.
            row_grand_totals (Optional[bool Indicates whether to show row grand totals.
            column_grand_totals (Optional[bool]): Indicates whether to show column grand.
            show_drill (Optional[bool]): Indicates whether to show drill indicators.
            show_row_headers (Optional[bool]): Indicates whether to show row headers.
            show_column_headers (Optional[bool]): Indicates whether to show column headers.
            show_row_stripes (Optional[bool]): Indicates whether to show row stripes.
            show_col_stripes (Optional[bool]): Indicates whether to show column stripes.
            show_last_column (Optional[bool]): Indicates whether to show the last column.
            use_auto_formatting (Optional[bool]): Indicates whether to use automatic formatting.
            page_over_then_down (Optional[bool]): Indicates whether pages should be ordered from top to bottom
                then left to right.
            merge_item (Optional[bool]): Indicates whether to merge items.
            compact_data (Optional[bool]): Indicates whether to show in a compact form.
            show_errorOptional[bool]): Indicates whether to show errors.
            pivot_table_style_name (Optional[str]): Specifies the style the pivot table.
        """
        ...

    def add_pivot_table(
        self,
        pivot_table: Optional[PivotTable | list[PivotTable]] = None,
        data_range: Optional[str] = None,
        pivot_table_range: Optional[str] = None,
        rows: Optional[list[PivotTableField]] = None,
        pivot_filter: Optional[list[PivotTableField]] = None,
        columns: Optional[list[PivotTableField]] = None,
        data: Optional[list[PivotTableField]] = None,
        row_grand_totals: Optional[bool] = None,
        column_grand_totals: Optional[bool] = None,
        show_drill: Optional[bool] = None,
        show_row_headers: Optional[bool] = None,
        show_column_headers: Optional[bool] = None,
        show_row_stripes: Optional[bool] = None,
        show_col_stripes: Optional[bool] = None,
        show_last_column: Optional[bool] = None,
        use_auto_formatting: Optional[bool] = None,
        page_over_then_down: Optional[bool] = None,
        merge_item: Optional[bool] = None,
        compact_data: Optional[bool] = None,
        show_error: Optional[bool] = None,
        pivot_table_style_name: Optional[str] = None,
    ) -> None:
        if pivot_table is not None:
            if isinstance(pivot_table, list):
                self._pivot_table_list.extend(
                    [pivot.model_dump(by_alias=True) for pivot in pivot_table],
                )
            else:
                self._pivot_table_list.append(pivot_table.model_dump(by_alias=True))
        elif data_range is not None and pivot_table_range is not None:
            pivot_table = PivotTable(
                data_range=data_range,
                pivot_table_range=pivot_table_range,
                rows=rows,
                pivot_filter=pivot_filter,
                columns=columns,
                data=data,
                row_grand_totals=row_grand_totals,
                column_grand_totals=column_grand_totals,
                show_drill=show_drill,
                show_row_headers=show_row_headers,
                show_column_headers=show_column_headers,
                show_row_stripes=show_row_stripes,
                show_col_stripes=show_col_stripes,
                show_last_column=show_last_column,
                use_auto_formatting=use_auto_formatting,
                page_over_then_down=page_over_then_down,
                merge_item=merge_item,
                compact_data=compact_data,
                show_error=show_error,
                pivot_table_style_name=pivot_table_style_name,
            )
            self._pivot_table_list.append(pivot_table.model_dump(by_alias=True))
