from __future__ import annotations

from openpyxl_style_writer import CustomStyle

from pyfastexcel.driver import ExcelDriver

from .utils import column_to_index, separate_alpha_numeric

# TODO: Implement a General Writer for all cases to use, and enable
#   the ability to use excel index or number index to set the value


class BaseWriter(ExcelDriver):
    """
    A base class for writing data to Excel files with custom styles.

    This class provides methods to set file properties, cell dimensions,
    merge cells, manipulate sheets, and more.

    Methods:
        remove_sheet(sheet: str) -> None:
            Removes a sheet from the Excel data.
        create_sheet(sheet_name: str) -> None:
            Creates a new sheet.
        switch_sheet(sheet_name: str) -> None:
            Set current self.sheet to a different sheet.
        set_file_props(key: str, value: str) -> None:
            Sets a file property.
        set_cell_width(sheet: str, col: str | int, value: int) -> None:
            Sets the width of a cell.
        set_cell_height(sheet: str, row: int, value: int) -> None:
            Sets the height of a cell.
        set_merge_cell(sheet: str, top_left_cell: str, bottom_right_cell: str) -> None:
            Sets a merge cell range in the specified sheet.
    """

    def remove_sheet(self, sheet: str) -> None:
        """
        Removes a sheet from the Excel data.

        Args:
            sheet (str): The name of the sheet to remove.
        """
        self.workbook.pop(sheet)
        # TODO: Make a function to set the current sheet to the first sheet
        self.sheet = 'Sheet1'

    def create_sheet(self, sheet_name: str) -> None:
        """
        Creates a new sheet, and set it as current self.sheet.

        Args:
            sheet_name (str): The name of the new sheet.
        """
        self.workbook[sheet_name] = self._get_default_sheet()
        self.sheet = sheet_name

    def switch_sheet(self, sheet_name: str) -> None:
        """
        Set current self.sheet to a different sheet. If sheet does not existed
        then raise error.

        Args:
            sheet_name (str): The name of the sheet to switch to.

        Raises:
            IndexError: If sheet does not exist.
        """
        self._check_if_sheet_exists(sheet_name)
        self.sheet = sheet_name

    def set_file_props(self, key: str, value: str) -> None:
        """
        Sets a file property.

        Args:
            key (str): The property key.
            value (str): The property value.

        Raises:
            ValueError: If the key is invalid.
        """
        if key not in self._FILE_PROPS:
            raise ValueError(f'Invalid file property: {key}')
        self.file_props[key] = value

    def set_cell_width(self, sheet: str, col: str | int, value: int) -> None:
        self._check_if_sheet_exists(sheet)
        if isinstance(col, str):
            col = column_to_index(col)
        if col < 1 or col > 16384:
            raise ValueError(f'Invalid column index: {col}')
        self.workbook[sheet]['Width'][col] = value

    def set_cell_height(self, sheet: str, row: int, value: int) -> None:
        self._check_if_sheet_exists(sheet)
        if row < 1 or row > 1048576:
            raise ValueError(f'Invalid row index: {row}')
        self.workbook[sheet]['Height'][row] = value

    def set_merge_cell(self, sheet: str, top_left_cell: str, bottom_right_cell: str) -> None:
        """
        Sets a merge cell range in the specified sheet.

        Args:
            sheet (str): The name of the sheet where the merge cell range will be set.
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
        self._check_if_sheet_exists(sheet)
        top_alpha, top_number = separate_alpha_numeric(top_left_cell)
        bottom_alpha, bottom_number = separate_alpha_numeric(bottom_right_cell)
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

        self.workbook[sheet]['MergeCells'].append((top_left_cell, bottom_right_cell))

    def create_single_header(self) -> None:
        pass

    def create_body(self) -> None:
        pass


class FastWriter(BaseWriter):
    """
    A class for fast writing data to Excel files with custom styles.

    Attributes:
        _row_list (list[list[Union[str, Tuple[str, str]]]]): A list of rows to
        be written to the Excel file.
        data (list[dict[str, str]]): The data to be written to the Excel file.

    Methods:
        __init__(data: list[dict[str, str]]): Initializes the FastWriter.
        row_append(value: str, style: str, row_idx: int, col_idx: int): Appends
            a value to a specific row and column.
        _pop_none_from_row_list(idx: int) -> None: Removes None values from
            the row list.
        apply_to_header(idx: int = 0): Applies the header row to the Excel data.
            create_row(idx): Creates a row in the Excel data.
    """

    def __init__(self, data: list[dict[str, str]]):
        """
        Initializes the FastWriter.

        Args:
            data (list[dict[str, str]]): The data to be written to the
            Excel file.
        """
        super().__init__()
        # The data is list[dict[str, str]] as default, if your data is other dtype
        # You should override the __init___ method to allocate correct space for __row_list
        self._row_list = [[None] * (len(data[0]) + 1) for _ in range(len(data) + 1)]
        self._original_row_list = self._row_list.copy()
        self.data = data
        self.current_row = 0

    def row_append(self, value: str, style: str, col_idx: int):
        """
        Appends a value to a specific row and column.

        Args:
            value (str): The value to be appended.
            style (str): The style of the value.
            row_idx (int): The index of the row.
            col_idx (int): The index of the column.
        """
        if isinstance(style, CustomStyle):
            style = self.style_map_name[style]
        self._row_list[self.current_row][col_idx] = (value, style)

    def create_sheet(self, sheet_name: str) -> None:
        super().create_sheet(sheet_name)
        self.reset_row_list()

    def switch_sheet(self, sheet_name: str) -> None:
        super().switch_sheet(sheet_name)
        self.reset_row_list()

    def reset_row_list(self):
        self._row_list = self._original_row_list.copy()
        self.current_row = 0

    def _pop_none_from_row_list(self, idx: int) -> None:
        """
        Removes None values from the row list.

        Args:
            idx (int): The index of the row.
        """
        for i in range(len(self._row_list[idx]) - 1, 0, -1):
            if self._row_list[idx][i] is None:
                self._row_list[idx].pop()
            else:
                break

    def create_row(self):
        """
        Creates a row in the Excel data.
        """
        self._pop_none_from_row_list(self.current_row)
        self.workbook[self.sheet]['Data'].append(
            self._row_list[self.current_row].copy(),
        )
        self.current_row += 1


class NormalWriter(BaseWriter):
    """
    A class for writing data to Excel files with or without custom styles.

    Attributes:
        _row_list (list[Tuple[str, str | CustomStyle]]): A list of tuples
            representing rows with values and styles.
        data (list[dict[str, str]]): The data to be written to the Excel file.

    Methods:
        __init__(data: list[dict[str, str]]): Initializes the NormalWriter.
        row_append(value: str, style: str | CustomStyle): Appends a value to
            the row list.
        create_row(is_header: bool = False): Creates a row in the Excel data.
    """

    def __init__(self, data: list[dict[str, str]]):
        """
        Initializes the NormalWriter.

        Args:
            data (list[dict[str, str]]): The data to be written to the
            Excel file.
        """
        super().__init__()
        self._row_list = []
        self.data = data

    def row_append(self, value: str, style: str | CustomStyle):
        """
        Appends a value to the row list.

        Args:
            value (str): The value to be appended.
            style (str | CustomStyle): The style of the value, can be either
                a style name or a CustomStyle object.
        """
        if isinstance(style, CustomStyle):
            style = self.style_map_name[style]
        self._row_list.append((value, style))

    def create_row(self, is_header: bool = False):
        """
        Creates a row in the Excel data, and clean the current _row_list.

        Args:
            is_header (bool, optional): Indicates whether the row is a header
                row. Defaults to False.
        """
        key = 'Header' if is_header is True else 'Data'
        self.workbook[self.sheet][key].append(self._row_list)
        self._row_list = []
