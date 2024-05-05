from __future__ import annotations

from openpyxl_style_writer import CustomStyle

from pyfastexcel.driver import ExcelDriver, WorkSheet

from .utils import style_validation


class Workbook(ExcelDriver):
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

    INDEX_SUPPORTED = True

    def remove_sheet(self, sheet: str) -> None:
        """
        Removes a sheet from the Excel data.

        Args:
            sheet (str): The name of the sheet to remove.
        """
        if len(self.workbook) == 1:
            raise ValueError('Cannot remove the only sheet in the workbook.')
        if self.workbook.get(sheet) is None:
            raise IndexError(f'Sheet {sheet} does not exist.')
        self.workbook.pop(sheet)
        self._sheet_list = tuple(self.workbook.keys())
        self.sheet = self._sheet_list[0]

    def create_sheet(self, sheet_name: str) -> None:
        """
        Creates a new sheet, and set it as current self.sheet.

        Args:
            sheet_name (str): The name of the new sheet.
        """
        if self.workbook.get(sheet_name) is not None:
            raise ValueError(f'Sheet {sheet_name} already exists.')
        self.workbook[sheet_name] = (
            WorkSheet(index_supported=True) if self.INDEX_SUPPORTED else WorkSheet()
        )
        self.sheet = sheet_name
        self._sheet_list = tuple([x for x in self._sheet_list] + [sheet_name])

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
        self.workbook[sheet].set_cell_width(col, value)

    def set_cell_height(self, sheet: str, row: int, value: int) -> None:
        self._check_if_sheet_exists(sheet)
        self.workbook[sheet].set_cell_height(row, value)

    def set_merge_cell(self, sheet: str, top_left_cell: str, bottom_right_cell: str) -> None:
        self._check_if_sheet_exists(sheet)
        self.workbook[sheet].set_merge_cell(top_left_cell, bottom_right_cell)


class FastWriter(Workbook):
    """
    A class for fast writing data to Excel files with custom styles.

    Attributes:
        INDEX_SUPPORTED (bool): Indicates whether the writer supports using
            index to access worksheet.
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

    INDEX_SUPPORTED = True

    def __init__(self, data: list[dict[str, str]]):
        """
        Initializes the FastWriter.

        Args:
            data (list[dict[str, str]]): The data to be written to the
            Excel file.
        """
        super().__init__()
        # The data is list[dict[str, str]] as default, if your data is other dtype
        # You should override the __init___ method to allocate correct space for _row_list
        self.max_rows = len(data) + 1
        self.max_cols = len(data[0]) + 1
        self._row_list = [[None] * (self.max_cols) for _ in range(self.max_rows)]
        self._original_row_list = self._row_list.copy()
        self.workbook[self.sheet].data = self._original_row_list.copy()
        self.data = data
        # The row and col index for streaming row_append method
        self.current_row = 0
        self.current_col = 0

    def row_append(self, value: str, style: str | CustomStyle = 'DEFAULT_STYLE'):
        """
        Appends a value to a specific row and column.

        Args:
            value (str): The value to be appended.
            style (str): The style of the value.
        """
        if isinstance(style, CustomStyle):
            if self._STYLE_NAME_MAP.get(style) is None:
                style_validation(style)
            style = self._STYLE_NAME_MAP[style]
        self._row_list[self.current_row][self.current_col] = (value, style)
        self.current_col += 1

    def create_sheet(self, sheet_name: str) -> None:
        super().create_sheet(sheet_name)
        self.workbook[self.sheet].data = self._original_row_list.copy()
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
        self.workbook[self.sheet].data.append(
            self._row_list[self.current_row].copy(),
        )

        # Row + 1, Column reset
        self.current_row += 1
        self.current_col = 0


class NormalWriter(Workbook):
    """
    A class for writing data to Excel files with or without custom styles.

    Attributes:
        INDEX_SUPPORTED (bool): Indicates whether the writer supports using
            index to access worksheet.
        _row_list (list[Tuple[str, str | CustomStyle]]): A list of tuples
            representing rows with values and styles.
        data (list[dict[str, str]]): The data to be written to the Excel file.

    Methods:
        __init__(data: list[dict[str, str]]): Initializes the NormalWriter.
        row_append(value: str, style: str | CustomStyle): Appends a value to
            the row list.
        create_row(is_header: bool = False): Creates a row in the Excel data.
    """

    INDEX_SUPPORTED = False

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

    def row_append(self, value: str, style: str | CustomStyle = 'DEFAULT_STYLE'):
        """
        Appends a value to the row list.

        Args:
            value (str): The value to be appended.
            style (str | CustomStyle): The style of the value, can be either
                a style name or a CustomStyle object.
        """
        if isinstance(style, CustomStyle):
            if self._STYLE_NAME_MAP.get(style) is None:
                style_validation(style)
            style = self._STYLE_NAME_MAP[style]
        self._row_list.append((value, style))

    def create_row(self):
        """
        Creates a row in the Excel data, and clean the current _row_list.

        Args:
            is_header (bool, optional): Indicates whether the row is a header
                row. Defaults to False.
        """
        self.workbook[self.sheet].data.append(self._row_list)
        self._row_list = []
