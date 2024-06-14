from __future__ import annotations

from pyfastexcel.driver import ExcelDriver, WorkSheet
from pyfastexcel.utils import deprecated_warning


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

    def create_sheet(self, sheet_name: str, plain_data: list[list] = None) -> None:
        """
        Creates a new sheet, and set it as current self.sheet.

        Args:
            sheet_name (str): The name of the new sheet.
        """
        if self.workbook.get(sheet_name) is not None:
            raise ValueError(f'Sheet {sheet_name} already exists.')
        self.workbook[sheet_name] = WorkSheet(plain_data=plain_data)
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

    def protect_workbook(
        self,
        algorithm: str,
        password: str,
        lock_structure: bool = False,
        lock_windows: bool = False,
    ):
        if algorithm not in self._PROTECT_ALGORITHM:
            raise ValueError(
                f'Invalid algorithm, the options are {self._PROTECT_ALGORITHM}',
            )
        self.protection['algorithm'] = algorithm
        self.protection['password'] = password
        self.protection['lock_structure'] = lock_structure
        self.protection['lock_windows'] = lock_windows

    def set_cell_width(self, sheet: str, col: str | int, value: int) -> None:
        self._check_if_sheet_exists(sheet)
        self.workbook[sheet].set_cell_width(col, value)

    def set_cell_height(self, sheet: str, row: int, value: int) -> None:
        self._check_if_sheet_exists(sheet)
        self.workbook[sheet].set_cell_height(row, value)

    def set_merge_cell(self, sheet: str, top_left_cell: str, bottom_right_cell: str) -> None:
        deprecated_warning(
            "This function is going to deprecated in v1.0.0. Please use 'wb.merge_cell' instead",
        )
        self.merge_cell(sheet, top_left_cell, bottom_right_cell)

    def merge_cell(self, sheet: str, top_left_cell: str, bottom_right_cell: str) -> None:
        self._check_if_sheet_exists(sheet)
        self.workbook[sheet].set_merge_cell(top_left_cell, bottom_right_cell)

    def auto_filter(self, sheet: str, target_range: str) -> None:
        self._check_if_sheet_exists(sheet)
        self.workbook[sheet].auto_filter(target_range)
