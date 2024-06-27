from __future__ import annotations

from typing import Literal

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

    def create_sheet(
        self,
        sheet_name: str,
        pre_allocate: dict[str, int] = None,
        plain_data: list[list] = None,
    ) -> None:
        """
        Creates a new sheet, and set it as current self.sheet.

        Args:
            sheet_name (str): The name of the new sheet.
            pre_allocate (dict[str, int], optional): A dictionary containing
                'n_rows' and 'n_cols' keys specifying the dimensions
                for pre-allocating data in new sheet.
            plain_data (list[list[str]], optional): A 2D list of strings
                representing initial data to populate new sheet.
        """
        if self.workbook.get(sheet_name) is not None:
            raise ValueError(f'Sheet {sheet_name} already exists.')
        self.workbook[sheet_name] = WorkSheet(pre_allocate=pre_allocate, plain_data=plain_data)
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

    def set_panes(
        self,
        sheet: str,
        freeze: bool = False,
        split: bool = False,
        x_split: int = 0,
        y_split: int = 0,
        top_left_cell: str = '',
        active_pane: Literal['bottomLeft', 'bottomRight', 'topLeft', 'topRight', ''] = '',
        selection: list[dict[str, str]] = None,
    ) -> None:
        self._check_if_sheet_exists(sheet)
        self.workbook[sheet].set_panes(
            freeze=freeze,
            split=split,
            x_split=x_split,
            y_split=y_split,
            top_left_cell=top_left_cell,
            active_pane=active_pane,
            selection=selection,
        )

    def set_data_validation(
        self,
        sheet: str,
        sq_ref: str = '',
        set_range: list[int | float] = None,
        input_msg: list[str] = None,
        drop_list: list[str] | str = None,
        error_msg: list[str] = None,
    ):
        self._check_if_sheet_exists(sheet)
        self.workbook[sheet].set_data_validation(
            sq_ref=sq_ref,
            set_range=set_range,
            input_msg=input_msg,
            drop_list=drop_list,
            error_msg=error_msg,
        )

    def add_comment(
        self,
        sheet: str,
        cell: str,
        author: str,
        text: str | dict[str, str] | list[str | dict[str, str]],
    ) -> None:
        """
        Adds a comment to the specified cell.
        Args:
            sheet (str): The name of the sheet.
            cell (str): The cell location to add the comment.
            author (str): The author of the comment.
            text (str | dict[str, str] | list[str | dict[str, str]]): The text of the comment.
        Raises:
            ValueError: If the cell location is invalid.
        Returns:
            None
        """
        self._check_if_sheet_exists(sheet)
        self.workbook[sheet].add_comment(cell, author, text)
