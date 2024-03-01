from __future__ import annotations

from openpyxl_style_writer import CustomStyle

from pyfastexcel.driver import ExcelDriver


class FastWriter(ExcelDriver):
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
        self._row_list = [[None] * (len(data[0])) for _ in range(len(data))]
        self.data = data

    def row_append(self, value: str, style: str, row_idx: int, col_idx: int):
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
        self._row_list[row_idx][col_idx] = (value, style)

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

    def apply_to_header(self, idx: int = 0):
        """
        Applies the header row to the Excel data.

        Args:
            idx (int, optional): The index of the header row. Defaults to 0.
        """
        original_len = len(self._row_list[idx])
        self._pop_none_from_row_list(idx)
        self.excel_data[self.sheet]['Header'] = self._row_list[idx]
        # Reset row_list for body creation
        self._row_list[idx] = [None] * original_len

    def create_row(self, idx):
        """
        Creates a row in the Excel data.

        Args:
            idx: The index of the row.
        """
        self._pop_none_from_row_list(idx)
        self.excel_data[self.sheet]['Data'].append(self._row_list[idx])


class NormalWriter(ExcelDriver):
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
        self.excel_data[self.sheet][key].append(self._row_list)
        self._row_list = []
