from __future__ import annotations

import base64
import ctypes
import logging
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import msgspec
from openpyxl_style_writer import CustomStyle

from .exceptions import CreateFileNotCalledError
from .utils import excel_index_to_list_index

BASE_DIR = Path(__file__).resolve().parent

# TODO: Implement a CustomStyle without the dependency of openpyxl_style_writer


class ExcelDriver:
    """
    A driver class to write data to Excel files using custom styles.

    ### Attributes:
        BORDER_TO_INDEX (dict[str, int]): Mapping of border styles to excelize's
        corresponding index.
        _FILE_PROPS (dict[str, str]): Default file properties for the Excel
        file.

    ### Methods:
        __init__(): Initializes the ExcelDriver.
        _read_lib(lib_path: str): Reads a library for Excel manipulation.
        read_lib_and_create_excel(lib_path: str = None): Reads the library and
                creates the Excel file.
        set_custom_style(cls, name: str, custom_style: CustomStyle): Set custom style
            by register method.
        _create_style(): Creates custom styles for the Excel file.
        _get_style_collections(): Gets collections of custom styles.
        _get_default_style(): Gets the default style.
        _update_style_map(style_name: str, custom_style: CustomStyle): Updates
            the style map.
        _get_font_style(style: CustomStyle): Gets the font style.
        _get_fill_style(style: CustomStyle): Gets the fill style.
        _get_border_style(style: CustomStyle): Gets the border style.
        _get_alignment_style(style: CustomStyle): Gets the alignment style.
        _get_protection_style(style: CustomStyle): Gets the protection style.
    """

    BORDER_TO_INDEX = {
        None: 0,
        'thick': 5,
        'slantDashDot': 13,
        'dotted': 4,
        'hair': 7,
        'dashed': 3,
        'double': 6,
        'mediumDashDotDot': 12,
        'medium': 2,
        'dashDotDot': 11,
        'thin': 1,
        'dashDot': 9,
        'mediumDashed': 8,
        'mediumDashDot': 10,
    }
    _FILE_PROPS = {
        'Category': '',
        'ContentStatus': '',
        'Created': '',
        'Creator': 'pyfastexcel',
        'Description': '',
        'Identifier': 'xlsx',
        'Keywords': 'spreadsheet',
        'LastModifiedBy': 'pyfastexcel',
        'Modified': '',
        'Revision': '0',
        'Subject': '',
        'Title': '',
        'Language': 'en-Us',
        'Version': '',
    }
    # The style retrieved from set_custom_style will be stored in
    # REGISTERED_STYLES temporarily. It will be created after any
    # Writer is initialized and calls the self._create_style() method.
    DEFAULT_STYLE = CustomStyle()
    REGISTERED_STYLES = {'DEFAULT_STYLE': DEFAULT_STYLE}
    _STYLE_NAME_MAP = {}
    _STYLE_ID = 0
    # The shared memory in the parent class that stores every CustomStyle
    # from different Writer classes.
    _style_map = {}

    def __init__(self):
        """
        Initializes the PyExcelizeDriver.

        It initializes the Excel data, file properties, default sheet,
        current sheet, and style mappings.
        """
        self.workbook = {
            'Sheet1': WorkSheet(index_supported=True) if self.INDEX_SUPPORTED else WorkSheet(),
        }
        self.file_props = self._get_default_file_props()
        self.sheet = 'Sheet1'
        self._sheet_list = tuple(['Sheet1'])

    @property
    def sheet_list(self):
        return self._sheet_list

    @classmethod
    def set_custom_style(cls, name: str, custom_style: CustomStyle):
        cls.REGISTERED_STYLES[name] = custom_style
        cls._STYLE_NAME_MAP[custom_style] = name

    def save(self, path: str = './pyfastexcel.xlsx') -> None:
        if not hasattr(self, 'decoded_bytes'):
            raise CreateFileNotCalledError(
                'Function read_lib_and_create_excel should be ' + 'called before saving the file.',
            )

        with open(path, 'wb') as file:
            file.write(self.decoded_bytes)

    def __getitem__(self, key: str) -> tuple:
        return self.workbook[key]

    def _check_if_sheet_exists(self, sheet_name: str) -> None:
        if sheet_name not in self.sheet_list:
            raise KeyError(f'{sheet_name} Sheet Does Not Exist.')

    def read_lib_and_create_excel(self, lib_path: str = None) -> bytes:
        """
        Reads the library and creates the Excel file.

        Args:
            lib_path (str, optional): The path to the library. Defaults to None.

        Returns:
            bytes: The byte data of the created Excel file.
        """
        pyfastexcel = self._read_lib(lib_path)
        self._create_style()

        # Transfer all WorkSheet Object to the sheet dictionary in the workbook.
        for sheet in self._sheet_list:
            self.workbook[sheet] = self.workbook[sheet]._transfer_to_dict()

        results = {
            'content': self.workbook,
            'file_props': self.file_props,
            'style': self._style_map,
        }
        json_data = msgspec.json.encode(results)
        create_excel = pyfastexcel.Export
        free_pointer = pyfastexcel.FreeCPointer
        free_pointer.argtypes = [ctypes.c_void_p]
        create_excel.argtypes = [ctypes.c_char_p]
        create_excel.restype = ctypes.c_void_p
        byte_data = create_excel(json_data)
        self.decoded_bytes = base64.b64decode(ctypes.cast(byte_data, ctypes.c_char_p).value)
        free_pointer(byte_data)
        return self.decoded_bytes

    def _read_lib(self, lib_path: str) -> ctypes.CDLL:
        """
        Reads a shared-library for writing Excel.

        Args:
            lib_path (str): The path to the library.

        Returns:
            ctypes.CDLL: The library object.
        """
        if lib_path is None:
            if sys.platform.startswith('linux'):
                lib_path = str(list(BASE_DIR.glob('**/*.so'))[0])
            elif sys.platform.startswith('win32'):
                lib_path = str(list(BASE_DIR.glob('**/*.dll'))[0])
        lib = ctypes.CDLL(lib_path, winmode=0)
        return lib

    def _get_default_file_props(self) -> dict[str, str]:
        now = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
        file_props = self._FILE_PROPS.copy()
        file_props['Created'] = now
        file_props['Modified'] = now
        return file_props

    def _create_style(self) -> None:
        """
        Creates custom styles for the Excel file.

        This method initializes custom styles for the Excel file based on
        predefined attributes.
        """
        style_collections = self._get_style_collections()
        self._STYLE_NAME_MAP.update({val: key for key, val in style_collections.items()})

        # Set the CustomStyle from the pre-defined class attributes.
        for key, val in style_collections.items():
            self._update_style_map(key, val)

        # Set the CustomStyle from the REGISTERED method.
        for key, val in self.REGISTERED_STYLES.items():
            self._update_style_map(key, val)

    def _get_style_collections(self) -> dict[str, CustomStyle]:
        """
        Gets collections of custom styles.

        Returns:
            dict[str, CustomStyle]: A dictionary containing custom style
                collections.
        """
        return {
            attr: getattr(self, attr)
            for attr in dir(self)
            if not callable(getattr(self, attr)) and isinstance(getattr(self, attr), CustomStyle)
        }

    def _get_default_style(self) -> dict[str, dict[str, Any] | str]:
        """
        Gets the default style.

        Returns:
            dict[str, dict[str, Any] | str]: A dictionary containing the
                default style settings.
        """
        return {
            'Font': {},
            'Fill': {},
            'Border': {},
            'Alignment': {},
            'Protection': {},
            'CustomNumFmt': 'general',
        }

    def _update_style_map(self, style_name: str, custom_style: CustomStyle) -> None:
        if self._style_map.get(style_name):
            logging.warning(f'{style_name} has already existed. Overiding the style settings.')
        self._style_map[style_name] = self._get_default_style()
        self._style_map[style_name]['Font'] = self._get_font_style(custom_style)
        self._style_map[style_name]['Fill'] = self._get_fill_style(custom_style)
        self._style_map[style_name]['Border'] = self._get_border_style(custom_style)
        self._style_map[style_name]['Alignment'] = self._get_alignment_style(custom_style)
        self._style_map[style_name]['Protection'] = self._get_protection_style(custom_style)
        self._style_map[style_name]['CustomNumFmt'] = custom_style.number_format

    def _get_font_style(self, style: CustomStyle) -> dict[str, str | int | bool | None]:
        font_style_map = {}
        if style.font.name:
            font_style_map['Family'] = style.font.name
        if style.font.sz:
            font_style_map['Size'] = style.font.sz
        if style.font.b:
            font_style_map['Bold'] = style.font.b
        if style.font.i:
            font_style_map['Italic'] = style.font.i
        if style.font.strike:
            font_style_map['Strike'] = style.font.strike
        if style.font.u:
            font_style_map['UnderLine'] = style.font.u
        if style.font.color.rgb:
            font_style_map['Color'] = f'#{style.font.color.rgb[2:]}'
        return font_style_map

    def _get_fill_style(self, style: CustomStyle) -> dict[str, str]:
        fill_style_map = {}
        if style.fill.fgColor.rgb:
            fill_style_map['Color'] = f'#{style.fill.fgColor.rgb[2:]}'
        fill_style_map['Type'] = 'pattern'
        fill_style_map['Pattern'] = 1
        return fill_style_map

    def _get_border_style(self, style: CustomStyle) -> dict[str, str]:
        border_style_map = {}
        direction = ['left', 'right', 'top', 'bottom']

        for d in direction:
            border_style_map[d] = {}

        if style.border.left.style:
            border_style_map['left']['Style'] = self.BORDER_TO_INDEX[style.border.left.style]
        if style.border.right.style:
            border_style_map['right']['Style'] = self.BORDER_TO_INDEX[style.border.right.style]
        if style.border.top.style:
            border_style_map['top']['Style'] = self.BORDER_TO_INDEX[style.border.top.style]
        if style.border.bottom.style:
            border_style_map['bottom']['Style'] = self.BORDER_TO_INDEX[style.border.bottom.style]

        if style.border.left.color.rgb:
            border_style_map['left']['Color'] = f'#{style.border.left.color.rgb[2:]}'
        if style.border.right.color.rgb:
            border_style_map['right']['Color'] = f'#{style.border.right.color.rgb[2:]}'
        if style.border.top.color.rgb:
            border_style_map['top']['Color'] = f'#{style.border.top.color.rgb[2:]}'
        if style.border.bottom.color.rgb:
            border_style_map['bottom']['Color'] = f'#{style.border.bottom.color.rgb[2:]}'
        return border_style_map

    def _get_alignment_style(self, style: CustomStyle) -> dict[str, str]:
        ali_style_map = {}

        if style.ali.horizontal:
            ali_style_map['Horizontal'] = style.ali.horizontal
        if style.ali.vertical:
            ali_style_map['Vertical'] = style.ali.vertical
        if style.ali.wrapText:
            ali_style_map['WrapText'] = style.ali.wrapText
        if style.ali.shrinkToFit:
            ali_style_map['ShrinkToFit'] = style.ali.shrinkToFit
        if style.ali.indent:
            ali_style_map['Indent'] = style.ali.indent
        if style.ali.readingOrder:
            ali_style_map['ReadingOrder'] = style.ali.readingOrder
        if style.ali.textRotation:
            ali_style_map['TextRotation'] = style.ali.textRotation
        if style.ali.justifyLastLine:
            ali_style_map['JustifyLastLine'] = style.ali.justifyLastLine
        if style.ali.relativeIndent:
            ali_style_map['RelativeIndent'] = style.ali.relativeIndent

        return ali_style_map

    def _get_protection_style(self, style: CustomStyle) -> dict[str, str]:
        protection_style_map = {}
        if style.protection.locked:
            protection_style_map['Locked'] = style.protection.locked
        if style.protection.hidden:
            protection_style_map['Hidden'] = style.protection.hidden
        return protection_style_map


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
        index_supported (bool): A flag indicating whether index-based
            access is supported.

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

        __getitem__(key: str | slice) -> tuple | list[tuple]:
            If index_supported is True, retrieves the cell value at the
            specified index. Raises TypeError if index_supported is False.

        __setitem__(key: str | slice, value: Any) -> None:
            If index_supported is True, sets the cell value at the specified
            index. Raises TypeError if index_supported is False.
    """

    def __init__(self, index_supported: bool = False):
        """
        Initializes a WorkSheet instance.

        Args:
            index_supported (bool, optional): A flag indicating whether
                index-based access is supported. Defaults to False.
        """
        self.sheet = self._get_default_sheet()
        self.data = [[]]
        self.header = []
        self.merge_cells = []
        self.width = {}
        self.height = {}
        self.index_supported = index_supported

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

    def _expand_row_and_cols(self, target_row: int, target_col: int):
        current_row = len(self.data)
        current_col = len(self.data[0])
        current_row = max(current_row, target_row + 1)
        current_col = max(current_col, target_col + 1)
        new_data = [[('', 'DEFAULT_STYLE')] * current_col for _ in range(current_row)]
        for i, _ in enumerate(self.data):
            for j, _ in enumerate(self.data[i]):
                new_data[i][j] = self.data[i][j]
        self.data = new_data

    def _transfer_to_dict(self) -> None:
        self.sheet = {
            'Header': self.header,
            'Data': self.data,
            'MergeCells': self.merge_cells,
            'Width': self.width,
            'Height': self.height,
        }
        return self.sheet

    def _get_default_sheet(self) -> dict[str, dict[str, list]]:
        return {
            'Header': [],
            'Data': [],
            'MergeCells': [],
            'Width': {},
            'Height': {},
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
            value = (f'{value}', 'DEFAULT_STYLE')
        elif not isinstance(value[1], (str, CustomStyle)):
            raise TypeError('Style should be a string or CustomStyle object.')
        return value

    def __getitem__(self, key: str | slice) -> tuple | list[tuple]:
        if self.index_supported:
            if isinstance(key, slice):
                return self._get_cell_by_slice(key)
            elif isinstance(key, str):
                return self._get_cell_by_location(key)
        else:
            raise IndexError('Index is not supported in this Writer.')

    def __setitem__(self, key: str | slice | int, value: Any) -> None:
        if self.index_supported:
            if isinstance(key, slice):
                self._set_cell_by_slice(key, value)
            elif isinstance(key, str):
                self._set_cell_by_location(key, value)
            else:
                raise TypeError('Key should be a string or slice.')
        else:
            raise IndexError('Index is not supported in this Writer.')

    def _get_cell_by_slice(self, cell_slice: slice) -> list[tuple]:
        if cell_slice.start[1] == cell_slice.stop[1]:
            return self.data[int(cell_slice.start[1]) - 1]
        row_start, _ = excel_index_to_list_index(cell_slice.start)
        row_stop, _ = excel_index_to_list_index(cell_slice.stop)
        return self.data[row_start:row_stop]

    def _get_cell_by_location(self, key: str) -> tuple:
        row, col = excel_index_to_list_index(key)
        return self.data[row][col]

    def _set_cell_by_slice(self, cell_slice: slice, value: Any) -> None:
        if cell_slice.start[1] != cell_slice.stop[1]:
            raise ValueError('Only support row-wise slicing.')
        start_row, start_col = excel_index_to_list_index(cell_slice.start)
        _, col_stop = excel_index_to_list_index(cell_slice.stop)
        self._expand_row_and_cols(start_row, col_stop)
        for idx, col in enumerate(range(start_col, col_stop + 1)):
            val = self._validate_value_and_set_default(value[idx])
            self.data[start_row][col] = val

    def _set_cell_by_location(self, key: str, value: Any) -> None:
        row, col = excel_index_to_list_index(key)
        value = self._validate_value_and_set_default(value)
        try:
            self.data[row][col] = value
        except IndexError:
            self._expand_row_and_cols(row, col)
            self.data[row][col] = value
