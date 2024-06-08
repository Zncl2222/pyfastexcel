from __future__ import annotations

import base64
import ctypes
import sys
from datetime import datetime, timezone
from pathlib import Path

import msgspec
from openpyxl_style_writer import CustomStyle

from .style import StyleManager
from .worksheet import WorkSheet

BASE_DIR = Path(__file__).resolve().parent


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
    """

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
    _PROTECT_ALGORITHM = (
        'XOR',
        'MD4',
        'MD5',
        'SHA-1',
        'SHA-256',
        'SHA-384',
        'SHA-512',
    )

    def __init__(self):
        """
        Initializes the PyExcelizeDriver.

        It initializes the Excel data, file properties, default sheet,
        current sheet, and style mappings.
        """
        self.workbook = {
            'Sheet1': WorkSheet(),
        }
        self.file_props = self._get_default_file_props()
        self.sheet = 'Sheet1'
        self._sheet_list = tuple(['Sheet1'])
        self._dict_wb = {}
        self.protection = {}
        self.style = StyleManager()

    @property
    def sheet_list(self):
        return self._sheet_list

    def save(self, path: str = './pyfastexcel.xlsx') -> None:
        if not hasattr(self, 'decoded_bytes'):
            self.read_lib_and_create_excel()

        with open(path, 'wb') as file:
            file.write(self.decoded_bytes)

    def __getitem__(self, key: str) -> WorkSheet:
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
            self._dict_wb[sheet] = self.workbook[sheet]._transfer_to_dict()

        results = {
            'content': self._dict_wb,
            'file_props': self.file_props,
            'style': self.style._style_map,
            'protection': self.protection,
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
        StyleManager.reset_style_configs()
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

    def _create_style(self) -> None:
        """
        Creates custom styles for the Excel file.

        This method initializes custom styles for the Excel file based on
        predefined attributes.
        """
        style_collections = self._get_style_collections()
        self.style._STYLE_NAME_MAP.update({val: key for key, val in style_collections.items()})

        # Set the CustomStyle from the pre-defined class attributes.
        for key, val in style_collections.items():
            self.style._update_style_map(key, val)

        # Set the CustomStyle from the REGISTERED method.
        for key, val in self.style.REGISTERED_STYLES.items():
            self.style._update_style_map(key, val)
