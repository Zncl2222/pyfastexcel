from __future__ import annotations

import base64
import ctypes
import logging
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import overload

import msgspec

from ._typing import Writable
from .logformatter import formatter
from .manager import StyleManager
from .style import CustomStyle
from .validators import TableFinalValidation
from .worksheet import WorkSheet

BASE_DIR = Path(__file__).resolve().parent

logger = logging.getLogger(__name__)
style_formatter = logging.StreamHandler()
style_formatter.setFormatter(formatter)

logger.addHandler(style_formatter)
logger.propagate = False


class ExcelDriver:
    """
    A driver class to write data to Excel files using custom styles.

    ### Attributes:
        _FILE_PROPS (dict[str, str]): Default file properties for the Excel
        file.
        _PROTECT_ALGORITHM (tuple[str]): Algorithm for the workbook protection
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
    DEBUG = False

    def __init__(self, pre_allocate: dict[str, int] = None, plain_data: list[list[str]] = None):
        """
        Initializes the Workbook with default settings and initializes Sheet1.

        It initializes the workbook structure with Sheet1 as the default sheet,
        sets default file properties, initializes current sheet and sheet list,
        and initializes dictionaries for workbook and protection settings.

        Args:
            pre_allocate (dict[str, int], optional): A dictionary containing 'n_rows' and 'n_cols'
                keys specifying the dimensions for pre-allocating data in Sheet1.
            plain_data (list[list[str]], optional): A 2D list of strings representing initial data
                to populate Sheet1.
        """
        self.workbook = {
            'Sheet1': WorkSheet(pre_allocate=pre_allocate, plain_data=plain_data),
        }
        self.file_props = self._get_default_file_props()
        self.sheet = 'Sheet1'
        self._sheet_list = tuple(['Sheet1'])
        self._dict_wb = {}
        self.protection = {}
        self.style = StyleManager()

    @property
    def sheet_list(self):
        return list(self._sheet_list)

    @overload
    def save(self, file: Writable) -> None:
        """
        Saves the workbook to a writable object.

        Args:
            file (Writable): Writable object that has .write() function.
        """
        ...

    @overload
    def save(self, path: str) -> None:
        """
        Saves the workbook to a file.

        Args:
            path (str): A path to save the file.
        """
        ...

    def save(self, file_or_path: Writable | str) -> None:
        if not hasattr(self, 'decoded_bytes'):
            self.read_lib_and_create_excel()

        if isinstance(file_or_path, str):
            with open(file_or_path, 'wb') as file:
                file.write(self.decoded_bytes)
        else:
            file_or_path.write(self.decoded_bytes)

    def __getitem__(self, key: str) -> WorkSheet:
        return self.workbook[key]

    def _check_if_sheet_exists(self, sheet_name: str) -> None:
        if sheet_name not in self.sheet_list:
            raise KeyError(f'{sheet_name} Sheet Does Not Exist.')

    def read_lib_and_create_excel(
        self, lib_path: str = None, ignore_go_panic: bool = True
    ) -> bytes:
        """
        Reads the library and creates the Excel file.

        Args:
            lib_path (str, optional): The path to the library. Defaults to None.
            ignore_go_panic (bool): The flag to determine should trigger panic in go.

        Returns:
            bytes: The byte data of the created Excel file.
        """
        ignore_go_panic = 0 if ignore_go_panic is False else 1
        pyfastexcel = self._read_lib(lib_path)
        self._create_style()

        # Transfer all WorkSheet Object to the sheet dictionary in the workbook.
        for sheet in self._sheet_list:
            self._dict_wb[sheet] = self.workbook[sheet]._transfer_to_dict()
            if len(self.workbook[sheet]._table_list) != 0:
                TableFinalValidation(
                    data=self.workbook[sheet]._data,
                    table_list=self.workbook[sheet]._table_list,
                )

        results = {
            'content': self._dict_wb,
            'file_props': self.file_props,
            'style': self.style._style_map,
            'protection': self.protection,
            'sheet_order': self._sheet_list,
        }
        json_data = msgspec.json.encode(results)
        create_excel = pyfastexcel.Export
        free_pointer = pyfastexcel.FreeCPointer
        free_pointer.argtypes = [ctypes.c_void_p, ctypes.c_int64]
        create_excel.argtypes = [ctypes.c_char_p, ctypes.c_int64]
        create_excel.restype = ctypes.c_void_p
        byte_data = create_excel(json_data, ignore_go_panic)
        self.decoded_bytes = base64.b64decode(ctypes.cast(byte_data, ctypes.c_char_p).value)
        free_pointer(byte_data, 1 if self.DEBUG else 0)
        StyleManager.reset_style_configs()

        return self.decoded_bytes

    def _read_lib(self, lib_path: str) -> ctypes.CDLL:  # pragma: no cover
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
            elif sys.platform.startswith('darwin'):
                lib_path = str(list(BASE_DIR.glob('**/*.dylib'))[0])

        # On macOS, there is no winmode parameter, so we should not pass it
        if sys.platform.startswith('win32') or sys.platform.startswith('linux'):
            lib = ctypes.CDLL(lib_path, winmode=0)
        else:
            lib = ctypes.CDLL(lib_path)
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
