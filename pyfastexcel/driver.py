from __future__ import annotations

import base64
import ctypes
import logging
import os
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, overload

from ._typing import Writable
from .logformatter import formatter
from .manager import StyleManager
from .style import CustomStyle
from .validators import TableFinalValidation
from .wire import encode_payload
from .worksheet import WorkSheet

BASE_DIR = Path(__file__).resolve().parent

# Set once the first native export begins; PYFASTEXCEL_ZIP_LEVEL is read by
# the native library exactly once per process, so later changes are inert.
_NATIVE_EXPORT_STARTED = False


def native_export_started() -> bool:
    """Report whether this process has already run a native export."""
    return _NATIVE_EXPORT_STARTED


def _mark_native_export_started() -> None:
    global _NATIVE_EXPORT_STARTED
    _NATIVE_EXPORT_STARTED = True


logger = logging.getLogger(__name__)
style_formatter = logging.StreamHandler()
style_formatter.setFormatter(formatter)

logger.addHandler(style_formatter)
logger.propagate = False


# D203 conflicts with Ruff's formatter, which removes this blank line.
class NativeExcelClient:  # noqa: D203
    """Versioned ctypes boundary with explicit ownership for C allocations."""

    def __init__(self, library: ctypes.CDLL, *, debug: bool = False) -> None:
        """Bind the supported native export functions from ``library``."""
        self.library = library
        self.debug = debug
        self.free_pointer = library.FreeCPointer
        self._set_signature(
            self.free_pointer,
            [ctypes.c_void_p, ctypes.c_int64],
            None,
        )
        self.abi_version = self._get_abi_version()
        self.export_v2 = getattr(library, 'ExportV2', None) if self.abi_version >= 2 else None
        self.export_to_file_v2 = (
            getattr(library, 'ExportToFileV2', None) if self.abi_version >= 2 else None
        )

    @staticmethod
    def _set_signature(function, argtypes, restype) -> None:
        function.argtypes = argtypes
        function.restype = restype

    def _get_abi_version(self) -> int:
        get_version = getattr(self.library, 'GetABIVersion', None)
        if get_version is None:
            return 1
        self._set_signature(get_version, [], ctypes.c_int64)
        return int(get_version())

    @property
    def supports_v2_export(self) -> bool:
        return self.export_v2 is not None

    @property
    def supports_direct_file_export(self) -> bool:
        return self.export_to_file_v2 is not None

    def _free(self, pointer, *, debug: bool = False) -> None:
        if pointer:
            self.free_pointer(pointer, 1 if debug else 0)

    @staticmethod
    def _error_message(error_pointer: ctypes.c_char_p) -> str | None:
        if not error_pointer.value:
            return None
        return error_pointer.value.decode('utf-8', errors='replace')

    def export_bytes(self, payload: bytes, ignore_go_panic: int) -> bytes:
        _mark_native_export_started()
        if self.export_v2 is None:
            return self._export_legacy(payload, ignore_go_panic)

        self._set_signature(
            self.export_v2,
            [
                ctypes.c_void_p,
                ctypes.c_size_t,
                ctypes.c_int64,
                ctypes.POINTER(ctypes.c_size_t),
                ctypes.POINTER(ctypes.c_char_p),
            ],
            ctypes.c_void_p,
        )
        payload_pointer = ctypes.c_char_p(payload)
        output_length = ctypes.c_size_t()
        error_pointer = ctypes.c_char_p()
        output_pointer = self.export_v2(
            ctypes.cast(payload_pointer, ctypes.c_void_p),
            len(payload),
            ignore_go_panic,
            ctypes.byref(output_length),
            ctypes.byref(error_pointer),
        )
        try:
            error_message = self._error_message(error_pointer)
            if error_message is not None:
                raise RuntimeError(error_message)
            if not output_pointer:
                raise RuntimeError('pyfastexcel native export returned a null pointer.')
            if output_length.value == 0:
                raise RuntimeError('pyfastexcel native export returned an empty workbook.')
            return ctypes.string_at(output_pointer, output_length.value)
        finally:
            self._free(output_pointer, debug=self.debug)
            self._free(ctypes.cast(error_pointer, ctypes.c_void_p))

    def _export_legacy(self, payload: bytes, ignore_go_panic: int) -> bytes:
        create_excel = self.library.Export
        self._set_signature(
            create_excel,
            [ctypes.c_char_p, ctypes.c_int64],
            ctypes.c_void_p,
        )
        output_pointer = create_excel(payload, ignore_go_panic)
        try:
            if not output_pointer:
                raise RuntimeError('pyfastexcel native export returned a null pointer.')
            encoded_output = ctypes.cast(output_pointer, ctypes.c_char_p).value
            if encoded_output is None:
                raise RuntimeError('pyfastexcel native export returned no data.')
            return base64.b64decode(encoded_output)
        finally:
            self._free(output_pointer, debug=self.debug)

    def export_to_file(self, payload: bytes, path: str, ignore_go_panic: int) -> None:
        if self.export_to_file_v2 is None:
            raise RuntimeError('Direct file export is not supported by this native library.')
        _mark_native_export_started()

        self._set_signature(
            self.export_to_file_v2,
            [
                ctypes.c_void_p,
                ctypes.c_size_t,
                ctypes.c_char_p,
                ctypes.c_int64,
                ctypes.POINTER(ctypes.c_char_p),
            ],
            ctypes.c_int64,
        )
        payload_pointer = ctypes.c_char_p(payload)
        error_pointer = ctypes.c_char_p()
        status = self.export_to_file_v2(
            ctypes.cast(payload_pointer, ctypes.c_void_p),
            len(payload),
            os.fsencode(path),
            ignore_go_panic,
            ctypes.byref(error_pointer),
        )
        try:
            error_message = self._error_message(error_pointer)
            if status != 0 or error_message is not None:
                raise RuntimeError(error_message or f'pyfastexcel native export failed ({status}).')
        finally:
            self._free(ctypes.cast(error_pointer, ctypes.c_void_p))


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
        self.style = StyleManager()
        self.workbook = {
            'Sheet1': WorkSheet(
                pre_allocate=pre_allocate,
                plain_data=plain_data,
                style_manager=self.style,
            ),
        }
        self.file_props = self._get_default_file_props()
        self.sheet = 'Sheet1'
        self._sheet_list = tuple(['Sheet1'])
        self._dict_wb = {}
        self.protection = {}

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
        if isinstance(file_or_path, str) and '\x00' in file_or_path:
            raise ValueError('embedded null byte')
        if not hasattr(self, 'decoded_bytes'):
            if isinstance(file_or_path, str) and self._try_direct_file_export(file_or_path):
                return
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
        catch_panic = 0 if ignore_go_panic is False else 1
        native = NativeExcelClient(self._read_lib(lib_path), debug=self.DEBUG)
        export_data = self._build_export_data()
        payload = encode_payload(export_data, force_json=not native.supports_v2_export)
        self.decoded_bytes = native.export_bytes(payload, catch_panic)
        return self.decoded_bytes

    def _try_direct_file_export(self, path: str, ignore_go_panic: bool = True) -> bool:
        native = NativeExcelClient(self._read_lib(None), debug=self.DEBUG)
        if not native.supports_direct_file_export:
            return False

        catch_panic = 0 if ignore_go_panic is False else 1
        export_data = self._build_export_data()
        payload = encode_payload(export_data)
        native.export_to_file(payload, path, catch_panic)
        return True

    def _build_export_data(self) -> dict[str, Any]:
        self._create_style()
        workbook_data: dict[str, Any] = {}

        # Transfer all WorkSheet objects to the workbook dictionary.
        for sheet in self._sheet_list:
            worksheet = self.workbook[sheet]
            workbook_data[sheet] = worksheet._transfer_to_dict()
            if worksheet._table_list:
                TableFinalValidation(
                    data=worksheet._data,
                    table_list=worksheet._table_list,
                )

        self._dict_wb = workbook_data
        return {
            'content': workbook_data,
            'file_props': self.file_props,
            'style': self.style._style_map,
            'protection': self.protection,
            'sheet_order': self._sheet_list,
        }

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
        self.style.begin_style_build()
        style_collections = self._get_style_collections()
        # Class-attribute styles are part of the workbook-local overlay, so
        # they deterministically win a same-name process default.
        for key, val in style_collections.items():
            self.style.register_style(key, val)

        # Serialize one merged, workbook-local snapshot.
        for key, val in self.style.REGISTERED_STYLES.items():
            self.style._update_style_map(key, val)
