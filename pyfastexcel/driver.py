from __future__ import annotations

import sys

from datetime import datetime
from typing import Any
from openpyxl_style_writer import RowWriter, CustomStyle
from pathlib import Path

import ctypes
import msgspec
import base64

BASE_DIR = Path(__file__).resolve().parent


class PyExcelizeDriver(RowWriter):
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

    def __init__(self):
        self.excel_data = {
            'Sheet1': self._get_default_sheet(),
        }
        self.file_props = self._get_default_file_props()
        self.sheet = 'Sheet1'
        self.style_name_map = {}
        self._create_style()

    def _get_default_sheet(self) -> dict[str, dict[str, list]]:
        return {
            'Header': [],
            'Data': [],
            'MergeCells': [],
            'Width': [],
        }

    def _get_default_file_props(self) -> dict[str, str]:
        now = datetime.utcnow().strftime('%Y-%m-%dT%H:%M:%SZ')
        file_props = self._FILE_PROPS.copy()
        file_props['Created'] = now
        file_props['Modified'] = now
        return file_props

    def set_file_props(self, key: str, value: str) -> None:
        if key not in self._FILE_PROPS:
            raise ValueError(f'Invalid file property: {key}')
        self.file_props[key] = value

    def remove_sheet(self, sheet: str) -> None:
        self.excel_data.pop(sheet)

    def create_sheet(self, sheet_name: str) -> None:
        self.excel_data[sheet_name] = self._get_default_sheet()

    def switch_sheet(self, sheet_name: str) -> None:
        self.sheet = sheet_name
        if self.excel_data.get(sheet_name) is None:
            self.excel_data[sheet_name] = self._get_default_sheet()

    def create_single_header(self) -> None:
        pass

    def create_body(self) -> None:
        pass

    def _read_lib(self, lib_path: str) -> str:
        lib_path = f'{BASE_DIR}/pyfastexcel' if lib_path is None else lib_path
        if lib_path.split('.')[-1] not in ('so', 'dll'):
            if sys.platform.startswith('linux'):
                lib_path += '.so'
            elif sys.platform.startswith('win32'):
                lib_path += '.dll'
        lib = ctypes.CDLL(lib_path, winmode=0)
        return lib

    def _read_lib_and_create_excel(self, lib_path: str = None) -> bytes:
        pyfastexcel = self._read_lib(lib_path)
        results = {
            'content': self.excel_data,
            'file_props': self.file_props,
        }
        json_data = msgspec.json.encode(results)
        create_excel = pyfastexcel.Export
        free_pointer = pyfastexcel.FreeCPointer
        free_pointer.argtypes = [ctypes.c_void_p]
        create_excel.argtypes = [ctypes.c_char_p]
        create_excel.restype = ctypes.c_void_p
        byte_data = create_excel(json_data)
        decoded_bytes = base64.b64decode(ctypes.cast(byte_data, ctypes.c_char_p).value)
        free_pointer(byte_data)
        return decoded_bytes

    def _create_style(self) -> None:
        style_collections = self._get_style_collections()
        self.style_map_name = {val: key for key, val in style_collections.items()}
        self.style_map = {}
        for key, val in style_collections.items():
            self._update_style_map(key, val)
        self.excel_data['Style'] = self.style_map

    def _get_style_collections(self) -> dict[str, CustomStyle]:
        return {
            attr: getattr(self, attr)
            for attr in dir(self)
            if not callable(getattr(self, attr)) and isinstance(getattr(self, attr), CustomStyle)
        }

    def _get_default_style(self) -> dict[str, dict[str, Any] | str]:
        return {
            'Font': {},
            'Fill': {},
            'Border': {},
            'Alignment': {},
            'Protection': {},
            'CustomNumFmt': 'general',
        }

    def _update_style_map(self, style_name: str, custom_style: CustomStyle) -> None:
        self.style_map[style_name] = self._get_default_style()
        self.style_map[style_name]['Font'] = self._get_font_style(custom_style)
        self.style_map[style_name]['Fill'] = self._get_fill_style(custom_style)
        self.style_map[style_name]['Border'] = self._get_border_style(custom_style)
        self.style_map[style_name]['Alignment'] = self._get_alignment_style(custom_style)
        self.style_map[style_name]['Protection'] = self._get_protection_style(custom_style)
        self.style_map[style_name]['CustomNumFmt'] = custom_style.number_format

    def _get_font_style(self, style: CustomStyle) -> dict[str, str | int | bool | None]:
        font_style_map = {}
        if style.font.name:
            font_style_map['Name'] = style.font.name
        if style.font.sz:
            font_style_map['Size'] = style.font.sz
        if style.font.family:
            font_style_map['Family'] = style.font.family
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


class FastWriter(PyExcelizeDriver):
    def __init__(self, data: list[dict[str, str]]):
        super().__init__()
        # The data is list[dict[str, str]] as default, if your data is other dtype
        # You should override the __init___ method to allocate correct space for __row_list
        self._row_list = [[None] * (len(data[0])) for _ in range(len(data))]
        self.data = data

    def row_append(self, value: str, style: str, row_idx: int, col_idx: int):
        if isinstance(style, CustomStyle):
            style = self.style_map_name[style]
        self._row_list[row_idx][col_idx] = (value, style)

    def _pop_none_from_row_list(self, idx: int) -> None:
        for i in range(len(self._row_list[idx]) - 1, 0, -1):
            if self._row_list[idx][i] is None:
                self._row_list[idx].pop()
            else:
                break

    def apply_to_header(self, idx: int = 0):
        original_len = len(self._row_list[idx])
        self._pop_none_from_row_list(idx)
        self.excel_data[self.sheet]['Header'] = self._row_list[idx]
        # Reset row_list for body creation
        self._row_list[idx] = [None] * original_len

    def create_row(self, idx):
        self._pop_none_from_row_list(idx)
        self.excel_data[self.sheet]['Data'].append(self._row_list[idx])


class NormalWriter(PyExcelizeDriver):
    def __init__(self, data: list[dict[str, str]]):
        super().__init__()
        self._row_list = []
        self.data = data

    def row_append(self, value: str, style: str | CustomStyle):
        if isinstance(style, CustomStyle):
            style = self.style_map_name[style]
        self._row_list.append((value, style))

    def create_row(self, is_header: bool = False):
        key = 'Header' if is_header is True else 'Data'
        self.excel_data[self.sheet][key].append(self._row_list)
        self._row_list = []
