from __future__ import annotations

import logging
from pathlib import Path
from typing import Any

from openpyxl_style_writer import CustomStyle

BASE_DIR = Path(__file__).resolve().parent

logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

ch = logging.StreamHandler()
ch.setLevel(logging.DEBUG)

formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
ch.setFormatter(formatter)

if not logger.hasHandlers():
    logger.addHandler(ch)

# TODO: Implement a CustomStyle without the dependency of openpyxl_style_writer


class StyleManager:
    """
    A class to set custom styles for Excel files.

    ### Attributes:
        BORDER_TO_INDEX (dict[str, int]): Mapping of border styles to excelize's
        corresponding index.

    ### Methods:
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

    @classmethod
    def set_custom_style(cls, name: str, custom_style: CustomStyle):
        cls.REGISTERED_STYLES[name] = custom_style
        cls._STYLE_NAME_MAP[custom_style] = name

    @classmethod
    def reset_style_configs(cls):
        cls.REGISTERED_STYLES = {'DEFAULT_STYLE': cls.DEFAULT_STYLE}
        cls._STYLE_NAME_MAP = {}
        cls._STYLE_ID = 0
        cls._style_map = {}

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
            logger.warning(f'{style_name} has already existed. Overiding the style settings.')
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
