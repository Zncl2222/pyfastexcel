from __future__ import annotations

import logging
from pathlib import Path
from typing import Any

from pyfastexcel import CustomStyle

from .logformatter import formatter, log_warning

BASE_DIR = Path(__file__).resolve().parent

logger = logging.getLogger(__name__)
style_formatter = logging.StreamHandler()
style_formatter.setFormatter(formatter)

logger.addHandler(style_formatter)
logger.propagate = False


class StyleManager:
    """
    A class to set custom styles for Excel files.

    ### Attributes:
        BORDER_TO_INDEX (dict[str, int]): Mapping of border styles to excelize's
        corresponding index.

    ### Methods:
        set_custom_style(cls, name: str, custom_style: CustomStyle): Set custom style
        by register method.
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

    BORDER_NUMBER = {
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
        if cls.REGISTERED_STYLES.get(name):
            log_warning(
                logger,
                f'{name} has already existed. Overiding the style settings.',
            )
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
            log_warning(
                logger,
                f'{style_name} has already existed. Overriding the style settings.',
            )
        self._style_map[style_name] = self._get_default_style()
        self._style_map[style_name]['Font'] = self._get_font_style(custom_style)
        self._style_map[style_name]['Fill'] = self._get_fill_style(custom_style)
        self._style_map[style_name]['Border'] = self._get_border_style(custom_style)
        self._style_map[style_name]['Alignment'] = self._get_alignment_style(custom_style)
        self._style_map[style_name]['Protection'] = self._get_protection_style(custom_style)
        self._style_map[style_name]['CustomNumFmt'] = custom_style.number_format

    def _get_font_style(self, style: CustomStyle) -> dict[str, str | int | bool | None]:
        font = style.font.model_dump(by_alias=True)
        if font.get('FgColor'):
            font['Color'] = font.pop('FgColor')
        if font.get('Name'):
            font['Family'] = font.pop('Name')
        return style.font.model_dump(by_alias=True)

    def _get_fill_style(self, style: CustomStyle) -> dict[str, str]:
        fill = style.fill.model_dump(by_alias=True)
        # If using FgColor than override the Color
        if fill.get('FgColor'):
            fill['Color'] = fill.pop('FgColor')

        # Using Pattern as a string ensures backward compatibility with openpyxl_style_ writer.
        # If Pattern is set as a string, set Type to 'pattern' and Pattern to 1.
        if isinstance(fill.get('Pattern'), str):
            fill['Type'] = 'pattern'
            fill['Pattern'] = 1

        # Setting Pattern as an integer is the official configuration for Excelize.
        # If Pattern is an integer and Type is not set, set Type to 'pattern' by default.
        elif isinstance(fill.get('Pattern'), int) and fill.get('Type') is None:
            fill['Type'] = 'pattern'

        return fill

    def _get_border_style(self, style: CustomStyle) -> dict[str, str]:
        border = style.border.model_dump(by_alias=True)
        border['left']['Style'] = self.BORDER_NUMBER[style.border.left.style]
        border['right']['Style'] = self.BORDER_NUMBER[style.border.right.style]
        border['top']['Style'] = self.BORDER_NUMBER[style.border.top.style]
        border['bottom']['Style'] = self.BORDER_NUMBER[style.border.bottom.style]

        return border

    def _get_alignment_style(self, style: CustomStyle) -> dict[str, str]:
        return style.ali.model_dump(by_alias=True)

    def _get_protection_style(self, style: CustomStyle) -> dict[str, str]:
        return style.protection.model_dump(by_alias=True)
