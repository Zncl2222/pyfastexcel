from __future__ import annotations

import copy
import logging

from pydantic import BaseModel, Field

from typing import Optional, Literal, Any, ClassVar
from pathlib import Path

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
        if fill.get('FgColor'):
            fill['Color'] = fill.pop('FgColor')
        # TODO: Implement the pattern, type and color that corresponds to the excelize's
        fill['Type'] = 'pattern'
        fill['Pattern'] = 1
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


class Font(BaseModel):
    """
    Defines font settings for text elements in a chart.

    Attributes:
        bold (Optional[bool]): Specifies if the text is bold.
        color (Optional[str]): The color of the text.
        family (Optional[str]): The font family for the text.
        italic (Optional[bool]): Specifies if the text is italic.
        size (Optional[float]): The font size for the text.
        strike (Optional[bool]): Specifies if the text has a strikethrough.
        underline (Optional[str]): The style of underline for the text.
        vert_align (Optional[str]): Vertical alignment for the text, such as
            "baseline", "superscript" or "subscript".
    """

    bold: Optional[bool] = Field(False, serialization_alias='Bold')

    # fgColor is the backward compatibility for openpyxl_style_writer, it is the same as 'color'
    color: Optional[str] = Field('000000', serialization_alias='Color')
    fgColor: Optional[str] = Field(None, serialization_alias='FgColor')

    # name is equivalent to family, name is the backward compatibility for openpyxl_style_writer
    family: Optional[str] = Field(None, serialization_alias='Family')
    name: Optional[str] = Field(None, serialization_alias='Name')

    italic: Optional[bool] = Field(False, serialization_alias='Italic')
    size: Optional[float] = Field(11, serialization_alias='Size')
    strike: Optional[bool] = Field(False, serialization_alias='Strike')
    underline: Optional[str] = Field('none', serialization_alias='Underline')
    vert_align: Optional[str] = Field(None, serialization_alias='VertAlign')


class Fill(BaseModel):
    """
    Describes the fill settings.

    Attributes:
        ftype (Optional[Literal['pattern', 'gradient']]): The type of fill, either
            'pattern' or 'gradient'.
        pattern (Optional[int]): The pattern index for fill (between 0 and 18).
        color (Optional[str]): The fill color (Only support hex color value).
        shading (Optional[int]): The shading index for the fill (between 0 and 5).
    """

    # fgColor is the backward compatibility for openpyxl_style_writer, it is the same as 'color'
    color: Optional[str] = Field(None, serialization_alias='Color')
    fgColor: Optional[str] = Field(None, serialization_alias='FgColor')

    # pattern is the backward compatibility for openpyxl_style_writer, the real implementation of
    # excelize should use 'ftype(str)' and 'pattern(int)' both to represent the fill pattern
    # this conflict should be resolved in the future
    pattern: Optional[Literal['solid']] = Field('solid', serialization_alias='Pattern')

    # shading is not yet supported
    shading: Optional[int] = Field(None, serialization_alias='Shading', gt=-1, lt=6)
    # TODO: ftype is not yet supported, ftype has conflict to the 'pattern' in openpyxl_style_writer
    # we need to find a way to resolve this conflict
    # ftype: Optional[Literal['pattern', 'gradient', 'solid']] = Field('solid', serialization_alias='Type')


class Alignment(BaseModel):
    """
    Model representing a Alignment style in Excel.
    """

    horizontal: Optional[str] = Field(None, serialization_alias='Horizontal')
    vertical: Optional[str] = Field('bottom', serialization_alias='Vertical')
    text_rotation: Optional[int] = Field(0, serialization_alias='TextRotation')
    wrap_text: Optional[bool] = Field(False, serialization_alias='WrapText')
    shrink_to_fit: Optional[bool] = Field(False, serialization_alias='ShrinkToFit')
    indent: Optional[int] = Field(0, serialization_alias='Indent')
    reading_order: Optional[int] = Field(None, serialization_alias='ReadingOrder')
    justify_last_line: Optional[bool] = Field(None, serialization_alias='JustifyLastLine')
    relative_indent: Optional[int] = Field(None, serialization_alias='RelativeIndent')


class BorderStyle(BaseModel):
    """
    Model representing a border style in Excel.
    """

    style: Optional[str] = Field(None, serialization_alias='Style')
    color: Optional[str] = Field(None, serialization_alias='Color')


default_border_style = BorderStyle(style='thin', color='C0C0C0')


class Border(BaseModel):
    """
    Model representing a border style in Excel.
    """

    left: Optional[BorderStyle] = Field(default_border_style, serialization_alias='left')
    right: Optional[BorderStyle] = Field(default_border_style, serialization_alias='right')
    top: Optional[BorderStyle] = Field(default_border_style, serialization_alias='top')
    bottom: Optional[BorderStyle] = Field(default_border_style, serialization_alias='bottom')


class Protection(BaseModel):
    """
    Model representing a protection style in Excel.
    """

    locked: Optional[bool] = Field(False, serialization_alias='Locked')
    hidden: Optional[bool] = Field(False, serialization_alias='Hidden')


class DefaultStyle:
    # params
    font_params: ClassVar[Optional[dict[str, Any]]] = None
    fill_params: ClassVar[Optional[dict[str, Any]]] = None
    ali_params: ClassVar[Optional[dict[str, Any]]] = None
    border_params: ClassVar[Optional[dict[str, Any]]] = None
    protection_params: ClassVar[Optional[dict[str, Any]]] = None

    # font
    font_size: ClassVar[int] = 11
    font_name: ClassVar[str] = 'Calibri'
    font_bold: ClassVar[bool] = False
    font_italic: ClassVar[bool] = False
    font_underline: ClassVar[str] = 'none'
    font_strike: ClassVar[bool] = False
    font_vertAlign: ClassVar[Optional[str]] = None
    font_color: ClassVar[str] = '000000'

    # fill
    fill_pattern: ClassVar[str] = 'solid'
    fill_color: ClassVar[Optional[str]] = None

    # alignment
    ali_horizontal: ClassVar[Optional[str]] = None
    ali_vertical: ClassVar[Optional[str]] = 'bottom'
    ali_text_rotation: ClassVar[int] = 0
    ali_wrap_text: ClassVar[bool] = False
    ali_shrink_to_fit: ClassVar[bool] = False
    ali_indent: ClassVar[int] = 0

    # border
    border_style_top: ClassVar[str] = 'thin'
    border_style_right: ClassVar[str] = 'thin'
    border_style_left: ClassVar[str] = 'thin'
    border_style_bottom: ClassVar[str] = 'thin'
    border_color_top: ClassVar[str] = 'C0C0C0'
    border_color_right: ClassVar[str] = 'C0C0C0'
    border_color_left: ClassVar[str] = 'C0C0C0'
    border_color_bottom: ClassVar[str] = 'C0C0C0'

    # protect
    protect: ClassVar[bool] = False
    hidden: ClassVar[bool] = False

    # format
    number_format: ClassVar[str] = 'General'

    font: ClassVar[Font] = Font(
        size=font_size,
        name=font_name,
        bold=font_bold,
        color=font_color,
    )

    fill: ClassVar[Fill] = (Fill(color=fill_color, pattern=fill_pattern),)

    ali: ClassVar[Alignment] = Alignment(
        horizontal=ali_horizontal,
        vertical=ali_vertical,
        wrap_text=ali_wrap_text,
    )

    border: ClassVar[Border] = Border(
        top=BorderStyle(
            style=border_style_top,
            color=border_color_top,
        ),
        right=BorderStyle(
            style=border_style_right,
            color=border_color_right,
        ),
        left=BorderStyle(
            style=border_style_left,
            color=border_color_left,
        ),
        bottom=BorderStyle(
            style=border_style_bottom,
            color=border_color_bottom,
        ),
    )

    protection: ClassVar[Protection] = Protection(locked=protect, hidden=hidden)

    @classmethod
    def set_default(cls, **kwargs):
        cls.font_params = kwargs.get('font_params', cls.font_params)
        cls.fill_params = kwargs.get('fill_params', cls.fill_params)
        cls.ali_params = kwargs.get('ali_params', cls.ali_params)
        cls.border_params = kwargs.get('border_params', cls.border_params)
        cls.number_format = kwargs.get('number_format', cls.number_format)

        cls.protect = kwargs.get('protect', cls.protect)
        cls.hidden = kwargs.get('hidden', cls.hidden)

        cls.font_name = kwargs.get('font_name', cls.font_name)
        cls.font_color = kwargs.get('font_color', cls.font_color)
        cls.font_size = kwargs.get('font_size', cls.font_size)
        cls.font_bold = kwargs.get('font_bold', cls.font_bold)

        cls.fill_color = kwargs.get('fill_color', cls.fill_color)

        cls.ali_horizontal = kwargs.get('ali_horizontal', cls.ali_horizontal)
        cls.ali_vertical = kwargs.get('ali_vertical', cls.ali_vertical)
        cls.ali_wrap_text = kwargs.get('ali_wrap_text', cls.ali_wrap_text)

        cls.border_style_top = kwargs.get('border_style_top', cls.border_style_top)
        cls.border_style_right = kwargs.get('border_style_right', cls.border_style_right)
        cls.border_style_left = kwargs.get('border_style_left', cls.border_style_left)
        cls.border_style_bottom = kwargs.get('border_style_bottom', cls.border_style_bottom)
        cls.border_color_top = kwargs.get('border_color_top', cls.border_color_top)
        cls.border_color_right = kwargs.get('border_color_right', cls.border_color_right)
        cls.border_color_left = kwargs.get('border_color_left', cls.border_color_left)
        cls.border_color_bottom = kwargs.get('border_color_bottom', cls.border_color_bottom)

        cls.apply_settings()

    @classmethod
    def apply_settings(cls):
        cls.font = (
            Font(**cls.font_params)
            if cls.font_params
            else Font(
                size=cls.font_size, name=cls.font_name, bold=cls.font_bold, color=cls.font_color
            )
        )
        cls.fill = (
            Fill(**cls.fill_params)
            if cls.fill_params
            else Fill(ftype=cls.fill_pattern, color=cls.fill_color)
        )
        cls.ali = (
            Alignment(**cls.ali_params)
            if cls.ali_params
            else Alignment(
                horizontal=cls.ali_horizontal,
                vertical=cls.ali_vertical,
                wrap_text=cls.ali_wrap_text,
            )
        )
        cls.border = (
            Border(**cls.border_params)
            if cls.border_params
            else Border(
                top=BorderStyle(style=cls.border_style_top, color=cls.border_color_top),
                right=BorderStyle(style=cls.border_style_right, color=cls.border_color_right),
                left=BorderStyle(style=cls.border_style_left, color=cls.border_color_left),
                bottom=BorderStyle(style=cls.border_style_bottom, color=cls.border_color_bottom),
            )
        )
        cls.protection = (
            Protection(**cls.protection)
            if cls.protection_params
            else Protection(
                locked=cls.protect,
                hidden=cls.hidden,
            )
        )


class CustomStyle(DefaultStyle):
    def __init__(self, **kwargs):
        super().__init__()
        self.set_custom_style(**kwargs)

    def set_custom_style(self, **kwargs):
        self.font_params = kwargs.get('font_params', self.font_params)
        self.fill_params = kwargs.get('fill_params', self.fill_params)
        self.ali_params = kwargs.get('ali_params', self.ali_params)
        self.border_params = kwargs.get('border_params', self.border_params)
        self.number_format = kwargs.get('number_format', self.number_format)

        self.protect = kwargs.get('protect', self.protect)
        self.hidden = kwargs.get('hidden', self.hidden)

        self.font_name = kwargs.get('font_name', self.font_name)
        self.font_color = kwargs.get('font_color', self.font_color)
        self.font_size = kwargs.get('font_size', self.font_size)
        self.font_bold = kwargs.get('font_bold', self.font_bold)

        self.fill_color = kwargs.get('fill_color', self.fill_color)

        self.ali_horizontal = kwargs.get('ali_horizontal', self.ali_horizontal)
        self.ali_vertical = kwargs.get('ali_vertical', self.ali_vertical)
        self.ali_wrap_text = kwargs.get('ali_wrap_text', self.ali_wrap_text)

        self.border_style_top = kwargs.get('border_style_top', self.border_style_top)
        self.border_style_right = kwargs.get('border_style_right', self.border_style_right)
        self.border_style_left = kwargs.get('border_style_left', self.border_style_left)
        self.border_style_bottom = kwargs.get('border_style_bottom', self.border_style_bottom)
        self.border_color_top = kwargs.get('border_color_top', self.border_color_top)
        self.border_color_right = kwargs.get('border_color_right', self.border_color_right)
        self.border_color_left = kwargs.get('border_color_left', self.border_color_left)
        self.border_color_bottom = kwargs.get('border_color_bottom', self.border_color_bottom)

        self.apply_settings()

    def apply_settings(self):
        self.font = (
            Font(**self.font_params)
            if self.font_params
            else Font(
                size=self.font_size, name=self.font_name, bold=self.font_bold, color=self.font_color
            )
        )
        self.fill = (
            Fill(**self.fill_params)
            if self.fill_params
            else Fill(pattern=self.fill_pattern, color=self.fill_color)
        )
        self.ali = (
            Alignment(**self.ali_params)
            if self.ali_params
            else Alignment(
                horizontal=self.ali_horizontal,
                vertical=self.ali_vertical,
                wrap_text=self.ali_wrap_text,
            )
        )
        self.border = (
            Border(**self.border_params)
            if self.border_params
            else Border(
                top=BorderStyle(style=self.border_style_top, color=self.border_color_top),
                right=BorderStyle(style=self.border_style_right, color=self.border_color_right),
                left=BorderStyle(style=self.border_style_left, color=self.border_color_left),
                bottom=BorderStyle(style=self.border_style_bottom, color=self.border_color_bottom),
            )
        )
        self.protection = (
            Protection(**self.protection)
            if self.protection_params
            else Protection(
                locked=self.protect,
                hidden=self.hidden,
            )
        )

    def clone_and_modify(self, **kwargs):
        cloned_style = copy.deepcopy(self)
        cloned_style.set_custom_style(**kwargs)
        return cloned_style

    def __repr__(self) -> str:
        return f'CustomStyle(font={self.font}, fill={self.fill}, ali={self.ali}, border={self.border}, '
