from __future__ import annotations

import copy
from typing import Any, Callable, ClassVar, Literal, Optional

from pydantic import BaseModel, Field, model_serializer


class Font(BaseModel):
    """
    Model representing a Font style in Excel.
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

    @model_serializer(mode='wrap')
    def wrap_serializer(self, handler: Callable) -> dict[str, Any]:
        font = handler(self)
        if self.fgColor:
            font['Color'] = self.fgColor
        if self.name:
            font['Family'] = self.name
        return font


class Fill(BaseModel):
    """
    Model representing a Fill style in Excel.
    """

    # fgColor is the backward compatibility for openpyxl_style_writer, it is the same as 'color'
    color: Optional[str] = Field(None, serialization_alias='Color')
    fgColor: Optional[str] = Field(None, serialization_alias='FgColor')

    # pattern (str) is the backward compatibility for openpyxl_style_writer
    # pattern (int) + ftype(Literal['pattern', 'gradient']) is the implementation of excelize
    pattern: Optional[str | int] = Field('solid', serialization_alias='Pattern')
    ftype: Optional[Optional[Literal['pattern', 'gradient']]] = Field(
        'pattern', serialization_alias='Type'
    )

    shading: Optional[int] = Field(None, serialization_alias='Shading', gt=-1, lt=6)

    @model_serializer(mode='wrap')
    def wrap_serializer(self, handler: Callable) -> dict[str, Any]:
        # Prevenet None type input in ChartModel (The default value of Fill)
        if not self:
            return
        fill = handler(self)
        if self.fgColor:
            fill['Color'] = self.fgColor

        # Using Pattern as a string ensures backward compatibility with openpyxl_style_ writer.
        # If Pattern is set as a string, set Type to 'pattern' and Pattern to 1.
        if isinstance(self.pattern, str):
            fill['Type'] = 'pattern'
            fill['Pattern'] = 1

        # Setting Pattern as an integer is the official configuration for Excelize.
        # If Pattern is an integer and Type is not set, set Type to 'pattern' by default.
        elif isinstance(self.pattern, int) and self.ftype is None:
            fill['Type'] = 'pattern'

        return fill


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

    BORDER_NUMBER: ClassVar[dict[Optional[str], int]] = {
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
    left: Optional[BorderStyle] = Field(default_border_style, serialization_alias='left')
    right: Optional[BorderStyle] = Field(default_border_style, serialization_alias='right')
    top: Optional[BorderStyle] = Field(default_border_style, serialization_alias='top')
    bottom: Optional[BorderStyle] = Field(default_border_style, serialization_alias='bottom')

    @model_serializer(mode='wrap')
    def wrap_serializer(self, handler: Callable) -> dict[str, Any]:
        border = handler(self)

        border['left']['Style'] = self.BORDER_NUMBER[self.left.style]
        border['right']['Style'] = self.BORDER_NUMBER[self.right.style]
        border['top']['Style'] = self.BORDER_NUMBER[self.top.style]
        border['bottom']['Style'] = self.BORDER_NUMBER[self.bottom.style]

        return border


class Protection(BaseModel):
    """
    Model representing a protection style in Excel.
    """

    locked: Optional[bool] = Field(False, serialization_alias='Locked')
    hidden: Optional[bool] = Field(False, serialization_alias='Hidden')


class DefaultStyle:
    """
    Module for defining and customizing default and custom styles for formatting purposes.
    """

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
    fill_pattern: ClassVar[str | int] = 'solid'
    fill_type: ClassVar[Optional[str]] = None
    fill_color: ClassVar[Optional[str]] = None
    fill_shading: ClassVar[Optional[int]] = None

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

        cls._apply_default_settings()

    @classmethod
    def _apply_default_settings(cls):
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
            else Fill(
                ftype=cls.fill_type,
                color=cls.fill_color,
                pattern=cls.fill_pattern,
                shading=cls.fill_shading,
            )
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

    def __repr__(self) -> str:
        return (
            f'CustomStyle(font={self.font}, fill={self.fill}, ali={self.ali}, border={self.border}'
        )


class CustomStyle(DefaultStyle):
    def __init__(self, **kwargs):
        """
        Initialize a CustomStyle instance with optional custom styling attributes.

        Args:
            - font_params (dict): Advanced customization for font settings.
            - fill_params (dict): Advanced customization for fill settings.
            - ali_params (dict): Advanced customization for alignment settings.
            - border_params (dict): Advanced customization for border settings.
            - number_format (str): Custom number format.
            - protect (bool): Whether the cells are locked.
            - hidden (bool): Whether the cells are hidden.
            - font_name (str): Font name.
            - font_color (str): Font color in hex format (e.g., 'FF0000').
            - font_size (int): Font size.
            - font_bold (bool): Whether the font is bold.
            - fill_color (str): Fill color in hex format.
            - ali_horizontal (str): Horizontal alignment (e.g., 'center').
            - ali_vertical (str): Vertical alignment (e.g., 'top').
            - ali_wrap_text (bool): Whether text wrapping is enabled.
            - border_style_* (str): Border styles for top, right, left, and bottom.
            - border_color_* (str): Border colors for top, right, left, and bottom.
        """
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
        self.fill_pattern = kwargs.get('fill_pattern', self.fill_pattern)
        self.fill_shading = kwargs.get('fill_shading', self.fill_shading)
        self.fill_type = kwargs.get('fill_type', self.fill_type)

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

        self._apply_settings()

    def _apply_settings(self):
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
            else Fill(
                ftype=self.fill_type,
                color=self.fill_color,
                pattern=self.fill_pattern,
                shading=self.fill_shading,
            )
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
        """
        Create a deep copy of the current CustomStyle instance and modify it with
        the provided attributes.

        Args:
            **kwargs: Keyword arguments for the style customization.
                (Refer to `__init__` for supported parameters.)

        Returns:
            CustomStyle: A new CustomStyle instance with the modified attributes.
        """
        cloned_style = copy.deepcopy(self)
        cloned_style.set_custom_style(**kwargs)
        return cloned_style
