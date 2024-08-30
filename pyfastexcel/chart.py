from __future__ import annotations

from pydantic import BaseModel, Field, field_serializer
from typing import List, Literal, Optional

from .enums import ChartDataLabelPosition, ChartLineType, ChartType, MarkerSymbol


class FontModel(BaseModel):
    bold: Optional[bool] = Field(None, serialization_alias='Bold')
    color: Optional[str] = Field(None, serialization_alias='Color')
    family: Optional[str] = Field(None, serialization_alias='Family')
    italic: Optional[bool] = Field(None, serialization_alias='Italic')
    size: Optional[float] = Field(None, serialization_alias='Size')
    strike: Optional[bool] = Field(None, serialization_alias='Strike')
    underline: Optional[str] = Field(None, serialization_alias='Underline')
    vert_align: Optional[str] = Field(None, serialization_alias='VertAlign')


class FillModel(BaseModel):
    ftype: Optional[Literal['pattern', 'gradient']] = Field('pattern', serialization_alias='Type')
    pattern: Optional[int] = Field(None, serialization_alias='Pattern', gt=-1, lt=19)
    color: Optional[str] = Field(None, serialization_alias='Color')
    shading: Optional[int] = Field(None, serialization_alias='Shading', gt=-1, lt=6)


class MarkerModel(BaseModel):
    fill: Optional[FillModel] = Field(None, serialization_alias='Fill')
    symbol: Optional[str | MarkerSymbol] = Field(None, serialization_alias='Symbol')
    size: Optional[int] = Field(None, serialization_alias='Size')

    @field_serializer('symbol')
    @classmethod
    def marker_symbol_serializer(cls, symbol: str | MarkerSymbol | None) -> str | None:
        if symbol is None:
            return None
        if isinstance(symbol, MarkerSymbol):
            return symbol.value
        return MarkerSymbol.get_enum(symbol).value


class LineModel(BaseModel):
    ltype: Optional[str | ChartLineType] = Field(None, serialization_alias='Type')
    smooth: Optional[bool] = Field(None, serialization_alias='Smooth')
    width: Optional[float] = Field(None, serialization_alias='Width')
    show_marker_line: Optional[bool] = Field(None, serialization_alias='ShowMarkerLine')

    @field_serializer('ltype')
    @classmethod
    def line_type_validator(cls, ltype: str | ChartLineType | None) -> int:
        if ltype is None:
            return None
        if isinstance(ltype, ChartLineType):
            return ltype.value
        return ChartLineType.get_enum(ltype).value


class ChartLegendModel(BaseModel):
    position: Optional[Literal['none', 'top', 'bottom', 'left', 'right', 'top_right']] = Field(
        None, serialization_alias='Position'
    )
    show_legend_key: Optional[bool] = Field(None, serialization_alias='ShowLegendKey')


class RichTextRunModel(BaseModel):
    text: str = Field(..., serialization_alias='Text')
    font: Optional[FontModel] = Field(None, serialization_alias='Font')


class ChartCustomNumFmtModel(BaseModel):
    num_fmt: Optional[str] = Field(None, serialization_alias='CustomNumFmt')
    source_linked: Optional[bool] = Field(None, serialization_alias='SourceLinked')


class ChartAxisModel(BaseModel):
    none: Optional[bool] = Field(None, serialization_alias='None')
    font: Optional[FontModel] = Field(None, serialization_alias='Font')
    major_grid_lines: Optional[bool] = Field(None, serialization_alias='MajorGridLines')
    minor_grid_lines: Optional[bool] = Field(None, serialization_alias='MinorGridLines')
    major_unit: Optional[float] = Field(None, serialization_alias='MajorUnit')
    tick_label_skip: Optional[int] = Field(None, serialization_alias='TickLabelSkip')
    reverse_order: Optional[bool] = Field(None, serialization_alias='ReverseOrder')
    secondary: Optional[bool] = Field(None, serialization_alias='Secondary')
    maximum: Optional[float] = Field(None, serialization_alias='Maximum')
    minimum: Optional[float] = Field(None, serialization_alias='Minimum')
    log_base: Optional[float] = Field(None, serialization_alias='LogBase')
    num_fmt: Optional[ChartCustomNumFmtModel] = Field(None, serialization_alias='NumFmt')
    title: Optional[List[RichTextRunModel]] = Field(None, serialization_alias='Title')


class ChartPlotAreaModel(BaseModel):
    second_plot_values: Optional[int] = Field(None, serialization_alias='SecondPlotValues')
    show_bubble_size: Optional[bool] = Field(None, serialization_alias='ShowBubbleSize')
    show_cat_name: Optional[bool] = Field(None, serialization_alias='ShowCatName')
    show_leader_lines: Optional[bool] = Field(None, serialization_alias='ShowLeaderLines')
    show_percent: Optional[bool] = Field(None, serialization_alias='ShowPercent')
    show_ser_name: Optional[bool] = Field(None, serialization_alias='ShowSerName')
    show_val: Optional[bool] = Field(None, serialization_alias='ShowVal')
    fill: Optional[FillModel] = Field(None, serialization_alias='Fill')
    num_fmt: Optional[ChartCustomNumFmtModel] = Field(None, serialization_alias='NumFmt')


class GraphicOptionsModel(BaseModel):
    alt_text: Optional[str] = Field(None, serialization_alias='AltText')
    print_object: Optional[bool] = Field(None, serialization_alias='PrintObject')
    locked: Optional[bool] = Field(None, serialization_alias='Locked')
    lock_aspect_ratio: Optional[bool] = Field(None, serialization_alias='LockAspectRatio')
    auto_fit: Optional[bool] = Field(None, serialization_alias='AutoFit')
    offset_x: Optional[int] = Field(None, serialization_alias='OffsetX')
    offset_y: Optional[int] = Field(None, serialization_alias='OffsetY')
    scale_x: Optional[float] = Field(None, serialization_alias='ScaleX')
    scale_y: Optional[float] = Field(None, serialization_alias='ScaleY')
    hyperlink: Optional[str] = Field(None, serialization_alias='Hyperlink')
    hyperlink_type: Optional[str] = Field(None, serialization_alias='HyperlinkType')
    positioning: Optional[str] = Field(None, serialization_alias='Positioning')


class ChartDimensionModel(BaseModel):
    width: Optional[int] = Field(None, serialization_alias='Width')
    height: Optional[int] = Field(None, serialization_alias='Height')


class ChartSeriesModel(BaseModel):
    name: str = Field(..., serialization_alias='Name')
    categories: str = Field(..., serialization_alias='Categories')
    values: str = Field(..., serialization_alias='Values')
    sizes: Optional[str] = Field(None, serialization_alias='Sizes')
    fill: Optional[FillModel] = Field(None, serialization_alias='Fill')
    line: Optional[LineModel] = Field(None, serialization_alias='Line')
    marker: Optional[MarkerModel] = Field(None, serialization_alias='Marker')
    data_label_position: Optional[str | ChartDataLabelPosition] = Field(
        None, serialization_alias='DataLabelPosition'
    )

    @field_serializer('data_label_position')
    @classmethod
    def data_label_position_validator(cls, label: str | ChartDataLabelPosition | None) -> int:
        if label is None:
            return ChartDataLabelPosition.Unset.value
        if isinstance(label, ChartDataLabelPosition):
            return label.value
        return ChartDataLabelPosition.get_enum(label).value


class ChartModel(BaseModel):
    chart_type: str | ChartType = Field(..., serialization_alias='Type')
    series: List[ChartSeriesModel] | ChartSeriesModel = Field(None, serialization_alias='Series')
    format: Optional[GraphicOptionsModel] = Field(None, serialization_alias='Format')
    title: Optional[List[RichTextRunModel]] = Field(None, serialization_alias='Title')
    legend: Optional[ChartLegendModel] = Field(None, serialization_alias='Legend')
    dimension: Optional[ChartDimensionModel] = Field(None, serialization_alias='Dimension')
    vary_colors: Optional[bool] = Field(None, serialization_alias='VaryColors')
    x_axis: Optional[ChartAxisModel] = Field(None, serialization_alias='XAxis')
    y_axis: Optional[ChartAxisModel] = Field(None, serialization_alias='YAxis')
    plot_area: Optional[ChartPlotAreaModel] = Field(None, serialization_alias='PlotArea')
    fill: Optional[FillModel] = Field(None, serialization_alias='Fill')
    border: Optional[LineModel] = Field(None, serialization_alias='Border')
    show_blanks_as: Optional[str] = Field(None, serialization_alias='ShowBlanksAs')
    bubble_size: Optional[int] = Field(None, serialization_alias='BubbleSize')
    hole_size: Optional[int] = Field(None, serialization_alias='HoleSize')
    order: Optional[int] = Field(None, serialization_alias='order')

    @field_serializer('chart_type')
    @classmethod
    def chart_type_validator(cls, chart_type: str | ChartType) -> int:
        if isinstance(chart_type, ChartType):
            return chart_type.value
        return ChartType.get_enum(chart_type).value
