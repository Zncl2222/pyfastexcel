from __future__ import annotations

from pydantic import BaseModel, Field, field_serializer
from typing import List, Literal, Optional

from .enums import ChartDataLabelPosition, ChartLineType, ChartType, MarkerSymbol


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

    bold: Optional[bool] = Field(None, serialization_alias='Bold')
    color: Optional[str] = Field(None, serialization_alias='Color')
    family: Optional[str] = Field(None, serialization_alias='Family')
    italic: Optional[bool] = Field(None, serialization_alias='Italic')
    size: Optional[float] = Field(None, serialization_alias='Size')
    strike: Optional[bool] = Field(None, serialization_alias='Strike')
    underline: Optional[str] = Field(None, serialization_alias='Underline')
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

    ftype: Optional[Literal['pattern', 'gradient']] = Field('pattern', serialization_alias='Type')
    pattern: Optional[int] = Field(None, serialization_alias='Pattern', gt=-1, lt=19)
    color: Optional[str] = Field(None, serialization_alias='Color')
    shading: Optional[int] = Field(None, serialization_alias='Shading', gt=-1, lt=6)


class Marker(BaseModel):
    """
    Defines the appearance and style of markers used in charts.

    Attributes:
        fill (Optional[Fill]): Fill settings for the marker.
        symbol (Optional[str | MarkerSymbol]): The symbol used for the marker.
        size (Optional[int]): The size of the marker.
    """

    fill: Optional[Fill] = Field(None, serialization_alias='Fill')
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


class Line(BaseModel):
    """
    Represents line settings for chart elements.

    Attributes:
        ltype (Optional[str | ChartLineType]): The type of line.
        smooth (Optional[bool]): Specifies if the line should be smoothed.
        width (Optional[float]): The width of the line.
        show_marker_line (Optional[bool]): Indicates if the line should be shown on markers.
    """

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


class ChartLegend(BaseModel):
    """
    Defines settings for the chart legend.

    Attributes:
        position (Optional[Literal['none', 'top', 'bottom', 'left', 'right', 'top_right']]):
            The position of the legend.
        show_legend_key (Optional[bool]): Specifies if the legend key should be shown.
    """

    position: Optional[Literal['none', 'top', 'bottom', 'left', 'right', 'top_right']] = Field(
        None, serialization_alias='Position'
    )
    show_legend_key: Optional[bool] = Field(None, serialization_alias='ShowLegendKey')


class RichTextRun(BaseModel):
    """
    Represents a text run with rich text formatting.

    Attributes:
        text (str): The text content.
        font (Optional[Font]): Font settings for the text.
    """

    text: str = Field(..., serialization_alias='Text')
    font: Optional[Font] = Field(None, serialization_alias='Font')


class ChartCustomNumFmt(BaseModel):
    """
    Custom number formatting for chart elements.

    Attributes:
        num_fmt (Optional[str]): The custom number format.
        source_linked (Optional[bool]): Specifies if the format is linked to the source.
    """

    num_fmt: Optional[str] = Field(None, serialization_alias='CustomNumFmt')
    source_linked: Optional[bool] = Field(None, serialization_alias='SourceLinked')


class ChartAxis(BaseModel):
    """
    Defines the settings for chart axes.

    Attributes:
        none (Optional[bool]): Specifies if the axis should be hidden.
        font (Optional[Font]): Font settings for the axis labels.
        major_grid_lines (Optional[bool]): Specifies if major grid lines should be displayed.
        minor_grid_lines (Optional[bool]): Specifies if minor grid lines should be displayed.
        major_unit (Optional[float]): The interval between major grid lines.
        tick_label_skip (Optional[int]): Specifies the number of tick labels to skip between
            each drawn label.
        reverse_order (Optional[bool]): Indicates if the axis order should be reversed.
        secondary (Optional[bool]): Specifies if this is a secondary axis.
        maximum (Optional[float]): The maximum value for the axis.
        minimum (Optional[float]): The minimum value for the axis.
        log_base (Optional[float]): The logarithmic base for the axis scale.
        num_fmt (Optional[ChartCustomNumFmt]): Custom number format for the axis.
        title (Optional[List[RichTextRun]]): The title of the axis.
    """

    none: Optional[bool] = Field(None, serialization_alias='None')
    font: Optional[Font] = Field(None, serialization_alias='Font')
    major_grid_lines: Optional[bool] = Field(None, serialization_alias='MajorGridLines')
    minor_grid_lines: Optional[bool] = Field(None, serialization_alias='MinorGridLines')
    major_unit: Optional[float] = Field(None, serialization_alias='MajorUnit')
    tick_label_skip: Optional[int] = Field(None, serialization_alias='TickLabelSkip')
    reverse_order: Optional[bool] = Field(None, serialization_alias='ReverseOrder')
    secondary: Optional[bool] = Field(None, serialization_alias='Secondary')
    maximum: Optional[float] = Field(None, serialization_alias='Maximum')
    minimum: Optional[float] = Field(None, serialization_alias='Minimum')
    log_base: Optional[float] = Field(None, serialization_alias='LogBase')
    num_fmt: Optional[ChartCustomNumFmt] = Field(None, serialization_alias='NumFmt')
    title: Optional[List[RichTextRun]] = Field(None, serialization_alias='Title')


class ChartPlotArea(BaseModel):
    """
    Represents the plot area of a chart.

    Attributes:
        second_plot_values (Optional[int]): The number of values in a secondary plot
            (Only for pieOfPie and barOfPie chart).
        show_bubble_size (Optional[bool]): Indicates if bubble sizes should be displayed.
        show_cat_name (Optional[bool]): Specifies if category names should be shown.
        show_leader_lines (Optional[bool]): Indicates if leader lines should be shown in
            the data label.
        show_percent (Optional[bool]): Specifies if percentages should be shown in the data label.
        show_ser_name (Optional[bool]): Indicates if series names should be displayed in the
            data label.
        show_val (Optional[bool]): Specifies if values should be shown in the data label.
        fill (Optional[Fill]): Fill settings for the plot area.
        num_fmt (Optional[ChartCustomNumFmt]): Custom number format for the plot area.
    """

    second_plot_values: Optional[int] = Field(None, serialization_alias='SecondPlotValues')
    show_bubble_size: Optional[bool] = Field(None, serialization_alias='ShowBubbleSize')
    show_cat_name: Optional[bool] = Field(None, serialization_alias='ShowCatName')
    show_leader_lines: Optional[bool] = Field(None, serialization_alias='ShowLeaderLines')
    show_percent: Optional[bool] = Field(None, serialization_alias='ShowPercent')
    show_ser_name: Optional[bool] = Field(None, serialization_alias='ShowSerName')
    show_val: Optional[bool] = Field(None, serialization_alias='ShowVal')
    fill: Optional[Fill] = Field(None, serialization_alias='Fill')
    num_fmt: Optional[ChartCustomNumFmt] = Field(None, serialization_alias='NumFmt')


class GraphicOptions(BaseModel):
    """
    Defines various graphical options for chart elements.

    Attributes:
        alt_text (Optional[str]): Alternative text for accessibility.
        print_object (Optional[bool]): Indicates if the object should be printed.
        locked (Optional[bool]): Specifies if the object is locked.
        lock_aspect_ratio (Optional[bool]): Indicates if the aspect ratio should be locked.
        auto_fit (Optional[bool]): Specifies if the object should automatically fit its content.
        offset_x (Optional[int]): The horizontal offset of the object.
        offset_y (Optional[int]): The vertical offset of the object.
        scale_x (Optional[float]): The horizontal scale factor.
        scale_y (Optional[float]): The vertical scale factor.
        hyperlink (Optional[str]): A hyperlink associated with the object.
        hyperlink_type (Optional[str]): The type of hyperlink.
        positioning (Optional[str]): The positioning mode for the object.
    """

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


class ChartDimension(BaseModel):
    """
    Specifies the dimensions of a chart.

    Attributes:
        width (Optional[int]): The width of the chart.
        height (Optional[int]): The height of the chart.
    """

    width: Optional[int] = Field(None, serialization_alias='Width')
    height: Optional[int] = Field(None, serialization_alias='Height')


class ChartSeries(BaseModel):
    """
    Represents a data series in a chart.

    Attributes:
        name (str): The name of the series (Legend).
        categories (str): The categories for the series (X value).
        values (str): The values for the series (Y value).
        sizes (Optional[str]): The sizes for bubble charts.
        fill (Optional[Fill]): Fill settings for the series.
        line (Optional[Line]): Line settings for the series.
        marker (Optional[Marker]): Marker settings for the series.
        data_label_position (Optional[str | ChartDataLabelPosition]): The position
            of data labels for the series.
    """

    name: str = Field(..., serialization_alias='Name')
    categories: str = Field(..., serialization_alias='Categories')
    values: str = Field(..., serialization_alias='Values')
    sizes: Optional[str] = Field(None, serialization_alias='Sizes')
    fill: Optional[Fill] = Field(None, serialization_alias='Fill')
    line: Optional[Line] = Field(None, serialization_alias='Line')
    marker: Optional[Marker] = Field(None, serialization_alias='Marker')
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


class Chart(BaseModel):
    """
    Defines the configuration for a chart.

    Attributes:
        chart_type (str | ChartType): The type of chart, such as 'bar', 'line', etc.
        series (List[ChartSeries] | ChartSeries): The data series to be plotted
            in the chart.
        graph_format (Optional[GraphicOptions]): Graphical options for the chart.
        title (Optional[List[RichTextRun]]): The title of the chart.
        legend (Optional[ChartLegend]): The legend settings for the chart.
        dimension (Optional[ChartDimension]): The dimensions of the chart.
        vary_colors (Optional[bool]): Specifies if colors should vary by data point.
        x_axis (Optional[ChartAxis]): The configuration of the X-axis.
        y_axis (Optional[ChartAxis]): The configuration of the Y-axis.
        plot_area (Optional[ChartPlotArea]): The configuration of the plot area.
        fill (Optional[Fill]): The fill settings for the chart.
        border (Optional[Line]): The border settings for the chart.
        show_blanks_as (Optional[str]): Specifies how blanks should be shown in the chart.
        bubble_size (Optional[int]): The size of bubbles in a bubble chart.
        hole_size (Optional[int]): The size of the hole in a doughnut chart.
        order (Optional[int]): The order of the series in the chart.
    """

    chart_type: str | ChartType = Field(..., serialization_alias='Type')
    series: List[ChartSeries] | ChartSeries = Field(None, serialization_alias='Series')
    graph_format: Optional[GraphicOptions] = Field(None, serialization_alias='Format')
    title: Optional[List[RichTextRun]] = Field(None, serialization_alias='Title')
    legend: Optional[ChartLegend] = Field(None, serialization_alias='Legend')
    dimension: Optional[ChartDimension] = Field(None, serialization_alias='Dimension')
    vary_colors: Optional[bool] = Field(None, serialization_alias='VaryColors')
    x_axis: Optional[ChartAxis] = Field(None, serialization_alias='XAxis')
    y_axis: Optional[ChartAxis] = Field(None, serialization_alias='YAxis')
    plot_area: Optional[ChartPlotArea] = Field(None, serialization_alias='PlotArea')
    fill: Optional[Fill] = Field(None, serialization_alias='Fill')
    border: Optional[Line] = Field(None, serialization_alias='Border')
    show_blanks_as: Optional[str] = Field(None, serialization_alias='ShowBlanksAs')
    bubble_size: Optional[int] = Field(None, serialization_alias='BubbleSize')
    hole_size: Optional[int] = Field(None, serialization_alias='HoleSize')
    order: Optional[int] = Field(None, serialization_alias='order')

    @field_serializer('series')
    @classmethod
    def series_serializer(cls, series: List[ChartSeries] | ChartSeries) -> List[ChartSeries]:
        series = [series] if isinstance(series, ChartSeries) else series
        return series

    @field_serializer('chart_type')
    @classmethod
    def chart_type_validator(cls, chart_type: str | ChartType) -> int:
        if isinstance(chart_type, ChartType):
            return chart_type.value
        return ChartType.get_enum(chart_type).value
