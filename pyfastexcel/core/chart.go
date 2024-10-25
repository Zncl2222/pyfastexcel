package core

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

func getLineStyle(line map[string]interface{}) excelize.ChartLine {
	var chartLine excelize.ChartLine

	lineMappings := []fieldMapping{
		{Name: "Type", Type: "ChartLineType"},
		{Name: "Smooth", Type: "bool"},
		{Name: "Width", Type: "float64"},
		{Name: "ShowMarkerLine", Type: "bool"},
	}

	setField(&chartLine, line, lineMappings)

	return chartLine
}

func getMarkerStyle(marker map[string]interface{}) excelize.ChartMarker {
	var chartMarker excelize.ChartMarker

	markerMappings := []fieldMapping{
		{Name: "Fill", Type: "Fill"},
		{Name: "Symbol", Type: "string"},
		{Name: "Size", Type: "int"},
	}

	setField(&chartMarker, marker, markerMappings)

	return chartMarker
}

func getSeriesStruct(series []interface{}) []excelize.ChartSeries {
	var chartSeries []excelize.ChartSeries

	chartSeriesMappings := []fieldMapping{
		{Name: "Name", Type: "string"},
		{Name: "Categories", Type: "string"},
		{Name: "Values", Type: "string"},
		{Name: "Sizes", Type: "string"},
		{Name: "Fill", Type: "Fill"},
		{Name: "Line", Type: "ChartLine"},
		{Name: "Marker", Type: "ChartMarker"},
		{Name: "DataLabelPosition", Type: "ChartDataLabelPositionType"},
	}

	for _, seriesMap := range series {
		c := excelize.ChartSeries{}
		setField(&c, seriesMap.(map[string]interface{}), chartSeriesMappings)
		chartSeries = append(chartSeries, c)
	}

	return chartSeries
}

func getFormatStruct(format interface{}) excelize.GraphicOptions {
	var chartFormat excelize.GraphicOptions

	if format == nil {
		return chartFormat
	}

	formatMappings := []fieldMapping{
		{Name: "AltText", Type: "string"},
		{Name: "PrintObject", Type: "*bool"},
		{Name: "Locked", Type: "*bool"},
		{Name: "LockAspectRatio", Type: "bool"},
		{Name: "AutoFit", Type: "bool"},
		{Name: "OffsetX", Type: "int"},
		{Name: "OffsetY", Type: "int"},
		{Name: "ScaleX", Type: "float64"},
		{Name: "ScaleY", Type: "float64"},
		{Name: "Hyperlink", Type: "string"},
		{Name: "HyperlinkType", Type: "string"},
		{Name: "Positioning", Type: "string"},
	}

	setField(&chartFormat, format.(map[string]interface{}), formatMappings)

	return chartFormat
}

func getDimensionStruct(dim interface{}) excelize.ChartDimension {
	var chartDimension excelize.ChartDimension

	if dim == nil {
		return chartDimension
	}

	dimMappings := []fieldMapping{
		{Name: "Width", Type: "uint64"},
		{Name: "Height", Type: "uint64"},
	}

	setField(&chartDimension, dim.(map[string]interface{}), dimMappings)

	return chartDimension
}

func getLegendStruct(legend interface{}) excelize.ChartLegend {
	var chartLegend excelize.ChartLegend

	if legend == nil {
		return chartLegend
	}

	legendMappings := []fieldMapping{
		{Name: "Position", Type: "string"},
		{Name: "ShowLegendKey", Type: "bool"},
	}

	setField(&chartLegend, legend.(map[string]interface{}), legendMappings)

	return chartLegend
}

func getTitleStruct(titles interface{}) []excelize.RichTextRun {
	chartTitles := []excelize.RichTextRun{}
	if titles == nil {
		return chartTitles
	}

	for _, title := range titles.([]interface{}) {
		chartTitle := excelize.RichTextRun{}
		if title == nil {
			chartTitles = append(chartTitles, chartTitle)
			continue
		}

		titleMappings := []fieldMapping{
			{Name: "Font", Type: "*Font"},
			{Name: "Text", Type: "string"},
		}

		setField(&chartTitle, title.(map[string]interface{}), titleMappings)
		chartTitles = append(chartTitles, chartTitle)
	}

	return chartTitles
}

func getChartNumFmtStruct(numFmt interface{}) excelize.ChartNumFmt {
	var chartNumFmt excelize.ChartNumFmt

	if numFmt == nil {
		return chartNumFmt
	}

	numFmtMappings := []fieldMapping{
		{Name: "CustomNumFmt", Type: "string"},
		{Name: "SourceLinked", Type: "bool"},
	}

	setField(&chartNumFmt, numFmt.(map[string]interface{}), numFmtMappings)
	return chartNumFmt
}

func getAxisStruct(Axis interface{}) excelize.ChartAxis {
	var chartAxis excelize.ChartAxis

	if Axis == nil {
		return chartAxis
	}

	AxisMappings := []fieldMapping{
		{Name: "None", Type: "bool"},
		{Name: "MajorGridLines", Type: "bool"},
		{Name: "MinorGridLines", Type: "bool"},
		{Name: "MajorUnit", Type: "float64"},
		{Name: "TickLabelSkip", Type: "int"},
		{Name: "ReverseOrder", Type: "bool"},
		{Name: "Secondary", Type: "bool"},
		{Name: "Maximum", Type: "*float64"},
		{Name: "Minimum", Type: "*float64"},
		{Name: "Font", Type: "Font"},
		{Name: "LogBase", Type: "float64"},
		{Name: "NumFmt", Type: "ChartNumFmt"},
		{Name: "Title", Type: "[]RichTextRun"},
		{Name: "axID", Type: "int"},
	}

	setField(&chartAxis, Axis.(map[string]interface{}), AxisMappings)
	return chartAxis
}

func getPlotAreaStruct(plotArea interface{}) excelize.ChartPlotArea {
	var chartPlotArea excelize.ChartPlotArea

	if plotArea == nil {
		return chartPlotArea
	}

	plotAreaMappings := []fieldMapping{
		{Name: "SecondPlotValues", Type: "int"},
		{Name: "ShowBubbleSize", Type: "bool"},
		{Name: "ShowCatName", Type: "bool"},
		{Name: "ShowLeaderLines", Type: "bool"},
		{Name: "ShowPercent", Type: "bool"},
		{Name: "ShowSerName", Type: "bool"},
		{Name: "ShowVal", Type: "bool"},
		{Name: "Fill", Type: "Fill"},
		{Name: "NumFmt", Type: "ChartNumFmt"},
	}

	setField(&chartPlotArea, plotArea.(map[string]interface{}), plotAreaMappings)
	return chartPlotArea
}

func getChartFill(fill interface{}) excelize.Fill {
	if fill == nil {
		return excelize.Fill{}
	}

	return getFillStyle(fill.(map[string]interface{}))
}

func getBorderStruct(border interface{}) excelize.ChartLine {
	var chartBorder excelize.ChartLine

	if border == nil {
		return chartBorder
	}

	borderMappings := []fieldMapping{
		{Name: "Type", Type: "ChartLineType"},
		{Name: "Smooth", Type: "bool"},
		{Name: "Width", Type: "float64"},
	}

	setField(&chartBorder, border.(map[string]interface{}), borderMappings)
	return chartBorder
}

func (ew *ExcelWriter) addChart(sheet string, charts []interface{}) {
	for _, chart := range charts {
		chart := chart.(map[string]interface{})
		chartData := chart["chart"].([]interface{})
		cell := chart["cell"].(string)
		comboCharts := []*excelize.Chart{}
		for _, c := range chartData {
			c := c.(map[string]interface{})
			chartType := excelize.ChartType(uint8(c["Type"].(float64)))
			series := getSeriesStruct(c["Series"].([]interface{}))
			format := getFormatStruct(c["Format"])
			dimension := getDimensionStruct(c["Dimension"])
			legend := getLegendStruct(c["Legend"])
			title := getTitleStruct(c["Title"])
			xAxis := getAxisStruct(c["XAxis"])
			yAxis := getAxisStruct(c["YAxis"])
			plotArea := getPlotAreaStruct(c["PlotArea"])
			fill := getChartFill(c["Fill"])
			border := getBorderStruct(c["Border"])
			comboCharts = append(comboCharts, &excelize.Chart{
				Type:       chartType,
				Series:     series,
				Format:     format,
				Legend:     legend,
				Dimension:  dimension,
				Title:      title,
				XAxis:      xAxis,
				YAxis:      yAxis,
				PlotArea:   plotArea,
				Fill:       fill,
				Border:     border,
				BubbleSize: int(getFloat64Value(c, "BubbleSize", 100.0)),
				HoleSize:   int(getFloat64Value(c, "HoleSize", 75.0)),
			})
		}
		if err := ew.File.AddChart(sheet, cell, comboCharts[0], comboCharts[1:]...); err != nil {
			fmt.Println(err)
		}
	}
}
