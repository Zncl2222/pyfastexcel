package core

import (
	"testing"

	"github.com/xuri/excelize/v2"
)

func TestGetLineStyle(t *testing.T) {
	lineData := map[string]interface{}{
		"Type":   0.0,
		"Smooth": true,
		"Width":  2.0,
	}

	expected := excelize.ChartLine{
		Type:   excelize.ChartLineSolid,
		Smooth: true,
		Width:  2.0,
	}
	result := getLineStyle(lineData)
	if result != expected {
		t.Errorf("getLineStyle() = %v, want %v", result, expected)
	}
}

func TestGetMarkerStyle(t *testing.T) {
	markerData := map[string]interface{}{
		"Fill":   map[string]interface{}{"Color": "FF0000"},
		"Symbol": "circle",
		"Size":   10.0,
	}

	result := getMarkerStyle(markerData)
	if result.Fill.Color[0] != "FF0000" {
		t.Errorf("getMarkerStyle() = %v, want %v", result.Fill.Color, []string{"FF0000"})
	}
	if result.Symbol != "circle" {
		t.Errorf("getMarkerStyle() = %v, want %v", result.Symbol, "circle")
	}
	if result.Size != 10 {
		t.Errorf("getMarkerStyle() = %v, want %v", result.Size, 10)
	}
}

func TestGetSeriesStruct(t *testing.T) {
	seriesData := []interface{}{
		map[string]interface{}{
			"Name":       "Series1",
			"Categories": "A1:A10",
			"Values":     "B1:B10",
			"Sizes":      "C1:C10",
			"Fill":       map[string]interface{}{"Color": "0000FF"},
			"Line":       map[string]interface{}{"Type": 0.0, "Width": 1.0},
			"Marker":     map[string]interface{}{"Symbol": "square", "Size": 5.0},
		},
	}

	expected := []excelize.ChartSeries{
		{
			Name:       "Series1",
			Categories: "A1:A10",
			Values:     "B1:B10",
			Sizes:      "C1:C10",
			Fill:       excelize.Fill{Color: []string{"0000FF"}},
			Line:       excelize.ChartLine{Type: excelize.ChartLineSolid, Width: 1.0},
			Marker:     excelize.ChartMarker{Symbol: "square", Size: 5},
		},
	}

	result := getSeriesStruct(seriesData)
	if len(result) != len(expected) {
		t.Fatalf("getSeriesStruct() = %v, want %v", result, expected)
	}
	for i := range result {
		if result[i].Name != expected[i].Name {
			t.Errorf("getSeriesStruct() = %v, want %v", result[i].Name, expected[i].Name)
		}
		if result[i].Categories != expected[i].Categories {
			t.Errorf("getSeriesStruct() = %v, want %v", result[i].Categories, expected[i].Categories)
		}
		if result[i].Values != expected[i].Values {
			t.Errorf("getSeriesStruct() = %v, want %v", result[i].Values, expected[i].Values)
		}
		if result[i].Sizes != expected[i].Sizes {
			t.Errorf("getSeriesStruct() = %v, want %v", result[i].Sizes, expected[i].Sizes)
		}
		if result[i].Fill.Color[0] != expected[i].Fill.Color[0] {
			t.Errorf("getSeriesStruct() = %v, want %v", result[i].Fill.Color, expected[i].Fill.Color)
		}
		if result[i].Line.Type != expected[i].Line.Type {
			t.Errorf("getSeriesStruct() = %v, want %v", result[i].Line.Type, expected[i].Line.Type)
		}
		if result[i].Line.Width != expected[i].Line.Width {
			t.Errorf("getSeriesStruct() = %v, want %v", result[i].Line.Width, expected[i].Line.Width)
		}
		if result[i].Marker.Symbol != expected[i].Marker.Symbol {
			t.Errorf("getSeriesStruct() = %v, want %v", result[i].Marker.Symbol, expected[i].Marker.Symbol)
		}
		if result[i].Marker.Size != expected[i].Marker.Size {
			t.Errorf("getSeriesStruct() = %v, want %v", result[i].Marker.Size, expected[i].Marker.Size)
		}
	}
}

func TestGetFormatStruct(t *testing.T) {
	formatData := map[string]interface{}{
		"AltText":         "Some text",
		"PrintObject":     true,
		"Locked":          false,
		"LockAspectRatio": true,
		"AutoFit":         true,
		"OffsetX":         10.0,
		"OffsetY":         20.0,
		"ScaleX":          1.5,
		"ScaleY":          1.5,
		"Hyperlink":       "http://example.com",
		"HyperlinkType":   "external",
		"Positioning":     "absolute",
	}

	printObject := true
	locked := false

	expected := excelize.GraphicOptions{
		AltText:         "Some text",
		PrintObject:     &printObject,
		Locked:          &locked,
		LockAspectRatio: true,
		AutoFit:         true,
		OffsetX:         10,
		OffsetY:         20,
		ScaleX:          1.5,
		ScaleY:          1.5,
		Hyperlink:       "http://example.com",
		HyperlinkType:   "external",
		Positioning:     "absolute",
	}

	result := getFormatStruct(formatData)
	if result.AltText != expected.AltText {
		t.Errorf("getFormatStruct() = %v, want %v", result.AltText, expected.AltText)
	}
	if *result.PrintObject != *expected.PrintObject {
		t.Errorf("getFormatStruct() = %v, want %v", *result.PrintObject, true)
	}
	if *result.Locked != *expected.Locked {
		t.Errorf("getFormatStruct() = %v, want %v", *result.Locked, false)
	}
	if result.LockAspectRatio != expected.LockAspectRatio {
		t.Errorf("getFormatStruct() = %v, want %v", result.LockAspectRatio, expected.LockAspectRatio)
	}
	if result.AutoFit != expected.AutoFit {
		t.Errorf("getFormatStruct() = %v, want %v", result.AutoFit, expected.AutoFit)
	}
	if result.OffsetX != expected.OffsetX {
		t.Errorf("getFormatStruct() = %v, want %v", result.OffsetX, expected.OffsetY)
	}
	if result.OffsetY != expected.OffsetY {
		t.Errorf("getFormatStruct() = %v, want %v", result.OffsetY, expected.OffsetY)
	}
	if result.ScaleX != expected.ScaleX {
		t.Errorf("getFormatStruct() = %v, want %v", result.ScaleX, expected.ScaleX)
	}
	if result.ScaleY != expected.ScaleY {
		t.Errorf("getFormatStruct() = %v, want %v", result.ScaleY, expected.ScaleY)
	}
	if result.Hyperlink != expected.Hyperlink {
		t.Errorf("getFormatStruct() = %v, want %v", result.Hyperlink, expected.Hyperlink)
	}
	if result.HyperlinkType != expected.HyperlinkType {
		t.Errorf("getFormatStruct() = %v, want %v", result.HyperlinkType, expected.HyperlinkType)
	}
	if result.Positioning != expected.Positioning {
		t.Errorf("getFormatStruct() = %v, want %v", result.Positioning, expected.Positioning)
	}
}

func TestGetChartNumFmtStruct(t *testing.T) {
	numFmtData := map[string]interface{}{
		"CustomNumFmt": "0.00",
		"SourceLinked": true,
	}

	expected := excelize.ChartNumFmt{
		CustomNumFmt: "0.00",
		SourceLinked: true,
	}

	result := getChartNumFmtStruct(numFmtData)
	if result.CustomNumFmt != expected.CustomNumFmt {
		t.Errorf("getChartNumFmtStruct() = %v, want %v", result.CustomNumFmt, expected.CustomNumFmt)
	}
	if result.SourceLinked != expected.SourceLinked {
		t.Errorf("getChartNumFmtStruct() = %v, want %v", result.SourceLinked, expected.SourceLinked)
	}
}

func TestGetAxisStruct(t *testing.T) {
	majorUnit := 10.5
	maximum := float64(100)
	minimum := float64(0)

	axisData := map[string]interface{}{
		"None":           true,
		"MajorGridLines": true,
		"MinorGridLines": false,
		"MajorUnit":      majorUnit,
		"TickLabelSkip":  2.0,
		"ReverseOrder":   false,
		"Secondary":      true,
		"Maximum":        maximum,
		"Minimum":        minimum,
		"LogBase":        10.0,
		"NumFmt":         map[string]interface{}{"CustomNumFmt": "0.00"},
	}

	expected := excelize.ChartAxis{
		None:           true,
		MajorGridLines: true,
		MinorGridLines: false,
		MajorUnit:      10.5,
		TickLabelSkip:  2,
		ReverseOrder:   false,
		Secondary:      true,
		Maximum:        &maximum,
		Minimum:        &minimum,
		LogBase:        10.0,
		NumFmt:         excelize.ChartNumFmt{CustomNumFmt: "0.00"},
	}

	result := getAxisStruct(axisData)
	if result.None != expected.None {
		t.Errorf("getAxisStruct() None = %v, want %v", result.None, expected.None)
	}
	if result.MajorGridLines != expected.MajorGridLines {
		t.Errorf("getAxisStruct() MajorGridLines = %v, want %v", result.MajorGridLines, expected.MajorGridLines)
	}
	if result.MinorGridLines != expected.MinorGridLines {
		t.Errorf("getAxisStruct() MinorGridLines = %v, want %v", result.MinorGridLines, expected.MinorGridLines)
	}
	if result.MajorUnit != expected.MajorUnit {
		t.Errorf("getAxisStruct() MajorUnit = %v, want %v", result.MajorUnit, expected.MajorUnit)
	}
	if result.TickLabelSkip != expected.TickLabelSkip {
		t.Errorf("getAxisStruct() TickLabelSkip = %v, want %v", result.TickLabelSkip, expected.TickLabelSkip)
	}
	if result.ReverseOrder != expected.ReverseOrder {
		t.Errorf("getAxisStruct() ReverseOrder = %v, want %v", result.ReverseOrder, expected.ReverseOrder)
	}
	if result.Secondary != expected.Secondary {
		t.Errorf("getAxisStruct() Secondary = %v, want %v", result.Secondary, expected.Secondary)
	}
	if *result.Maximum != *expected.Maximum {
		t.Errorf("getAxisStruct() Maximum = %v, want %v", *result.Maximum, *expected.Maximum)
	}
	if *result.Minimum != *expected.Minimum {
		t.Errorf("getAxisStruct() Minimum = %v, want %v", *result.Minimum, *expected.Minimum)
	}
	if result.LogBase != expected.LogBase {
		t.Errorf("getAxisStruct() LogBase = %v, want %v", result.LogBase, expected.LogBase)
	}
	if result.NumFmt.CustomNumFmt != expected.NumFmt.CustomNumFmt {
		t.Errorf("getAxisStruct() NumFmt.CustomNumFmt = %v, want %v", result.NumFmt.CustomNumFmt, expected.NumFmt.CustomNumFmt)
	}
}

func TestGetPlotAreaStruct(t *testing.T) {
	plotAreaData := map[string]interface{}{
		"SecondPlotValues": 1.0,
		"ShowBubbleSize":   true,
		"ShowCatName":      false,
		"ShowLeaderLines":  true,
		"ShowPercent":      false,
		"ShowSerName":      true,
		"ShowVal":          false,
		"Fill":             map[string]interface{}{"Pattern": 1.0},
		"NumFmt":           map[string]interface{}{"CustomNumFmt": "0.00"},
	}

	expected := excelize.ChartPlotArea{
		SecondPlotValues: 1,
		ShowBubbleSize:   true,
		ShowCatName:      false,
		ShowLeaderLines:  true,
		ShowPercent:      false,
		ShowSerName:      true,
		ShowVal:          false,
		Fill:             excelize.Fill{Pattern: 1},
		NumFmt:           excelize.ChartNumFmt{CustomNumFmt: "0.00"},
	}

	result := getPlotAreaStruct(plotAreaData)
	if result.SecondPlotValues != expected.SecondPlotValues {
		t.Errorf("getPlotAreaStruct() SecondPlotValues = %v, want %v", result.SecondPlotValues, expected.SecondPlotValues)
	}
	if result.ShowBubbleSize != expected.ShowBubbleSize {
		t.Errorf("getPlotAreaStruct() ShowBubbleSize = %v, want %v", result.ShowBubbleSize, expected.ShowBubbleSize)
	}
	if result.ShowCatName != expected.ShowCatName {
		t.Errorf("getPlotAreaStruct() ShowCatName = %v, want %v", result.ShowCatName, expected.ShowCatName)
	}
	if result.ShowLeaderLines != expected.ShowLeaderLines {
		t.Errorf("getPlotAreaStruct() ShowLeaderLines = %v, want %v", result.ShowLeaderLines, expected.ShowLeaderLines)
	}
	if result.ShowPercent != expected.ShowPercent {
		t.Errorf("getPlotAreaStruct() ShowPercent = %v, want %v", result.ShowPercent, expected.ShowPercent)
	}
	if result.ShowSerName != expected.ShowSerName {
		t.Errorf("getPlotAreaStruct() ShowSerName = %v, want %v", result.ShowSerName, expected.ShowSerName)
	}
	if result.ShowVal != expected.ShowVal {
		t.Errorf("getPlotAreaStruct() ShowVal = %v, want %v", result.ShowVal, expected.ShowVal)
	}
	if result.Fill.Pattern != expected.Fill.Pattern {
		t.Errorf("getPlotAreaStruct() Fill.Pattern = %v, want %v", result.Fill.Pattern, expected.Fill.Pattern)
	}
	if result.NumFmt.CustomNumFmt != expected.NumFmt.CustomNumFmt {
		t.Errorf("getPlotAreaStruct() NumFmt.CustomNumFmt = %v, want %v", result.NumFmt.CustomNumFmt, expected.NumFmt.CustomNumFmt)
	}
}

func TestGetChartFill(t *testing.T) {
	fillData := map[string]interface{}{
		"Pattern": 1.0,
	}

	expected := excelize.Fill{
		Pattern: 1,
	}

	result := getChartFill(fillData)
	if result.Pattern != expected.Pattern {
		t.Errorf("getChartFill() Pattern = %v, want %v", result.Pattern, expected.Pattern)
	}
}

func TestGetBorderStruct(t *testing.T) {
	borderData := map[string]interface{}{
		"Type":   0.0,
		"Smooth": true,
		"Width":  2.5,
	}

	expected := excelize.ChartLine{
		Type:   excelize.ChartLineSolid,
		Smooth: true,
		Width:  2.5,
	}

	result := getBorderStruct(borderData)
	if result.Type != expected.Type {
		t.Errorf("getBorderStruct() Type = %v, want %v", result.Type, expected.Type)
	}
	if result.Smooth != expected.Smooth {
		t.Errorf("getBorderStruct() Smooth = %v, want %v", result.Smooth, expected.Smooth)
	}
	if result.Width != expected.Width {
		t.Errorf("getBorderStruct() Width = %v, want %v", result.Width, expected.Width)
	}
}

func TestGetDimensionStruct(t *testing.T) {
	dimData := map[string]interface{}{
		"Width":  800.0,
		"Height": 600.0,
	}

	expected := excelize.ChartDimension{
		Width:  800,
		Height: 600,
	}

	result := getDimensionStruct(dimData)
	if result.Width != expected.Width {
		t.Errorf("getDimensionStruct() Width = %v, want %v", result, expected.Width)
	}
	if result.Height != expected.Height {
		t.Errorf("getDimensionStruct() Height = %v, want %v", result.Height, expected.Height)
	}
}

func TestGetLegendStruct(t *testing.T) {
	legendData := map[string]interface{}{
		"Position":      "bottom",
		"ShowLegendKey": true,
	}

	expected := excelize.ChartLegend{
		Position:      "bottom",
		ShowLegendKey: true,
	}

	result := getLegendStruct(legendData)
	if result.Position != expected.Position {
		t.Errorf("getLegendStruct() Position = %v, want %v", result.Position, expected.Position)
	}
	if result.ShowLegendKey != expected.ShowLegendKey {
		t.Errorf("getLegendStruct() ShowLegendKey = %v, want %v", result.ShowLegendKey, expected.ShowLegendKey)
	}
}

func TestGetTitleStruct(t *testing.T) {
	titleData := []interface{}{
		map[string]interface{}{
			"Font": nil,
			"Text": "Title1",
		},
		map[string]interface{}{
			"Font": nil,
			"Text": "Title2",
		},
	}

	expected := []excelize.RichTextRun{
		{Text: "Title1"},
		{Text: "Title2"},
	}

	result := getTitleStruct(titleData)
	for i, title := range result {
		if title.Text != expected[i].Text {
			t.Errorf("getTitleStruct() Title[%d].Text = %v, want %v", i, title.Text, expected[i].Text)
		}
	}
}

func TestAddChart(t *testing.T) {
	file := excelize.NewFile()
	sheet := "Sheet1"
	charts := []interface{}{
		map[string]interface{}{
			"chart": []interface{}{
				map[string]interface{}{
					"Type": 1.0,
					"Series": []interface{}{
						map[string]interface{}{
							"Name":       "Series1",
							"Categories": "A1:A10",
							"Values":     "B1:B10",
						},
					},
					"Format":     nil,
					"Dimension":  nil,
					"Legend":     nil,
					"Title":      nil,
					"XAxis":      nil,
					"YAxis":      nil,
					"PlotArea":   nil,
					"Fill":       nil,
					"Border":     nil,
					"BubbleSize": nil,
					"HoleSize":   nil,
				},
			},
			"cell": "A1",
		},
	}

	addChart(file, sheet, charts)
	file.Close()
}
