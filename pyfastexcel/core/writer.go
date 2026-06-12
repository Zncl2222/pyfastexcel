package core

// #include <stdlib.h>
import (
	"C"
)
import (
	"encoding/base64"
	"fmt"
	"strconv"
	"strings"

	"github.com/perimeterx/marshmallow"
	"github.com/xuri/excelize/v2"
)

var styleMap map[string]int

type (
	// StyleWrapper is used by marshmallow for partial JSON unmarshalling:
	// the style section is decoded into a typed map while all other fields
	// are returned as map[string]interface{} by marshmallow.Unmarshal.
	StyleWrapper struct {
		Style map[string]map[string]interface{} `json:"Style" binding:"required"`
	}
)

type ExcelWriter struct {
	File       *excelize.File
	StyleMap   map[string]interface{}
	Content    map[string]interface{}
	FileProps  map[string]interface{}
	Protection map[string]interface{}
	SheetOrder []interface{}
	Engine     interface{}
}

// WriteExcelRaw parses the JSON payload using marshmallow (backed by easyjson,
// the fastest option for large nested payloads) and returns raw Excel bytes
// without any base64 encoding.
//
// Args:
//
//	data (string): JSON data representing the Excel file information.
//
// Returns:
//
//	[]byte: Raw bytes of the generated Excel file.
func WriteExcelRaw(data string) []byte {
	var styleStruct StyleWrapper
	strJson, err := marshmallow.Unmarshal([]byte(data), &styleStruct)
	if err != nil {
		panic(err)
	}
	writer := ExcelWriter{
		File:       excelize.NewFile(),
		StyleMap:   strJson["style"].(map[string]interface{}),
		Content:    strJson["content"].(map[string]interface{}),
		FileProps:  strJson["file_props"].(map[string]interface{}),
		Protection: strJson["protection"].(map[string]interface{}),
		SheetOrder: strJson["sheet_order"].([]interface{}),
	}
	return writer.writeExcelRaw()
}

// WriteExcel takes a JSON string and returns a base64 encoded Excel file.
// Kept for backward compatibility with Go tests.
func WriteExcel(data string) string {
	return base64.StdEncoding.EncodeToString(WriteExcelRaw(data))
}

// writeExcelRaw writes all sheets and returns the raw Excel bytes.
func (ew *ExcelWriter) writeExcelRaw() []byte {
	styleMap = CreateStyle(ew.File, ew.StyleMap)
	ew.setFileProps(ew.FileProps)
	if len(ew.Protection) != 0 {
		ew.setProtection(ew.Protection)
	}

	sheetCount := 1
	hasSheet1 := false
	var pivotTableList [][]interface{}
	for s := range ew.Content {
		if s == "Sheet1" {
			hasSheet1 = true
		}
	}
	for _, sheet := range ew.SheetOrder {
		sheet := sheet.(string)
		sheetData := ew.Content[sheet].(map[string]interface{})

		if !hasSheet1 && sheetCount == 1 {
			ew.File.SetSheetName("Sheet1", sheet)
			hasSheet1 = true
		} else {
			ew.File.NewSheet(sheet)
			sheetCount++
		}
		if sheetData["WriterEngine"] == "NormalWriter" {
			ew.performNormalWrite(sheet, sheetData)
			ew.createTable(sheet, sheetData["Table"].([]interface{}))
		} else {
			streamWriter := ew.performStreamWrite(sheet, sheetData)
			streamCreateTable(streamWriter, sheetData["Table"].([]interface{}))
			if err := streamWriter.Flush(); err != nil {
				fmt.Println(err)
			}
		}
		pivotTableList = append(pivotTableList, sheetData["PivotTable"].([]interface{}))
		if err := ew.File.SetSheetVisible(sheet, sheetData["SheetVisible"].(bool)); err != nil {
			fmt.Println(err)
		}
	}

	for _, pivot := range pivotTableList {
		ew.seedPivotSourceHeaders(pivot)
		ew.createPivotTable(pivot)
	}

	buffer, _ := ew.File.WriteToBuffer()
	return buffer.Bytes()
}

// writeExcel is kept for backward compatibility with Go tests.
func (ew *ExcelWriter) writeExcel() string {
	return base64.StdEncoding.EncodeToString(ew.writeExcelRaw())
}

func getCellScalarValue(item interface{}) interface{} {
	if cell, ok := item.([]interface{}); ok && len(cell) > 0 {
		return cell[0]
	}
	if cell, ok := item.(excelize.Cell); ok {
		return cell.Value
	}
	return item
}

func (ew *ExcelWriter) seedPivotSourceHeaders(pivotData []interface{}) {
	for _, pivot := range pivotData {
		pivotMap := pivot.(map[string]interface{})
		dataRange, ok := pivotMap["DataRange"].(string)
		if !ok || !strings.Contains(dataRange, "!") {
			continue
		}

		rangeParts := strings.SplitN(dataRange, "!", 2)
		sheetName := strings.Trim(rangeParts[0], "'")
		sourceRange := strings.ReplaceAll(rangeParts[1], "$", "")
		cellRefs := strings.SplitN(sourceRange, ":", 2)
		if len(cellRefs) != 2 {
			continue
		}

		startCol, startRow, err := excelize.CellNameToCoordinates(cellRefs[0])
		if err != nil {
			continue
		}
		endCol, endRow, err := excelize.CellNameToCoordinates(cellRefs[1])
		if err != nil {
			continue
		}
		if endCol < startCol {
			startCol, endCol = endCol, startCol
		}
		if endRow < startRow {
			startRow = endRow
		}

		sheetData, ok := ew.Content[sheetName].(map[string]interface{})
		if !ok {
			continue
		}
		dataRows, ok := sheetData["Data"].([]interface{})
		if !ok || len(dataRows) < startRow {
			continue
		}
		headerRow, ok := dataRows[startRow-1].([]interface{})
		if !ok {
			continue
		}

		for col := startCol; col <= endCol; col++ {
			headerIndex := col - startCol
			if headerIndex >= len(headerRow) {
				continue
			}
			cell, err := excelize.CoordinatesToCellName(col, startRow)
			if err != nil {
				continue
			}
			if err := ew.File.SetCellValue(sheetName, cell, getCellScalarValue(headerRow[headerIndex])); err != nil {
				fmt.Println(err)
			}
		}
	}
}

// streamWriter writes content to different sheets in the Excel file based on provided data.
//
// Args:
//
//	file (*excelize.File): The Excel file object.
//	data (map[string]interface{}): Map containing data for each sheet.
func (ew *ExcelWriter) performStreamWrite(sheet string, sheetData map[string]interface{}) *excelize.StreamWriter {
	// Add Chart
	ew.addChart(sheet, sheetData["Chart"].([]interface{}))

	// Set DataValidations
	ew.setDataValidation(sheet, sheetData["DataValidation"].([]interface{}))

	// Add Comment
	ew.addComment(sheet, sheetData["Comment"].([]interface{}))

	// Set Panes
	panes := ew.Content[sheet].(map[string]interface{})["Panes"].(map[string]interface{})
	ew.setPanes(sheet, panes)

	// Set AutoFilters
	autoFilters := ew.Content[sheet].(map[string]interface{})["AutoFilter"].([]interface{})
	ew.setAutoFilter(sheet, autoFilters)

	streamWriter, _ := ew.File.NewStreamWriter(sheet)

	// CellWidtrh should be set before SetRow
	// Height should be set with SetRow in StreamWriter
	setCellWidth(streamWriter, sheetData)
	rowHeightMap := getRowHeightMap(sheetData)

	mergeCell(streamWriter, sheetData["MergeCells"].([]interface{}))

	// Write Data
	startedRow := 1
	excelData := sheetData["Data"].([]interface{})
	if sheetData["NoStyle"] == false {
		for i, rowData := range excelData {
			for j, cellData := range rowData.([]interface{}) {
				if cellData == nil {
					continue
				}
				excelData[i].([]interface{})[j] = createCell(cellData.([]interface{}))
			}
		}
	}

	for i, rowData := range excelData {
		row := rowData.([]interface{})
		cell, _ := excelize.CoordinatesToCellName(1, i+startedRow)

		// Write cell with Height if rowHeightMap key found
		if rowHeight, ok := rowHeightMap[strconv.Itoa(i+startedRow)]; ok {
			if err := streamWriter.SetRow(cell, row, rowHeight); err != nil {
				fmt.Println(err)
			}
		} else {
			if err := streamWriter.SetRow(cell, row); err != nil {
				fmt.Println(err)
			}
		}
	}
	return streamWriter
}

// normalWriter writes content to different sheets in the Excel file based on provided data.
//
// Args:
//
//	file (*excelize.File): The Excel file object.
//	data (map[string]interface{}): Map containing data for each sheet.
func (ew *ExcelWriter) performNormalWrite(sheet string, sheetData map[string]interface{}) {

	// Add Chart
	ew.addChart(sheet, sheetData["Chart"].([]interface{}))

	// Set DataValidations
	ew.setDataValidation(sheet, sheetData["DataValidation"].([]interface{}))

	// Add Comment
	ew.addComment(sheet, sheetData["Comment"].([]interface{}))

	// Set Panes
	panes := ew.Content[sheet].(map[string]interface{})["Panes"].(map[string]interface{})
	ew.setPanes(sheet, panes)

	// Set AutoFilters
	autoFilters := ew.Content[sheet].(map[string]interface{})["AutoFilter"].([]interface{})
	ew.setAutoFilter(sheet, autoFilters)

	// Set Cell Width and Height
	ew.setCellWidthNormalWriter(sheet, sheetData)
	ew.setCellHeightNormalWriter(sheet, sheetData)

	// Merge Cell
	ew.mergeCellNormalWriter(sheet, sheetData["MergeCells"].([]interface{}))

	// Group col and row
	if sheetData["GroupedRow"] != nil {
		ew.groupRow(sheet, sheetData["GroupedRow"].([]interface{}))
	}
	if sheetData["GroupedCol"] != nil {
		ew.groupCol(sheet, sheetData["GroupedCol"].([]interface{}))
	}

	// Write Data
	startedRow := 1
	excelData := sheetData["Data"].([]interface{})
	for i, rowData := range excelData {
		row := rowData.([]interface{})

		for col, item := range row {
			colCell, _ := excelize.CoordinatesToCellName(col+startedRow, i+startedRow)
			v := item.([]interface{})
			if len(v) == 0 {
				if err := ew.File.SetCellValue(sheet, colCell, ""); err != nil {
					fmt.Println(err)
				}
				if err := ew.File.SetCellStyle(sheet, colCell, colCell, styleMap["DEFAULT_STYLE"]); err != nil {
					fmt.Println(err)
				}
			} else {
				switch value := v[0].(type) {
				case string:
					if strings.HasPrefix(value, "=") {
						if err := ew.File.SetCellFormula(sheet, colCell, normalizeFormula(value)); err != nil {
							fmt.Println(err)
						}
					} else {
						if err := ew.File.SetCellValue(sheet, colCell, value); err != nil {
							fmt.Println(err)
						}
					}
				default:
					if err := ew.File.SetCellValue(sheet, colCell, value); err != nil {
						fmt.Println(err)
					}
				}
				if err := ew.File.SetCellStyle(sheet, colCell, colCell, styleMap[item.([]interface{})[1].(string)]); err != nil {
					fmt.Println(err)
				}
			}
		}
	}
}
