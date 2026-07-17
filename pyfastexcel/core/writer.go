package core

import (
	"encoding/base64"
	"errors"
	"fmt"
	"io"
	"sort"
	"strconv"
	"strings"

	"github.com/perimeterx/marshmallow"
	"github.com/xuri/excelize/v2"
)

type (
	StyleWrapper struct {
		Style map[string]map[string]interface{} `json:"Style" binding:"required"`
	}
)

type ExcelWriter struct {
	File               *excelize.File
	StyleMap           map[string]interface{}
	StyleIDs           map[string]int
	WireStyleIDs       []int
	Content            map[string]interface{}
	FileProps          map[string]interface{}
	Protection         map[string]interface{}
	SheetOrder         []interface{}
	Engine             interface{}
	PivotSourceHeaders map[string]map[int][]interface{}
	// WireRowStream references the MessagePack row bytes of a PFX2 payload;
	// sheet_offsets index into it for concurrent per-sheet decoding.
	WireRowStream []byte
}

// WriteExcel takes a JSON string containing file properties, styles,
// and content and returns a base64 encoded string of the generated Excel file.
//
// Args:
//
//	data (string): JSON data representing the Excel file information.
//
// Returns:
//
//	string: Base64 encoded string of the generated Excel file.
//
// Panics:
//   - panics on errors during JSON unmarshalling or cell conversion.
func WriteExcel(data string) string {
	result, err := WriteExcelBytes(data)
	if err != nil {
		panic(err)
	}
	return base64.StdEncoding.EncodeToString(result)
}

// WriteExcelBytes generates an XLSX workbook from the legacy JSON wire format.
// Unlike WriteExcel, it returns ordinary errors and raw workbook bytes.
func WriteExcelBytes(data string) (result []byte, err error) {
	defer recoverAsError(&err)

	writer, err := newExcelWriter([]byte(data))
	if err != nil {
		return nil, err
	}
	defer func() {
		err = errors.Join(err, writer.File.Close())
	}()

	if err = writer.buildLegacyWorkbook(); err != nil {
		return nil, err
	}
	return writer.writeToBytes()
}

func newExcelWriter(data []byte) (*ExcelWriter, error) {
	configureZipCompression()
	var StyleStruct StyleWrapper
	strJson, err := marshmallow.Unmarshal(data, &StyleStruct)
	if err != nil {
		return nil, fmt.Errorf("decode workbook metadata: %w", err)
	}
	styleMap, ok := strJson["style"].(map[string]interface{})
	if !ok {
		return nil, fmt.Errorf("workbook metadata field %q must be an object", "style")
	}
	content, ok := strJson["content"].(map[string]interface{})
	if !ok {
		return nil, fmt.Errorf("workbook metadata field %q must be an object", "content")
	}
	fileProps, ok := strJson["file_props"].(map[string]interface{})
	if !ok {
		return nil, fmt.Errorf("workbook metadata field %q must be an object", "file_props")
	}
	protection, ok := strJson["protection"].(map[string]interface{})
	if !ok {
		return nil, fmt.Errorf("workbook metadata field %q must be an object", "protection")
	}
	sheetOrder, ok := strJson["sheet_order"].([]interface{})
	if !ok {
		return nil, fmt.Errorf("workbook metadata field %q must be an array", "sheet_order")
	}
	writer := &ExcelWriter{
		File:       excelize.NewFile(),
		StyleMap:   styleMap,
		Content:    content,
		FileProps:  fileProps,
		Protection: protection,
		SheetOrder: sheetOrder,
		// Engine:     strJson["engine"],
	}
	return writer, nil
}

func (ew *ExcelWriter) writeExcel() string {
	if err := ew.buildLegacyWorkbook(); err != nil {
		panic(err)
	}
	result, err := ew.writeToBytes()
	if err != nil {
		panic(err)
	}
	return base64.StdEncoding.EncodeToString(result)
}

func (ew *ExcelWriter) initializeStyles(styleNames []string) error {
	styleIDs, wireStyleIDs, err := createStylesOrdered(ew.File, ew.StyleMap, styleNames)
	if err != nil {
		return err
	}
	ew.StyleIDs = styleIDs
	ew.WireStyleIDs = wireStyleIDs
	return nil
}

func (ew *ExcelWriter) buildLegacyWorkbook() error {
	styleNames := make([]string, 0, len(ew.StyleMap))
	for name := range ew.StyleMap {
		styleNames = append(styleNames, name)
	}
	sort.Strings(styleNames)
	if err := ew.initializeStyles(styleNames); err != nil {
		return err
	}
	if err := ew.setFileProps(ew.FileProps); err != nil {
		return err
	}
	if len(ew.Protection) != 0 {
		if err := ew.setProtection(ew.Protection); err != nil {
			return err
		}
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
			if err := ew.File.SetSheetName("Sheet1", sheet); err != nil {
				return fmt.Errorf("rename first sheet to %q: %w", sheet, err)
			}
			hasSheet1 = true
		} else {
			if _, err := ew.File.NewSheet(sheet); err != nil {
				return fmt.Errorf("create sheet %q: %w", sheet, err)
			}
			sheetCount++
		}
		if sheetData["WriterEngine"] == "NormalWriter" {
			if err := ew.performNormalWrite(sheet, sheetData); err != nil {
				return err
			}
			// Excelize should create table with the existed row.
			if err := ew.createTable(sheet, sheetData["Table"].([]interface{})); err != nil {
				return err
			}
		} else {
			streamWriter, err := ew.performStreamWrite(sheet, sheetData)
			if err != nil {
				return err
			}
			// Create Stream Table
			// Excelize should create table with the existed row.
			if err := streamCreateTable(streamWriter, sheetData["Table"].([]interface{})); err != nil {
				return fmt.Errorf("create tables on stream sheet %q: %w", sheet, err)
			}

			if err := streamWriter.Flush(); err != nil {
				return fmt.Errorf("flush stream sheet %q: %w", sheet, err)
			}
		}
		// To prevent the pivot table from being created before the data is written
		// we store the pivot table data in a list and create it after the data is written
		pivotTableList = append(pivotTableList, sheetData["PivotTable"].([]interface{}))

		// Set Sheet Visible
		if err := ew.File.SetSheetVisible(sheet, sheetData["SheetVisible"].(bool)); err != nil {
			return fmt.Errorf("set visibility for sheet %q: %w", sheet, err)
		}

	}

	// Create pivot tables after every sheet has been written and flushed. Large
	// streamed worksheets can spill to temp files; seed the source header row in
	// memory so excelize can validate PivotTableOptions.DataRange.
	for _, pivot := range pivotTableList {
		if err := ew.seedPivotSourceHeaders(pivot); err != nil {
			return err
		}
		if err := ew.createPivotTable(pivot); err != nil {
			return err
		}
	}
	return nil
}

func (ew *ExcelWriter) writeToBytes() ([]byte, error) {
	buffer, err := ew.File.WriteToBuffer()
	if err != nil {
		return nil, fmt.Errorf("serialize workbook: %w", err)
	}
	return buffer.Bytes(), nil
}

func (ew *ExcelWriter) writeTo(output io.Writer) error {
	if err := ew.File.Write(output); err != nil {
		return fmt.Errorf("serialize workbook: %w", err)
	}
	return nil
}

func recoverAsError(err *error) {
	if recovered := recover(); recovered != nil {
		*err = errors.Join(*err, fmt.Errorf("pyfastexcel panic: %v", recovered))
	}
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

func (ew *ExcelWriter) seedPivotSourceHeaders(pivotData []interface{}) error {
	for pivotIndex, pivot := range pivotData {
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
			return fmt.Errorf("parse pivot table %d source start cell %q: %w", pivotIndex+1, cellRefs[0], err)
		}
		endCol, endRow, err := excelize.CellNameToCoordinates(cellRefs[1])
		if err != nil {
			return fmt.Errorf("parse pivot table %d source end cell %q: %w", pivotIndex+1, cellRefs[1], err)
		}
		if endCol < startCol {
			startCol, endCol = endCol, startCol
		}
		if endRow < startRow {
			startRow = endRow
		}

		var headerRow []interface{}
		if capturedRows, ok := ew.PivotSourceHeaders[sheetName]; ok {
			headerRow = capturedRows[startRow]
		}
		if headerRow == nil {
			sheetData, ok := ew.Content[sheetName].(map[string]interface{})
			if !ok {
				continue
			}
			dataRows, ok := sheetData["Data"].([]interface{})
			if !ok || len(dataRows) < startRow {
				continue
			}
			headerRow, ok = dataRows[startRow-1].([]interface{})
			if !ok {
				continue
			}
		}

		for col := startCol; col <= endCol; col++ {
			headerIndex := col - 1
			if headerIndex >= len(headerRow) {
				continue
			}
			cell, err := excelize.CoordinatesToCellName(col, startRow)
			if err != nil {
				return fmt.Errorf(
					"format pivot table %d source header at column %d row %d: %w",
					pivotIndex+1,
					col,
					startRow,
					err,
				)
			}
			if err := ew.File.SetCellValue(sheetName, cell, getCellScalarValue(headerRow[headerIndex])); err != nil {
				return fmt.Errorf(
					"seed pivot table %d source header on sheet %q at cell %q: %w",
					pivotIndex+1,
					sheetName,
					cell,
					err,
				)
			}
		}
	}
	return nil
}

// streamWriter writes content to different sheets in the Excel file based on provided data.
//
// Args:
//
//	file (*excelize.File): The Excel file object.
//	data (map[string]interface{}): Map containing data for each sheet.
func (ew *ExcelWriter) performStreamWrite(
	sheet string,
	sheetData map[string]interface{},
) (*excelize.StreamWriter, error) {
	streamWriter, rowHeightMap, err := ew.prepareStreamWrite(sheet, sheetData)
	if err != nil {
		return nil, err
	}

	// Write Data
	startedRow := 1
	excelData := sheetData["Data"].([]interface{})
	if sheetData["NoStyle"] == false {
		for i, rowData := range excelData {
			for j, cellData := range rowData.([]interface{}) {
				if cellData == nil {
					continue
				}
				cell, err := createCell(cellData.([]interface{}), ew.StyleIDs)
				if err != nil {
					return nil, fmt.Errorf("sheet %q row %d column %d: %w", sheet, i+1, j+1, err)
				}
				excelData[i].([]interface{})[j] = cell
			}
		}
	}

	for i, rowData := range excelData {
		row := rowData.([]interface{})
		cell, _ := excelize.CoordinatesToCellName(1, i+startedRow)

		// Write cell with Height if rowHeightMap key found
		if rowHeight, ok := rowHeightMap[strconv.Itoa(i+startedRow)]; ok {
			if err := streamWriter.SetRow(cell, row, rowHeight); err != nil {
				return nil, fmt.Errorf("write stream sheet %q row %d: %w", sheet, i+startedRow, err)
			}
		} else {
			if err := streamWriter.SetRow(cell, row); err != nil {
				return nil, fmt.Errorf("write stream sheet %q row %d: %w", sheet, i+startedRow, err)
			}
		}
	}
	return streamWriter, nil
}

// prepareSheetFeatures applies worksheet features shared by both writer
// engines. Keep this order stable because some features must exist before row
// data is written.
func (ew *ExcelWriter) prepareSheetFeatures(
	sheet string,
	sheetData map[string]interface{},
) error {
	// Add Chart
	if err := ew.addChart(sheet, sheetData["Chart"].([]interface{})); err != nil {
		return err
	}

	// Set DataValidations
	if err := ew.setDataValidation(sheet, sheetData["DataValidation"].([]interface{})); err != nil {
		return err
	}

	// Add Comment
	if err := ew.addComment(sheet, sheetData["Comment"].([]interface{})); err != nil {
		return err
	}

	// Set Panes
	panes := ew.Content[sheet].(map[string]interface{})["Panes"].(map[string]interface{})
	if err := ew.setPanes(sheet, panes); err != nil {
		return err
	}

	// Set AutoFilters
	autoFilters := ew.Content[sheet].(map[string]interface{})["AutoFilter"].([]interface{})
	if err := ew.setAutoFilter(sheet, autoFilters); err != nil {
		return err
	}

	return nil
}

func (ew *ExcelWriter) prepareStreamWrite(
	sheet string,
	sheetData map[string]interface{},
) (*excelize.StreamWriter, map[string]excelize.RowOpts, error) {
	if err := ew.prepareSheetFeatures(sheet, sheetData); err != nil {
		return nil, nil, err
	}

	streamWriter, err := ew.File.NewStreamWriter(sheet)
	if err != nil {
		return nil, nil, fmt.Errorf("create stream writer for sheet %q: %w", sheet, err)
	}

	// CellWidtrh should be set before SetRow
	// Height should be set with SetRow in StreamWriter
	if err := setCellWidth(streamWriter, sheetData); err != nil {
		return nil, nil, fmt.Errorf("set widths on stream sheet %q: %w", sheet, err)
	}
	rowHeightMap := getRowHeightMap(sheetData)

	if err := mergeCell(streamWriter, sheetData["MergeCells"].([]interface{})); err != nil {
		return nil, nil, fmt.Errorf("merge cells on stream sheet %q: %w", sheet, err)
	}
	return streamWriter, rowHeightMap, nil
}

// normalWriter writes content to different sheets in the Excel file based on provided data.
//
// Args:
//
//	file (*excelize.File): The Excel file object.
//	data (map[string]interface{}): Map containing data for each sheet.
func (ew *ExcelWriter) performNormalWrite(sheet string, sheetData map[string]interface{}) error {
	if err := ew.prepareNormalWrite(sheet, sheetData); err != nil {
		return err
	}

	excelData := sheetData["Data"].([]interface{})
	for i, rowData := range excelData {
		row := rowData.([]interface{})
		if noStyle, _ := sheetData["NoStyle"].(bool); !noStyle {
			for column, item := range row {
				if item == nil {
					continue
				}
				cell, err := createCell(item.([]interface{}), ew.StyleIDs)
				if err != nil {
					return fmt.Errorf("sheet %q row %d column %d: %w", sheet, i+1, column+1, err)
				}
				row[column] = cell
			}
		}
		if err := ew.writeDecodedNormalRow(sheet, i+1, row); err != nil {
			return err
		}
	}
	return nil
}

func (ew *ExcelWriter) prepareNormalWrite(sheet string, sheetData map[string]interface{}) error {
	if err := ew.prepareSheetFeatures(sheet, sheetData); err != nil {
		return err
	}

	// Set Cell Width and Height
	if err := ew.setCellWidthNormalWriter(sheet, sheetData); err != nil {
		return err
	}
	if err := ew.setCellHeightNormalWriter(sheet, sheetData); err != nil {
		return err
	}

	// Merge Cell
	if err := ew.mergeCellNormalWriter(sheet, sheetData["MergeCells"].([]interface{})); err != nil {
		return err
	}

	// Group col and row
	if sheetData["GroupedRow"] != nil {
		if err := ew.groupRow(sheet, sheetData["GroupedRow"].([]interface{})); err != nil {
			return err
		}
	}
	if sheetData["GroupedCol"] != nil {
		if err := ew.groupCol(sheet, sheetData["GroupedCol"].([]interface{})); err != nil {
			return err
		}
	}
	return nil
}
