package core

// #include <stdlib.h>
import (
	"C"
)
import (
	"encoding/base64"
	"fmt"
	"strconv"

	"github.com/perimeterx/marshmallow"

	"github.com/xuri/excelize/v2"
)

var styleMap map[string]int

type (
	StyleWrapper struct {
		Style map[string]map[string]interface{} `json:"Style" binding:"required"`
	}
)

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
	var StyleStruct StyleWrapper
	byteJson := []byte(data)

	strJson, err := marshmallow.Unmarshal(byteJson, &StyleStruct)
	if err != nil {
		panic(err)
	}

	file := excelize.NewFile()
	styleMap = CreateStyle(file, strJson["style"].(map[string]interface{}))
	setFileProps(file, strJson["file_props"].(map[string]interface{}))
	writeContentBySheet(file, strJson["content"].(map[string]interface{}))

	// Save data in buffer and encode binary data to base64
	buffer, _ := file.WriteToBuffer()
	byteResults := []byte(buffer.Bytes())
	encodedString := base64.StdEncoding.EncodeToString(byteResults)

	return encodedString
}

// setFileProps sets document properties of the Excel file based on a map of key-value pairs.
//
// Args:
//
//	file (*excelize.File): The Excel file object.
//	config (map[string]interface{}): Map containing key-value pairs for document properties.
func setFileProps(file *excelize.File, config map[string]interface{}) {
	err := file.SetDocProps(&excelize.DocProperties{
		Category:       config["Category"].(string),
		ContentStatus:  config["ContentStatus"].(string),
		Created:        config["Created"].(string),
		Creator:        config["Creator"].(string),
		Description:    config["Description"].(string),
		Identifier:     config["Identifier"].(string),
		Keywords:       config["Keywords"].(string),
		LastModifiedBy: config["LastModifiedBy"].(string),
		Modified:       config["Modified"].(string),
		Revision:       config["Revision"].(string),
		Subject:        config["Subject"].(string),
		Title:          config["Title"].(string),
		Language:       config["Language"].(string),
		Version:        config["Version"].(string),
	})

	if err != nil {
		fmt.Println(err)
	}
}

// setCellWidthAndHeight sets the width and height of cells in the Excel file based on a map of key-value pairs.
//
// Args:
//
//	file (*excelize.File): The Excel file object.
//	config (map[string]interface{}): Map containing key-value pairs for cell width and height.
func setCellWidth(streamWriter *excelize.StreamWriter, config map[string]interface{}) {
	if config["Width"] == nil {
		return
	}
	width := config["Width"].(map[string]interface{})
	for col := range width {
		cidx, _ := strconv.Atoi(col)
		streamWriter.SetColWidth(cidx, cidx, width[col].(float64))
	}
}

// getRowHeightMap returns a map of row heights based on a map of key-value pairs.
//
// Args:
//
//	config (map[string]interface{}): Map containing key-value pairs for row heights.
//
// Returns:
//
//	map[string]excelize.RowOpts: Map of row heights.
func getRowHeightMap(config map[string]interface{}) map[string]excelize.RowOpts {
	rowHeightMap := make(map[string]excelize.RowOpts)

	if config["Height"] == nil {
		return rowHeightMap
	}
	height := config["Height"].(map[string]interface{})
	for row := range height {
		rowHeightMap[row] = excelize.RowOpts{Height: height[row].(float64), Hidden: false}
	}
	return rowHeightMap
}

// mergeCell merges cells in an Excel worksheet using the provided StreamWriter.
//
// Args:
//
//     sw (excelize.StreamWriter): The StreamWriter to use for merging cells.
//     cell ([]interface{}): A slice of cell ranges to merge, where each cell range is
//         represented as a pair of strings (top-left and bottom-right cells).
func mergeCell(sw *excelize.StreamWriter, cell []interface{}) {
	for _, col := range cell {
		cellRange := col.([]interface{})
		topLeft := cellRange[0].(string)
		bottomRight := cellRange[1].(string)
		sw.MergeCell(topLeft, bottomRight)
	}
}

// writeContentBySheet writes content to different sheets in the Excel file based on provided data.
//
// Args:
//
//	file (*excelize.File): The Excel file object.
//	data (map[string]interface{}): Map containing data for each sheet.
func writeContentBySheet(file *excelize.File, data map[string]interface{}) {
	for sheet := range data {
		sheetData := data[sheet].(map[string]interface{})
		// Create Sheet and Wrtie Header
		file.NewSheet(sheet)
		streamWriter, _ := file.NewStreamWriter(sheet)

		// CellWidtrh should be set before SetRow
		// Height should be set with SetRow in StreamWriter
		setCellWidth(streamWriter, sheetData)
		rowHeightMap := getRowHeightMap(sheetData)

		mergeCell(streamWriter, sheetData["MergeCells"].([]interface{}))

		startedRow := 1
		cell, _ := excelize.CoordinatesToCellName(1, startedRow)
		for i, _ := range sheetData["Header"].([]interface{}) {
			sheetData["Header"].([]interface{})[i] = createCell(sheetData["Header"].([]interface{})[i].([]interface{}))
		}

		if len(sheetData["Header"].([]interface{})) != 0 {
			if err := streamWriter.SetRow(cell, sheetData["Header"].([]interface{})); err != nil {
				fmt.Println(err)
			}
			startedRow += 1
		}

		// Write Data
		excelData := sheetData["Data"].([]interface{})
		for i, rowData := range excelData {
			row := rowData.([]interface{})

			processedRow := make([]interface{}, len(row))
			for j, cellData := range row {
				cell := cellData.([]interface{})
				processedRow[j] = createCell(cell)
			}

			cell, err := excelize.CoordinatesToCellName(1, i+startedRow)
			if err != nil {
				panic(err)
			}

			// Write cell with Height if rowHeightMap key found
			if rowHeight, ok := rowHeightMap[strconv.Itoa(i+startedRow)]; ok {
				if err := streamWriter.SetRow(cell, processedRow, rowHeight); err != nil {
					fmt.Println(err)
				}
			} else {
				if err := streamWriter.SetRow(cell, processedRow); err != nil {
					fmt.Println(err)
				}
			}

		}

		if err := streamWriter.Flush(); err != nil {
			fmt.Println(err)
		}
	}
}
