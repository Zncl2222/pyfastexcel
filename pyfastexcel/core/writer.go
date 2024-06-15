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
	if len(strJson["protection"].(map[string]interface{})) != 0 {
		setProtection(file, strJson["protection"].(map[string]interface{}))
	}
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

// setProtection protect workbook with password
//
// Args:
//
//	file (*excelize.File): The Excel file object.
//	config (map[string]interface{}): Map containing key-value pairs for protection properties.
func setProtection(file *excelize.File, config map[string]interface{}) {
	err := file.ProtectWorkbook(&excelize.WorkbookProtectionOptions{
		AlgorithmName: config["algorithm"].(string),
		Password:      config["password"].(string),
		LockStructure: config["lock_structure"].(bool),
		LockWindows:   config["lock_windows"].(bool),
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
//	sw (excelize.StreamWriter): The StreamWriter to use for merging cells.
//	cell ([]interface{}): A slice of cell ranges to merge, where each cell range is
//	    represented as a pair of strings (top-left and bottom-right cells).
func mergeCell(sw *excelize.StreamWriter, cell []interface{}) {
	for _, col := range cell {
		cellRange := col.([]interface{})
		topLeft := cellRange[0].(string)
		bottomRight := cellRange[1].(string)
		sw.MergeCell(topLeft, bottomRight)
	}
}

// setAutoFilter applies an auto filter to a specific sheet in an Excel file using the provided Excelize file.
//
// Args:
//
//	file (*excelize.File): The Excelize file.
//	sheet (string): The name of the sheet to apply the auto filter.
//	autoFilters ([]interface{}): A slice of cell ranges where the auto filter will be applied.
func setAutoFilter(file *excelize.File, sheet string, autoFilters []interface{}) {
	for _, filter := range autoFilters {
		file.AutoFilter(sheet, filter.(string), []excelize.AutoFilterOptions{})
	}
}

// writeContentBySheet writes content to different sheets in the Excel file based on provided data.
//
// Args:
//
//	file (*excelize.File): The Excel file object.
//	data (map[string]interface{}): Map containing data for each sheet.
func writeContentBySheet(file *excelize.File, data map[string]interface{}) {
	sheetCount := 1
	hasSheet1 := false
	for s := range data {
		if s == "Sheet1" {
			hasSheet1 = true
		}
	}
	for sheet := range data {
		sheetData := data[sheet].(map[string]interface{})
		// Create Sheet and Wrtie Header
		if !hasSheet1 && sheetCount == 1 {
			file.SetSheetName("Sheet1", sheet)
			hasSheet1 = true
		} else {
			file.NewSheet(sheet)
			sheetCount++
		}

		// Set AutoFilters
		autoFilters := data[sheet].(map[string]interface{})["AutoFilter"].([]interface{})
		setAutoFilter(file, sheet, autoFilters)

		streamWriter, _ := file.NewStreamWriter(sheet)

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

		if err := streamWriter.Flush(); err != nil {
			fmt.Println(err)
		}
	}
}
