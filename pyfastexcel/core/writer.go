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

	if strJson["engine"] == "normalWriter" {
		normalWriter(file, strJson["content"].(map[string]interface{}))
	} else {
		streamWriter(file, strJson["content"].(map[string]interface{}))
	}

	// Save data in buffer and encode binary data to base64
	buffer, _ := file.WriteToBuffer()
	byteResults := []byte(buffer.Bytes())
	encodedString := base64.StdEncoding.EncodeToString(byteResults)

	return encodedString
}

// streamCreateTable adds multiple tables to an Excel sheet using a StreamWriter.
//
// This function takes a StreamWriter and a list of tables, each represented by a map of key-value pairs.
// It iterates over the list of tables and adds each one to the sheet using the StreamWriter's AddTable method.
// If an error occurs while adding a table, it prints the error.
//
// Args:
//
//	sw (*excelize.StreamWriter): The StreamWriter object used to write data to the Excel sheet.
//	tables ([]interface{}): A slice of maps where each map represents a table's properties.
//	  The map should contain the following keys:
//	  - "range" (string): The cell range for the table (e.g., "A1:C10").
//	  - "name" (string): The name of the table.
//	  - "style_name" (string): The style name for the table.
//	  - "show_first_column" (bool): Whether to highlight the first column.
//	  - "show_last_column" (bool): Whether to highlight the last column.
//	  - "show_row_stripes" (bool): Whether to display row stripes for better readability.
//	  - "show_column_stripes" (bool): Whether to display column stripes for better readability.
//
// Example:
//
//	tables := []interface{}{
//	  map[string]interface{}{
//	    "range": "A1:C10", "name": "Table1", "style_name": "TableStyleMedium9",
//	    "show_first_column": true, "show_last_column": false,
//	    "show_row_stripes": true, "show_column_stripes": false,
//	  },
//	  // Add more tables as needed
//	}
//	streamCreateTable(sw, tables)
func streamCreateTable(sw *excelize.StreamWriter, tables []interface{}) {
	for _, table := range tables {
		t := table.(map[string]interface{})
		showRowStripes := t["show_row_stripes"].(bool)
		err := sw.AddTable(&excelize.Table{
			Range:             t["range"].(string),
			Name:              t["name"].(string),
			StyleName:         t["style_name"].(string),
			ShowFirstColumn:   t["show_first_column"].(bool),
			ShowLastColumn:    t["show_last_column"].(bool),
			ShowRowStripes:    &showRowStripes,
			ShowColumnStripes: t["show_column_stripes"].(bool),
		})

		if err != nil {
			fmt.Println(err)
		}
	}
}

// createTable adds multiple tables to a specified sheet in an Excel file.
//
// This function takes an Excel file object, a sheet name, and a list of tables.
// Each table is represented by a map of key-value pairs defining its properties.
// It iterates over the list of tables and adds each one to the specified sheet using the file's AddTable method.
// If an error occurs while adding a table, the function prints the error.
//
// Args:
//
//	file (*excelize.File): The Excel file object to which tables will be added.
//	sheet (string): The name of the sheet in which to create the tables.
//	tables ([]interface{}): A slice of maps where each map represents a table's properties.
//	  The map should contain the following keys:
//	  - "range" (string): The cell range for the table (e.g., "A1:C10").
//	  - "name" (string): The name of the table.
//	  - "style_name" (string): The style name for the table.
//	  - "show_first_column" (bool): Whether to highlight the first column.
//	  - "show_last_column" (bool): Whether to highlight the last column.
//	  - "show_row_stripes" (bool): Whether to display row stripes for better readability.
//	  - "show_column_stripes" (bool): Whether to display column stripes for better readability.
//
// Example:
//
//	tables := []interface{}{
//	  map[string]interface{}{
//	    "range": "A1:C10", "name": "Table1", "style_name": "TableStyleMedium9",
//	    "show_first_column": true, "show_last_column": false,
//	    "show_row_stripes": true, "show_column_stripes": false,
//	  },
//	  // Add more tables as needed
//	}
//	createTable(file, "Sheet1", tables)
func createTable(file *excelize.File, sheet string, tables []interface{}) {
	for _, table := range tables {
		t := table.(map[string]interface{})
		showRowStripes := t["show_row_stripes"].(bool)
		err := file.AddTable(sheet, &excelize.Table{
			Range:             t["range"].(string),
			Name:              t["name"].(string),
			StyleName:         t["style_name"].(string),
			ShowFirstColumn:   t["show_first_column"].(bool),
			ShowLastColumn:    t["show_last_column"].(bool),
			ShowRowStripes:    &showRowStripes,
			ShowColumnStripes: t["show_column_stripes"].(bool),
		})

		if err != nil {
			fmt.Println(err)
		}
	}
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

// setPanes configures the pane settings for a specific sheet in an Excel file using the provided Excelize file.
//
// Args:
//
//	file (*excelize.File): The Excelize file.
//	sheet (string): The name of the sheet to configure the panes.
//	panes (map[string]interface{}): A map containing the pane settings, including freeze, split, x_split, y_split, top_left_cell, active_pane, and selection.
func setPanes(file *excelize.File, sheet string, panes map[string]interface{}) {
	if len(panes) != 0 {
		var selection []excelize.Selection
		for _, val := range panes["selection"].([]interface{}) {
			selection = append(selection, excelize.Selection{
				SQRef:      val.(map[string]interface{})["sq_ref"].(string),
				ActiveCell: val.(map[string]interface{})["active_cell"].(string),
				Pane:       val.(map[string]interface{})["pane"].(string),
			})
		}
		file.SetPanes(sheet, &excelize.Panes{
			Freeze:      panes["freeze"].(bool),
			Split:       panes["split"].(bool),
			XSplit:      int(panes["x_split"].(float64)),
			YSplit:      int(panes["y_split"].(float64)),
			TopLeftCell: panes["top_left_cell"].(string),
			ActivePane:  panes["active_pane"].(string),
			Selection:   selection,
		})
	}
}

// setDataValidation configures the data validation rules for a specific sheet in an Excel file using the provided Excelize file.
//
// Args:
//
//	file (*excelize.File): The Excelize file.
//	sheet (string): The name of the sheet to configure the data validation rules.
//	validation ([]interface{}): A slice containing the data validation settings,
//		each represented as a map with keys like sq_ref, set_range_start, set_range_stop,
//		input_title, input_body, error_title, error_body, drop_list, and sqref_drop_list.
func setDataValidation(file *excelize.File, sheet string, validation []interface{}) {
	dv := excelize.NewDataValidation(true)
	for _, v := range validation {
		dv.SetSqref(v.(map[string]interface{})["sq_ref"].(string))

		_, setRangeStart := v.(map[string]interface{})["set_range_start"]
		_, setRangeStope := v.(map[string]interface{})["set_range_stop"]
		if setRangeStart && setRangeStope {
			dv.SetRange(
				v.(map[string]interface{})["set_range_start"],
				v.(map[string]interface{})["set_range_stop"],
				excelize.DataValidationTypeWhole,
				excelize.DataValidationOperatorBetween,
			)
		}

		_, inputTitle := v.(map[string]interface{})["input_title"]
		_, inputBody := v.(map[string]interface{})["input_body"]
		if inputTitle && inputBody {
			dv.SetInput(
				v.(map[string]interface{})["input_title"].(string),
				v.(map[string]interface{})["input_body"].(string),
			)
		}

		_, errorTitle := v.(map[string]interface{})["error_title"]
		_, errorBody := v.(map[string]interface{})["error_body"]
		if errorTitle && errorBody {
			dv.SetError(
				excelize.DataValidationErrorStyleStop,
				v.(map[string]interface{})["error_title"].(string),
				v.(map[string]interface{})["error_body"].(string),
			)
		}

		if _, ok := v.(map[string]interface{})["drop_list"]; ok {
			dropList := make([]string, len(v.(map[string]interface{})["drop_list"].([]interface{})))
			for _, dropItem := range v.(map[string]interface{})["drop_list"].([]interface{}) {
				dropList = append(dropList, dropItem.(string))
			}
			dv.SetDropList(dropList)
		}
		if _, ok := v.(map[string]interface{})["sqref_drop_list"]; ok {
			dv.SetSqrefDropList(v.(map[string]interface{})["sqref_drop_list"].(string))
		}

		file.AddDataValidation(sheet, dv)
	}
}

// addComment adds comments to specific cells in an Excel file using the provided Excelize file.
//
// Args:
//
// file (*excelize.File): The Excelize file.
// sheet (string): The name of the sheet to add the comments.
// comment ([]interface{}): An array containing the comment data, including the cell, author, and paragraph.
func addComment(file *excelize.File, sheet string, comment []interface{}) {
	for _, c := range comment {
		paragraph := make([]excelize.RichTextRun, 0)
		commentData := c.(map[string]interface{})
		for _, p := range commentData["paragraph"].([]interface{}) {
			fontStyle := getFontStyle(p.(map[string]interface{}))
			paragraph = append(
				paragraph,
				excelize.RichTextRun{
					Text: p.(map[string]interface{})["text"].(string),
					Font: fontStyle,
				},
			)
		}
		file.AddComment(sheet, excelize.Comment{
			Cell:      commentData["cell"].(string),
			Author:    commentData["author"].(string),
			Paragraph: paragraph,
		},
		)
	}
}

// streamWriter writes content to different sheets in the Excel file based on provided data.
//
// Args:
//
//	file (*excelize.File): The Excel file object.
//	data (map[string]interface{}): Map containing data for each sheet.
func streamWriter(file *excelize.File, data map[string]interface{}) {
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

		// Add Chart
		addChart(file, sheet, sheetData["Chart"].([]interface{}))

		// Set DataValidations
		setDataValidation(file, sheet, sheetData["DataValidation"].([]interface{}))

		// Add Comment
		addComment(file, sheet, sheetData["Comment"].([]interface{}))

		// Set Panes
		panes := data[sheet].(map[string]interface{})["Panes"].(map[string]interface{})
		setPanes(file, sheet, panes)

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

		// Create Stream Table
		// Excelize should create table with the existed row.
		streamCreateTable(streamWriter, sheetData["Table"].([]interface{}))

		if err := streamWriter.Flush(); err != nil {
			fmt.Println(err)
		}
	}
}

// groupRow groups rows in an Excel worksheet using the provided file.
//
// Args:
//
//	file (*excelize.File): The excelize file.
//	sheet (string): The name of the worksheet.
//	group ([]interface{}): A slice of row groups, where each group is represented
//	    as a map containing "start_row" (float64), "end_row" (float64, optional),
//	    "outline_level" (float64), and "hidden" (bool).
func groupRow(file *excelize.File, sheet string, group []interface{}) {
	var endRow int
	for _, g := range group {
		startRow := int(g.(map[string]interface{})["start_row"].(float64))
		_, ok := g.(map[string]interface{})["end_row"].(float64)
		outlineLevel := uint8(g.(map[string]interface{})["outline_level"].(float64))
		hidden := g.(map[string]interface{})["hidden"].(bool)
		if !ok {
			endRow = startRow
		} else {
			endRow = int(g.(map[string]interface{})["end_row"].(float64))
		}
		for i := startRow; i <= endRow; i++ {
			file.SetRowOutlineLevel(sheet, i, outlineLevel)
			if hidden {
				file.SetRowVisible(sheet, i, false)
			}
		}
	}

}

// groupCol groups columns in an Excel worksheet using the provided file.
//
// Args:
//
//	file (*excelize.File): The excelize file.
//	sheet (string): The name of the worksheet.
//	group ([]interface{}): A slice of column groups, where each group is represented
//	    as a map containing "start_col" (string), "end_col" (string, optional),
//	    "outline_level" (float64), and "hidden" (bool).
func groupCol(file *excelize.File, sheet string, group []interface{}) {
	for _, g := range group {
		startCol := g.(map[string]interface{})["start_col"].(string)
		endCol, ok := g.(map[string]interface{})["end_col"].(string)
		outlineLevel := uint8(g.(map[string]interface{})["outline_level"].(float64))
		hidden := g.(map[string]interface{})["hidden"].(bool)
		if !ok {
			endCol = startCol
		}
		startColNum, _ := excelize.ColumnNameToNumber(startCol)
		endColNum, _ := excelize.ColumnNameToNumber(endCol)
		for i := startColNum; i <= endColNum; i++ {
			col, _ := excelize.ColumnNumberToName(i)
			file.SetColOutlineLevel(sheet, col, outlineLevel)
		}
		file.SetColVisible(sheet, startCol+":"+endCol, !hidden)
	}
}

// mergeCellNormalWriter merges cells in an Excel worksheet using the provided file.
//
// Args:
//
//	file (*excelize.File): The excelize file.
//	cell ([]interface{}): A slice of cell ranges to merge, where each cell range is
//	    represented as a pair of strings (top-left and bottom-right cells).
func mergeCellNormalWriter(file *excelize.File, sheet string, cell []interface{}) {
	for _, col := range cell {
		cellRange := col.([]interface{})
		topLeft := cellRange[0].(string)
		bottomRight := cellRange[1].(string)
		file.MergeCell(sheet, topLeft, bottomRight)
	}
}

// setCellWidthNormalWriter sets the width of columns in an Excel worksheet using the provided file.
//
// Args:
//
//	file (*excelize.File): The excelize file.
//	sheet (string): The name of the worksheet.
//	config (map[string]interface{}): A map containing column width configurations, where the key is the column
//	    index as a string and the value is the width as a float64.
func setCellWidthNormalWriter(file *excelize.File, sheet string, config map[string]interface{}) {
	if config["Width"] == nil {
		return
	}
	if width := config["Width"].(map[string]interface{}); width != nil {
		for col := range width {
			colIndex, _ := strconv.Atoi(col)
			colName, _ := excelize.ColumnNumberToName(colIndex)
			file.SetColWidth(sheet, colName, colName, width[col].(float64))
		}
	}
}

// setCellHeightNormalWriter sets the height of rows in an Excel worksheet using the provided file.
//
// Args:
//
//	file (*excelize.File): The excelize file.
//	sheet (string): The name of the worksheet.
//	config (map[string]interface{}): A map containing row height configurations, where the key is the row
//	    index as a string and the value is the height as a float64.
func setCellHeightNormalWriter(file *excelize.File, sheet string, config map[string]interface{}) {
	if config["Height"] == nil {
		return
	}
	if height := config["Height"].(map[string]interface{}); height != nil {
		for row := range height {
			rowIndex, _ := strconv.Atoi(row)
			file.SetRowHeight(sheet, rowIndex, height[row].(float64))
		}
	}
}

// normalWriter writes content to different sheets in the Excel file based on provided data.
//
// Args:
//
//	file (*excelize.File): The Excel file object.
//	data (map[string]interface{}): Map containing data for each sheet.
func normalWriter(file *excelize.File, data map[string]interface{}) {
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

		addChart(file, sheet, sheetData["Chart"].([]interface{}))

		// Set DataValidations
		setDataValidation(file, sheet, sheetData["DataValidation"].([]interface{}))

		// Add Comment
		addComment(file, sheet, sheetData["Comment"].([]interface{}))

		// Set Panes
		panes := data[sheet].(map[string]interface{})["Panes"].(map[string]interface{})
		setPanes(file, sheet, panes)

		// Set AutoFilters
		autoFilters := data[sheet].(map[string]interface{})["AutoFilter"].([]interface{})
		setAutoFilter(file, sheet, autoFilters)

		// Set Cell Width and Height
		setCellWidthNormalWriter(file, sheet, sheetData)
		setCellHeightNormalWriter(file, sheet, sheetData)

		// Merge Cell
		mergeCellNormalWriter(file, sheet, sheetData["MergeCells"].([]interface{}))

		// Group col and row
		if sheetData["GroupedRow"] != nil {
			groupRow(file, sheet, sheetData["GroupedRow"].([]interface{}))
		}
		if sheetData["GroupedCol"] != nil {
			groupCol(file, sheet, sheetData["GroupedCol"].([]interface{}))
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
					file.SetCellValue(sheet, colCell, "")
					file.SetCellStyle(sheet, colCell, colCell, styleMap["DEFAULT_STYLE"])
				} else {
					switch value := v[0].(type) {
					case string:
						if strings.HasPrefix(value, "=") {
							file.SetCellFormula(sheet, colCell, value)
						} else {
							file.SetCellValue(sheet, colCell, value)
						}
					default:
						file.SetCellValue(sheet, colCell, value)
					}
					file.SetCellStyle(sheet, colCell, colCell, styleMap[item.([]interface{})[1].(string)])
				}
			}
		}
		// Excelize should create table with the existed row.
		createTable(file, sheet, sheetData["Table"].([]interface{}))
	}
}
