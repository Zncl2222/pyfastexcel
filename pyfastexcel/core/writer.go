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

type ExcelWriter struct {
	File       *excelize.File
	StyleMap   map[string]interface{}
	Content    map[string]interface{}
	FileProps  map[string]interface{}
	Protection map[string]interface{}
	SheetOrder []interface{}
	Engine     interface{}
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
	var StyleStruct StyleWrapper
	byteJson := []byte(data)

	strJson, err := marshmallow.Unmarshal(byteJson, &StyleStruct)
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
		Engine:     strJson["engine"],
	}
	return writer.writeExcel()
}

func (ew *ExcelWriter) writeExcel() string {
	styleMap = CreateStyle(ew.File, ew.StyleMap)
	ew.setFileProps(ew.FileProps)
	if len(ew.Protection) != 0 {
		ew.setProtection(ew.Protection)
	}

	if ew.Engine == "normalWriter" {
		ew.performNormalWrite()
	} else {
		ew.performStreamWrite()
	}

	// Save data in buffer and encode binary data to base64
	buffer, _ := ew.File.WriteToBuffer()
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
func (ew *ExcelWriter) createTable(sheet string, tables []interface{}) {
	for _, table := range tables {
		t := table.(map[string]interface{})
		showRowStripes := t["show_row_stripes"].(bool)
		err := ew.File.AddTable(sheet, &excelize.Table{
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
func (ew *ExcelWriter) setFileProps(config map[string]interface{}) {
	err := ew.File.SetDocProps(&excelize.DocProperties{
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
func (ew *ExcelWriter) setProtection(config map[string]interface{}) {
	err := ew.File.ProtectWorkbook(&excelize.WorkbookProtectionOptions{
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
func (ew *ExcelWriter) setAutoFilter(sheet string, autoFilters []interface{}) {
	for _, filter := range autoFilters {
		ew.File.AutoFilter(sheet, filter.(string), []excelize.AutoFilterOptions{})
	}
}

// setPanes configures the pane settings for a specific sheet in an Excel file using the provided Excelize file.
//
// Args:
//
//	file (*excelize.File): The Excelize file.
//	sheet (string): The name of the sheet to configure the panes.
//	panes (map[string]interface{}): A map containing the pane settings, including freeze, split, x_split, y_split, top_left_cell, active_pane, and selection.
func (ew *ExcelWriter) setPanes(sheet string, panes map[string]interface{}) {
	if len(panes) != 0 {
		var selection []excelize.Selection
		for _, val := range panes["selection"].([]interface{}) {
			selection = append(selection, excelize.Selection{
				SQRef:      val.(map[string]interface{})["sq_ref"].(string),
				ActiveCell: val.(map[string]interface{})["active_cell"].(string),
				Pane:       val.(map[string]interface{})["pane"].(string),
			})
		}
		ew.File.SetPanes(sheet, &excelize.Panes{
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
func (ew *ExcelWriter) setDataValidation(sheet string, validation []interface{}) {
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

		ew.File.AddDataValidation(sheet, dv)
	}
}

// addComment adds comments to specific cells in an Excel file using the provided Excelize file.
//
// Args:
//
// file (*excelize.File): The Excelize file.
// sheet (string): The name of the sheet to add the comments.
// comment ([]interface{}): An array containing the comment data, including the cell, author, and paragraph.
func (ew *ExcelWriter) addComment(sheet string, comment []interface{}) {
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
		ew.File.AddComment(sheet, excelize.Comment{
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
func (ew *ExcelWriter) performStreamWrite() {
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
		// Create Sheet and Wrtie Header
		if !hasSheet1 && sheetCount == 1 {
			ew.File.SetSheetName("Sheet1", sheet)
			hasSheet1 = true
		} else {
			ew.File.NewSheet(sheet)
			sheetCount++
		}

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

		// Create Stream Table
		// Excelize should create table with the existed row.
		streamCreateTable(streamWriter, sheetData["Table"].([]interface{}))

		if err := streamWriter.Flush(); err != nil {
			fmt.Println(err)
		}

		// To prevent the pivot table from being created before the data is written
		// we store the pivot table data in a list and create it after the data is written
		pivotTableList = append(pivotTableList, sheetData["PivotTable"].([]interface{}))

		// Set Sheet Visible
		if err := ew.File.SetSheetVisible(sheet, sheetData["SheetVisible"].(bool)); err != nil {
			fmt.Println(err)
		}
	}

	// Create Pivot Table. It should Create after the data is written
	for _, pivot := range pivotTableList {
		ew.createPivotTable(pivot)
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
func (ew *ExcelWriter) groupRow(sheet string, group []interface{}) {
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
			ew.File.SetRowOutlineLevel(sheet, i, outlineLevel)
			if hidden {
				ew.File.SetRowVisible(sheet, i, false)
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
func (ew *ExcelWriter) groupCol(sheet string, group []interface{}) {
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
			ew.File.SetColOutlineLevel(sheet, col, outlineLevel)
		}
		ew.File.SetColVisible(sheet, startCol+":"+endCol, !hidden)
	}
}

// mergeCellNormalWriter merges cells in an Excel worksheet using the provided file.
//
// Args:
//
//	file (*excelize.File): The excelize file.
//	cell ([]interface{}): A slice of cell ranges to merge, where each cell range is
//	    represented as a pair of strings (top-left and bottom-right cells).
func (ew *ExcelWriter) mergeCellNormalWriter(sheet string, cell []interface{}) {
	for _, col := range cell {
		cellRange := col.([]interface{})
		topLeft := cellRange[0].(string)
		bottomRight := cellRange[1].(string)
		ew.File.MergeCell(sheet, topLeft, bottomRight)
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
func (ew *ExcelWriter) setCellWidthNormalWriter(sheet string, config map[string]interface{}) {
	if config["Width"] == nil {
		return
	}
	if width := config["Width"].(map[string]interface{}); width != nil {
		for col := range width {
			colIndex, _ := strconv.Atoi(col)
			colName, _ := excelize.ColumnNumberToName(colIndex)
			ew.File.SetColWidth(sheet, colName, colName, width[col].(float64))
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
func (ew *ExcelWriter) setCellHeightNormalWriter(sheet string, config map[string]interface{}) {
	if config["Height"] == nil {
		return
	}
	if height := config["Height"].(map[string]interface{}); height != nil {
		for row := range height {
			rowIndex, _ := strconv.Atoi(row)
			ew.File.SetRowHeight(sheet, rowIndex, height[row].(float64))
		}
	}
}

// normalWriter writes content to different sheets in the Excel file based on provided data.
//
// Args:
//
//	file (*excelize.File): The Excel file object.
//	data (map[string]interface{}): Map containing data for each sheet.
func (ew *ExcelWriter) performNormalWrite() {
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
		// Create Sheet and Wrtie Header
		if !hasSheet1 && sheetCount == 1 {
			ew.File.SetSheetName("Sheet1", sheet)
			hasSheet1 = true
		} else {
			ew.File.NewSheet(sheet)
			sheetCount++
		}

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
					ew.File.SetCellValue(sheet, colCell, "")
					ew.File.SetCellStyle(sheet, colCell, colCell, styleMap["DEFAULT_STYLE"])
				} else {
					switch value := v[0].(type) {
					case string:
						if strings.HasPrefix(value, "=") {
							ew.File.SetCellFormula(sheet, colCell, value)
						} else {
							ew.File.SetCellValue(sheet, colCell, value)
						}
					default:
						ew.File.SetCellValue(sheet, colCell, value)
					}
					ew.File.SetCellStyle(sheet, colCell, colCell, styleMap[item.([]interface{})[1].(string)])
				}
			}
		}
		// Excelize should create table with the existed row.
		ew.createTable(sheet, sheetData["Table"].([]interface{}))

		// To prevent the pivot table from being created before the data is written
		// we store the pivot table data in a list and create it after the data is written
		pivotTableList = append(pivotTableList, sheetData["PivotTable"].([]interface{}))

		// Set Sheet Visible
		if err := ew.File.SetSheetVisible(sheet, sheetData["SheetVisible"].(bool)); err != nil {
			fmt.Println(err)
		}
	}

	// Create Pivot Table. It should Create after the data is written
	for _, pivot := range pivotTableList {
		ew.createPivotTable(pivot)
	}

}
