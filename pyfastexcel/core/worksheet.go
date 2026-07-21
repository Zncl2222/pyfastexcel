package core

import (
	"fmt"
	"strconv"

	"github.com/xuri/excelize/v2"
)

// streamCreateTable adds multiple tables to an Excel sheet using a StreamWriter.
//
// This function takes a StreamWriter and a list of tables, each represented by a map of key-value pairs.
// It iterates over the list of tables and adds each one to the sheet using the StreamWriter's AddTable method.
// It returns the first error reported by the stream writer.
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
func streamCreateTable(sw *excelize.StreamWriter, tables []interface{}) error {
	for index, table := range tables {
		t := table.(map[string]interface{})
		showRowStripes := t["show_row_stripes"].(bool)
		if err := sw.AddTable(&excelize.Table{
			Range:             t["range"].(string),
			Name:              t["name"].(string),
			StyleName:         t["style_name"].(string),
			ShowFirstColumn:   t["show_first_column"].(bool),
			ShowLastColumn:    t["show_last_column"].(bool),
			ShowRowStripes:    &showRowStripes,
			ShowColumnStripes: t["show_column_stripes"].(bool),
		}); err != nil {
			return fmt.Errorf(
				"add stream table %d %q over range %q: %w",
				index+1,
				t["name"],
				t["range"],
				err,
			)
		}
	}
	return nil
}

// createTable adds multiple tables to a specified sheet in an Excel file.
//
// This function takes an Excel file object, a sheet name, and a list of tables.
// Each table is represented by a map of key-value pairs defining its properties.
// It iterates over the list of tables and adds each one to the specified sheet using the file's AddTable method.
// It returns the first error reported by excelize.
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
func (ew *ExcelWriter) createTable(sheet string, tables []interface{}) error {
	for index, table := range tables {
		t := table.(map[string]interface{})
		showRowStripes := t["show_row_stripes"].(bool)
		if err := ew.File.AddTable(sheet, &excelize.Table{
			Range:             t["range"].(string),
			Name:              t["name"].(string),
			StyleName:         t["style_name"].(string),
			ShowFirstColumn:   t["show_first_column"].(bool),
			ShowLastColumn:    t["show_last_column"].(bool),
			ShowRowStripes:    &showRowStripes,
			ShowColumnStripes: t["show_column_stripes"].(bool),
		}); err != nil {
			return fmt.Errorf(
				"add table %d %q to sheet %q over range %q: %w",
				index+1,
				t["name"],
				sheet,
				t["range"],
				err,
			)
		}
	}
	return nil
}

// setFileProps sets document properties of the Excel file based on a map of key-value pairs.
//
// Args:
//
//	file (*excelize.File): The Excel file object.
//	config (map[string]interface{}): Map containing key-value pairs for document properties.
func (ew *ExcelWriter) setFileProps(config map[string]interface{}) error {
	if err := ew.File.SetDocProps(&excelize.DocProperties{
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
	}); err != nil {
		return fmt.Errorf("set document properties: %w", err)
	}
	return nil
}

// setProtection protect workbook with password
//
// Args:
//
//	file (*excelize.File): The Excel file object.
//	config (map[string]interface{}): Map containing key-value pairs for protection properties.
func (ew *ExcelWriter) setProtection(config map[string]interface{}) error {
	algorithm := config["algorithm"].(string)
	// XOR is part of the existing Python API, but excelize v2.9 rejects it
	// when hashing a non-empty workbook password. Preserve that public input
	// while emitting a supported modern workbook-protection hash.
	if algorithm == "XOR" {
		algorithm = "SHA-512"
	}
	if err := ew.File.ProtectWorkbook(&excelize.WorkbookProtectionOptions{
		AlgorithmName: algorithm,
		Password:      config["password"].(string),
		LockStructure: config["lock_structure"].(bool),
		LockWindows:   config["lock_windows"].(bool),
	}); err != nil {
		return fmt.Errorf("protect workbook: %w", err)
	}
	return nil
}

// setCellWidthAndHeight sets the width and height of cells in the Excel file based on a map of key-value pairs.
//
// Args:
//
//	file (*excelize.File): The Excel file object.
//	config (map[string]interface{}): Map containing key-value pairs for cell width and height.
func setCellWidth(streamWriter *excelize.StreamWriter, config map[string]interface{}) error {
	if config["Width"] == nil {
		return nil
	}
	width := config["Width"].(map[string]interface{})
	for col := range width {
		columnIndex, err := strconv.Atoi(col)
		if err != nil {
			return fmt.Errorf("parse stream column index %q: %w", col, err)
		}
		if err := streamWriter.SetColWidth(columnIndex, columnIndex, width[col].(float64)); err != nil {
			return fmt.Errorf("set stream column %d width: %w", columnIndex, err)
		}
	}
	return nil
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
func mergeCell(sw *excelize.StreamWriter, cell []interface{}) error {
	for index, col := range cell {
		cellRange := col.([]interface{})
		topLeft := cellRange[0].(string)
		bottomRight := cellRange[1].(string)
		if err := sw.MergeCell(topLeft, bottomRight); err != nil {
			return fmt.Errorf("merge stream range %d %q:%q: %w", index+1, topLeft, bottomRight, err)
		}
	}
	return nil
}

// setAutoFilter applies an auto filter to a specific sheet in an Excel file using the provided Excelize file.
//
// Args:
//
//	file (*excelize.File): The Excelize file.
//	sheet (string): The name of the sheet to apply the auto filter.
//	autoFilters ([]interface{}): A slice of cell ranges where the auto filter will be applied.
func (ew *ExcelWriter) setAutoFilter(sheet string, autoFilters []interface{}) error {
	for index, filter := range autoFilters {
		cellRange := filter.(string)
		if err := ew.File.AutoFilter(sheet, cellRange, []excelize.AutoFilterOptions{}); err != nil {
			return fmt.Errorf("set auto filter %d on sheet %q over range %q: %w", index+1, sheet, cellRange, err)
		}
	}
	return nil
}

// setPanes configures the pane settings for a specific sheet in an Excel file using the provided Excelize file.
//
// Args:
//
//	file (*excelize.File): The Excelize file.
//	sheet (string): The name of the sheet to configure the panes.
//	panes (map[string]interface{}): A map containing the pane settings, including freeze, split, x_split, y_split, top_left_cell, active_pane, and selection.
func (ew *ExcelWriter) setPanes(sheet string, panes map[string]interface{}) error {
	if len(panes) != 0 {
		var selection []excelize.Selection
		for _, val := range panes["selection"].([]interface{}) {
			selection = append(selection, excelize.Selection{
				SQRef:      val.(map[string]interface{})["sq_ref"].(string),
				ActiveCell: val.(map[string]interface{})["active_cell"].(string),
				Pane:       val.(map[string]interface{})["pane"].(string),
			})
		}
		if err := ew.File.SetPanes(sheet, &excelize.Panes{
			Freeze:      panes["freeze"].(bool),
			Split:       panes["split"].(bool),
			XSplit:      int(panes["x_split"].(float64)),
			YSplit:      int(panes["y_split"].(float64)),
			TopLeftCell: panes["top_left_cell"].(string),
			ActivePane:  panes["active_pane"].(string),
			Selection:   selection,
		}); err != nil {
			return fmt.Errorf("set panes on sheet %q: %w", sheet, err)
		}
	}
	return nil
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
func (ew *ExcelWriter) setDataValidation(sheet string, validation []interface{}) error {
	for index, v := range validation {
		dv := excelize.NewDataValidation(true)
		validationData := v.(map[string]interface{})
		dv.SetSqref(validationData["sq_ref"].(string))

		_, setRangeStart := validationData["set_range_start"]
		_, setRangeStope := validationData["set_range_stop"]
		if setRangeStart && setRangeStope {
			if err := dv.SetRange(
				validationData["set_range_start"],
				validationData["set_range_stop"],
				excelize.DataValidationTypeWhole,
				excelize.DataValidationOperatorBetween,
			); err != nil {
				return fmt.Errorf("configure data validation %d range on sheet %q: %w", index+1, sheet, err)
			}
		}

		_, inputTitle := validationData["input_title"]
		_, inputBody := validationData["input_body"]
		if inputTitle && inputBody {
			dv.SetInput(
				validationData["input_title"].(string),
				validationData["input_body"].(string),
			)
		}

		_, errorTitle := validationData["error_title"]
		_, errorBody := validationData["error_body"]
		if errorTitle && errorBody {
			dv.SetError(
				excelize.DataValidationErrorStyleStop,
				validationData["error_title"].(string),
				validationData["error_body"].(string),
			)
		}

		if rawDropList, ok := validationData["drop_list"].([]interface{}); ok {
			dropList := make([]string, 0, len(rawDropList))
			for _, dropItem := range rawDropList {
				dropList = append(dropList, dropItem.(string))
			}
			if err := dv.SetDropList(dropList); err != nil {
				return fmt.Errorf("configure data validation %d drop list on sheet %q: %w", index+1, sheet, err)
			}
		}
		if sqrefDropList, ok := validationData["sqref_drop_list"].(string); ok {
			dv.SetSqrefDropList(sqrefDropList)
		}

		if err := ew.File.AddDataValidation(sheet, dv); err != nil {
			return fmt.Errorf("add data validation %d to sheet %q: %w", index+1, sheet, err)
		}
	}
	return nil
}

// addComment adds comments to specific cells in an Excel file using the provided Excelize file.
//
// Args:
//
// file (*excelize.File): The Excelize file.
// sheet (string): The name of the sheet to add the comments.
// comment ([]interface{}): An array containing the comment data, including the cell, author, and paragraph.
func (ew *ExcelWriter) addComment(sheet string, comment []interface{}) error {
	for index, c := range comment {
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
		if err := ew.File.AddComment(sheet, excelize.Comment{
			Cell:      commentData["cell"].(string),
			Author:    commentData["author"].(string),
			Paragraph: paragraph,
		}); err != nil {
			return fmt.Errorf(
				"add comment %d to sheet %q at cell %q: %w",
				index+1,
				sheet,
				commentData["cell"],
				err,
			)
		}
	}
	return nil
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
func (ew *ExcelWriter) groupRow(sheet string, group []interface{}) error {
	var endRow int
	for groupIndex, g := range group {
		startRow := int(g.(map[string]interface{})["start_row"].(float64))
		_, ok := g.(map[string]interface{})["end_row"].(float64)
		outlineLevel := normalizeOutlineLevel(g.(map[string]interface{})["outline_level"].(float64))
		hidden := g.(map[string]interface{})["hidden"].(bool)
		if !ok {
			endRow = startRow
		} else {
			endRow = int(g.(map[string]interface{})["end_row"].(float64))
		}
		for i := startRow; i <= endRow; i++ {
			if err := ew.File.SetRowOutlineLevel(sheet, i, outlineLevel); err != nil {
				return fmt.Errorf("set row group %d outline on sheet %q row %d: %w", groupIndex+1, sheet, i, err)
			}
			if hidden {
				if err := ew.File.SetRowVisible(sheet, i, false); err != nil {
					return fmt.Errorf("hide row group %d on sheet %q row %d: %w", groupIndex+1, sheet, i, err)
				}
			}
		}
	}
	return nil
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
func (ew *ExcelWriter) groupCol(sheet string, group []interface{}) error {
	for groupIndex, g := range group {
		startCol := g.(map[string]interface{})["start_col"].(string)
		endCol, ok := g.(map[string]interface{})["end_col"].(string)
		outlineLevel := normalizeOutlineLevel(g.(map[string]interface{})["outline_level"].(float64))
		hidden := g.(map[string]interface{})["hidden"].(bool)
		if !ok {
			endCol = startCol
		}
		startColNum, err := columnReferenceToNumber(startCol)
		if err != nil {
			return fmt.Errorf("parse column group %d start %q on sheet %q: %w", groupIndex+1, startCol, sheet, err)
		}
		endColNum, err := columnReferenceToNumber(endCol)
		if err != nil {
			return fmt.Errorf("parse column group %d end %q on sheet %q: %w", groupIndex+1, endCol, sheet, err)
		}
		startColName, err := excelize.ColumnNumberToName(startColNum)
		if err != nil {
			return fmt.Errorf("format column group %d start index %d on sheet %q: %w", groupIndex+1, startColNum, sheet, err)
		}
		endColName, err := excelize.ColumnNumberToName(endColNum)
		if err != nil {
			return fmt.Errorf("format column group %d end index %d on sheet %q: %w", groupIndex+1, endColNum, sheet, err)
		}
		for i := startColNum; i <= endColNum; i++ {
			col, err := excelize.ColumnNumberToName(i)
			if err != nil {
				return fmt.Errorf("format column group %d index %d on sheet %q: %w", groupIndex+1, i, sheet, err)
			}
			if err := ew.File.SetColOutlineLevel(sheet, col, outlineLevel); err != nil {
				return fmt.Errorf("set column group %d outline on sheet %q column %q: %w", groupIndex+1, sheet, col, err)
			}
		}
		if err := ew.File.SetColVisible(sheet, startColName+":"+endColName, !hidden); err != nil {
			return fmt.Errorf("set column group %d visibility on sheet %q range %q:%q: %w", groupIndex+1, sheet, startColName, endColName, err)
		}
	}
	return nil
}

func normalizeOutlineLevel(level float64) uint8 {
	if level < 1 {
		return 1
	}
	if level > 7 {
		return 7
	}
	return uint8(level)
}

func columnReferenceToNumber(reference string) (int, error) {
	column, err := excelize.ColumnNameToNumber(reference)
	if err == nil {
		return column, nil
	}
	column, _, cellErr := excelize.CellNameToCoordinates(reference)
	if cellErr != nil {
		return 0, fmt.Errorf("not a column name or cell reference: %w", cellErr)
	}
	return column, nil
}

// mergeCellNormalWriter merges cells in an Excel worksheet using the provided file.
//
// Args:
//
//	file (*excelize.File): The excelize file.
//	cell ([]interface{}): A slice of cell ranges to merge, where each cell range is
//	    represented as a pair of strings (top-left and bottom-right cells).
func (ew *ExcelWriter) mergeCellNormalWriter(sheet string, cell []interface{}) error {
	for index, col := range cell {
		cellRange := col.([]interface{})
		topLeft := cellRange[0].(string)
		bottomRight := cellRange[1].(string)
		if err := ew.File.MergeCell(sheet, topLeft, bottomRight); err != nil {
			return fmt.Errorf("merge range %d %q:%q on sheet %q: %w", index+1, topLeft, bottomRight, sheet, err)
		}
	}
	return nil
}

// setCellWidthNormalWriter sets the width of columns in an Excel worksheet using the provided file.
//
// Args:
//
//	file (*excelize.File): The excelize file.
//	sheet (string): The name of the worksheet.
//	config (map[string]interface{}): A map containing column width configurations, where the key is the column
//	    index as a string and the value is the width as a float64.
func (ew *ExcelWriter) setCellWidthNormalWriter(sheet string, config map[string]interface{}) error {
	if config["Width"] == nil {
		return nil
	}
	if width := config["Width"].(map[string]interface{}); width != nil {
		for col := range width {
			colIndex, err := strconv.Atoi(col)
			if err != nil {
				return fmt.Errorf("parse column index %q for sheet %q: %w", col, sheet, err)
			}
			colName, err := excelize.ColumnNumberToName(colIndex)
			if err != nil {
				return fmt.Errorf("format column index %d for sheet %q: %w", colIndex, sheet, err)
			}
			if err := ew.File.SetColWidth(sheet, colName, colName, width[col].(float64)); err != nil {
				return fmt.Errorf("set column %q width on sheet %q: %w", colName, sheet, err)
			}
		}
	}
	return nil
}

// setCellHeightNormalWriter sets the height of rows in an Excel worksheet using the provided file.
//
// Args:
//
//	file (*excelize.File): The excelize file.
//	sheet (string): The name of the worksheet.
//	config (map[string]interface{}): A map containing row height configurations, where the key is the row
//	    index as a string and the value is the height as a float64.
func (ew *ExcelWriter) setCellHeightNormalWriter(sheet string, config map[string]interface{}) error {
	if config["Height"] == nil {
		return nil
	}
	if height := config["Height"].(map[string]interface{}); height != nil {
		for row := range height {
			rowIndex, err := strconv.Atoi(row)
			if err != nil {
				return fmt.Errorf("parse row index %q for sheet %q: %w", row, sheet, err)
			}
			if err := ew.File.SetRowHeight(sheet, rowIndex, height[row].(float64)); err != nil {
				return fmt.Errorf("set row %d height on sheet %q: %w", rowIndex, sheet, err)
			}
		}
	}
	return nil
}
