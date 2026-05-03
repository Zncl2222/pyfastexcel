package core

import (
	"archive/zip"
	"bytes"
	"encoding/base64"
	"encoding/json"
	"fmt"
	"strings"
	"testing"

	"github.com/xuri/excelize/v2"
)

var data map[string]interface{}
var dataNormalWriter map[string]interface{}

func init() {
	data = map[string]interface{}{
		"style": map[string]interface{}{
			"style1": map[string]interface{}{
				"Font": map[string]interface{}{
					"Bold": true,
				},
				"Fill": map[string]interface{}{
					"Type":    "pattern",
					"Color":   "#FFFFFF",
					"Pattern": 1,
					"Shading": 100,
				},
				"Border": map[string]interface{}{
					"left": map[string]interface{}{
						"Color": "FF0000",
						"Style": 1,
					},
					"top": map[string]interface{}{
						"Color": "00FF00",
						"Style": 2,
					},
				},
				"Alignment": map[string]interface{}{
					"Horizontal":      "center",
					"Vertical":        "middle",
					"Indent":          0,
					"JustifyLastLine": false,
					"ReadingOrder":    0,
					"RelativeIndent":  0,
					"ShrinkToFit":     false,
					"TextRotation":    0,
					"WrapText":        false,
				},
				"Protection": map[string]interface{}{
					"Hidden": true,
					"Locked": false,
				},
				"CustomNumFmt": "0.00",
			},
		},
		"file_props": map[string]interface{}{
			"Title":          "Test Excel File",
			"Creator":        "Test User",
			"Category":       "Test Category",
			"ContentStatus":  "Draft",
			"Description":    "Test Description",
			"Keywords":       "Test Keywords",
			"Language":       "en-US",
			"LastModifiedBy": "Test User",
			"Revision":       "1",
			"Subject":        "Test Subject",
			"Version":        "1.0",
			"Identifier":     "",
			"Created":        "",
			"Modified":       "",
		},
		"protection": map[string]interface{}{
			"algorithm":      "XOR",
			"password":       "12345",
			"lock_structure": true,
			"lock_windows":   false,
		},
		"sheet_order": []interface{}{"TestingSheet2", "Sheet2WithNoWidth", "Sheet3WithNoHeight", "Sheet4WithNoWidthAndHeight"},
		"content": map[string]interface{}{
			"TestingSheet2": map[string]interface{}{
				"Header": [][]string{
					{"Column1", "Column2", "Column3"},
				},
				"Data": [][][]string{
					{{"Data1", "style1"}, {"Data2", "style1"}, {"Data3", "style1"}},
					{{"Data4", "style1"}, {"Data5", "style1"}, {"Data6", "style1"}},
				},
				"Height":         map[string]int{"3": 252},
				"Width":          map[string]int{"1": 25, "2": 26, "3": 6},
				"MergeCells":     []interface{}{},
				"AutoFilter":     []interface{}{},
				"Panes":          map[string]interface{}{},
				"DataValidation": []interface{}{},
				"Comment": []interface{}{map[string]interface{}{
					"cell":      "A1",
					"author":    "author",
					"paragraph": []interface{}{map[string]interface{}{"text": "text", "bold": true}}}},
				"NoStyle":      false,
				"Table":        []interface{}{},
				"Chart":        []interface{}{},
				"PivotTable":   []interface{}{},
				"SheetVisible": true,
				"WriterEngine": "StreamWriter",
			},
			"Sheet2WithNoWidth": map[string]interface{}{
				"Header": [][]string{
					{"Column1", "Column2", "Column3"},
				},
				"Data": [][][]string{
					{{"Data1", "style1"}, {"Data2", "style1"}, {"Data3", "style1"}, {}},
					{{"Data4", "style1"}, {"Data5", "style1"}, {"Data6", "style1"}, {}},
				},
				"Height":     map[string]int{"3": 252},
				"MergeCells": [][]interface{}{{"A1", "A2"}, {"B2", "C3"}},
				"AutoFilter": []interface{}{},
				"Panes":      map[string]interface{}{},
				"DataValidation": []interface{}{map[string]interface{}{
					"sq_ref":    "A1",
					"set_range": "B1",
					"drop_list": []string{"123", "qwe"}}},
				"NoStyle":      false,
				"Comment":      []interface{}{},
				"Table":        []interface{}{},
				"Chart":        []interface{}{},
				"PivotTable":   []interface{}{},
				"SheetVisible": true,
				"WriterEngine": "StreamWriter",
			},
			"Sheet3WithNoHeight": map[string]interface{}{
				"Header": [][]string{
					{"Column1", "Column2", "Column3"},
				},
				"Data": [][][]string{
					{{"Data1", "style1"}, {"Data2", "style1"}, {"Data3", "style1"}},
					{{"Data4", "style1"}, {"Data5", "style1"}, {"Data6", "style1"}},
				},
				"Width":          map[string]int{"1": 25, "2": 26, "3": 6},
				"MergeCells":     []interface{}{},
				"AutoFilter":     []interface{}{},
				"Panes":          map[string]interface{}{},
				"DataValidation": []interface{}{map[string]interface{}{"sq_ref": "A1", "sqref_drop_list": "A1:B1"}},
				"Comment":        []interface{}{},
				"NoStyle":        false,
				"Table":          []interface{}{map[string]interface{}{"range": "A1:B3", "name": "test", "style_name": "", "show_first_column": true, "show_last_column": true, "show_row_stripes": false, "show_column_stripes": true}},
				"Chart":          []interface{}{},
				"PivotTable":     []interface{}{map[string]interface{}{"DataRange": "Sheet1$A1:C1", "PivotTableRange": "Sheet1$D1:F1", "ShowDrill": true, "Rows": []interface{}{}, "Filter": []interface{}{}, "Columns": []interface{}{}, "Data": []interface{}{}}},
				"SheetVisible":   false,
				"WriterEngine":   "StreamWriter",
			},
			"Sheet4WithNoWidthAndHeight": map[string]interface{}{
				"Header": [][]string{
					{"Column1", "Column2", "Column3"},
				},
				"Data": [][][]string{
					{{"Data1", "style1"}, {"Data2", "style1"}, {"Data3", "style1"}},
					{{"Data4", "style1"}, {"Data5", "style1"}, {"Data6", "style1"}},
				},
				"MergeCells": [][]interface{}{{"A1", "A2"}, {"B2", "C3"}},
				"AutoFilter": []interface{}{"A1:C1"},
				"Panes":      map[string]interface{}{},
				"DataValidation": []interface{}{map[string]interface{}{
					"sq_ref":      "A1",
					"error_title": "err_test",
					"error_body":  "err_body",
					"input_title": "input_test",
					"input_body":  "input_body"}},
				"Comment":      []interface{}{},
				"NoStyle":      false,
				"Table":        []interface{}{map[string]interface{}{"range": "A1:B3", "name": "test", "style_name": "", "show_first_column": true, "show_last_column": true, "show_row_stripes": false, "show_column_stripes": true}},
				"Chart":        []interface{}{},
				"PivotTable":   []interface{}{},
				"SheetVisible": true,
				"WriterEngine": "StreamWriter",
			},
		},
	}
	dataNormalWriter = map[string]interface{}{
		"style": map[string]interface{}{
			"style1": map[string]interface{}{
				"Font": map[string]interface{}{
					"Bold": true,
				},
				"Fill": map[string]interface{}{
					"Type":    "pattern",
					"Color":   "#FFFFFF",
					"Pattern": 1,
					"Shading": 100,
				},
				"Border": map[string]interface{}{
					"left": map[string]interface{}{
						"Color": "FF0000",
						"Style": 1,
					},
					"top": map[string]interface{}{
						"Color": "00FF00",
						"Style": 2,
					},
				},
				"Alignment": map[string]interface{}{
					"Horizontal":      "center",
					"Vertical":        "middle",
					"Indent":          0,
					"JustifyLastLine": false,
					"ReadingOrder":    0,
					"RelativeIndent":  0,
					"ShrinkToFit":     false,
					"TextRotation":    0,
					"WrapText":        false,
				},
				"Protection": map[string]interface{}{
					"Hidden": true,
					"Locked": false,
				},
				"CustomNumFmt": "0.00",
			},
		},
		"file_props": map[string]interface{}{
			"Title":          "Test Excel File",
			"Creator":        "Test User",
			"Category":       "Test Category",
			"ContentStatus":  "Draft",
			"Description":    "Test Description",
			"Keywords":       "Test Keywords",
			"Language":       "en-US",
			"LastModifiedBy": "Test User",
			"Revision":       "1",
			"Subject":        "Test Subject",
			"Version":        "1.0",
			"Identifier":     "",
			"Created":        "",
			"Modified":       "",
		},
		"protection": map[string]interface{}{
			"algorithm":      "XOR",
			"password":       "12345",
			"lock_structure": true,
			"lock_windows":   false,
		},
		"engine":      "normalWriter",
		"sheet_order": []interface{}{"TestingSheet2", "Sheet2WithNoWidth", "Sheet3WithNoHeight", "Sheet4WithNoWidthAndHeight"},
		"content": map[string]interface{}{
			"TestingSheet2": map[string]interface{}{
				"Header": [][]string{
					{"Column1", "Column2", "Column3"},
				},
				"Data": [][][]string{
					{{"Data1", "style1"}, {"Data2", "style1"}, {"Data3", "style1"}, {}},
					{{"Data4", "style1"}, {"Data5", "style1"}, {"Data6", "style1"}, {}},
				},
				"Height":         map[string]int{"3": 252},
				"Width":          map[string]int{"1": 25, "2": 26, "3": 6},
				"MergeCells":     []interface{}{},
				"AutoFilter":     []interface{}{},
				"Panes":          map[string]interface{}{},
				"DataValidation": []interface{}{},
				"Comment": []interface{}{map[string]interface{}{
					"cell":      "A1",
					"author":    "author",
					"paragraph": []interface{}{map[string]interface{}{"text": "text", "bold": true}}}},
				"NoStyle":      false,
				"GroupedRow":   []interface{}{map[string]interface{}{"start_row": 1.0, "end_row": 3.0, "outline_level": 1, "hidden": false}},
				"GroupedCol":   []interface{}{map[string]interface{}{"start_col": "A", "end_col": "C", "outline_level": 1, "hidden": true}},
				"Table":        []interface{}{},
				"Chart":        []interface{}{},
				"PivotTable":   []interface{}{},
				"SheetVisible": true,
				"WriterEngine": "NormalWriter",
			},
			"Sheet2WithNoWidth": map[string]interface{}{
				"Header": [][]string{
					{"Column1", "Column2", "Column3"},
				},
				"Data": [][][]string{
					{{"Data1", "style1"}, {"Data2", "style1"}, {"Data3", "style1"}},
					{{"Data4", "style1"}, {"Data5", "style1"}, {"Data6", "style1"}},
				},
				"Height":     map[string]int{"3": 252},
				"MergeCells": [][]interface{}{{"A1", "A2"}, {"B2", "C3"}},
				"AutoFilter": []interface{}{},
				"Panes":      map[string]interface{}{},
				"DataValidation": []interface{}{map[string]interface{}{
					"sq_ref":    "A1",
					"set_range": "B1",
					"drop_list": []string{"123", "qwe"}}},
				"NoStyle":      false,
				"Comment":      []interface{}{},
				"GroupedRow":   []interface{}{map[string]interface{}{"start_row": 1.0, "end_row": 1.0, "outline_level": 9, "hidden": true}},
				"GroupedCol":   []interface{}{map[string]interface{}{"start_col": "A", "end_col": "A", "outline_level": 2, "hidden": true}},
				"Table":        []interface{}{},
				"Chart":        []interface{}{},
				"PivotTable":   []interface{}{},
				"SheetVisible": true,
				"WriterEngine": "NormalWriter",
			},
			"Sheet3WithNoHeight": map[string]interface{}{
				"Header": [][]string{
					{"Column1", "Column2", "Column3"},
				},
				"Data": [][][]string{
					{{"Data1", "style1"}, {"Data2", "style1"}, {"Data3", "style1"}},
					{{"Data4", "style1"}, {"Data5", "style1"}, {"Data6", "style1"}},
				},
				"Width":          map[string]int{"1": 25, "2": 26, "3": 6},
				"MergeCells":     []interface{}{},
				"AutoFilter":     []interface{}{},
				"Panes":          map[string]interface{}{},
				"DataValidation": []interface{}{map[string]interface{}{"sq_ref": "A1", "sqref_drop_list": "A1:B1"}},
				"Comment":        []interface{}{},
				"NoStyle":        false,
				"GroupedRow":     []interface{}{},
				"GroupedCol":     []interface{}{},
				"Table":          []interface{}{},
				"Chart":          []interface{}{},
				"PivotTable":     []interface{}{},
				"SheetVisible":   false,
				"WriterEngine":   "NormalWriter",
			},
			"Sheet4WithNoWidthAndHeight": map[string]interface{}{
				"Header": [][]string{
					{"Column1", "Column2", "Column3"},
				},
				"Data": [][][]string{
					{{"Data1", "style1"}, {"Data2", "style1"}, {"Data3", "style1"}},
					{{"Data4", "style1"}, {"Data5", "style1"}, {"Data6", "style1"}},
				},
				"MergeCells": [][]interface{}{{"A1", "A2"}, {"B2", "C3"}},
				"AutoFilter": []interface{}{"A1:C1"},
				"Panes":      map[string]interface{}{},
				"DataValidation": []interface{}{map[string]interface{}{
					"sq_ref":      "A1",
					"error_title": "err_test",
					"error_body":  "err_body",
					"input_title": "input_test",
					"input_body":  "input_body"}},
				"Comment":      []interface{}{},
				"NoStyle":      false,
				"GroupedRow":   []interface{}{map[string]interface{}{"start_row": 1.0, "end_row": 1.0, "outline_level": 9, "hidden": true}},
				"GroupedCol":   []interface{}{},
				"Table":        []interface{}{map[string]interface{}{"range": "A1:B3", "name": "test", "style_name": "", "show_first_column": true, "show_last_column": true, "show_row_stripes": false, "show_column_stripes": true}},
				"Chart":        []interface{}{},
				"PivotTable":   []interface{}{},
				"SheetVisible": false,
				"WriterEngine": "NormalWriter",
			},
		},
	}
}

func TestWriteExcel(t *testing.T) {
	// Mock input data
	jsonData, err := json.Marshal(data)
	if err != nil {
		fmt.Println("Error:", err)
		return
	}

	encodedExcel := WriteExcel(string(jsonData))
	decodedExcel, err := base64.StdEncoding.DecodeString(encodedExcel)
	if err != nil {
		t.Fatalf("Failed to decode encoded Excel data: %v", err)
	}

	// Assert the expected result
	if len(decodedExcel) == 0 {
		t.Error("Encoded Excel data is empty")
	}
}

func TestWriteExcel2(t *testing.T) {
	// Mock input data
	data["content"] = map[string]interface{}{
		"Sheet1": map[string]interface{}{
			"Header": [][]string{
				{"Column1", "Column2", "Column3"},
			},
			"Data": [][][]string{
				{{"Data1", "style1"}, {"Data2", "style1"}, {"Data3", "style1"}},
				{{"Data4", "style1"}, {"Data5", "style1"}, {"Data6", "style1"}},
			},
			"MergeCells":     [][]interface{}{{"A1", "A2"}, {"B2", "C3"}},
			"AutoFilter":     []interface{}{"A1:C1"},
			"Panes":          map[string]interface{}{},
			"DataValidation": []interface{}{},
			"Comment":        []interface{}{},
			"NoStyle":        false,
			"Table":          []interface{}{},
			"Chart":          []interface{}{},
			"PivotTable":     []interface{}{},
			"SheetVisible":   true,
			"WriterEngine":   "StreamWriter",
		},
	}
	data["sheet_order"] = []interface{}{"Sheet1"}
	jsonData, err := json.Marshal(data)
	if err != nil {
		fmt.Println("Error:", err)
		return
	}

	encodedExcel := WriteExcel(string(jsonData))
	decodedExcel, err := base64.StdEncoding.DecodeString(encodedExcel)
	if err != nil {
		t.Fatalf("Failed to decode encoded Excel data: %v", err)
	}

	// Assert the expected result
	if len(decodedExcel) == 0 {
		t.Error("Encoded Excel data is empty")
	}

}

func TestWriteExcelCreatesPivotTableForLargeStreamedDataRange(t *testing.T) {
	const rowCount = 2600
	payload := strings.Repeat("x", 7000)
	sheetRows := make([]interface{}, 0, rowCount+1)
	sheetRows = append(sheetRows, []interface{}{"DEPT_NO", "AMOUNT", "PAYLOAD"})
	for rowIndex := 0; rowIndex < rowCount; rowIndex++ {
		sheetRows = append(
			sheetRows,
			[]interface{}{fmt.Sprintf("D%04d", rowIndex%10), rowIndex, payload},
		)
	}

	pivotTable := map[string]interface{}{
		"DataRange":       fmt.Sprintf("Sheet1!A1:C%d", rowCount+1),
		"PivotTableRange": "Pivot!A3:C20",
		"Rows": []interface{}{
			map[string]interface{}{"Data": "DEPT_NO", "Name": "Dept No"},
		},
		"Filter":  []interface{}{},
		"Columns": []interface{}{},
		"Data": []interface{}{
			map[string]interface{}{"Data": "AMOUNT", "Name": "Amount", "Subtotal": "Sum"},
		},
		"RowGrandTotals": true,
		"ColGrandTotals": true,
		"ShowDrill":      true,
		"ShowRowHeaders": true,
		"ShowColHeaders": true,
		"ShowLastColumn": false,
		"ClassicLayout":  true,
	}

	file := excelize.NewFile()
	defer func() {
		if err := file.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	writer := ExcelWriter{
		File:       file,
		StyleMap:   map[string]interface{}{},
		FileProps:  newFileProps(),
		Protection: map[string]interface{}{},
		SheetOrder: []interface{}{"Sheet1", "Pivot"},
		Content: map[string]interface{}{
			"Sheet1": newStreamWriterSheet(sheetRows, []interface{}{pivotTable}),
			"Pivot":  newStreamWriterSheet([]interface{}{}, []interface{}{}),
		},
	}

	decodedExcel, err := base64.StdEncoding.DecodeString(writer.writeExcel())
	if err != nil {
		t.Fatalf("Failed to decode encoded Excel data: %v", err)
	}

	archive, err := zip.NewReader(bytes.NewReader(decodedExcel), int64(len(decodedExcel)))
	if err != nil {
		t.Fatalf("Failed to read generated Excel archive: %v", err)
	}

	var pivotTableCount int
	var pivotCacheCount int
	var cacheDefinition string
	for _, file := range archive.File {
		switch {
		case strings.HasPrefix(file.Name, "xl/pivotTables/pivotTable"):
			pivotTableCount++
		case strings.HasPrefix(file.Name, "xl/pivotCache/pivotCacheDefinition"):
			pivotCacheCount++
			cacheDefinition = readZipFile(t, file)
		}
	}

	if pivotTableCount != 1 {
		t.Fatalf("Expected 1 pivot table, got %d", pivotTableCount)
	}
	if pivotCacheCount != 1 {
		t.Fatalf("Expected 1 pivot cache definition, got %d", pivotCacheCount)
	}
	if !strings.Contains(cacheDefinition, `sheet="Sheet1"`) {
		t.Fatalf("Expected pivot cache source sheet to be Sheet1: %s", cacheDefinition)
	}
	expectedRef := fmt.Sprintf(`ref="A1:C%d"`, rowCount+1)
	if !strings.Contains(cacheDefinition, expectedRef) {
		t.Fatalf("Expected pivot cache source ref %s: %s", expectedRef, cacheDefinition)
	}
}

func newStreamWriterSheet(dataRows []interface{}, pivotTables []interface{}) map[string]interface{} {
	return map[string]interface{}{
		"Data":           dataRows,
		"Width":          map[string]interface{}{},
		"Height":         map[string]interface{}{},
		"MergeCells":     []interface{}{},
		"AutoFilter":     []interface{}{},
		"Panes":          map[string]interface{}{},
		"DataValidation": []interface{}{},
		"Comment":        []interface{}{},
		"NoStyle":        true,
		"Table":          []interface{}{},
		"Chart":          []interface{}{},
		"PivotTable":     pivotTables,
		"SheetVisible":   true,
		"WriterEngine":   "StreamWriter",
	}
}

func newFileProps() map[string]interface{} {
	return map[string]interface{}{
		"Title":          "Test Excel File",
		"Creator":        "Test User",
		"Category":       "Test Category",
		"ContentStatus":  "Draft",
		"Description":    "Test Description",
		"Keywords":       "Test Keywords",
		"Language":       "en-US",
		"LastModifiedBy": "Test User",
		"Revision":       "1",
		"Subject":        "Test Subject",
		"Version":        "1.0",
		"Identifier":     "",
		"Created":        "",
		"Modified":       "",
	}
}

func readZipFile(t *testing.T, file *zip.File) string {
	t.Helper()
	reader, err := file.Open()
	if err != nil {
		t.Fatalf("Failed to open %s: %v", file.Name, err)
	}
	defer reader.Close()
	buffer := new(bytes.Buffer)
	if _, err := buffer.ReadFrom(reader); err != nil {
		t.Fatalf("Failed to read %s: %v", file.Name, err)
	}
	return buffer.String()
}

func TestWriteExcelNormalWriter(t *testing.T) {
	// Mock input data
	jsonData, err := json.Marshal(dataNormalWriter)
	if err != nil {
		fmt.Println("Error:", err)
		return
	}

	encodedExcel := WriteExcel(string(jsonData))
	decodedExcel, err := base64.StdEncoding.DecodeString(encodedExcel)
	if err != nil {
		t.Fatalf("Failed to decode encoded Excel data: %v", err)
	}

	// Assert the expected result
	if len(decodedExcel) == 0 {
		t.Error("Encoded Excel data is empty")
	}
}

func TestWriteExcel2NormalWriter(t *testing.T) {
	// Mock input data
	dataNormalWriter["content"] = map[string]interface{}{
		"Sheet1": map[string]interface{}{
			"Header": [][]string{
				{"Column1", "Column2", "Column3"},
			},
			"Data": [][][]string{
				{{"Data1", "style1"}, {"Data2", "style1"}, {"Data3", "style1"}},
				{{"Data4", "style1"}, {"Data5", "style1"}, {"Data6", "style1"}},
			},
			"MergeCells":     [][]interface{}{{"A1", "A2"}, {"B2", "C3"}},
			"AutoFilter":     []interface{}{"A1:C1"},
			"Panes":          map[string]interface{}{},
			"DataValidation": []interface{}{},
			"Comment":        []interface{}{},
			"NoStyle":        false,
			"Table":          []interface{}{},
			"Chart":          []interface{}{},
			"PivotTable":     []interface{}{},
			"SheetVisible":   true,
			"WriterEngine":   "NormalWriter",
		},
	}
	dataNormalWriter["sheet_order"] = []interface{}{"Sheet1"}
	jsonData, err := json.Marshal(data)
	if err != nil {
		fmt.Println("Error:", err)
		return
	}

	encodedExcel := WriteExcel(string(jsonData))
	decodedExcel, err := base64.StdEncoding.DecodeString(encodedExcel)
	if err != nil {
		t.Fatalf("Failed to decode encoded Excel data: %v", err)
	}

	// Assert the expected result
	if len(decodedExcel) == 0 {
		t.Error("Encoded Excel data is empty")
	}

}
