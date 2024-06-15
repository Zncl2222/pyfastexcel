package core

import (
	"encoding/base64"
	"encoding/json"
	"fmt"
	"testing"
)

var data map[string]interface{}

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
		"content": map[string]interface{}{
			"TestingSheet2": map[string]interface{}{
				"Header": [][]string{
					{"Column1", "Column2", "Column3"},
				},
				"Data": [][][]string{
					{{"Data1", "style1"}, {"Data2", "style1"}, {"Data3", "style1"}},
					{{"Data4", "style1"}, {"Data5", "style1"}, {"Data6", "style1"}},
				},
				"Height":     map[string]int{"3": 252},
				"Width":      map[string]int{"1": 25, "2": 26, "3": 6},
				"MergeCells": []interface{}{},
				"AutoFilter": []interface{}{},
				"NoStyle":    false,
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
				"NoStyle":    false,
			},
			"Sheet3WithNoHeight": map[string]interface{}{
				"Header": [][]string{
					{"Column1", "Column2", "Column3"},
				},
				"Data": [][][]string{
					{{"Data1", "style1"}, {"Data2", "style1"}, {"Data3", "style1"}},
					{{"Data4", "style1"}, {"Data5", "style1"}, {"Data6", "style1"}},
				},
				"Width":      map[string]int{"1": 25, "2": 26, "3": 6},
				"MergeCells": []interface{}{},
				"AutoFilter": []interface{}{},
				"NoStyle":    false,
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
				"NoStyle":    false,
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
			"MergeCells": [][]interface{}{{"A1", "A2"}, {"B2", "C3"}},
			"AutoFilter": []interface{}{"A1:C1"},
		},
	}
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
