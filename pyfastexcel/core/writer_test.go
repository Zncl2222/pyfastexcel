package core

import (
	"encoding/base64"
	"testing"
)

func TestWriteExcel(t *testing.T) {
	// Mock input data
	inputData := `{
		"style": {
			"style1": {
				"Font": {
					"Bold": true
				},
				"Fill": {
					"Type":    "pattern",
					"Color":   "#FFFFFF",
					"Pattern": 1,
					"Shading": 100
				},
				"Border": {
					"left": {
						"Color": "FF0000",
						"Style": 1
					},
					"top": {
						"Color": "00FF00",
						"Style": 2
					}
				},
				"Alignment": {
					"Horizontal":      "center",
					"Vertical":        "middle",
					"Indent":          0,
					"JustifyLastLine": false,
					"ReadingOrder":    0,
					"RelativeIndent":  0,
					"ShrinkToFit":     false,
					"TextRotation":    0,
					"WrapText":        false
				},
				"Protection": {
					"Hidden": true,
					"Locked": false
				},
				"CustomNumFmt": "0.00"
			}
		},
		"file_props": {
			"Title": "Test Excel File",
			"Creator": "Test User",
			"Category": "Test Category",
			"ContentStatus": "Draft",
			"Description": "Test Description",
			"Keywords": "Test Keywords",
			"Language": "en-US",
			"LastModifiedBy": "Test User",
			"Revision": "1",
			"Subject": "Test Subject",
			"Version": "1.0",
			"Identifier": "",
			"Created": "",
			"Modified": ""
		},
		"content": {
			"Sheet1": {
				"Header": [
					["Column1", "Column2", "Column3"]
				],
				"Data": [
					[["Data1", "style1"], ["Data2", "style1"], ["Data3", "style1"]],
					[["Data4", "style1"], ["Data5", "style1"], ["Data6", "style1"]]
				],
				"Height": {"3": 252},
				"Width": {"1": 25, "2": 26, "3": 6}
			},
			"Sheet2WithNoWidth": {
				"Header": [
					["Column1", "Column2", "Column3"]
				],
				"Data": [
					[["Data1", "style1"], ["Data2", "style1"], ["Data3", "style1"]],
					[["Data4", "style1"], ["Data5", "style1"], ["Data6", "style1"]]
				],
				"Height": {"3": 252}
			},
			"Sheet3WithNoHeight": {
				"Header": [
					["Column1", "Column2", "Column3"]
				],
				"Data": [
					[["Data1", "style1"], ["Data2", "style1"], ["Data3", "style1"]],
					[["Data4", "style1"], ["Data5", "style1"], ["Data6", "style1"]]
				],
				"Width": {"1": 25, "2": 26, "3": 6}
			},
			"Sheet4WithNoWidthAndHeight": {
				"Header": [
					["Column1", "Column2", "Column3"]
				],
				"Data": [
					[["Data1", "style1"], ["Data2", "style1"], ["Data3", "style1"]],
					[["Data4", "style1"], ["Data5", "style1"], ["Data6", "style1"]]
				]
			}
		}
	}`
	encodedExcel := WriteExcel(inputData)
	decodedExcel, err := base64.StdEncoding.DecodeString(encodedExcel)
	if err != nil {
		t.Fatalf("Failed to decode encoded Excel data: %v", err)
	}

	// Assert the expected result
	if len(decodedExcel) == 0 {
		t.Error("Encoded Excel data is empty")
	}
}
