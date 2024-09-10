package core

import (
	"fmt"
	"testing"

	"github.com/xuri/excelize/v2"
)

func TestGetPivotTableField(t *testing.T) {
	fieldData := []interface{}{
		map[string]interface{}{
			"Compact":         true,
			"Data":            "SampleData",
			"Name":            "SampleName",
			"Outline":         true,
			"Subtotal":        "SampleSubtotal",
			"DefaultSubtotal": false,
		},
	}

	expected := []excelize.PivotTableField{
		{
			Compact:         true,
			Data:            "SampleData",
			Name:            "SampleName",
			Outline:         true,
			Subtotal:        "SampleSubtotal",
			DefaultSubtotal: false,
		},
	}

	result := getPivotTableField(fieldData)

	if len(result) != len(expected) {
		t.Errorf("Expected length %d, but got %d", len(expected), len(result))
	}

	for i := range expected {
		if result[i] != expected[i] {
			t.Errorf("Expected %+v, but got %+v", expected[i], result[i])
		}
	}
}

func TestCreatePivotTable(t *testing.T) {
	// Initialize an excel file
	file := excelize.NewFile()
	defer func() {
		if err := file.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	// Mock pivot table data
	pivotData := []interface{}{
		map[string]interface{}{
			"DataRange":       "Sheet1!A1:E31",
			"PivotTableRange": "Sheet1!G2:M34",
			"Rows": []interface{}{
				map[string]interface{}{
					"Name": "RowField",
				},
			},
			"Filter": []interface{}{
				map[string]interface{}{
					"Name": "FilterField",
				},
			},
			"Columns": []interface{}{
				map[string]interface{}{
					"Name": "ColumnField",
				},
			},
			"Data": []interface{}{
				map[string]interface{}{
					"Name": "DataField",
				},
			},
			"RowGrandTotals": true,
			"ColGrandTotals": true,
			"ShowDrill":      false,
			"ShowRowHeaders": true,
			"ShowColHeaders": true,
			"ShowLastColumn": false,
		},
	}

	// Call the function to test
	createPivotTable(file, pivotData)
}
