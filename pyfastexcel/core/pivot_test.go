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
	for row, values := range map[string][]interface{}{
		"A1": {"RowField", "FilterField", "ColumnField", "DataField", "Other"},
		"A2": {"row", "filter", "column", 10, "other"},
	} {
		if err := file.SetSheetRow("Sheet1", row, &values); err != nil {
			t.Fatalf("seed pivot source row %s: %v", row, err)
		}
	}

	// Mock pivot table data
	pivotData := []interface{}{
		map[string]interface{}{
			"DataRange":       "Sheet1!A1:E2",
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
			"ClassicLayout":  false,
		},
	}
	ew := ExcelWriter{
		File: file,
	}

	// Call the function to test
	if err := ew.createPivotTable(pivotData); err != nil {
		t.Fatalf("create pivot table: %v", err)
	}
}

func TestSeedPivotSourceHeadersUsesAbsoluteSourceColumns(t *testing.T) {
	file := excelize.NewFile()
	defer file.Close()
	ew := ExcelWriter{
		File: file,
		Content: map[string]interface{}{
			"Sheet1": map[string]interface{}{
				"Data": []interface{}{
					[]interface{}{
						[]interface{}{"ignored", "DEFAULT_STYLE"},
						[]interface{}{"Month", "DEFAULT_STYLE"},
						[]interface{}{"Sales", "DEFAULT_STYLE"},
					},
				},
			},
		},
	}

	if err := ew.seedPivotSourceHeaders([]interface{}{
		map[string]interface{}{"DataRange": "Sheet1!B1:C3"},
	}); err != nil {
		t.Fatalf("seed pivot source headers: %v", err)
	}

	for cell, expected := range map[string]string{"B1": "Month", "C1": "Sales"} {
		actual, err := file.GetCellValue("Sheet1", cell)
		if err != nil {
			t.Fatalf("read %s: %v", cell, err)
		}
		if actual != expected {
			t.Errorf("%s: expected %q, got %q", cell, expected, actual)
		}
	}
}
