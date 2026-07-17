package core

import (
	"strings"
	"testing"

	"github.com/xuri/excelize/v2"
)

func TestSetProtectionPreservesSupportedAlgorithmsAndRejectsUnknownOnes(t *testing.T) {
	config := func(algorithm string) map[string]interface{} {
		return map[string]interface{}{
			"algorithm":      algorithm,
			"password":       "password",
			"lock_structure": true,
			"lock_windows":   false,
		}
	}

	file := excelize.NewFile()
	if err := (&ExcelWriter{File: file}).setProtection(config("MD5")); err != nil {
		t.Fatalf("supported protection algorithm changed: %v", err)
	}
	_ = file.Close()

	file = excelize.NewFile()
	err := (&ExcelWriter{File: file}).setProtection(config("unknown"))
	_ = file.Close()
	if err == nil || !strings.Contains(err.Error(), "protect workbook: unsupported hash algorithm") {
		t.Fatalf("expected contextual unsupported-algorithm error, got %v", err)
	}
}

func TestSetDataValidationDoesNotPrependOrLeakDropListItems(t *testing.T) {
	file := excelize.NewFile()
	defer file.Close()
	writer := ExcelWriter{File: file}

	err := writer.setDataValidation("Sheet1", []interface{}{
		map[string]interface{}{
			"sq_ref":    "A1",
			"drop_list": []interface{}{"alpha", "beta"},
		},
		map[string]interface{}{
			"sq_ref": "B1",
		},
	})
	if err != nil {
		t.Fatalf("set data validations: %v", err)
	}

	validations, err := file.GetDataValidations("Sheet1")
	if err != nil {
		t.Fatalf("get data validations: %v", err)
	}
	if len(validations) != 2 {
		t.Fatalf("expected two independent validations, got %d", len(validations))
	}
	if formula := validations[0].Formula1; formula != `"alpha,beta"` {
		t.Fatalf("drop list contains leading blanks: %q", formula)
	}
	if formula := validations[1].Formula1; formula != "" {
		t.Fatalf("drop list leaked into the next validation: %q", formula)
	}
}

func TestGroupOutlineLevelsAreClampedForCompatibility(t *testing.T) {
	file := excelize.NewFile()
	defer file.Close()
	writer := ExcelWriter{File: file}

	if err := writer.groupRow("Sheet1", []interface{}{
		map[string]interface{}{
			"start_row":     float64(1),
			"outline_level": float64(12),
			"hidden":        false,
		},
		map[string]interface{}{
			"start_row":     float64(2),
			"outline_level": float64(-3),
			"hidden":        false,
		},
	}); err != nil {
		t.Fatalf("group rows with legacy outline levels: %v", err)
	}
	if err := writer.groupCol("Sheet1", []interface{}{
		map[string]interface{}{
			"start_col":     "A",
			"outline_level": float64(12),
			"hidden":        false,
		},
		map[string]interface{}{
			"start_col":     "B",
			"outline_level": float64(-3),
			"hidden":        false,
		},
	}); err != nil {
		t.Fatalf("group columns with legacy outline levels: %v", err)
	}

	for row, expected := range map[int]uint8{1: 7, 2: 1} {
		actual, err := file.GetRowOutlineLevel("Sheet1", row)
		if err != nil {
			t.Fatalf("get row %d outline level: %v", row, err)
		}
		if actual != expected {
			t.Errorf("row %d outline level = %d, want %d", row, actual, expected)
		}
	}
	for column, expected := range map[string]uint8{"A": 7, "B": 1} {
		actual, err := file.GetColOutlineLevel("Sheet1", column)
		if err != nil {
			t.Fatalf("get column %s outline level: %v", column, err)
		}
		if actual != expected {
			t.Errorf("column %s outline level = %d, want %d", column, actual, expected)
		}
	}
}

func TestGroupColumnsAcceptsCellReferencesAndRejectsInvalidReferences(t *testing.T) {
	file := excelize.NewFile()
	defer file.Close()
	writer := ExcelWriter{File: file}

	err := writer.groupCol("Sheet1", []interface{}{
		map[string]interface{}{
			"start_col":     "F1",
			"outline_level": float64(1),
			"hidden":        false,
		},
	})
	if err != nil {
		t.Fatalf("group legacy cell-reference column: %v", err)
	}
	level, err := file.GetColOutlineLevel("Sheet1", "F")
	if err != nil {
		t.Fatalf("get normalized column outline level: %v", err)
	}
	if level != 1 {
		t.Fatalf("column F outline level = %d, want 1", level)
	}

	err = writer.groupCol("Sheet1", []interface{}{
		map[string]interface{}{
			"start_col":     "not-a-column",
			"outline_level": float64(1),
			"hidden":        false,
		},
	})
	if err == nil || !strings.Contains(err.Error(), `parse column group 1 start "not-a-column" on sheet "Sheet1"`) {
		t.Fatalf("expected contextual invalid column error, got %v", err)
	}
}
