package core

import (
	"reflect"
	"testing"

	"github.com/xuri/excelize/v2"
)

func TestCreateCell(t *testing.T) {
	styles := map[string]int{"DEFAULT_STYLE": 7, "styleID": 11}
	tests := []struct {
		name   string
		input  []interface{}
		expect excelize.Cell
	}{
		{
			name:   "StringWithValue",
			input:  []interface{}{"test", "styleID"},
			expect: excelize.Cell{StyleID: styles["styleID"], Value: "test"},
		},
		{
			name:   "StringWithFormula",
			input:  []interface{}{"=SUM(A1:A10)", "styleID"},
			expect: excelize.Cell{StyleID: styles["styleID"], Formula: "SUM(A1:A10)"},
		},
		{
			name:   "NonString",
			input:  []interface{}{123, "styleID"},
			expect: excelize.Cell{StyleID: styles["styleID"], Value: 123},
		},
		{
			name:   "EmptyInterface",
			input:  []interface{}{},
			expect: excelize.Cell{StyleID: styles["DEFAULT_STYLE"], Value: ""},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			actual, err := createCell(tt.input, styles)
			if err != nil {
				t.Fatalf("createCell returned an unexpected error: %v", err)
			}
			if !reflect.DeepEqual(actual, tt.expect) {
				t.Errorf("Expected %#v but got %#v", tt.expect, actual)
			}
		})
	}
}

func TestCreateCellRejectsUnknownStyle(t *testing.T) {
	_, err := createCell([]interface{}{"value", "missing"}, map[string]int{})
	if err == nil {
		t.Fatal("expected an unknown style error")
	}
}
