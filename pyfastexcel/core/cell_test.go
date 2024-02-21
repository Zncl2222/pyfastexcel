package core

import (
	"reflect"
	"testing"

	"github.com/xuri/excelize/v2"
)

func TestCreateCell(t *testing.T) {
	tests := []struct {
		name   string
		input  []interface{}
		expect excelize.Cell
	}{
		{
			name:   "StringWithValue",
			input:  []interface{}{"test", "styleID"},
			expect: excelize.Cell{StyleID: styleMap["styleID"], Value: "test"},
		},
		{
			name:   "StringWithFormula",
			input:  []interface{}{"=SUM(A1:A10)", "styleID"},
			expect: excelize.Cell{StyleID: styleMap["styleID"], Formula: "=SUM(A1:A10)"},
		},
		{
			name:   "NonString",
			input:  []interface{}{123, "styleID"},
			expect: excelize.Cell{StyleID: styleMap["styleID"], Value: 123},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			actual := createCell(tt.input)
			if !reflect.DeepEqual(actual, tt.expect) {
				t.Errorf("Expected %#v but got %#v", tt.expect, actual)
			}
		})
	}
}
