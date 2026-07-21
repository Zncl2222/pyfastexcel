package core

import (
	"fmt"
	"strings"

	"github.com/xuri/excelize/v2"
)

// createCell takes a slice of interface{} representing a cell data and returns an excelize.Cell object.
//
// Args:
//
//	v ([]interface{}): A slice containing cell data. The first element represents the value,
//					   and the second element (optional) represents the style name.
//
// Returns:
//
//	excelize.Cell: An excelize.Cell object representing the cell with appropriate value and style.
//
// Notes:
//   - The function checks the type of the first element in the slice (`v[0]`).
//   - If it's a string:
//   - If the string starts with "=" (formula): the cell is created with
//     the formula and the style ID from the second element (`v[1]`).
//   - Otherwise, the cell is created with the string value and the style ID.
//   - For any other type, the cell is created with the value and the style ID.
func createCell(v []interface{}, styles map[string]int) (excelize.Cell, error) {
	if len(v) == 0 {
		styleID := styles["DEFAULT_STYLE"]
		return excelize.Cell{StyleID: styleID, Value: ""}, nil
	}
	if len(v) != 2 {
		return excelize.Cell{}, fmt.Errorf("styled cell must have 0 or 2 elements, got %d", len(v))
	}
	styleName, ok := v[1].(string)
	if !ok {
		return excelize.Cell{}, fmt.Errorf("cell style must be a string, got %T", v[1])
	}
	styleID, ok := styles[styleName]
	if !ok {
		return excelize.Cell{}, fmt.Errorf("style %q is not defined", styleName)
	}
	switch value := v[0].(type) {
	case string:
		if strings.HasPrefix(value, "=") {
			return excelize.Cell{StyleID: styleID, Formula: normalizeFormula(value)}, nil
		}
		return excelize.Cell{StyleID: styleID, Value: value}, nil
	default:
		return excelize.Cell{StyleID: styleID, Value: value}, nil
	}
}

func normalizeFormula(formula string) string {
	return strings.TrimPrefix(formula, "=")
}
