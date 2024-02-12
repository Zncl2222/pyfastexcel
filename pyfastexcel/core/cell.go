package core

import (
	"strings"

	"github.com/xuri/excelize/v2"
)

func createCell(v []interface{}) excelize.Cell {
	switch value := v[0].(type) {
	case string:
		if strings.HasPrefix(value, "=") {
			return excelize.Cell{StyleID: styleMap[v[1].(string)], Formula: value}
		}
		return excelize.Cell{StyleID: styleMap[v[1].(string)], Value: value}
	default:
		return excelize.Cell{StyleID: styleMap[v[1].(string)], Value: value}
	}
}
