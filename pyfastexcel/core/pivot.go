package core

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

func getPivotTableField(field interface{}) []excelize.PivotTableField {

	var pivotTableFields []excelize.PivotTableField

	fieldMappings := []fieldMapping{
		{Name: "Compact", Type: "bool"},
		{Name: "Data", Type: "string"},
		{Name: "Name", Type: "string"},
		{Name: "Outline", Type: "bool"},
		{Name: "Subtotal", Type: "string"},
		{Name: "DefaultSubtotal", Type: "bool"},
	}

	for _, f := range field.([]interface{}) {
		pivotTableField := excelize.PivotTableField{}
		setField(&pivotTableField, f.(map[string]interface{}), fieldMappings)
		pivotTableFields = append(pivotTableFields, pivotTableField)
	}
	return pivotTableFields
}

func (ew *ExcelWriter) createPivotTable(pivot_data []interface{}) {
	for _, pivot := range pivot_data {
		pivotMap := pivot.(map[string]interface{})
		err := ew.File.AddPivotTable(&excelize.PivotTableOptions{
			DataRange:       pivotMap["DataRange"].(string),
			PivotTableRange: pivotMap["PivotTableRange"].(string),
			Rows:            getPivotTableField(pivotMap["Rows"]),
			Filter:          getPivotTableField(pivotMap["Filter"]),
			Columns:         getPivotTableField(pivotMap["Columns"]),
			Data:            getPivotTableField(pivotMap["Data"]),
			RowGrandTotals:  getBoolValue(pivotMap, "RowGrandTotals", false),
			ColGrandTotals:  getBoolValue(pivotMap, "ColGrandTotals", false),
			ShowDrill:       getBoolValue(pivotMap, "ShowDrill", false),
			ShowRowHeaders:  getBoolValue(pivotMap, "ShowRowHeaders", false),
			ShowColHeaders:  getBoolValue(pivotMap, "ShowColHeaders", false),
			ShowLastColumn:  getBoolValue(pivotMap, "ShowLastColumn", false),
			ClassicLayout:   getBoolValue(pivotMap, "ClassicLayout", false),
		})
		if err != nil {
			fmt.Println(err)
		}
	}
}
