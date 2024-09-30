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

		rowGrandTotals := false
		if pivotMap["RowGrandTotals"] != nil {
			rowGrandTotals = pivotMap["RowGrandTotals"].(bool)
		}
		colGrandTotals := false
		if pivotMap["ColGrandTotals"] != nil {
			colGrandTotals = pivotMap["ColGrandTotals"].(bool)
		}
		showDrill := false
		if pivotMap["ShowDrill"] != nil {
			showDrill = pivotMap["ShowDrill"].(bool)
		}
		showRowHeaders := false
		if pivotMap["ShowRowHeaders"] != nil {
			showRowHeaders = pivotMap["ShowRowHeaders"].(bool)
		}
		showColHeaders := false
		if pivotMap["ShowColHeaders"] != nil {
			showColHeaders = pivotMap["ShowColHeaders"].(bool)
		}
		showLastColumn := false
		if pivotMap["ShowLastColumn"] != nil {
			showLastColumn = pivotMap["ShowLastColumn"].(bool)
		}
		classicLayout := false
		if pivotMap["ClassicLayout"] != nil {
			classicLayout = pivotMap["ClassicLayout"].(bool)
		}

		err := ew.File.AddPivotTable(&excelize.PivotTableOptions{
			DataRange:       pivotMap["DataRange"].(string),
			PivotTableRange: pivotMap["PivotTableRange"].(string),
			Rows:            getPivotTableField(pivotMap["Rows"]),
			Filter:          getPivotTableField(pivotMap["Filter"]),
			Columns:         getPivotTableField(pivotMap["Columns"]),
			Data:            getPivotTableField(pivotMap["Data"]),
			RowGrandTotals:  rowGrandTotals,
			ColGrandTotals:  colGrandTotals,
			ShowDrill:       showDrill,
			ShowRowHeaders:  showRowHeaders,
			ShowColHeaders:  showColHeaders,
			ShowLastColumn:  showLastColumn,
			ClassicLayout:   classicLayout,
		})
		if err != nil {
			fmt.Println(err)
		}
	}
}
