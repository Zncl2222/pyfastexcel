package core

import (
	"errors"
	"reflect"

	"github.com/xuri/excelize/v2"
)

// setField sets a field in a struct based on a map and field mappings.
//
// Args:
//
//	obj (interface{}): The struct object to set the field on.
//	fieldMap (map[string]interface{}): A map containing key-value pairs for field values.
//	mappings ([]fieldMapping): An array of field mappings specifying expected names and types.
//
//	error: Any error encountered during reflection or type conversion.
func setField(obj interface{}, fieldMap map[string]interface{}, mappings []fieldMapping) error {
	for _, mapping := range mappings {
		if value, ok := fieldMap[mapping.Name]; ok {
			if value == nil {
				continue
			}
			field := reflect.ValueOf(obj).Elem().FieldByName(mapping.Name)
			switch mapping.Type {
			case "string":
				field.SetString(value.(string))
			case "int":
				value = int(value.(float64))
				field.SetInt(int64(value.(int)))
			case "uint64":
				value = uint64(value.(float64))
				field.SetUint(uint64(value.(uint64)))
			case "bool":
				field.SetBool(value.(bool))
			case "*bool":
				v := value.(bool)
				field.Set(reflect.ValueOf(&v))
			case "float64":
				field.SetFloat(value.(float64))
			case "*float64":
				v := value.(float64)
				field.Set(reflect.ValueOf(&v))
			case "[]string":
				value = []string{value.(string)}
				field.Set(reflect.ValueOf(value))
			case "Fill":
				fillStyle := getFillStyle(value.(map[string]interface{}))
				field.Set(reflect.ValueOf(fillStyle))
			case "*Font":
				fillStyle := getFontStyle(value.(map[string]interface{}))
				field.Set(reflect.ValueOf(fillStyle))
			case "Font":
				fillStyle := getFontStyle(value.(map[string]interface{}))
				field.Set(reflect.ValueOf(*fillStyle))
			case "ChartDataLabelPositionType":
				positionType := excelize.ChartDataLabelPositionType(value.(float64))
				field.Set(reflect.ValueOf(positionType))
			case "ChartLine":
				lineStyle := getLineStyle(value.(map[string]interface{}))
				field.Set(reflect.ValueOf(lineStyle))
			case "ChartLineType":
				chartLineType := excelize.ChartLineType(value.(float64))
				field.Set(reflect.ValueOf(chartLineType))
			case "ChartMarker":
				markerStyle := getMarkerStyle(value.(map[string]interface{}))
				field.Set(reflect.ValueOf(markerStyle))
			case "[]RichTextRun":
				richTextRun := getTitleStruct(value)
				field.Set(reflect.ValueOf(richTextRun))
			case "ChartNumFmt":
				numFmt := getChartNumFmtStruct(value)
				field.Set(reflect.ValueOf(numFmt))
			default:
				return errors.New("unsupported field type")
			}
		}
	}
	return nil
}
