package core

import (
	"errors"
	"reflect"

	"github.com/xuri/excelize/v2"
)

type fieldMapping struct {
	Name  string
	Type  string
	Value interface{}
}

func setField(obj interface{}, fieldMap map[string]interface{}, mappings []fieldMapping) error {
	for _, mapping := range mappings {
		if value, ok := fieldMap[mapping.Name]; ok {
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
			case "float64":
				field.SetFloat(value.(float64))
			case "[]string":
				value = []string{value.(string)}
				field.Set(reflect.ValueOf(value))
			default:
				return errors.New("unsupported field type")
			}
		}
	}
	return nil
}

func getFontStyle(fontMap map[string]interface{}) *excelize.Font {
	var fontStyle excelize.Font

	mappings := []fieldMapping{
		{Name: "Bold", Type: "bool"},
		{Name: "Italic", Type: "bool"},
		{Name: "Underline", Type: "string"},
		{Name: "Family", Type: "string"},
		{Name: "Size", Type: "float64"},
		{Name: "Strike", Type: "bool"},
		{Name: "Color", Type: "string"},
	}
	setField(&fontStyle, fontMap, mappings)

	return &fontStyle
}

func getFillStyle(fillMap map[string]interface{}) excelize.Fill {
	var fillStyle excelize.Fill

	mappings := []fieldMapping{
		{Name: "Type", Type: "string"},
		{Name: "Color", Type: "[]string"},
		{Name: "Pattern", Type: "int"},
		{Name: "Shading", Type: "int"},
	}
	setField(&fillStyle, fillMap, mappings)

	return fillStyle
}

func getBorderStyle(borderMap map[string]interface{}) []excelize.Border {
	var borderStyle []excelize.Border

	direction := []string{"left", "top", "bottom", "right"}

	for _, dir := range direction {
		if bd, ok := borderMap[dir].(map[string]interface{}); ok {
			border := excelize.Border{Type: dir}
			mappings := []fieldMapping{
				{Name: "Color", Type: "string"},
				{Name: "Style", Type: "int"},
			}
			setField(&border, bd, mappings)
			borderStyle = append(borderStyle, border)
		}
	}

	return borderStyle
}

func getAlignmentStyle(alignmentMap map[string]interface{}) *excelize.Alignment {
	var alignmentStyle excelize.Alignment

	mappings := []fieldMapping{
		{Name: "Horizontal", Type: "string"},
		{Name: "Indent", Type: "int"},
		{Name: "JustifyLastLine", Type: "bool"},
		{Name: "ReadingOrder", Type: "uint64"},
		{Name: "RelativeIndent", Type: "int"},
		{Name: "ShrinkToFit", Type: "bool"},
		{Name: "TextRotation", Type: "int"},
		{Name: "Vertical", Type: "string"},
		{Name: "WrapText", Type: "bool"},
	}

	setField(&alignmentStyle, alignmentMap, mappings)

	return &alignmentStyle
}

func getProtectionStyle(protectionMap map[string]interface{}) *excelize.Protection {
	var protectionStyle excelize.Protection

	mappings := []fieldMapping{
		{Name: "Hidden", Type: "bool"},
		{Name: "Locked", Type: "bool"},
	}

	setField(&protectionStyle, protectionMap, mappings)

	return &protectionStyle
}

func CreateStyle(file *excelize.File, styleSettings map[string]map[string]interface{}) map[string]int {
	styleMap := make(map[string]int)

	for key, style := range styleSettings {
		customNumFmt := style["CustomNumFmt"].(string)
		customStyle, err := file.NewStyle(&excelize.Style{
			Font:         getFontStyle(style["Font"].(map[string]interface{})),
			Fill:         getFillStyle(style["Fill"].(map[string]interface{})),
			Border:       getBorderStyle(style["Border"].(map[string]interface{})),
			Alignment:    getAlignmentStyle(style["Alignment"].(map[string]interface{})),
			Protection:   getProtectionStyle(style["Protection"].(map[string]interface{})),
			CustomNumFmt: &customNumFmt,
		})
		if err != nil {
			panic(err)
		}

		styleMap[key] = customStyle
	}

	return styleMap
}
