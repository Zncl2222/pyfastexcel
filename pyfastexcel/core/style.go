package core

import (
	"errors"
	"fmt"
	"reflect"
	"sort"

	"github.com/xuri/excelize/v2"
)

type fieldMapping struct {
	Name  string
	Type  string
	Value interface{}
}

func setMappedValue(field reflect.Value, value interface{}, mappingType string) error {
	switch mappingType {
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

	return nil
}

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
			field := reflect.ValueOf(obj).Elem().FieldByNameFunc(func(fieldName string) bool {
				return fieldName == mapping.Name
			})
			if err := setMappedValue(field, value, mapping.Type); err != nil {
				return err
			}
		}
	}
	return nil
}

// getFontStyle extracts font style information from a map and returns an excelize.Font object.
//
// Args:
//
//	fontMap (map[string]interface{}): A map containing key-value pairs for font styles.
//
// Returns:
//
//	*excelize.Font: A pointer to an excelize.Font object representing the extracted style.
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

// getFillStyle extracts fill style information from a map and returns an excelize.Fill object.
//
// Args:
//
//	fillMap (map[string]interface{}): A map containing key-value pairs for fill styles.
//
// Returns:
//
//	excelize.Fill: An excelize.Fill object representing the extracted style.
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

// getBorderStyle extracts border style information from a map and returns a slice of excelize.Border objects.
//
// Args:
//
//	borderMap (map[string]interface{}): A map containing key-value pairs for border styles.
//
// Returns:
//
//	[]excelize.Border: A slice of excelize.Border objects representing the extracted styles for each direction.
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

// getAlignmentStyle extracts alignment style information from a map and returns an excelize.Alignment object.
//
// Args:
//
//	alignmentMap (map[string]interface{}): A map containing key-value pairs for alignment styles.
//
// Returns:
//
//	*excelize.Alignment: A pointer to an excelize.Alignment object representing the extracted style.
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

// getProtectionStyle extracts protection style information from a map and returns an excelize.Protection object.
//
// Args:
//
//	protectionMap (map[string]interface{}): A map containing key-value pairs for protection styles.
//
// Returns:
//
//	*excelize.Protection: A pointer to an excelize.Protection object representing the extracted style.
func getProtectionStyle(protectionMap map[string]interface{}) *excelize.Protection {
	var protectionStyle excelize.Protection

	mappings := []fieldMapping{
		{Name: "Hidden", Type: "bool"},
		{Name: "Locked", Type: "bool"},
	}

	setField(&protectionStyle, protectionMap, mappings)

	return &protectionStyle
}

// CreateStyle creates styles in an Excel file based on a map of style settings.
//
// Args:
//
//	file (*excelize.File): The Excel file object.
//	styleSettings (map[string]map[string]interface{}): A map containing key-value pairs for
//													   style names and their individual settings.
//
// Returns:
//
//	map[string]int: A map linking style names to their corresponding style index in the Excel file.
func CreateStyle(file *excelize.File, styleSettings map[string]interface{}) map[string]int {
	styleNames := make([]string, 0, len(styleSettings))
	for name := range styleSettings {
		styleNames = append(styleNames, name)
	}
	sort.Strings(styleNames)
	styleMap, _, err := createStylesOrdered(file, styleSettings, styleNames)
	if err != nil {
		panic(err)
	}
	return styleMap
}

// createStylesOrdered creates workbook styles in the supplied wire order. The
// returned slice maps a compact wire style ID to excelize's workbook-local
// style ID.
func createStylesOrdered(
	file *excelize.File,
	styleSettings map[string]interface{},
	styleNames []string,
) (map[string]int, []int, error) {
	styleMap := make(map[string]int, len(styleNames))
	styleIDs := make([]int, len(styleNames))
	seen := make(map[string]struct{}, len(styleNames))

	for wireID, key := range styleNames {
		if _, duplicate := seen[key]; duplicate {
			return nil, nil, fmt.Errorf("duplicate style name %q", key)
		}
		seen[key] = struct{}{}
		style, ok := styleSettings[key]
		if !ok {
			return nil, nil, fmt.Errorf("style %q is missing from metadata", key)
		}
		customNumFmt := style.(map[string]interface{})["CustomNumFmt"].(string)
		customStyle, err := file.NewStyle(&excelize.Style{
			Font:         getFontStyle(style.(map[string]interface{})["Font"].(map[string]interface{})),
			Fill:         getFillStyle(style.(map[string]interface{})["Fill"].(map[string]interface{})),
			Border:       getBorderStyle(style.(map[string]interface{})["Border"].(map[string]interface{})),
			Alignment:    getAlignmentStyle(style.(map[string]interface{})["Alignment"].(map[string]interface{})),
			Protection:   getProtectionStyle(style.(map[string]interface{})["Protection"].(map[string]interface{})),
			CustomNumFmt: &customNumFmt,
		})
		if err != nil {
			return nil, nil, fmt.Errorf("create style %q: %w", key, err)
		}

		styleMap[key] = customStyle
		styleIDs[wireID] = customStyle
	}

	if len(seen) != len(styleSettings) {
		return nil, nil, fmt.Errorf(
			"style order contains %d names, metadata contains %d styles",
			len(seen),
			len(styleSettings),
		)
	}

	return styleMap, styleIDs, nil
}
