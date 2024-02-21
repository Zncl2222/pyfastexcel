package core

import (
	"reflect"
	"testing"

	"github.com/xuri/excelize/v2"
)

func TestCreateStyle(t *testing.T) {
	// Mock styleSettings map
	styleSettings := map[string]map[string]interface{}{
		"style1": {
			"Font": map[string]interface{}{
				"Bold": true,
			},
			"Fill": map[string]interface{}{
				"Type":    "pattern",
				"Color":   "#FFFFFF",
				"Pattern": float64(1),
				"Shading": float64(100),
			},
			"Border": map[string]interface{}{
				"left": map[string]interface{}{
					"Color": "FF0000",
					"Style": float64(1),
				},
				"top": map[string]interface{}{
					"Color": "00FF00",
					"Style": float64(2),
				},
			},
			"Alignment": map[string]interface{}{
				"Horizontal":      "center",
				"Vertical":        "middle",
				"Indent":          float64(0),
				"JustifyLastLine": false,
				"ReadingOrder":    float64(0),
				"RelativeIndent":  float64(0),
				"ShrinkToFit":     false,
				"TextRotation":    float64(0),
				"WrapText":        false,
			},
			"Protection": map[string]interface{}{
				"Hidden": true,
				"Locked": false,
			},
			"CustomNumFmt": "0.00",
		},
		"style2": {
			"Font": map[string]interface{}{
				"Bold": true,
				"Size": float64(12),
			},
			"Fill": map[string]interface{}{
				"Type":    "gradient",
				"Color":   "#FFFFFF",
				"Pattern": float64(2),
				"Shading": float64(50),
			},
			"Border": map[string]interface{}{
				"left": map[string]interface{}{
					"Color": "0000FF",
					"Style": float64(3),
				},
				"bottom": map[string]interface{}{
					"Color": "FFFF00",
					"Style": float64(4),
				},
			},
			"Alignment": map[string]interface{}{
				"Horizontal": "right",
				"Vertical":   "top",
			},
			"Protection": map[string]interface{}{
				"Hidden": false,
				"Locked": true,
			},
			"CustomNumFmt": "0.000",
		},
	}

	// Mock excelize.File
	file := excelize.NewFile()

	// Call the function to be tested
	styleMap := CreateStyle(file, styleSettings)

	// Verify the created styles
	expectedStyles := map[string]int{
		"style1": 1,
		"style2": 2,
	}

	if !reflect.DeepEqual(styleMap, expectedStyles) {
		t.Errorf("Expected style map %#v, but got %#v", expectedStyles, styleMap)
	}
}
