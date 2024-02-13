package core

// #include <stdlib.h>
import (
	"C"
)
import (
	"encoding/base64"
	"fmt"

	"github.com/perimeterx/marshmallow"

	"github.com/xuri/excelize/v2"
)

var styleMap map[string]int

type (
	StyleWrapper struct {
		Style map[string]map[string]interface{} `json:"Style" binding:"required"`
	}
)

func WriteExcel(data string) string {
	var StyleStruct StyleWrapper
	byteJson := []byte(data)

	strJson, err := marshmallow.Unmarshal(byteJson, &StyleStruct)
	if err != nil {
		panic(err)
	}

	file := excelize.NewFile()
	styleMap = CreateStyle(file, StyleStruct.Style)

	writeContentBySheet(file, strJson)

	// Save data in buffer and encode binary data to base64
	buffer, _ := file.WriteToBuffer()
	byteResults := []byte(buffer.Bytes())
	encodedString := base64.StdEncoding.EncodeToString(byteResults)

	return encodedString
}

func writeContentBySheet(file *excelize.File, data map[string]interface{}) {
	for sheet := range data {
		if sheet == "Style" {
			continue
		}
		sheetData := data[sheet].(map[string]interface{})

		// Create Sheet and Wrtie Header
		file.NewSheet(sheet)
		streamWriter, _ := file.NewStreamWriter(sheet)
		startedRow := 1
		cell, _ := excelize.CoordinatesToCellName(1, startedRow)
		for i, _ := range sheetData["Header"].([]interface{}) {
			sheetData["Header"].([]interface{})[i] = createCell(sheetData["Header"].([]interface{})[i].([]interface{}))
		}

		if len(sheetData["Header"].([]interface{})) != 0 {
			if err := streamWriter.SetRow(cell, sheetData["Header"].([]interface{})); err != nil {
				fmt.Println(err)
			}
			startedRow += 1
		}

		// Write Data
		excelData := sheetData["Data"].([]interface{})
		for i, rowData := range excelData {
			row := rowData.([]interface{})

			processedRow := make([]interface{}, len(row))
			for j, cellData := range row {
				cell := cellData.([]interface{})
				processedRow[j] = createCell(cell)
			}

			cell, err := excelize.CoordinatesToCellName(1, i+startedRow)
			if err != nil {
				panic(err)
			}

			if err := streamWriter.SetRow(cell, processedRow); err != nil {
				fmt.Println(err)
			}
		}

		if err := streamWriter.Flush(); err != nil {
			fmt.Println(err)
		}
	}
}
