package core

import (
	"bytes"
	"encoding/binary"
	"encoding/json"
	"fmt"
	"strconv"
	"strings"
	"testing"

	"github.com/vmihailenco/msgpack/v5"
	"github.com/xuri/excelize/v2"
)

// newMultiSheetPFX2Payload builds a payload whose sheets all use the
// StreamWriter engine and whose wire metadata carries sheet_offsets, which is
// what routes WriteExcelV2 through the concurrent per-sheet path.
func newMultiSheetPFX2Payload(
	t testing.TB,
	rowsBySheet [][]interface{},
	mutateWire func(map[string]interface{}),
) []byte {
	t.Helper()

	content := map[string]interface{}{}
	sheetOrder := make([]string, 0, len(rowsBySheet))
	rowCounts := make([]int, 0, len(rowsBySheet))
	for index, rows := range rowsBySheet {
		name := "Sheet" + strconv.Itoa(index+1)
		sheet := newStreamWriterSheet([]interface{}{}, []interface{}{})
		sheet["NoStyle"] = false
		sheet["WriterEngine"] = "StreamWriter"
		content[name] = sheet
		sheetOrder = append(sheetOrder, name)
		rowCounts = append(rowCounts, len(rows))
	}

	var encodedRows bytes.Buffer
	encoder := msgpack.NewEncoder(&encodedRows)
	sheetOffsets := make([]int64, 0, len(rowsBySheet))
	for _, rows := range rowsBySheet {
		sheetOffsets = append(sheetOffsets, int64(encodedRows.Len()))
		for _, row := range rows {
			if err := encoder.Encode(row); err != nil {
				t.Fatalf("encode test row: %v", err)
			}
		}
	}

	wire := map[string]interface{}{
		"version":       wireVersion,
		"style_names":   []string{"DEFAULT_STYLE", "accent"},
		"row_counts":    rowCounts,
		"sheet_offsets": sheetOffsets,
	}
	if mutateWire != nil {
		mutateWire(wire)
	}
	metadata := map[string]interface{}{
		"style": map[string]interface{}{
			"DEFAULT_STYLE": testStyleDefinition("000000"),
			"accent":        testStyleDefinition("FF0000"),
		},
		"file_props":        newFileProps(),
		"protection":        map[string]interface{}{},
		"sheet_order":       sheetOrder,
		"content":           content,
		"_pyfastexcel_wire": wire,
	}
	metadataBytes, err := json.Marshal(metadata)
	if err != nil {
		t.Fatalf("marshal PFX2 metadata: %v", err)
	}

	payload := make([]byte, wireHeaderSize+len(metadataBytes)+encodedRows.Len())
	copy(payload[:4], wireMagic[:])
	binary.BigEndian.PutUint64(payload[4:12], uint64(len(metadataBytes)))
	copy(payload[12:], metadataBytes)
	copy(payload[12+len(metadataBytes):], encodedRows.Bytes())
	return payload
}

func multiSheetTestRows(sheets, rowsPerSheet int) [][]interface{} {
	rowsBySheet := make([][]interface{}, sheets)
	for sheetIndex := range rowsBySheet {
		rows := make([]interface{}, rowsPerSheet)
		for rowIndex := range rows {
			rows[rowIndex] = []interface{}{
				[]interface{}{
					fmt.Sprintf("s%d-r%d", sheetIndex+1, rowIndex+1),
					uint32((sheetIndex + rowIndex) % 2),
				},
				[]interface{}{int64(sheetIndex*rowsPerSheet + rowIndex), uint32(0)},
			}
		}
		rowsBySheet[sheetIndex] = rows
	}
	return rowsBySheet
}

func TestWriteExcelV2ParallelSheetsRoundTrip(t *testing.T) {
	const sheets = 3
	const rowsPerSheet = 40
	payload := newMultiSheetPFX2Payload(t, multiSheetTestRows(sheets, rowsPerSheet), nil)

	workbookBytes, err := WriteExcelV2(payload)
	if err != nil {
		t.Fatalf("WriteExcelV2 returned an error: %v", err)
	}
	assertMultiSheetContent(t, workbookBytes, sheets, rowsPerSheet)
}

func assertMultiSheetContent(t *testing.T, workbookBytes []byte, sheets, rowsPerSheet int) {
	t.Helper()
	workbook, err := excelize.OpenReader(bytes.NewReader(workbookBytes))
	if err != nil {
		t.Fatalf("open generated workbook: %v", err)
	}
	defer workbook.Close()

	for sheetIndex := 0; sheetIndex < sheets; sheetIndex++ {
		sheet := "Sheet" + strconv.Itoa(sheetIndex+1)
		for _, rowNumber := range []int{1, rowsPerSheet / 2, rowsPerSheet} {
			expected := fmt.Sprintf("s%d-r%d", sheetIndex+1, rowNumber)
			actual, err := workbook.GetCellValue(sheet, "A"+strconv.Itoa(rowNumber))
			if err != nil {
				t.Fatalf("read %s!A%d: %v", sheet, rowNumber, err)
			}
			if actual != expected {
				t.Errorf("%s!A%d: expected %q, got %q", sheet, rowNumber, expected, actual)
			}
			expectedNumber := strconv.Itoa(sheetIndex*rowsPerSheet + rowNumber - 1)
			actual, err = workbook.GetCellValue(sheet, "B"+strconv.Itoa(rowNumber))
			if err != nil {
				t.Fatalf("read %s!B%d: %v", sheet, rowNumber, err)
			}
			if actual != expectedNumber {
				t.Errorf("%s!B%d: expected %q, got %q", sheet, rowNumber, actual, expectedNumber)
			}
		}
	}
}

func TestWriteExcelV2ParallelSheetsRejectBadOffsets(t *testing.T) {
	rows := multiSheetTestRows(2, 4)
	for name, mutate := range map[string]func(map[string]interface{}){
		"non-zero first offset": func(wire map[string]interface{}) {
			wire["sheet_offsets"] = []int64{4, 100}
		},
		"decreasing offsets": func(wire map[string]interface{}) {
			wire["sheet_offsets"] = []int64{0, 1 << 40}
		},
		"count mismatch": func(wire map[string]interface{}) {
			wire["sheet_offsets"] = []int64{0}
		},
	} {
		t.Run(name, func(t *testing.T) {
			payload := newMultiSheetPFX2Payload(t, rows, mutate)
			if _, err := WriteExcelV2(payload); err == nil {
				t.Fatal("expected an error for invalid sheet_offsets")
			} else if !strings.Contains(err.Error(), "sheet_offsets") {
				t.Fatalf("expected a sheet_offsets error, got: %v", err)
			}
		})
	}
}

// Offsets that are in bounds but point at the wrong rows must fail loudly
// (each worker validates its own segment), never silently misplace data.
func TestWriteExcelV2ParallelSheetsRejectMisalignedOffsets(t *testing.T) {
	payload := newMultiSheetPFX2Payload(t, multiSheetTestRows(2, 4), func(wire map[string]interface{}) {
		offsets := wire["sheet_offsets"].([]int64)
		wire["sheet_offsets"] = []int64{0, offsets[1] - 1}
	})
	if _, err := WriteExcelV2(payload); err == nil {
		t.Fatal("expected an error for misaligned sheet_offsets")
	}
}

func TestWriteExcelV2WithoutSheetOffsetsStillDecodes(t *testing.T) {
	const sheets = 2
	const rowsPerSheet = 6
	payload := newMultiSheetPFX2Payload(t, multiSheetTestRows(sheets, rowsPerSheet), func(wire map[string]interface{}) {
		delete(wire, "sheet_offsets")
	})
	workbookBytes, err := WriteExcelV2(payload)
	if err != nil {
		t.Fatalf("WriteExcelV2 returned an error: %v", err)
	}
	assertMultiSheetContent(t, workbookBytes, sheets, rowsPerSheet)
}
