package core

import (
	"archive/zip"
	"bytes"
	"encoding/binary"
	"encoding/json"
	"fmt"
	"reflect"
	"strings"
	"testing"
	"time"

	"github.com/vmihailenco/msgpack/v5"
	"github.com/vmihailenco/msgpack/v5/msgpcode"
	"github.com/xuri/excelize/v2"
)

// A pivot source does not have to start in A1. This exercises the complete
// concurrent PFX2 path, including capturing a header from a non-first row and
// selecting only the columns covered by the source range.
func TestWriteExcelV2ParallelPivotUsesOffsetSourceHeader(t *testing.T) {
	t.Setenv("PYFASTEXCEL_SEQUENTIAL", "")

	styled := func(value interface{}) interface{} {
		return []interface{}{value, uint32(0)}
	}
	rowsBySheet := [][]interface{}{
		{
			[]interface{}{styled("ignored"), styled("not a header"), styled("also ignored")},
			[]interface{}{styled("row label"), styled("Category"), styled("Amount")},
			[]interface{}{styled("first"), styled("Hardware"), styled(int64(10))},
			[]interface{}{styled("second"), styled("Software"), styled(int64(20))},
		},
		{},
	}
	payload := newMultiSheetPFX2Payload(t, rowsBySheet, nil)
	payload = mutatePFX2TestMetadata(t, payload, func(metadata map[string]interface{}) {
		content := metadata["content"].(map[string]interface{})
		pivotSheet := content["Sheet2"].(map[string]interface{})
		pivotSheet["PivotTable"] = []interface{}{
			map[string]interface{}{
				"DataRange":       "Sheet1!B2:C4",
				"PivotTableRange": "Sheet2!A3:C20",
				"Rows": []interface{}{
					map[string]interface{}{"Data": "Category", "Name": "Category"},
				},
				"Filter":  []interface{}{},
				"Columns": []interface{}{},
				"Data": []interface{}{
					map[string]interface{}{"Data": "Amount", "Name": "Amount", "Subtotal": "Sum"},
				},
				"RowGrandTotals": true,
				"ColGrandTotals": true,
				"ShowDrill":      true,
				"ShowRowHeaders": true,
				"ShowColHeaders": true,
				"ShowLastColumn": false,
				"ClassicLayout":  true,
			},
		}
	})

	workbookBytes, err := WriteExcelV2(payload)
	if err != nil {
		t.Fatalf("write parallel PFX2 workbook with offset pivot source: %v", err)
	}
	archive, err := zip.NewReader(bytes.NewReader(workbookBytes), int64(len(workbookBytes)))
	if err != nil {
		t.Fatalf("open generated workbook: %v", err)
	}

	var cacheDefinition string
	var pivotDefinition string
	for _, file := range archive.File {
		switch {
		case strings.HasPrefix(file.Name, "xl/pivotCache/pivotCacheDefinition"):
			cacheDefinition = readZipFile(t, file)
		case strings.HasPrefix(file.Name, "xl/pivotTables/pivotTable"):
			pivotDefinition = readZipFile(t, file)
		}
	}
	if cacheDefinition == "" || pivotDefinition == "" {
		t.Fatalf("expected a pivot cache and table, cache=%t table=%t", cacheDefinition != "", pivotDefinition != "")
	}
	for _, expected := range []string{
		`sheet="Sheet1"`,
		`ref="B2:C4"`,
		`cacheField name="Category"`,
		`cacheField name="Amount"`,
	} {
		if !strings.Contains(cacheDefinition, expected) {
			t.Fatalf("pivot cache does not contain %q: %s", expected, cacheDefinition)
		}
	}
	for _, unexpected := range []string{`cacheField name="row label"`, `cacheField name="not a header"`} {
		if strings.Contains(cacheDefinition, unexpected) {
			t.Fatalf("pivot cache used a cell outside B2:C4 as a field: %s", cacheDefinition)
		}
	}
	for _, expected := range []string{
		`<rowFields count="1"><field x="0"`,
		`<dataFields count="1"><dataField name="Amount" fld="1" subtotal="sum"`,
	} {
		if !strings.Contains(pivotDefinition, expected) {
			t.Fatalf("pivot table does not contain %q: %s", expected, pivotDefinition)
		}
	}
}

func TestWriteExcelV2RejectsExcelResourceLimits(t *testing.T) {
	t.Run("row count", func(t *testing.T) {
		payload := newPFX2TestPayload(t, "StreamWriter", true, nil, func(wire map[string]interface{}) {
			wire["row_counts"] = []int{maxExcelRows + 1}
		})

		_, err := WriteExcelV2(payload)
		if err == nil || !strings.Contains(err.Error(), "row count 1048577 is outside Excel limits") {
			t.Fatalf("expected an Excel row-limit error, got %v", err)
		}
	})

	t.Run("Array32 column count", func(t *testing.T) {
		payload := newPFX2TestPayload(
			t,
			"StreamWriter",
			true,
			[]interface{}{[]interface{}{"unused"}},
			nil,
		)
		metadataLength := int(binary.BigEndian.Uint64(payload[len(wireMagic):wireHeaderSize]))
		metadataEnd := wireHeaderSize + metadataLength
		oversizedRow := make([]byte, 5)
		oversizedRow[0] = msgpcode.Array32
		binary.BigEndian.PutUint32(oversizedRow[1:], uint32(maxExcelCols+1))
		payload = append(append([]byte(nil), payload[:metadataEnd]...), oversizedRow...)

		_, err := WriteExcelV2(payload)
		if err == nil || !strings.Contains(err.Error(), "column count 16385 is outside Excel limits") {
			t.Fatalf("expected an Excel column-limit error, got %v", err)
		}
	})
}

func TestLegacyJSONAndPFX2HaveEquivalentWorkbookSemantics(t *testing.T) {
	for _, engine := range []string{"StreamWriter", "NormalWriter"} {
		t.Run(engine, func(t *testing.T) {
			legacyPayload, wirePayload := newEquivalentLegacyAndPFX2Payloads(t, engine)
			legacyBytes, err := WriteExcelV2(legacyPayload)
			if err != nil {
				t.Fatalf("write legacy JSON workbook: %v", err)
			}
			wireBytes, err := WriteExcelV2(wirePayload)
			if err != nil {
				t.Fatalf("write PFX2 workbook: %v", err)
			}

			legacy := readWorkbookSemantics(t, legacyBytes)
			wire := readWorkbookSemantics(t, wireBytes)
			expected := workbookSemantics{
				Sheets: []string{"Sheet1"},
				Rows: [][]string{
					{"Category", "Amount", "Computed"},
					{"Hardware", "7", ""},
					{"Software", "11", ""},
				},
				Formulas:   map[string]string{"C2": "B2*2", "C3": "B3*2"},
				FontColors: map[string]string{"A1": "FF0000", "A2": "000000"},
				Tables:     []string{"Records:A1:C3:TableStyleMedium2"},
			}
			if !reflect.DeepEqual(legacy, expected) {
				t.Fatalf("legacy JSON workbook violates the semantic contract\nexpected: %#v\nactual:   %#v", expected, legacy)
			}
			if !reflect.DeepEqual(legacy, wire) {
				t.Fatalf("legacy JSON and PFX2 differ semantically\nlegacy: %#v\nPFX2:  %#v", legacy, wire)
			}
		})
	}
}

// Corrupt one worker's segment while another worker has substantial valid
// input. The failing worker must cancel its peer and return an ordinary error,
// rather than leaving WriteExcelV2 blocked on the worker wait group.
func TestWriteExcelV2ParallelSegmentFailureReturnsPromptly(t *testing.T) {
	payload := newMultiSheetPFX2Payload(t, multiSheetTestRows(2, 2_000), nil)
	metadataLength := int(binary.BigEndian.Uint64(payload[len(wireMagic):wireHeaderSize]))
	rowStreamStart := wireHeaderSize + metadataLength
	payload[rowStreamStart] = 0xc1 // MessagePack's reserved, never-used code.

	result := make(chan error, 1)
	go func() {
		_, err := WriteExcelV2(payload)
		result <- err
	}()

	select {
	case err := <-result:
		if err == nil {
			t.Fatal("corrupt worker segment unexpectedly succeeded")
		}
		if !strings.Contains(err.Error(), `decode sheet "Sheet1" row 1`) {
			t.Fatalf("expected the first segment's decode error, got %v", err)
		}
	case <-time.After(5 * time.Second):
		t.Fatal("parallel PFX2 decode did not cancel workers after a segment failure")
	}
}

type workbookSemantics struct {
	Sheets     []string
	Rows       [][]string
	Formulas   map[string]string
	FontColors map[string]string
	Tables     []string
}

func newEquivalentLegacyAndPFX2Payloads(t testing.TB, engine string) ([]byte, []byte) {
	t.Helper()

	type cellSpec struct {
		value interface{}
		style string
	}
	logicalRows := [][]cellSpec{
		{{"Category", "accent"}, {"Amount", "accent"}, {"Computed", "accent"}},
		{{"Hardware", "DEFAULT_STYLE"}, {int64(7), "DEFAULT_STYLE"}, {"=B2*2", "DEFAULT_STYLE"}},
		{{"Software", "DEFAULT_STYLE"}, {uint64(11), "DEFAULT_STYLE"}, {"=B3*2", "DEFAULT_STYLE"}},
	}
	styleNames := []string{"DEFAULT_STYLE", "accent"}
	styleIndexes := map[string]uint32{"DEFAULT_STYLE": 0, "accent": 1}
	legacyRows := make([]interface{}, len(logicalRows))
	wireRows := make([]interface{}, len(logicalRows))
	for rowIndex, row := range logicalRows {
		legacyRow := make([]interface{}, len(row))
		wireRow := make([]interface{}, len(row))
		for columnIndex, cell := range row {
			legacyRow[columnIndex] = []interface{}{cell.value, cell.style}
			wireRow[columnIndex] = []interface{}{cell.value, styleIndexes[cell.style]}
		}
		legacyRows[rowIndex] = legacyRow
		wireRows[rowIndex] = wireRow
	}

	table := map[string]interface{}{
		"range":               "A1:C3",
		"name":                "Records",
		"style_name":          "TableStyleMedium2",
		"show_first_column":   false,
		"show_last_column":    false,
		"show_row_stripes":    true,
		"show_column_stripes": false,
	}
	newMetadata := func(rows []interface{}) map[string]interface{} {
		sheet := newStyledStreamWriterSheet(rows, []interface{}{})
		sheet["WriterEngine"] = engine
		sheet["Table"] = []interface{}{table}
		return map[string]interface{}{
			"style": map[string]interface{}{
				"DEFAULT_STYLE": testStyleDefinition("000000"),
				"accent":        testStyleDefinition("FF0000"),
			},
			"file_props":        newFileProps(),
			"protection":        map[string]interface{}{},
			"sheet_order":       []string{"Sheet1"},
			"content":           map[string]interface{}{"Sheet1": sheet},
			"_pyfastexcel_wire": map[string]interface{}{},
		}
	}

	legacyMetadata := newMetadata(legacyRows)
	delete(legacyMetadata, "_pyfastexcel_wire")
	legacyPayload, err := json.Marshal(legacyMetadata)
	if err != nil {
		t.Fatalf("marshal equivalent legacy metadata: %v", err)
	}

	wireMetadata := newMetadata([]interface{}{})
	wireMetadata["_pyfastexcel_wire"] = map[string]interface{}{
		"version":     wireVersion,
		"style_names": styleNames,
		"row_counts":  []int{len(wireRows)},
	}
	wireMetadataBytes, err := json.Marshal(wireMetadata)
	if err != nil {
		t.Fatalf("marshal equivalent PFX2 metadata: %v", err)
	}
	var rowStream bytes.Buffer
	encoder := msgpack.NewEncoder(&rowStream)
	for _, row := range wireRows {
		if err := encoder.Encode(row); err != nil {
			t.Fatalf("encode equivalent PFX2 row: %v", err)
		}
	}
	wirePayload := make([]byte, wireHeaderSize+len(wireMetadataBytes)+rowStream.Len())
	copy(wirePayload[:len(wireMagic)], wireMagic[:])
	binary.BigEndian.PutUint64(wirePayload[len(wireMagic):wireHeaderSize], uint64(len(wireMetadataBytes)))
	copy(wirePayload[wireHeaderSize:], wireMetadataBytes)
	copy(wirePayload[wireHeaderSize+len(wireMetadataBytes):], rowStream.Bytes())

	return legacyPayload, wirePayload
}

func readWorkbookSemantics(t testing.TB, data []byte) workbookSemantics {
	t.Helper()
	workbook, err := excelize.OpenReader(bytes.NewReader(data))
	if err != nil {
		t.Fatalf("open generated workbook: %v", err)
	}
	defer func() {
		if err := workbook.Close(); err != nil {
			t.Errorf("close generated workbook: %v", err)
		}
	}()

	rows, err := workbook.GetRows("Sheet1")
	if err != nil {
		t.Fatalf("read workbook rows: %v", err)
	}
	formulas := make(map[string]string)
	for _, cell := range []string{"C2", "C3"} {
		formula, err := workbook.GetCellFormula("Sheet1", cell)
		if err != nil {
			t.Fatalf("read formula %s: %v", cell, err)
		}
		formulas[cell] = formula
	}
	fontColors := make(map[string]string)
	for _, cell := range []string{"A1", "A2"} {
		styleID, err := workbook.GetCellStyle("Sheet1", cell)
		if err != nil {
			t.Fatalf("read style for %s: %v", cell, err)
		}
		style, err := workbook.GetStyle(styleID)
		if err != nil {
			t.Fatalf("resolve style for %s: %v", cell, err)
		}
		if style.Font == nil {
			t.Fatalf("style for %s has no font", cell)
		}
		fontColors[cell] = style.Font.Color
	}
	tables, err := workbook.GetTables("Sheet1")
	if err != nil {
		t.Fatalf("read tables: %v", err)
	}
	tableSemantics := make([]string, len(tables))
	for index, table := range tables {
		tableSemantics[index] = fmt.Sprintf("%s:%s:%s", table.Name, table.Range, table.StyleName)
	}

	return workbookSemantics{
		Sheets:     workbook.GetSheetList(),
		Rows:       rows,
		Formulas:   formulas,
		FontColors: fontColors,
		Tables:     tableSemantics,
	}
}
