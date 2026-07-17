package core

import (
	"archive/zip"
	"bytes"
	"encoding/binary"
	"encoding/json"
	"fmt"
	"io"
	"math"
	"os"
	"path/filepath"
	"runtime"
	"strings"
	"sync"
	"testing"

	"github.com/vmihailenco/msgpack/v5"
	"github.com/xuri/excelize/v2"
)

func TestWriteExcelV2RoundTripsSupportedRows(t *testing.T) {
	for _, engine := range []string{"StreamWriter", "NormalWriter"} {
		t.Run(engine, func(t *testing.T) {
			rows := []interface{}{
				[]interface{}{
					[]interface{}{"text", uint32(1)},
					[]interface{}{int64(-7), uint32(0)},
					[]interface{}{uint64(9), uint32(0)},
					[]interface{}{1.5, uint32(0)},
					[]interface{}{true, uint32(0)},
					[]interface{}{nil, uint32(0)},
					[]interface{}{"=SUM(B1:C1)", uint32(1)},
					[]interface{}{},
					nil,
					[]interface{}{int64(9_007_199_254_740_993), uint32(0)},
				},
			}
			payload := newPFX2TestPayload(t, engine, false, rows, nil)

			workbookBytes, err := WriteExcelV2(payload)
			if err != nil {
				t.Fatalf("WriteExcelV2 returned an error: %v", err)
			}
			if !bytes.HasPrefix(workbookBytes, []byte("PK")) {
				t.Fatalf("expected a ZIP workbook, got %x", workbookBytes[:2])
			}
			archive, err := zip.NewReader(bytes.NewReader(workbookBytes), int64(len(workbookBytes)))
			if err != nil {
				t.Fatalf("open generated ZIP: %v", err)
			}
			var worksheetXML []byte
			for _, file := range archive.File {
				if file.Name != "xl/worksheets/sheet1.xml" {
					continue
				}
				reader, err := file.Open()
				if err != nil {
					t.Fatalf("open worksheet XML: %v", err)
				}
				worksheetXML, err = io.ReadAll(reader)
				_ = reader.Close()
				if err != nil {
					t.Fatalf("read worksheet XML: %v", err)
				}
			}
			if !bytes.Contains(worksheetXML, []byte("<v>9007199254740993</v>")) {
				t.Fatalf("large integer lost precision in worksheet XML: %s", worksheetXML)
			}

			workbook, err := excelize.OpenReader(bytes.NewReader(workbookBytes))
			if err != nil {
				t.Fatalf("open generated workbook: %v", err)
			}
			defer workbook.Close()

			for cell, expected := range map[string]string{
				"A1": "text",
				"B1": "-7",
				"C1": "9",
				"D1": "1.5",
				"E1": "TRUE",
				"F1": "",
				"H1": "",
				"I1": "",
			} {
				actual, err := workbook.GetCellValue("Sheet1", cell)
				if err != nil {
					t.Fatalf("read %s: %v", cell, err)
				}
				if actual != expected {
					t.Errorf("%s: expected %q, got %q", cell, expected, actual)
				}
			}
			formula, err := workbook.GetCellFormula("Sheet1", "G1")
			if err != nil {
				t.Fatalf("read formula: %v", err)
			}
			if formula != "SUM(B1:C1)" {
				t.Fatalf("expected normalized formula, got %q", formula)
			}
			styleID, err := workbook.GetCellStyle("Sheet1", "A1")
			if err != nil {
				t.Fatalf("read style: %v", err)
			}
			style, err := workbook.GetStyle(styleID)
			if err != nil {
				t.Fatalf("read style definition: %v", err)
			}
			if style.Font == nil || style.Font.Color != "FF0000" {
				t.Fatalf("expected red font, got %#v", style.Font)
			}
		})
	}
}

func TestWriteExcelV2NoStyleAndJSONFallback(t *testing.T) {
	rows := []interface{}{
		[]interface{}{"plain", int64(42), true, nil, "=literal"},
	}
	payload := newPFX2TestPayload(t, "StreamWriter", true, rows, nil)
	workbookBytes, err := WriteExcelV2(payload)
	if err != nil {
		t.Fatalf("write no-style PFX2 workbook: %v", err)
	}
	workbook, err := excelize.OpenReader(bytes.NewReader(workbookBytes))
	if err != nil {
		t.Fatalf("open no-style workbook: %v", err)
	}
	defer workbook.Close()
	if formula, err := workbook.GetCellFormula("Sheet1", "E1"); err != nil || formula != "" {
		t.Fatalf("NoStyle formula-like string must remain literal, formula=%q err=%v", formula, err)
	}
	if value, _ := workbook.GetCellValue("Sheet1", "E1"); value != "=literal" {
		t.Fatalf("expected literal formula-like string, got %q", value)
	}

	legacyJSON := newLegacyTestPayload(t, "fallback", "00AA00", 1)
	if _, err := WriteExcelV2(legacyJSON); err != nil {
		t.Fatalf("ExportV2 JSON fallback returned an error: %v", err)
	}
}

func TestWriteExcelPreservesLegacyXORProtectionInput(t *testing.T) {
	payload := newLegacyTestPayload(t, "accent", "FF0000", 1)
	var metadata map[string]interface{}
	if err := json.Unmarshal(payload, &metadata); err != nil {
		t.Fatalf("decode legacy payload: %v", err)
	}
	metadata["protection"] = map[string]interface{}{
		"algorithm":      "XOR",
		"password":       "legacy-password",
		"lock_structure": true,
		"lock_windows":   false,
	}
	payload, err := json.Marshal(metadata)
	if err != nil {
		t.Fatalf("encode legacy payload: %v", err)
	}

	workbookBytes, err := WriteExcelBytes(string(payload))
	if err != nil {
		t.Fatalf("write XOR-compatible workbook: %v", err)
	}
	archive, err := zip.NewReader(bytes.NewReader(workbookBytes), int64(len(workbookBytes)))
	if err != nil {
		t.Fatalf("open generated workbook: %v", err)
	}
	var workbookXML []byte
	for _, file := range archive.File {
		if file.Name != "xl/workbook.xml" {
			continue
		}
		reader, err := file.Open()
		if err != nil {
			t.Fatalf("open workbook XML: %v", err)
		}
		workbookXML, err = io.ReadAll(reader)
		_ = reader.Close()
		if err != nil {
			t.Fatalf("read workbook XML: %v", err)
		}
		break
	}
	if len(workbookXML) == 0 {
		t.Fatal("generated workbook does not contain xl/workbook.xml")
	}
	if algorithm := xmlAttribute(workbookXML, "workbookAlgorithmName"); algorithm != "SHA-512" {
		t.Fatalf("legacy XOR input was not normalized, algorithm=%q", algorithm)
	}
	if hash := xmlAttribute(workbookXML, "workbookHashValue"); hash == "" {
		t.Fatal("normalized workbook protection has no password hash")
	}
}

func TestWriteExcelV2PropagatesInvalidTableErrors(t *testing.T) {
	styledRows := []interface{}{
		[]interface{}{[]interface{}{"header", uint32(0)}},
		[]interface{}{[]interface{}{"value", uint32(0)}},
	}
	tests := []struct {
		name    string
		payload []byte
		engine  string
		context string
	}{
		{
			name:    "legacy stream",
			payload: newLegacyTestPayload(t, "accent", "FF0000", 2),
			engine:  "StreamWriter",
			context: `add stream table 1 "1Table" over range "A1:A2"`,
		},
		{
			name:    "legacy normal",
			payload: newLegacyTestPayload(t, "accent", "FF0000", 2),
			engine:  "NormalWriter",
			context: `add table 1 "1Table" to sheet "Sheet1" over range "A1:A2"`,
		},
		{
			name:    "PFX2 stream",
			payload: newPFX2TestPayload(t, "StreamWriter", false, styledRows, nil),
			engine:  "StreamWriter",
			context: `add stream table 1 "1Table" over range "A1:A2"`,
		},
		{
			name:    "PFX2 normal",
			payload: newPFX2TestPayload(t, "NormalWriter", false, styledRows, nil),
			engine:  "NormalWriter",
			context: `add table 1 "1Table" to sheet "Sheet1" over range "A1:A2"`,
		},
	}

	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			payload := withInvalidTable(t, test.payload, test.engine)
			_, err := WriteExcelV2(payload)
			if err == nil {
				t.Fatal("invalid table unexpectedly succeeded")
			}
			for _, expected := range []string{test.context, `invalid name "1Table"`} {
				if !strings.Contains(err.Error(), expected) {
					t.Fatalf("expected error containing %q, got %v", expected, err)
				}
			}
		})
	}
}

func TestWriteExcelV2RejectsMalformedPayloads(t *testing.T) {
	validRows := []interface{}{[]interface{}{[]interface{}{"ok", uint32(0)}}}
	valid := newPFX2TestPayload(t, "StreamWriter", false, validRows, nil)

	tests := []struct {
		name    string
		payload func(t *testing.T) []byte
		match   string
	}{
		{
			name: "truncated header",
			payload: func(t *testing.T) []byte {
				return []byte("PFX2")
			},
			match: "truncated",
		},
		{
			name: "metadata length exceeds payload",
			payload: func(t *testing.T) []byte {
				payload := append([]byte(nil), valid[:12]...)
				binary.BigEndian.PutUint64(payload[4:12], uint64(len(valid)))
				return payload
			},
			match: "exceeds",
		},
		{
			name: "metadata length exceeds safety limit",
			payload: func(t *testing.T) []byte {
				payload := append([]byte(nil), valid[:12]...)
				binary.BigEndian.PutUint64(payload[4:12], math.MaxUint64)
				return payload
			},
			match: "limit",
		},
		{
			name: "wrong wire version",
			payload: func(t *testing.T) []byte {
				return newPFX2TestPayload(t, "StreamWriter", false, validRows, func(wire map[string]interface{}) {
					wire["version"] = 99
				})
			},
			match: "unsupported",
		},
		{
			name: "truncated row",
			payload: func(t *testing.T) []byte {
				return valid[:len(valid)-1]
			},
			match: "decode sheet",
		},
		{
			name: "style out of range",
			payload: func(t *testing.T) []byte {
				rows := []interface{}{[]interface{}{[]interface{}{"bad", uint32(200)}}}
				return newPFX2TestPayload(t, "StreamWriter", false, rows, nil)
			},
			match: "out of range",
		},
		{
			name: "negative style",
			payload: func(t *testing.T) []byte {
				rows := []interface{}{[]interface{}{[]interface{}{"bad", int64(-1)}}}
				return newPFX2TestPayload(t, "StreamWriter", false, rows, nil)
			},
			match: "negative",
		},
		{
			name: "nested cell value",
			payload: func(t *testing.T) []byte {
				rows := []interface{}{[]interface{}{[]interface{}{[]interface{}{1}, uint32(0)}}}
				return newPFX2TestPayload(t, "StreamWriter", false, rows, nil)
			},
			match: "unsupported MessagePack scalar",
		},
		{
			name: "bad cell arity",
			payload: func(t *testing.T) []byte {
				rows := []interface{}{[]interface{}{[]interface{}{"one"}}}
				return newPFX2TestPayload(t, "StreamWriter", false, rows, nil)
			},
			match: "0 or 2",
		},
		{
			name: "trailing msgpack",
			payload: func(t *testing.T) []byte {
				return append(append([]byte(nil), valid...), 0xc0)
			},
			match: "trailing",
		},
		{
			name: "unknown writer engine",
			payload: func(t *testing.T) []byte {
				return mutatePFX2TestMetadata(t, valid, func(metadata map[string]interface{}) {
					sheet := metadata["content"].(map[string]interface{})["Sheet1"].(map[string]interface{})
					sheet["WriterEngine"] = "UnknownWriter"
				})
			},
			match: "WriterEngine must be",
		},
		{
			name: "missing default style",
			payload: func(t *testing.T) []byte {
				return mutatePFX2TestMetadata(t, valid, func(metadata map[string]interface{}) {
					delete(metadata["style"].(map[string]interface{}), "DEFAULT_STYLE")
					wire := metadata["_pyfastexcel_wire"].(map[string]interface{})
					wire["style_names"] = []interface{}{"accent"}
				})
			},
			match: "must define DEFAULT_STYLE",
		},
		{
			name: "content and order mismatch",
			payload: func(t *testing.T) []byte {
				return mutatePFX2TestMetadata(t, valid, func(metadata map[string]interface{}) {
					content := metadata["content"].(map[string]interface{})
					content["Extra"] = content["Sheet1"]
				})
			},
			match: "content has 2 sheets",
		},
	}

	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			_, err := WriteExcelV2(test.payload(t))
			if err == nil || !strings.Contains(err.Error(), test.match) {
				t.Fatalf("expected error containing %q, got %v", test.match, err)
			}
		})
	}
}

func FuzzWriteExcelV2(f *testing.F) {
	valid := newPFX2TestPayload(
		f,
		"StreamWriter",
		false,
		[]interface{}{[]interface{}{[]interface{}{"seed", uint32(0)}}},
		nil,
	)
	f.Add(valid)
	f.Add([]byte("PFX2"))
	f.Add([]byte("{}"))
	f.Add([]byte{0xc1})

	f.Fuzz(func(t *testing.T, payload []byte) {
		if len(payload) > 1<<20 {
			t.Skip()
		}
		workbookBytes, err := WriteExcelV2(payload)
		if err != nil {
			return
		}
		if !bytes.HasPrefix(workbookBytes, []byte("PK")) {
			t.Fatalf("successful decoder returned non-ZIP bytes: %x", workbookBytes)
		}
		workbook, err := excelize.OpenReader(bytes.NewReader(workbookBytes))
		if err != nil {
			t.Fatalf("successful decoder returned an invalid workbook: %v", err)
		}
		if err := workbook.Close(); err != nil {
			t.Fatalf("close decoded workbook: %v", err)
		}
	})
}

func TestWriteExcelV2ToFileAllowsArbitraryExtension(t *testing.T) {
	payload := newPFX2TestPayload(
		t,
		"StreamWriter",
		false,
		[]interface{}{[]interface{}{[]interface{}{"file", uint32(0)}}},
		nil,
	)
	path := filepath.Join(t.TempDir(), "工作簿.without-xlsx")
	if err := WriteExcelV2ToFile(payload, path); err != nil {
		t.Fatalf("WriteExcelV2ToFile returned an error: %v", err)
	}
	workbook, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatalf("open direct output: %v", err)
	}
	defer workbook.Close()
	if value, _ := workbook.GetCellValue("Sheet1", "A1"); value != "file" {
		t.Fatalf("expected file cell value, got %q", value)
	}
}

func TestWriteExcelV2ToFilePreservesLegacyDestinationSemantics(t *testing.T) {
	payload := newPFX2TestPayload(
		t,
		"StreamWriter",
		false,
		[]interface{}{[]interface{}{[]interface{}{"file", uint32(0)}}},
		nil,
	)
	directory := t.TempDir()
	target := filepath.Join(directory, "existing.data")
	if err := os.WriteFile(target, []byte("original"), 0o600); err != nil {
		t.Fatal(err)
	}

	if err := WriteExcelV2ToFile([]byte("invalid payload"), target); err == nil {
		t.Fatal("invalid generation unexpectedly succeeded")
	}
	if content, err := os.ReadFile(target); err != nil || string(content) != "original" {
		t.Fatalf("generation failure changed target: content=%q err=%v", content, err)
	}

	if err := WriteExcelV2ToFile(payload, target); err != nil {
		t.Fatalf("overwrite existing target: %v", err)
	}
	info, err := os.Stat(target)
	if err != nil {
		t.Fatal(err)
	}
	if info.Mode().Perm() != 0o600 {
		t.Fatalf("existing mode changed to %o", info.Mode().Perm())
	}

	if runtime.GOOS == "windows" {
		return
	}
	realTarget := filepath.Join(directory, "real-target")
	linkTarget := filepath.Join(directory, "linked-output")
	if err := os.WriteFile(realTarget, []byte("old"), 0o600); err != nil {
		t.Fatal(err)
	}
	if err := os.Symlink(realTarget, linkTarget); err != nil {
		t.Fatal(err)
	}
	if err := WriteExcelV2ToFile(payload, linkTarget); err != nil {
		t.Fatalf("write through symlink: %v", err)
	}
	linkInfo, err := os.Lstat(linkTarget)
	if err != nil {
		t.Fatal(err)
	}
	if linkInfo.Mode()&os.ModeSymlink == 0 {
		t.Fatal("direct export replaced the destination symlink")
	}
	if workbook, err := excelize.OpenFile(realTarget); err != nil {
		t.Fatalf("symlink target is not a workbook: %v", err)
	} else {
		_ = workbook.Close()
	}
}

func TestWriteExcelConcurrentStyleIsolation(t *testing.T) {
	type workload struct {
		payload []byte
		color   string
	}
	workloads := []workload{
		{newLegacyTestPayload(t, "red-style", "FF0000", 300), "FF0000"},
		{newLegacyTestPayload(t, "blue-style", "0000FF", 300), "0000FF"},
	}

	const goroutines = 12
	start := make(chan struct{})
	errorsByWorker := make(chan error, goroutines)
	var wait sync.WaitGroup
	for worker := 0; worker < goroutines; worker++ {
		workload := workloads[worker%len(workloads)]
		wait.Add(1)
		go func() {
			defer wait.Done()
			<-start
			workbookBytes, err := WriteExcelBytes(string(workload.payload))
			if err != nil {
				errorsByWorker <- err
				return
			}
			workbook, err := excelize.OpenReader(bytes.NewReader(workbookBytes))
			if err != nil {
				errorsByWorker <- err
				return
			}
			defer workbook.Close()
			styleID, err := workbook.GetCellStyle("Sheet1", "A1")
			if err != nil {
				errorsByWorker <- err
				return
			}
			style, err := workbook.GetStyle(styleID)
			if err != nil {
				errorsByWorker <- err
				return
			}
			if style.Font == nil || style.Font.Color != workload.color {
				errorsByWorker <- fmt.Errorf("expected color %s, got %#v", workload.color, style.Font)
			}
		}()
	}
	close(start)
	wait.Wait()
	close(errorsByWorker)
	for err := range errorsByWorker {
		if err != nil {
			t.Error(err)
		}
	}
}

func newPFX2TestPayload(
	t testing.TB,
	engine string,
	noStyle bool,
	rows []interface{},
	mutateWire func(map[string]interface{}),
) []byte {
	t.Helper()
	sheet := newStreamWriterSheet([]interface{}{}, []interface{}{})
	sheet["NoStyle"] = noStyle
	sheet["WriterEngine"] = engine
	styles := map[string]interface{}{
		"DEFAULT_STYLE": testStyleDefinition("000000"),
		"accent":        testStyleDefinition("FF0000"),
	}
	wire := map[string]interface{}{
		"version":     wireVersion,
		"style_names": []string{"DEFAULT_STYLE", "accent"},
		"row_counts":  []int{len(rows)},
	}
	if mutateWire != nil {
		mutateWire(wire)
	}
	metadata := map[string]interface{}{
		"style":             styles,
		"file_props":        newFileProps(),
		"protection":        map[string]interface{}{},
		"sheet_order":       []string{"Sheet1"},
		"content":           map[string]interface{}{"Sheet1": sheet},
		"_pyfastexcel_wire": wire,
	}
	metadataBytes, err := json.Marshal(metadata)
	if err != nil {
		t.Fatalf("marshal PFX2 metadata: %v", err)
	}

	var encodedRows bytes.Buffer
	encoder := msgpack.NewEncoder(&encodedRows)
	for _, row := range rows {
		if err := encoder.Encode(row); err != nil {
			t.Fatalf("encode test row: %v", err)
		}
	}

	payload := make([]byte, wireHeaderSize+len(metadataBytes)+encodedRows.Len())
	copy(payload[:4], wireMagic[:])
	binary.BigEndian.PutUint64(payload[4:12], uint64(len(metadataBytes)))
	copy(payload[12:], metadataBytes)
	copy(payload[12+len(metadataBytes):], encodedRows.Bytes())
	return payload
}

func mutatePFX2TestMetadata(
	t *testing.T,
	payload []byte,
	mutate func(map[string]interface{}),
) []byte {
	t.Helper()
	metadataLength := int(binary.BigEndian.Uint64(payload[4:wireHeaderSize]))
	metadataEnd := wireHeaderSize + metadataLength
	var metadata map[string]interface{}
	if err := json.Unmarshal(payload[wireHeaderSize:metadataEnd], &metadata); err != nil {
		t.Fatalf("decode test PFX2 metadata: %v", err)
	}
	mutate(metadata)
	metadataBytes, err := json.Marshal(metadata)
	if err != nil {
		t.Fatalf("encode test PFX2 metadata: %v", err)
	}

	result := make([]byte, wireHeaderSize+len(metadataBytes)+len(payload[metadataEnd:]))
	copy(result[:4], wireMagic[:])
	binary.BigEndian.PutUint64(result[4:wireHeaderSize], uint64(len(metadataBytes)))
	copy(result[wireHeaderSize:], metadataBytes)
	copy(result[wireHeaderSize+len(metadataBytes):], payload[metadataEnd:])
	return result
}

func newLegacyTestPayload(t *testing.T, styleName, color string, rowCount int) []byte {
	t.Helper()
	rows := make([]interface{}, rowCount)
	for index := range rows {
		rows[index] = []interface{}{[]interface{}{fmt.Sprintf("row-%d", index), styleName}}
	}
	sheet := newStyledStreamWriterSheet(rows, []interface{}{})
	metadata := map[string]interface{}{
		"style": map[string]interface{}{
			"DEFAULT_STYLE": testStyleDefinition("000000"),
			styleName:       testStyleDefinition(color),
		},
		"file_props":  newFileProps(),
		"protection":  map[string]interface{}{},
		"sheet_order": []string{"Sheet1"},
		"content":     map[string]interface{}{"Sheet1": sheet},
	}
	payload, err := json.Marshal(metadata)
	if err != nil {
		t.Fatalf("marshal legacy test payload: %v", err)
	}
	return payload
}

func withInvalidTable(t *testing.T, payload []byte, engine string) []byte {
	t.Helper()
	isPFX2 := bytes.HasPrefix(payload, wireMagic[:])
	metadataBytes := payload
	var rowBytes []byte
	if isPFX2 {
		metadataLength := int(binary.BigEndian.Uint64(payload[4:wireHeaderSize]))
		metadataEnd := wireHeaderSize + metadataLength
		metadataBytes = payload[wireHeaderSize:metadataEnd]
		rowBytes = payload[metadataEnd:]
	}

	var metadata map[string]interface{}
	if err := json.Unmarshal(metadataBytes, &metadata); err != nil {
		t.Fatalf("decode test metadata: %v", err)
	}
	sheet := metadata["content"].(map[string]interface{})["Sheet1"].(map[string]interface{})
	sheet["WriterEngine"] = engine
	sheet["Table"] = []interface{}{
		map[string]interface{}{
			"range":               "A1:A2",
			"name":                "1Table",
			"style_name":          "",
			"show_first_column":   false,
			"show_last_column":    false,
			"show_row_stripes":    true,
			"show_column_stripes": false,
		},
	}
	metadataBytes, err := json.Marshal(metadata)
	if err != nil {
		t.Fatalf("encode test metadata: %v", err)
	}
	if !isPFX2 {
		return metadataBytes
	}

	result := make([]byte, wireHeaderSize+len(metadataBytes)+len(rowBytes))
	copy(result[:4], wireMagic[:])
	binary.BigEndian.PutUint64(result[4:wireHeaderSize], uint64(len(metadataBytes)))
	copy(result[wireHeaderSize:], metadataBytes)
	copy(result[wireHeaderSize+len(metadataBytes):], rowBytes)
	return result
}

func xmlAttribute(document []byte, name string) string {
	marker := []byte(name + `="`)
	start := bytes.Index(document, marker)
	if start < 0 {
		return ""
	}
	value := document[start+len(marker):]
	end := bytes.IndexByte(value, '"')
	if end < 0 {
		return ""
	}
	return string(value[:end])
}

func testStyleDefinition(color string) map[string]interface{} {
	return map[string]interface{}{
		"Font": map[string]interface{}{
			"Color": color,
		},
		"Fill": map[string]interface{}{
			"Type":    "pattern",
			"Color":   "#FFFFFF",
			"Pattern": 0,
			"Shading": 0,
		},
		"Border":    map[string]interface{}{},
		"Alignment": map[string]interface{}{},
		"Protection": map[string]interface{}{
			"Hidden": false,
			"Locked": false,
		},
		"CustomNumFmt": "general",
	}
}
