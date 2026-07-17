package core

import (
	"bytes"
	"encoding/binary"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"sync"

	"github.com/vmihailenco/msgpack/v5"
	"github.com/vmihailenco/msgpack/v5/msgpcode"
	"github.com/xuri/excelize/v2"
)

const (
	wireVersion          = 2
	maxExcelRows         = 1_048_576
	maxExcelCols         = 16_384
	maxWireMetadataBytes = 64 << 20
	wireHeaderSize       = 12
)

var wireMagic = [4]byte{'P', 'F', 'X', '2'}

type wireConfiguration struct {
	Version    int      `json:"version"`
	StyleNames []string `json:"style_names"`
	RowCounts  []int    `json:"row_counts"`
	// SheetOffsets holds each sheet's starting byte offset into the row
	// stream. Optional: encoders that omit it (or older payloads) still
	// decode through the sequential path; when present it enables one
	// decoder per sheet so multi-sheet workbooks are written concurrently.
	SheetOffsets []int64 `json:"sheet_offsets"`
}

type wireMetadata struct {
	Wire wireConfiguration `json:"_pyfastexcel_wire"`
}

// WriteExcelV2 generates raw XLSX bytes from either the PFX2 wire format or
// the complete legacy JSON payload used by the debugging escape hatch.
func WriteExcelV2(payload []byte) (result []byte, err error) {
	defer recoverAsError(&err)

	writer, build, err := prepareWorkbookPayload(payload)
	if err != nil {
		return nil, err
	}
	defer func() {
		err = errors.Join(err, writer.File.Close())
	}()

	if err = build(); err != nil {
		return nil, err
	}
	return writer.writeToBytes()
}

// WriteExcelV2ToFile writes a workbook without routing ZIP bytes through the
// cgo boundary. The ZIP is completed in a private temporary file before
// the destination is opened, preserving the legacy save path on generation
// failure while retaining open(2) semantics for symlinks and existing files.
func WriteExcelV2ToFile(payload []byte, path string) (err error) {
	defer recoverAsError(&err)

	writer, build, err := prepareWorkbookPayload(payload)
	if err != nil {
		return err
	}
	writerOpen := true
	defer func() {
		if writerOpen {
			err = errors.Join(err, writer.File.Close())
		}
	}()

	if err = build(); err != nil {
		return err
	}
	temporaryPath, err := writer.writeToTemporary(path)
	if err != nil {
		return err
	}
	defer func() {
		if removeErr := os.Remove(temporaryPath); removeErr != nil && !errors.Is(removeErr, os.ErrNotExist) {
			err = errors.Join(err, fmt.Errorf("remove temporary output: %w", removeErr))
		}
	}()

	writerOpen = false
	if err := writer.File.Close(); err != nil {
		return fmt.Errorf("close generated workbook: %w", err)
	}
	return copyTemporaryWorkbook(temporaryPath, path)
}

func (ew *ExcelWriter) writeToTemporary(path string) (temporaryPath string, err error) {
	temporary, err := os.CreateTemp("", "pyfastexcel-*.tmp")
	if err != nil {
		return "", fmt.Errorf("create temporary output for %q: %w", path, err)
	}
	temporaryPath = temporary.Name()
	temporaryOpen := true
	keepTemporary := false
	defer func() {
		if temporaryOpen {
			err = errors.Join(err, temporary.Close())
		}
		if !keepTemporary {
			if removeErr := os.Remove(temporaryPath); removeErr != nil && !errors.Is(removeErr, os.ErrNotExist) {
				err = errors.Join(err, fmt.Errorf("remove temporary output: %w", removeErr))
			}
		}
	}()

	if writeErr := ew.writeTo(temporary); writeErr != nil {
		return "", writeErr
	}
	closeErr := temporary.Close()
	temporaryOpen = false
	if closeErr != nil {
		return "", fmt.Errorf("close temporary output: %w", closeErr)
	}
	keepTemporary = true
	return temporaryPath, nil
}

func copyTemporaryWorkbook(temporaryPath, path string) (err error) {
	input, err := os.Open(temporaryPath)
	if err != nil {
		return fmt.Errorf("reopen temporary output: %w", err)
	}
	defer func() {
		err = errors.Join(err, input.Close())
	}()

	output, err := os.OpenFile(filepath.Clean(path), os.O_WRONLY|os.O_CREATE|os.O_TRUNC, 0o600)
	if err != nil {
		return fmt.Errorf("open output file %q: %w", path, err)
	}
	defer func() {
		err = errors.Join(err, output.Close())
	}()

	if _, err := io.Copy(output, input); err != nil {
		return fmt.Errorf("copy completed workbook to %q: %w", path, err)
	}
	return nil
}

func prepareWorkbookPayload(payload []byte) (*ExcelWriter, func() error, error) {
	if !bytes.HasPrefix(payload, wireMagic[:]) {
		writer, err := newExcelWriter(payload)
		if err != nil {
			return nil, nil, err
		}
		return writer, writer.buildLegacyWorkbook, nil
	}

	if len(payload) < wireHeaderSize {
		return nil, nil, fmt.Errorf("PFX2 payload is truncated before metadata length")
	}
	metadataLength := binary.BigEndian.Uint64(payload[len(wireMagic):wireHeaderSize])
	if metadataLength > maxWireMetadataBytes {
		return nil, nil, fmt.Errorf(
			"PFX2 metadata length %d exceeds limit %d",
			metadataLength,
			maxWireMetadataBytes,
		)
	}
	remaining := len(payload) - wireHeaderSize
	if metadataLength > uint64(remaining) {
		return nil, nil, fmt.Errorf(
			"PFX2 metadata length %d exceeds remaining payload length %d",
			metadataLength,
			remaining,
		)
	}
	metadataEnd := wireHeaderSize + int(metadataLength)
	metadataBytes := payload[wireHeaderSize:metadataEnd]

	var metadata wireMetadata
	if err := json.Unmarshal(metadataBytes, &metadata); err != nil {
		return nil, nil, fmt.Errorf("decode PFX2 wire metadata: %w", err)
	}
	if metadata.Wire.Version != wireVersion {
		return nil, nil, fmt.Errorf(
			"unsupported PFX2 wire version %d (expected %d)",
			metadata.Wire.Version,
			wireVersion,
		)
	}

	writer, err := newExcelWriter(metadataBytes)
	if err != nil {
		return nil, nil, err
	}
	rowStream := payload[metadataEnd:]
	if err := validateWireMetadata(writer, metadata.Wire, int64(len(rowStream))); err != nil {
		_ = writer.File.Close()
		return nil, nil, err
	}

	writer.WireRowStream = rowStream
	decoder := msgpack.NewDecoder(bytes.NewReader(rowStream))
	build := func() error {
		return writer.buildWireWorkbook(decoder, metadata.Wire)
	}
	return writer, build, nil
}

func validateWireMetadata(writer *ExcelWriter, wire wireConfiguration, rowStreamLength int64) error {
	if len(wire.SheetOffsets) != 0 {
		if len(wire.SheetOffsets) != len(wire.RowCounts) {
			return fmt.Errorf(
				"PFX2 sheet_offsets has %d entries for %d row counts",
				len(wire.SheetOffsets),
				len(wire.RowCounts),
			)
		}
		previous := int64(0)
		for index, offset := range wire.SheetOffsets {
			if index == 0 && offset != 0 {
				return fmt.Errorf("PFX2 sheet_offsets must start at 0, got %d", offset)
			}
			if offset < previous || offset > rowStreamLength {
				return fmt.Errorf(
					"PFX2 sheet_offsets entry %d (%d) is out of order or out of bounds",
					index,
					offset,
				)
			}
			previous = offset
		}
	}
	if len(wire.RowCounts) != len(writer.SheetOrder) {
		return fmt.Errorf(
			"PFX2 row_counts has %d entries for %d sheets",
			len(wire.RowCounts),
			len(writer.SheetOrder),
		)
	}
	if len(wire.StyleNames) != len(writer.StyleMap) {
		return fmt.Errorf(
			"PFX2 style_names has %d entries for %d styles",
			len(wire.StyleNames),
			len(writer.StyleMap),
		)
	}
	if len(wire.StyleNames) > excelize.MaxCellStyles {
		return fmt.Errorf(
			"PFX2 style count %d exceeds Excel limit %d",
			len(wire.StyleNames),
			excelize.MaxCellStyles,
		)
	}
	if len(writer.Content) != len(writer.SheetOrder) {
		return fmt.Errorf(
			"PFX2 content has %d sheets but sheet_order has %d entries",
			len(writer.Content),
			len(writer.SheetOrder),
		)
	}

	seenSheets := make(map[string]struct{}, len(writer.SheetOrder))
	for index, item := range writer.SheetOrder {
		sheet, ok := item.(string)
		if !ok || sheet == "" {
			return fmt.Errorf("PFX2 sheet_order entry %d must be a non-empty string", index)
		}
		if _, duplicate := seenSheets[sheet]; duplicate {
			return fmt.Errorf("PFX2 sheet_order contains duplicate sheet %q", sheet)
		}
		seenSheets[sheet] = struct{}{}
		if wire.RowCounts[index] < 0 || wire.RowCounts[index] > maxExcelRows {
			return fmt.Errorf(
				"PFX2 sheet %q row count %d is outside Excel limits",
				sheet,
				wire.RowCounts[index],
			)
		}
		sheetData, ok := writer.Content[sheet].(map[string]interface{})
		if !ok {
			return fmt.Errorf("PFX2 metadata has no object for sheet %q", sheet)
		}
		engine, ok := sheetData["WriterEngine"].(string)
		if !ok || (engine != "StreamWriter" && engine != "NormalWriter") {
			return fmt.Errorf(
				"PFX2 sheet %q WriterEngine must be StreamWriter or NormalWriter",
				sheet,
			)
		}
		noStyle, ok := sheetData["NoStyle"].(bool)
		if !ok {
			return fmt.Errorf("PFX2 sheet %q NoStyle must be a boolean", sheet)
		}
		if !noStyle {
			if _, ok := writer.StyleMap["DEFAULT_STYLE"]; !ok {
				return fmt.Errorf("PFX2 styled metadata must define DEFAULT_STYLE")
			}
		}
		data, ok := sheetData["Data"].([]interface{})
		if !ok || len(data) != 0 {
			return fmt.Errorf("PFX2 metadata sheet %q Data must be an empty array", sheet)
		}
	}
	return nil
}

// wireRowResult carries one decoded row (or its decode error) from the
// decoder goroutine to the sheet-writing loop.
type wireRowResult struct {
	row []interface{}
	err error
}

// startWireRowDecoder decodes every row of every sheet, in payload order, on
// its own goroutine so MessagePack decoding overlaps excelize's row
// serialization. The decoder goroutine owns `decoder` exclusively: it stops at
// the first error, checks for trailing data after the final row, and always
// closes the channel. Closing `cancel` releases the goroutine early when the
// consumer bails out first.
func startWireRowDecoder(
	ew *ExcelWriter,
	decoder *msgpack.Decoder,
	wire wireConfiguration,
	noStyleBySheet []bool,
	cancel <-chan struct{},
) <-chan wireRowResult {
	results := make(chan wireRowResult, 256)
	go func() {
		defer close(results)
		for sheetIndex := range wire.RowCounts {
			noStyle := noStyleBySheet[sheetIndex]
			for rowIndex := 0; rowIndex < wire.RowCounts[sheetIndex]; rowIndex++ {
				row, err := ew.decodeWireRow(decoder, noStyle, nil)
				select {
				case results <- wireRowResult{row: row, err: err}:
				case <-cancel:
					return
				}
				if err != nil {
					return
				}
			}
		}
		var trailingErr error
		if _, err := decoder.PeekCode(); err == nil {
			trailingErr = fmt.Errorf("PFX2 payload contains trailing MessagePack data")
		} else if !errors.Is(err, io.EOF) {
			trailingErr = fmt.Errorf("check PFX2 payload end: %w", err)
		}
		if trailingErr != nil {
			select {
			case results <- wireRowResult{err: trailingErr}:
			case <-cancel:
			}
		}
	}()
	return results
}

func (ew *ExcelWriter) buildWireWorkbook(decoder *msgpack.Decoder, wire wireConfiguration) error {
	if err := ew.initializeStyles(wire.StyleNames); err != nil {
		return err
	}
	if err := ew.setFileProps(ew.FileProps); err != nil {
		return err
	}
	if len(ew.Protection) != 0 {
		if err := ew.setProtection(ew.Protection); err != nil {
			return err
		}
	}
	ew.markPivotSourceHeaders()

	// NoStyle is validated up front because the decoder goroutine needs the
	// complete schedule before the first sheet is written.
	noStyleBySheet := make([]bool, len(ew.SheetOrder))
	allStreamSheets := true
	for sheetIndex, item := range ew.SheetOrder {
		sheet := item.(string)
		sheetData := ew.Content[sheet].(map[string]interface{})
		noStyle, ok := sheetData["NoStyle"].(bool)
		if !ok {
			return fmt.Errorf("PFX2 sheet %q NoStyle must be a boolean", sheet)
		}
		noStyleBySheet[sheetIndex] = noStyle
		if sheetData["WriterEngine"] == "NormalWriter" {
			allStreamSheets = false
		}
	}

	// Multiple stream sheets can serialize rows concurrently (excelize
	// StreamWriter.SetRow only touches per-sheet state); sheet_offsets let
	// each worker decode its own slice of the row stream. Workbooks that
	// contain a NormalWriter sheet keep the sequential path because normal
	// writes go through shared *excelize.File methods.
	// PYFASTEXCEL_SEQUENTIAL=1 is a debugging escape hatch.
	if allStreamSheets &&
		len(ew.SheetOrder) > 1 &&
		len(wire.SheetOffsets) == len(ew.SheetOrder) &&
		os.Getenv("PYFASTEXCEL_SEQUENTIAL") == "" {
		return ew.buildWireSheetsParallel(wire, noStyleBySheet)
	}

	sheetCount := 1
	hasSheet1 := false
	for sheet := range ew.Content {
		if sheet == "Sheet1" {
			hasSheet1 = true
			break
		}
	}

	cancel := make(chan struct{})
	defer close(cancel)
	decodedRows := startWireRowDecoder(ew, decoder, wire, noStyleBySheet, cancel)
	nextRow := func(sheet string, rowIndex int) ([]interface{}, error) {
		result, ok := <-decodedRows
		if !ok {
			return nil, fmt.Errorf("decode sheet %q row %d: row stream ended early", sheet, rowIndex+1)
		}
		if result.err != nil {
			return nil, fmt.Errorf("decode sheet %q row %d: %w", sheet, rowIndex+1, result.err)
		}
		return result.row, nil
	}

	var pivotTableList [][]interface{}
	for sheetIndex, item := range ew.SheetOrder {
		sheet := item.(string)
		sheetData := ew.Content[sheet].(map[string]interface{})
		if !hasSheet1 && sheetCount == 1 {
			if err := ew.File.SetSheetName("Sheet1", sheet); err != nil {
				return fmt.Errorf("rename first sheet to %q: %w", sheet, err)
			}
			hasSheet1 = true
		} else {
			if _, err := ew.File.NewSheet(sheet); err != nil {
				return fmt.Errorf("create sheet %q: %w", sheet, err)
			}
			sheetCount++
		}

		rowCount := wire.RowCounts[sheetIndex]
		if sheetData["WriterEngine"] == "NormalWriter" {
			if err := ew.prepareNormalWrite(sheet, sheetData); err != nil {
				return err
			}
			for rowIndex := 0; rowIndex < rowCount; rowIndex++ {
				row, err := nextRow(sheet, rowIndex)
				if err != nil {
					return err
				}
				ew.capturePivotSourceHeader(sheet, rowIndex+1, row)
				if err := ew.writeDecodedNormalRow(sheet, rowIndex+1, row); err != nil {
					return err
				}
			}
			if err := ew.createTable(sheet, sheetData["Table"].([]interface{})); err != nil {
				return err
			}
		} else {
			streamWriter, rowHeightMap, err := ew.prepareStreamWrite(sheet, sheetData)
			if err != nil {
				return err
			}
			for rowIndex := 0; rowIndex < rowCount; rowIndex++ {
				row, err := nextRow(sheet, rowIndex)
				if err != nil {
					return err
				}
				ew.capturePivotSourceHeader(sheet, rowIndex+1, row)
				cell := "A" + strconv.Itoa(rowIndex+1)
				if rowHeight, ok := rowHeightMap[strconv.Itoa(rowIndex+1)]; ok {
					err = streamWriter.SetRow(cell, row, rowHeight)
				} else {
					err = streamWriter.SetRow(cell, row)
				}
				if err != nil {
					return fmt.Errorf("write stream sheet %q row %d: %w", sheet, rowIndex+1, err)
				}
			}
			if err := streamCreateTable(streamWriter, sheetData["Table"].([]interface{})); err != nil {
				return fmt.Errorf("create tables on stream sheet %q: %w", sheet, err)
			}
			if err := streamWriter.Flush(); err != nil {
				return fmt.Errorf("flush stream sheet %q: %w", sheet, err)
			}
		}

		pivotTableList = append(pivotTableList, sheetData["PivotTable"].([]interface{}))
		if err := ew.File.SetSheetVisible(sheet, sheetData["SheetVisible"].(bool)); err != nil {
			return fmt.Errorf("set visibility for sheet %q: %w", sheet, err)
		}
	}

	// A trailing-data error is queued after the final row; anything else on
	// the channel at this point is that verdict.
	if result, ok := <-decodedRows; ok && result.err != nil {
		return result.err
	}

	for _, pivots := range pivotTableList {
		if err := ew.seedPivotSourceHeaders(pivots); err != nil {
			return err
		}
		if err := ew.createPivotTable(pivots); err != nil {
			return err
		}
	}
	return nil
}

// wireParallelControl propagates the first failure across the decoder and
// sheet-writer goroutines and releases everything blocked on a channel.
type wireParallelControl struct {
	cancel chan struct{}
	once   sync.Once
	mu     sync.Mutex
	err    error
}

func newWireParallelControl() *wireParallelControl {
	return &wireParallelControl{cancel: make(chan struct{})}
}

func (control *wireParallelControl) fail(err error) {
	control.mu.Lock()
	if control.err == nil && err != nil {
		control.err = err
	}
	control.mu.Unlock()
	control.once.Do(func() { close(control.cancel) })
}

func (control *wireParallelControl) firstError() error {
	control.mu.Lock()
	defer control.mu.Unlock()
	return control.err
}

type preparedStreamSheet struct {
	name         string
	data         map[string]interface{}
	streamWriter *excelize.StreamWriter
	rowHeights   map[string]excelize.RowOpts
}

// buildWireSheetsParallel writes multi-sheet, all-StreamWriter workbooks with
// one worker goroutine per sheet, each decoding its own sheet_offsets slice
// of the row stream and serializing rows as it goes. Everything that touches
// shared *excelize.File state stays on this goroutine: sheet creation and
// preparation happen before the workers start, tables/Flush/visibility and
// pivot tables after they finish.
func (ew *ExcelWriter) buildWireSheetsParallel(
	wire wireConfiguration,
	noStyleBySheet []bool,
) error {
	hasSheet1 := false
	for sheet := range ew.Content {
		if sheet == "Sheet1" {
			hasSheet1 = true
			break
		}
	}

	prepared := make([]preparedStreamSheet, len(ew.SheetOrder))
	sheetCount := 1
	for sheetIndex, item := range ew.SheetOrder {
		sheet := item.(string)
		sheetData := ew.Content[sheet].(map[string]interface{})
		if !hasSheet1 && sheetCount == 1 {
			if err := ew.File.SetSheetName("Sheet1", sheet); err != nil {
				return fmt.Errorf("rename first sheet to %q: %w", sheet, err)
			}
			hasSheet1 = true
		} else {
			if _, err := ew.File.NewSheet(sheet); err != nil {
				return fmt.Errorf("create sheet %q: %w", sheet, err)
			}
			sheetCount++
		}
		streamWriter, rowHeightMap, err := ew.prepareStreamWrite(sheet, sheetData)
		if err != nil {
			return err
		}
		prepared[sheetIndex] = preparedStreamSheet{
			name:         sheet,
			data:         sheetData,
			streamWriter: streamWriter,
			rowHeights:   rowHeightMap,
		}
	}

	control := newWireParallelControl()
	stream := ew.WireRowStream
	var workers sync.WaitGroup
	for sheetIndex := range prepared {
		segmentStart := wire.SheetOffsets[sheetIndex]
		segmentEnd := int64(len(stream))
		if sheetIndex+1 < len(wire.SheetOffsets) {
			segmentEnd = wire.SheetOffsets[sheetIndex+1]
		}
		workers.Add(1)
		go func(sheet *preparedStreamSheet, segment []byte, rowCount int, noStyle bool) {
			defer workers.Done()
			ew.writeStreamSheetSegment(sheet, segment, rowCount, noStyle, control)
		}(
			&prepared[sheetIndex],
			stream[segmentStart:segmentEnd],
			wire.RowCounts[sheetIndex],
			noStyleBySheet[sheetIndex],
		)
	}
	workers.Wait()
	if err := control.firstError(); err != nil {
		return err
	}

	var pivotTableList [][]interface{}
	for index := range prepared {
		sheet := &prepared[index]
		if err := streamCreateTable(sheet.streamWriter, sheet.data["Table"].([]interface{})); err != nil {
			return fmt.Errorf("create tables on stream sheet %q: %w", sheet.name, err)
		}
		if err := sheet.streamWriter.Flush(); err != nil {
			return fmt.Errorf("flush stream sheet %q: %w", sheet.name, err)
		}
		pivotTableList = append(pivotTableList, sheet.data["PivotTable"].([]interface{}))
		if err := ew.File.SetSheetVisible(sheet.name, sheet.data["SheetVisible"].(bool)); err != nil {
			return fmt.Errorf("set visibility for sheet %q: %w", sheet.name, err)
		}
	}

	for _, pivots := range pivotTableList {
		if err := ew.seedPivotSourceHeaders(pivots); err != nil {
			return err
		}
		if err := ew.createPivotTable(pivots); err != nil {
			return err
		}
	}
	return nil
}

// writeStreamSheetSegment decodes one sheet's slice of the row stream and
// serializes its rows. Workers share no mutable state: the decoder and row
// buffer are worker-local, SetRow only touches per-sheet excelize state, and
// capturePivotSourceHeader only mutates this sheet's own header map.
func (ew *ExcelWriter) writeStreamSheetSegment(
	sheet *preparedStreamSheet,
	segment []byte,
	rowCount int,
	noStyle bool,
	control *wireParallelControl,
) {
	decoder := msgpack.NewDecoder(bytes.NewReader(segment))
	var rowBuffer []interface{}
	for rowIndex := 0; rowIndex < rowCount; rowIndex++ {
		select {
		case <-control.cancel:
			return
		default:
		}
		row, err := ew.decodeWireRow(decoder, noStyle, rowBuffer)
		if err != nil {
			control.fail(fmt.Errorf("decode sheet %q row %d: %w", sheet.name, rowIndex+1, err))
			return
		}
		rowBuffer = row
		ew.capturePivotSourceHeader(sheet.name, rowIndex+1, row)
		cell := "A" + strconv.Itoa(rowIndex+1)
		if rowHeight, ok := sheet.rowHeights[strconv.Itoa(rowIndex+1)]; ok {
			err = sheet.streamWriter.SetRow(cell, row, rowHeight)
		} else {
			err = sheet.streamWriter.SetRow(cell, row)
		}
		if err != nil {
			control.fail(fmt.Errorf("write stream sheet %q row %d: %w", sheet.name, rowIndex+1, err))
			return
		}
	}
	if _, err := decoder.PeekCode(); err == nil {
		control.fail(fmt.Errorf(
			"PFX2 sheet %q segment contains trailing MessagePack data",
			sheet.name,
		))
	} else if !errors.Is(err, io.EOF) {
		control.fail(fmt.Errorf("check PFX2 sheet %q segment end: %w", sheet.name, err))
	}
}

func (ew *ExcelWriter) decodeWireRow(
	decoder *msgpack.Decoder,
	noStyle bool,
	reuse []interface{},
) ([]interface{}, error) {
	columnCount, err := decoder.DecodeArrayLen()
	if err != nil {
		return nil, err
	}
	if columnCount < 0 || columnCount > maxExcelCols {
		return nil, fmt.Errorf("column count %d is outside Excel limits", columnCount)
	}

	var row []interface{}
	if columnCount <= cap(reuse) {
		clear(reuse)
		row = reuse[:columnCount]
	} else {
		row = make([]interface{}, columnCount)
	}
	for column := 0; column < columnCount; column++ {
		if noStyle {
			value, err := decodeWireScalar(decoder)
			if err != nil {
				return nil, fmt.Errorf("column %d: %w", column+1, err)
			}
			row[column] = value
			continue
		}
		cell, err := ew.decodeWireCell(decoder)
		if err != nil {
			return nil, fmt.Errorf("column %d: %w", column+1, err)
		}
		row[column] = cell
	}
	return row, nil
}

func (ew *ExcelWriter) decodeWireCell(decoder *msgpack.Decoder) (interface{}, error) {
	code, err := decoder.PeekCode()
	if err != nil {
		return nil, err
	}
	if code == msgpcode.Nil {
		if err := decoder.DecodeNil(); err != nil {
			return nil, err
		}
		return nil, nil
	}

	cellLength, err := decoder.DecodeArrayLen()
	if err != nil {
		return nil, fmt.Errorf("styled cell must be an array: %w", err)
	}
	if cellLength == 0 {
		styleID := ew.StyleIDs["DEFAULT_STYLE"]
		return excelize.Cell{StyleID: styleID, Value: ""}, nil
	}
	if cellLength != 2 {
		return nil, fmt.Errorf("styled cell must have 0 or 2 elements, got %d", cellLength)
	}

	value, err := decodeWireScalar(decoder)
	if err != nil {
		return nil, fmt.Errorf("decode cell value: %w", err)
	}
	wireStyleID, err := decodeWireStyleID(decoder)
	if err != nil {
		return nil, err
	}
	if wireStyleID >= uint64(len(ew.WireStyleIDs)) {
		return nil, fmt.Errorf(
			"style ID %d is out of range for %d styles",
			wireStyleID,
			len(ew.WireStyleIDs),
		)
	}
	styleID := ew.WireStyleIDs[wireStyleID]
	if stringValue, ok := value.(string); ok && strings.HasPrefix(stringValue, "=") {
		return excelize.Cell{StyleID: styleID, Formula: normalizeFormula(stringValue)}, nil
	}
	return excelize.Cell{StyleID: styleID, Value: value}, nil
}

func decodeWireStyleID(decoder *msgpack.Decoder) (uint64, error) {
	code, err := decoder.PeekCode()
	if err != nil {
		return 0, err
	}
	if !isIntegerCode(code) {
		return 0, fmt.Errorf("style ID must be an unsigned integer")
	}
	value, err := decoder.DecodeInterfaceLoose()
	if err != nil {
		return 0, err
	}
	switch value := value.(type) {
	case int64:
		if value < 0 {
			return 0, fmt.Errorf("style ID must not be negative")
		}
		return uint64(value), nil
	case uint64:
		return value, nil
	default:
		return 0, fmt.Errorf("style ID must be an unsigned integer, got %T", value)
	}
}

func decodeWireScalar(decoder *msgpack.Decoder) (interface{}, error) {
	code, err := decoder.PeekCode()
	if err != nil {
		return nil, err
	}
	switch {
	case code == msgpcode.Nil:
		return nil, decoder.DecodeNil()
	case code == msgpcode.False || code == msgpcode.True:
		return decoder.DecodeBool()
	case isIntegerCode(code):
		return decoder.DecodeInterfaceLoose()
	case code == msgpcode.Float || code == msgpcode.Double:
		return decoder.DecodeFloat64()
	case msgpcode.IsString(code):
		return decoder.DecodeString()
	default:
		return nil, fmt.Errorf("unsupported MessagePack scalar code 0x%02x", code)
	}
}

func isIntegerCode(code byte) bool {
	if msgpcode.IsFixedNum(code) {
		return true
	}
	switch code {
	case msgpcode.Uint8, msgpcode.Uint16, msgpcode.Uint32, msgpcode.Uint64,
		msgpcode.Int8, msgpcode.Int16, msgpcode.Int32, msgpcode.Int64:
		return true
	default:
		return false
	}
}

func (ew *ExcelWriter) writeDecodedNormalRow(sheet string, rowNumber int, row []interface{}) error {
	for column, item := range row {
		if item == nil {
			continue
		}
		cellName, err := excelize.CoordinatesToCellName(column+1, rowNumber)
		if err != nil {
			return fmt.Errorf("resolve sheet %q row %d column %d: %w", sheet, rowNumber, column+1, err)
		}
		if cell, ok := item.(excelize.Cell); ok {
			if cell.Formula != "" {
				err = ew.File.SetCellFormula(sheet, cellName, cell.Formula)
			} else {
				err = ew.File.SetCellValue(sheet, cellName, cell.Value)
			}
			if err != nil {
				return fmt.Errorf("write normal sheet %q cell %s: %w", sheet, cellName, err)
			}
			if err := ew.File.SetCellStyle(sheet, cellName, cellName, cell.StyleID); err != nil {
				return fmt.Errorf("style normal sheet %q cell %s: %w", sheet, cellName, err)
			}
			continue
		}
		if err := ew.File.SetCellValue(sheet, cellName, item); err != nil {
			return fmt.Errorf("write normal sheet %q cell %s: %w", sheet, cellName, err)
		}
	}
	return nil
}

func (ew *ExcelWriter) markPivotSourceHeaders() {
	ew.PivotSourceHeaders = make(map[string]map[int][]interface{})
	for _, content := range ew.Content {
		sheetData, ok := content.(map[string]interface{})
		if !ok {
			continue
		}
		pivots, ok := sheetData["PivotTable"].([]interface{})
		if !ok {
			continue
		}
		for _, pivot := range pivots {
			pivotMap, ok := pivot.(map[string]interface{})
			if !ok {
				continue
			}
			dataRange, ok := pivotMap["DataRange"].(string)
			if !ok {
				continue
			}
			sheet, _, row, ok := pivotRangeStart(dataRange)
			if !ok {
				continue
			}
			if ew.PivotSourceHeaders[sheet] == nil {
				ew.PivotSourceHeaders[sheet] = make(map[int][]interface{})
			}
			ew.PivotSourceHeaders[sheet][row] = nil
		}
	}
}

func (ew *ExcelWriter) capturePivotSourceHeader(sheet string, rowNumber int, row []interface{}) {
	rows, ok := ew.PivotSourceHeaders[sheet]
	if !ok {
		return
	}
	if _, wanted := rows[rowNumber]; !wanted {
		return
	}
	header := make([]interface{}, len(row))
	for index, value := range row {
		header[index] = getCellScalarValue(value)
	}
	rows[rowNumber] = header
}

func pivotRangeStart(dataRange string) (sheet string, column int, row int, ok bool) {
	if !strings.Contains(dataRange, "!") {
		return "", 0, 0, false
	}
	parts := strings.SplitN(dataRange, "!", 2)
	sheet = strings.Trim(parts[0], "'")
	cellRange := strings.ReplaceAll(parts[1], "$", "")
	cellRefs := strings.SplitN(cellRange, ":", 2)
	column, row, err := excelize.CellNameToCoordinates(cellRefs[0])
	if err != nil {
		return "", 0, 0, false
	}
	return sheet, column, row, true
}
