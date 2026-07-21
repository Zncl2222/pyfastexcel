package main

// #include <stdlib.h>
import (
	"C"
)
import (
	"bytes"
	"encoding/binary"
	"fmt"
	"io/fs"
	"math"
	"os"
	"path/filepath"
	"strings"
	"unsafe"

	"encoding/base64"
	"testing"

	"github.com/Zncl2222/pyfastexcel/pyfastexcel/core"
)

const maxV2ErrorBytes = 4096

// Export takes a C char pointer containing JSON data for an Excel file and returns a base64 encoded string of the generated Excel file.
//
// Args:
//
//	data (*C.char): A C char pointer containing JSON data for the Excel file.
//
// Returns:
//
//	*C.char: A C char pointer containing the base64 encoded string of the generated Excel file.
//
// Notes:
//   - This function does not directly interact with C code.
//   - Remember to free the memory allocated for the returned pointer using `C.free`.
//
//export Export
func Export(data *C.char, useCatchPanic int64) *C.char {
	if useCatchPanic != 0 {
		defer catchPanic()
	}
	goStringData := C.GoString(data)
	result := core.WriteExcel(goStringData)
	encodedRes := C.CString(result)
	return encodedRes
}

// GetABIVersion reports the highest supported pyfastexcel C ABI version.
//
//export GetABIVersion
func GetABIVersion() int64 {
	return 2
}

// ExportV2 accepts a length-delimited PFX2 or legacy JSON payload and returns
// raw XLSX bytes allocated by C. The caller owns the returned allocation and
// must release it with FreeCPointer.
//
//export ExportV2
func ExportV2(
	data unsafe.Pointer,
	dataLen C.size_t,
	useCatchPanic int64,
	outLen *C.size_t,
	outError **C.char,
) (result unsafe.Pointer) {
	_ = useCatchPanic // ABI compatibility: v2 always converts panics to errors.
	initializeV2Outputs(outLen, outError)
	defer func() {
		if recovered := recover(); recovered != nil {
			if result != nil {
				C.free(result)
				result = nil
			}
			if outLen != nil {
				*outLen = 0
			}
			setV2Error(outError, fmt.Errorf("pyfastexcel panic: %v", recovered))
		}
	}()

	if outLen == nil {
		setV2Error(outError, fmt.Errorf("output length pointer must not be NULL"))
		return nil
	}
	payload, err := copyV2Payload(data, dataLen)
	if err != nil {
		setV2Error(outError, err)
		return nil
	}
	workbook, err := core.WriteExcelV2(payload)
	if err != nil {
		setV2Error(outError, err)
		return nil
	}
	if len(workbook) == 0 {
		setV2Error(outError, fmt.Errorf("generated workbook is empty"))
		return nil
	}
	result = C.CBytes(workbook)
	if result == nil {
		setV2Error(outError, fmt.Errorf("allocate C workbook buffer"))
		return nil
	}
	*outLen = C.size_t(len(workbook))
	return result
}

// ExportToFileV2 writes a PFX2 or legacy JSON payload to a local path. It
// returns zero on success and a non-zero status with a C-owned error string on
// failure.
//
//export ExportToFileV2
func ExportToFileV2(
	data unsafe.Pointer,
	dataLen C.size_t,
	path *C.char,
	useCatchPanic int64,
	outError **C.char,
) (status int64) {
	_ = useCatchPanic // ABI compatibility: v2 always converts panics to errors.
	initializeV2Outputs(nil, outError)
	status = 1
	defer func() {
		if recovered := recover(); recovered != nil {
			setV2Error(outError, fmt.Errorf("pyfastexcel panic: %v", recovered))
			status = 1
		}
	}()

	if path == nil {
		setV2Error(outError, fmt.Errorf("output path must not be NULL"))
		return status
	}
	payload, err := copyV2Payload(data, dataLen)
	if err != nil {
		setV2Error(outError, err)
		return status
	}
	if err := core.WriteExcelV2ToFile(payload, C.GoString(path)); err != nil {
		setV2Error(outError, err)
		return status
	}
	return 0
}

func copyV2Payload(data unsafe.Pointer, dataLen C.size_t) ([]byte, error) {
	if data == nil && dataLen != 0 {
		return nil, fmt.Errorf("payload pointer is NULL for %d bytes", uint64(dataLen))
	}
	if uint64(dataLen) > math.MaxInt32 {
		return nil, fmt.Errorf("payload length %d exceeds cgo copy limit", uint64(dataLen))
	}
	if dataLen == 0 {
		return []byte{}, nil
	}
	return C.GoBytes(data, C.int(dataLen)), nil
}

func initializeV2Outputs(outLen *C.size_t, outError **C.char) {
	if outLen != nil {
		*outLen = 0
	}
	if outError != nil {
		*outError = nil
	}
}

func setV2Error(outError **C.char, err error) {
	if outError == nil || err == nil {
		return
	}
	if *outError != nil {
		C.free(unsafe.Pointer(*outError))
	}
	message := strings.ReplaceAll(err.Error(), "\x00", "\\x00")
	if len(message) > maxV2ErrorBytes {
		message = message[:maxV2ErrorBytes-3] + "..."
	}
	*outError = C.CString(message)
}

func catchPanic() {
	if r := recover(); r != nil {
		fmt.Printf("Recovered from panic: %v\n", r)
	}
}

// FreeCPointer frees the memory allocated for a C char pointer.
//
// Args:
//
//	cptr (*C.char): The C char pointer to be freed.
//	printMsg (int64): The flag to print message.
//
//export FreeCPointer
func FreeCPointer(cptr *C.char, printMsg int64) {
	C.free(unsafe.Pointer(cptr))
	if printMsg == 1 {
		fmt.Println("C Pointer Free Successfully !")
	}
}

// testExport is a trick method to test cgo in golang standard test module
func testExport(t *testing.T) {
	// Mock input data
	inputData := `{
		"style": {
			"style1": {
				"Font": {
					"Bold": true
				},
				"Fill": {
					"Type":    "pattern",
					"Color":   "#FFFFFF",
					"Pattern": 1,
					"Shading": 100
				},
				"Border": {
					"left": {
						"Color": "FF0000",
						"Style": 1
					},
					"top": {
						"Color": "00FF00",
						"Style": 2
					}
				},
				"Alignment": {
					"Horizontal":      "center",
					"Vertical":        "middle",
					"Indent":          0,
					"JustifyLastLine": false,
					"ReadingOrder":    0,
					"RelativeIndent":  0,
					"ShrinkToFit":     false,
					"TextRotation":    0,
					"WrapText":        false
				},
				"Protection": {
					"Hidden": true,
					"Locked": false
				},
				"CustomNumFmt": "0.00"
			}
		},
		"protection": {},
		"file_props": {
			"Title": "Test Excel File",
			"Creator": "Test User",
			"Category": "Test Category",
			"ContentStatus": "Draft",
			"Description": "Test Description",
			"Keywords": "Test Keywords",
			"Language": "en-US",
			"LastModifiedBy": "Test User",
			"Revision": "1",
			"Subject": "Test Subject",
			"Version": "1.0",
			"Identifier": "",
			"Created": "",
			"Modified": ""
		},
		"sheet_order": ["Sheet1"],
		"content": {
			"Sheet1": {
				"Header": [
					["Column1", "Column2", "Column3"]
				],
				"Data": [
					[["Data1", "style1"], ["Data2", "style1"], ["Data3", "style1"], []],
					[["Data4", "style1"], ["Data5", "style1"], ["Data6", "style1"], []]
				],
				"Height": {"3": 252},
				"Width": {"1": 25, "2": 26, "3": 6},
				"MergeCells": [["A1", "A2"], ["B2","C3"]],
				"AutoFilter": [],
				"Panes":      {},
				"DataValidation": [],
				"Comment":    [],
				"NoStyle": false,
				"Table": [],
				"Chart": [],
				"PivotTable": [],
				"SheetVisible": true,
				"WriterEngine": "StreamWriter"
			}
		}
	}`

	// Convert input data to *C.char
	cInputData := C.CString(inputData)

	// Call the Export function
	encodedExcel := Export(cInputData, 1)

	// Free the allocated memory for cInputData
	FreeCPointer(cInputData, 1)

	// Convert the result back to a Go string
	goEncodedExcel := C.GoString(encodedExcel)
	defer FreeCPointer(encodedExcel, 0)

	// Decode the encoded Excel data
	decodedExcel, err := base64.StdEncoding.DecodeString(goEncodedExcel)
	if err != nil {
		t.Fatalf("Failed to decode encoded Excel data: %v", err)
	}

	// Assert the expected result
	if len(decodedExcel) == 0 {
		t.Error("Encoded Excel data is empty")
	}
}

func testExportV2(t *testing.T) {
	if version := GetABIVersion(); version != 2 {
		t.Fatalf("expected ABI version 2, got %d", version)
	}

	input := abiTestPFX2()
	cInput := C.CBytes(input)
	defer C.free(cInput)
	var outputLength C.size_t
	var outputError *C.char
	output := ExportV2(
		cInput,
		C.size_t(len(input)),
		1,
		&outputLength,
		&outputError,
	)
	if outputError != nil {
		defer FreeCPointer(outputError, 0)
		t.Fatalf("ExportV2 returned an error: %s", C.GoString(outputError))
	}
	if output == nil || outputLength == 0 {
		t.Fatal("ExportV2 returned an empty workbook")
	}
	defer FreeCPointer((*C.char)(output), 0)
	workbook := C.GoBytes(output, C.int(outputLength))
	if !bytes.HasPrefix(workbook, []byte("PK")) {
		t.Fatalf("ExportV2 did not return a ZIP workbook: %x", workbook[:2])
	}
	if bytes.IndexByte(workbook, 0) < 0 {
		t.Fatal("test workbook unexpectedly contains no NUL bytes")
	}

	invalid := []byte("not valid workbook metadata")
	cInvalid := C.CBytes(invalid)
	defer C.free(cInvalid)
	outputLength = 123
	outputError = nil
	invalidOutput := ExportV2(
		cInvalid,
		C.size_t(len(invalid)),
		1,
		&outputLength,
		&outputError,
	)
	if invalidOutput != nil {
		FreeCPointer((*C.char)(invalidOutput), 0)
		t.Fatal("invalid ExportV2 payload returned a workbook")
	}
	if outputLength != 0 || outputError == nil {
		t.Fatalf("invalid ExportV2 payload returned length=%d error=%v", outputLength, outputError)
	}
	FreeCPointer(outputError, 0)
}

func testV2ErrorBounds(t *testing.T) {
	var outputError *C.char
	setV2Error(
		&outputError,
		fmt.Errorf("bad\x00%s", strings.Repeat("x", maxV2ErrorBytes+100)),
	)
	if outputError == nil {
		t.Fatal("setV2Error returned no allocation")
	}
	defer FreeCPointer(outputError, 0)
	message := C.GoString(outputError)
	if len(message) > maxV2ErrorBytes {
		t.Fatalf("error message has %d bytes, maximum is %d", len(message), maxV2ErrorBytes)
	}
	if strings.ContainsRune(message, '\x00') {
		t.Fatal("error message contains an embedded NUL")
	}
}

func testExportToFileV2(t *testing.T) {
	input := []byte(abiTestJSON)
	cInput := C.CBytes(input)
	defer C.free(cInput)
	directory := t.TempDir()
	const outputName = "abi-output.no-xlsx-extension"
	path := filepath.Join(directory, outputName)
	cPath := C.CString(path)
	defer C.free(unsafe.Pointer(cPath))
	var outputError *C.char
	status := ExportToFileV2(
		cInput,
		C.size_t(len(input)),
		cPath,
		1,
		&outputError,
	)
	if outputError != nil {
		defer FreeCPointer(outputError, 0)
		t.Fatalf("ExportToFileV2 returned an error: %s", C.GoString(outputError))
	}
	if status != 0 {
		t.Fatalf("ExportToFileV2 returned status %d", status)
	}
	workbook, err := fs.ReadFile(os.DirFS(directory), outputName)
	if err != nil {
		t.Fatalf("read ExportToFileV2 output: %v", err)
	}
	if !bytes.HasPrefix(workbook, []byte("PK")) {
		t.Fatalf("ExportToFileV2 did not write a ZIP workbook: %x", workbook[:2])
	}
}

const abiTestJSON = `{
  "style": {},
  "protection": {},
  "file_props": {
    "Title": "ABI Test", "Creator": "pyfastexcel", "Category": "",
    "ContentStatus": "", "Description": "", "Keywords": "", "Language": "en-US",
    "LastModifiedBy": "", "Revision": "0", "Subject": "", "Version": "",
    "Identifier": "xlsx", "Created": "", "Modified": ""
  },
  "sheet_order": ["Sheet1"],
  "content": {
    "Sheet1": {
      "Data": [["raw", 42, true, null]],
      "Height": {}, "Width": {}, "MergeCells": [], "AutoFilter": [], "Panes": {},
      "DataValidation": [], "Comment": [], "NoStyle": true, "Table": [], "Chart": [],
      "PivotTable": [], "SheetVisible": true, "WriterEngine": "StreamWriter"
    }
  }
}`

func abiTestPFX2() []byte {
	metadata := strings.Replace(
		abiTestJSON,
		`"Data": [["raw", 42, true, null]]`,
		`"Data": []`,
		1,
	)
	metadata = strings.Replace(
		metadata,
		`"content": {`,
		`"_pyfastexcel_wire": {"version": 2, "style_names": [], "row_counts": [1]}, "content": {`,
		1,
	)
	row := []byte{0x94, 0xa3, 'r', 'a', 'w', 0x2a, 0xc3, 0xc0}
	payload := make([]byte, 12+len(metadata)+len(row))
	copy(payload[:4], "PFX2")
	binary.BigEndian.PutUint64(payload[4:12], uint64(len(metadata)))
	copy(payload[12:], metadata)
	copy(payload[12+len(metadata):], row)
	return payload
}

func main() {
}
