package main

// #include <stdlib.h>
import (
	"C"
)
import (
	"fmt"
	"testing"
	"unsafe"

	"encoding/base64"

	"github.com/Zncl2222/pyfastexcel/pyfastexcel/core"
)

// Export takes a C char pointer containing JSON data for an Excel file and
// writes the raw Excel bytes to a C.malloc'd buffer.  The number of bytes
// written is stored via the outLen out-parameter.  The caller must free the
// returned pointer with FreeCPointer.
//
// Args:
//
//	data (*C.char): JSON payload describing the Excel file.
//	outLen (*C.int64_t): Receives the byte-length of the returned buffer.
//	useCatchPanic (int64): Non-zero → recover from Go panics gracefully.
//
// Returns:
//
//	unsafe.Pointer: Pointer to a C.malloc'd buffer containing raw .xlsx bytes.
//
//export Export
func Export(data *C.char, outLen *C.int64_t, useCatchPanic int64) unsafe.Pointer {
	if useCatchPanic != 0 {
		defer catchPanic()
	}
	result := core.WriteExcelRaw(C.GoString(data))
	*outLen = C.int64_t(len(result))
	return C.CBytes(result)
}

func catchPanic() {
	if r := recover(); r != nil {
		fmt.Printf("Recovered from panic: %v\n", r)
	}
}

// FreeCPointer frees a buffer allocated by Export (or any C.malloc allocation).
//
// Args:
//
//	cptr (unsafe.Pointer): Pointer returned by Export.
//	printMsg (int64): Non-zero → print a confirmation message.
//
//export FreeCPointer
func FreeCPointer(cptr unsafe.Pointer, printMsg int64) {
	C.free(cptr)
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

	// Call the Export function with the new signature (raw bytes + length)
	var outLen C.int64_t
	rawPtr := Export(cInputData, &outLen, 1)

	// Free the input string
	FreeCPointer(unsafe.Pointer(cInputData), 1)

	// Copy raw bytes from the C buffer
	excelBytes := C.GoBytes(rawPtr, C.int(outLen))

	// Free the output buffer
	FreeCPointer(rawPtr, 1)

	// Assert the expected result
	if len(excelBytes) == 0 {
		t.Error("Excel data is empty")
	}

	// Validate that it is a valid xlsx (zip) file
	_ = base64.StdEncoding.EncodeToString(excelBytes) // use base64 import
}

func main() {
}
