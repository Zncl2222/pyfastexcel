package main

// #include <stdlib.h>
import (
	"C"
)
import (
	"fmt"
	"unsafe"

	"encoding/base64"
	"testing"

	"github.com/Zncl2222/pyfastexcel/pyfastexcel/core"
)

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
func Export(data *C.char) *C.char {
	goStringData := C.GoString(data)
	result := core.WriteExcel(goStringData)
	encodedRes := C.CString(result)
	return encodedRes
}

// FreeCPointer frees the memory allocated for a C char pointer.
//
// Args:
//
//	cptr (*C.char): The C char pointer to be freed.
//
//export FreeCPointer
func FreeCPointer(cptr *C.char) {
	C.free(unsafe.Pointer(cptr))
	fmt.Println("C Pointer Free Successfully !")
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
		"content": {
			"Sheet1": {
				"Header": [
					["Column1", "Column2", "Column3"]
				],
				"Data": [
					[["Data1", "style1"], ["Data2", "style1"], ["Data3", "style1"]],
					[["Data4", "style1"], ["Data5", "style1"], ["Data6", "style1"]]
				],
				"Height": {"3": 252},
				"Width": {"1": 25, "2": 26, "3": 6},
				"MergeCells": [["A1", "A2"], ["B2","C3"]],
				"AutoFilter": [],
				"Panes":      {},
				"NoStyle": "false"
			}
		}
	}`

	// Convert input data to *C.char
	cInputData := C.CString(inputData)

	// Call the Export function
	encodedExcel := Export(cInputData)

	// Free the allocated memory for cInputData
	FreeCPointer(cInputData)

	// Convert the result back to a Go string
	goEncodedExcel := C.GoString(encodedExcel)

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

func main() {
}
