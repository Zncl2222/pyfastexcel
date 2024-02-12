package main

// #include <stdlib.h>
import (
	"C"
)
import (
	"fmt"
	"unsafe"

	"github.com/Zncl2222/pyfastexcel/pyfastexcel/core"
)

//export Export
func Export(data *C.char) *C.char {
	goStringData := C.GoString(data)
	result := core.WriteExcel(goStringData)
	encodedRes := C.CString(result)
	return encodedRes
}

//export FreeCPointer
func FreeCPointer(cptr *C.char) {
	C.free(unsafe.Pointer(cptr))
	fmt.Println("C Pointer Free Successfully !")
}

func main() {
}
