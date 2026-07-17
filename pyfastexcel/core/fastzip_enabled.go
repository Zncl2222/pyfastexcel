//go:build pfx_fastzip

package core

import (
	"archive/zip"
	"io"
	"sync"
	_ "unsafe" // required for go:linkname

	kpflate "github.com/klauspost/compress/flate"
)

// archive/zip pre-registers the Deflate method and panics on re-registration,
// and excelize offers no hook to swap the compressor on the zip.Writer it
// creates internally. Storing into the package-level registry directly is the
// only way to substitute a faster DEFLATE without forking excelize. This
// requires building with -ldflags=-checklinkname=0 (see Makefile), which is
// why the whole file sits behind the pfx_fastzip build tag: plain `go build`
// and `go test` still work without the linker flag.
//
//go:linkname zipCompressors archive/zip.compressors
var zipCompressors sync.Map

const fastZipSupported = true

// setFastZipLevel routes all subsequent DEFLATE zip entries in this process
// through klauspost/compress at the given level (1 fastest .. 9 smallest).
// The output remains a standard, fully compatible DEFLATE stream; only the
// speed/size trade-off changes.
func setFastZipLevel(level int) {
	zipCompressors.Store(zip.Deflate, zip.Compressor(func(out io.Writer) (io.WriteCloser, error) {
		return kpflate.NewWriter(out, level)
	}))
}
