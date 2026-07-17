//go:build !pfx_fastzip

package core

// Stub used when the library is built without the pfx_fastzip tag (plain
// `go build` / `go test`). PYFASTEXCEL_ZIP_LEVEL is reported as unsupported.
const fastZipSupported = false

func setFastZipLevel(int) {}
