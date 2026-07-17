package core

import (
	"fmt"
	"os"
	"strconv"
	"sync"
)

// zipLevelEnvVar selects the DEFLATE speed/size trade-off for workbook
// output. Unset keeps the Go standard library compressor (level 5), which
// produces byte-for-byte the same archives as previous releases. Levels 1-9
// switch to klauspost/compress: level 1 is fastest, level 9 is smallest;
// level 6 compresses the reference 1.5M-cell workload about 3x faster than
// the standard library for roughly 20% larger files.
const zipLevelEnvVar = "PYFASTEXCEL_ZIP_LEVEL"

var configureZipOnce sync.Once

// configureZipCompression applies PYFASTEXCEL_ZIP_LEVEL once per process.
// Registration happens before any workbook bytes are produced, and the
// underlying registry swap is atomic, so concurrent exports are safe.
func configureZipCompression() {
	configureZipOnce.Do(func() {
		raw := os.Getenv(zipLevelEnvVar)
		if raw == "" {
			return
		}
		level, err := strconv.Atoi(raw)
		if err != nil || level < 1 || level > 9 {
			fmt.Fprintf(
				os.Stderr,
				"pyfastexcel: ignoring %s=%q (expected an integer from 1 to 9)\n",
				zipLevelEnvVar,
				raw,
			)
			return
		}
		if !fastZipSupported {
			fmt.Fprintf(
				os.Stderr,
				"pyfastexcel: %s is set but this native library was built without fast-zip support\n",
				zipLevelEnvVar,
			)
			return
		}
		setFastZipLevel(level)
	})
}
