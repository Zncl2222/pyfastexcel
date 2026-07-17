# Benchmark

The following results show the performance comparison between `pyfastexcel` and `openpyxl` when writing a data to an Excel file in different scenario.

## Performance and memory regression benchmark

`perf_memory.py` measures pyfastexcel itself, including workbook construction,
native export, total wall time, and process peak RSS. Every sample runs in a
fresh subprocess so one sample's high-water RSS does not contaminate the next
one. Timing runs never enable `tracemalloc`.

```bash
# Always rebuild first; the harness rejects a stale pre-v2 native library.
make build

# Default MessagePack wire path, 1.5 million styled cells, three samples.
uv run python benchmark/perf_memory.py --rows 50000 --cols 30 --repeat 3 \
  --output optimized.json

# Keep a JSON-wire baseline and compare a later run against it.
uv run python benchmark/perf_memory.py --wire json --output baseline.json
uv run python benchmark/perf_memory.py --compare baseline.json

# Exercise the direct-to-file path (the .bin suffix checks path compatibility).
uv run python benchmark/perf_memory.py --destination file
```

The JSON report records the workload, CPU, Python/Go/dependency versions, git
state, native ABI and shared-library SHA-256, all raw samples, and summary
statistics. Comparisons reject mismatched grid sizes or output destinations.
Use separate runs for time and any allocation profiler; allocation hooks
materially change the Python hot loop.

The Stage A acceptance reports are committed in [`benchmark/results`](results/):
the historical ABI-v1 baseline, ABI-v2 PFX2 result, and ABI-v2 compatibility-JSON
attribution run each contain three fresh-process samples.

- [Windows11](#benchmark-result-windows-11)
- [Windows11 WSL2 Ubuntu22.04](#benchmark-results-windows-11-wsl2-ubuntu-2204)

## Benchmark Environment

> - OS: Windows 11 & Windows 11 WSL2 Ubuntu 22.04
> - CPU: Intel(R) Core(TM) i7-12700 CPU
> - RAM: DDR4-3200 32GB
> - Hard Drive: Crucial P5 Plus 1TB Read: 6,600 MB/s Write: 5,000 MB/s
> - Python: 3.11.0
> - openpyxl: 3.1.2
> - pyfastexcel: 0.8.0

## Benchmark Results (Windows 11)

### Write 50 rows with 30 columns (Total 1500 cells)

<dev align='center'>
    <img src='../docs/images/50_30_horizontal_Windows11.png'>
</dev>

### Write 500 rows with 30 columns (Total 15000 cells)

<dev align='center'>
    <img src='../docs/images/500_30_horizontal_Windows11.png'>
</dev>

### Write 5000 rows with 30 columns (Total 150000 cells)

<dev align='center'>
    <img src='../docs/images/5000_30_horizontal_Windows11.png'>
</dev>

### Write 50000 rows with 30 columns (Total 1500000 cells)

<dev align='center'>
    <img src='../docs/images/50000_30_horizontal_Windows11.png'>
</dev>

## Benchmark Results (Windows 11 WSL2 Ubuntu 22.04)

### Write 50 rows with 30 columns (Total 1500 cells)

<dev align='center'>
    <img src='../docs/images/50_30_horizontal_WSL2-Ubuntu22.04.png'>
</dev>

### Write 500 rows with 30 columns (Total 15000 cells)

<dev align='center'>
    <img src='../docs/images/500_30_horizontal_WSL2-Ubuntu22.04.png'>
</dev>

### Write 5000 rows with 30 columns (Total 150000 cells)

<dev align='center'>
    <img src='../docs/images/5000_30_horizontal_WSL2-Ubuntu22.04.png'>
</dev>

### Write 50000 rows with 30 columns (Total 1500000 cells)

<dev align='center'>
    <img src='../docs/images/50000_30_horizontal_WSL2-Ubuntu22.04.png'>
</dev>
