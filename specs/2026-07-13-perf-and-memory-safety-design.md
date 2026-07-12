# pyfastexcel: Performance and Memory-Safety Redesign

**Date:** 2026-07-13
**Status:** Design approved in outline; two open questions remain (see end).

## Why this document exists

The starting hypothesis was: *"the JSON serialization bridge between Python and Go
is what makes pyfastexcel slow; replacing it with a better linking mechanism will
make it faster."*

**Measurement shows that hypothesis is false.** This document records what the cost
actually is, what to do about it, and — just as importantly — which plausible-sounding
ideas were measured and rejected, so they don't get proposed again.

All numbers below are measured on this machine (WSL2, Go 1.24.1) with a workload of
**50,000 rows x 30 columns = 1,500,000 cells**, every cell styled, single sheet, using
`StreamWriter`. Numbers are reproducible across runs.

---

## 1. Where the time actually goes

End-to-end: **3.1s**, peak RSS **617 MB** (for a 4.6 MB output file).

| Stage | Time | Share |
| --- | --- | --- |
| Go side (parse + excelize + base64) | 2.08s | **66%** |
| Python `row_append` loop | 0.97s | 31% |
| `msgspec.json.encode` | 0.06s | 2% |
| ctypes return + `base64.b64decode` | 0.015s | 0.5% |

Breaking the Go side down further:

| Go-side component | Time | Reducible? |
| --- | --- | --- |
| excelize `SetRow` loop | 0.53s | **No** — excelize's own work |
| excelize XML serialize + zip | 0.66s | **No** (but see §5.4) |
| `marshmallow.Unmarshal` -> `map[string]interface{}` | 0.56s | Yes |
| `createCell` + per-cell style-name string hash | ~0.17s | Yes |
| cgo boundary + `C.GoString`(33MB) + `C.CString` | ~0.14s | Mostly |

### The headline conclusions

1. **The JSON encode + FFI transfer is ~2.5% of runtime.** Replacing the *wire
   serialization* with msgpack / Arrow / shared memory wins essentially nothing on its
   own. The linking mechanism is not the problem.

2. **The problem is what Go does with the data after it arrives**: it explodes a 33 MB
   payload into a 215 MB `map[string]interface{}` tree (7x blowup, 1.24 GB total
   allocation, 20 GCs), then hashes a style-name string for every single cell.

3. **There is a hard floor.** excelize's own `SetRow` (0.53s) + XML/zip (0.66s) = 1.19s
   is 38% of current runtime and cannot be removed without changing backend. This is
   why the realistic target is ~1.7x, not 10x.

---

## 2. Correctness bug found: data race on the global style map

This is more important than any of the performance work.

`pyfastexcel/core/writer.go:17` declares a **package-level global**:

```go
var styleMap map[string]int
```

It is reassigned wholesale on every export (`writer.go:69`, `styleMap = CreateStyle(...)`)
and read per-cell (`cell.go:36`, `styleMap[v[1].(string)]`).

`ctypes.CDLL` **releases the GIL** for the duration of a foreign call. So two Python
threads calling `wb.save()` concurrently produce two goroutines concurrently reading and
writing the same global map.

**Verified** with a Go test that calls `core.WriteExcel` from 8 goroutines:

- without `-race`: passes, looks fine
- with `-race`: **8–11 `DATA RACE` reports on every run**, at `writer.go:69` (write),
  `style.go:246` (write), `cell.go:36` (read)

### Why this is severe

Concurrent map access in Go is undefined behaviour and can trigger the runtime's
`concurrent map read and map write` **fatal throw**. That is a `throw`, not a `panic`,
so the `recover()` in `catchPanic()` (`pyfastexcel.go:42`) **cannot catch it** — it
hard-kills the entire Python process.

### The same disease on the Python side

`StyleManager.REGISTERED_STYLES`, `_STYLE_NAME_MAP` and `_style_map` are **class-level**
(i.e. process-global) state. Worse, `read_lib_and_create_excel()` calls
`StyleManager.reset_style_configs()` on every save — so **saving workbook A wipes the
style registry of workbook B** in the same process. Observed symptom: under concurrency,
the style ID assigned to a given cell drifts nondeterministically between runs
(serial runs are stable).

> **Honesty note:** style-ID drift was observed, but in 10 concurrent trials the *final
> rendered font* still came out correct. So this is a latent race that is currently
> "getting away with it", not a demonstrated corruption. The race detector confirms the
> underlying race is real. Do not treat "it hasn't broken yet" as safety.

---

## 3. Rejected alternatives (measured, not assumed)

### 3.1 Rewrite the backend in Rust (`rust_xlsxwriter` + PyO3) — REJECTED

The intuition was "Rust is faster, and it removes the serialization boundary entirely."
It was also proposed to use Rust when no pivot table is present and fall back to Go when
one is (a dual backend).

**Measured, same 1.5M-cell workload, rust_xlsxwriter in its best configuration**
(`constant_memory` streaming mode = the analogue of excelize's `StreamWriter`, `zlib`
feature for zlib-ng, release + LTO):

| | write-cells loop | XML + zip | **total** |
| --- | --- | --- | --- |
| Go / excelize `StreamWriter` | 0.53s | 0.68s (**including base64**) | **1.21s** |
| rust_xlsxwriter (best config) | 0.61s | 1.06s | **1.67s** |

**rust_xlsxwriter is 1.4x SLOWER than excelize here** — and Go was additionally doing
base64 encoding that Rust was not.

Why:

- The cell-writing loop is **the same speed in both** (0.53s vs 0.61s). It is
  memory/allocation-bound, not CPU-bound; the language is not the variable.
- The dominant cost is XML serialization + deflate, and **Go's `compress/flate` beats
  rust_xlsxwriter's implementation**. Enabling zlib-ng did not help, which tells us the
  bottleneck is the XML serialization, not the compressor.

So a Rust backend — dual or total — would buy **negative** performance, before even
counting the cost of maintaining two style engines, two golden-file corpora, a doubled
wheel build matrix, and the surprise of "adding a pivot table silently changes both the
speed and the byte output of my file". The performance premise fails outright, so the
maintenance argument is not even needed.

**Do not revisit this without new measurements.**

### 3.2 Decode JSON into typed Go structs via reflection — REJECTED

Intuitively this should beat the generic `map[string]interface{}` path. Measured on the
1.5M-cell grid:

| Approach | Decode | Heap |
| --- | --- | --- |
| JSON + style names -> `[][]any` (**current**) | 0.902s | 204 MB |
| JSON + int style IDs -> `[][]any` | 0.643s | 169 MB |
| **JSON + int IDs -> typed struct w/ `UnmarshalJSON`** | **1.961s** | 127 MB |
| msgpack + int IDs -> `[][]any` | 0.476s | 123 MB |
| **msgpack + int IDs -> hand-written streaming decoder -> `[]excelize.Cell`** | **0.111s** | **74 MB** |

Reflection-based typed decoding is **2x slower than the code we already have**. The
per-cell `UnmarshalJSON` call and reflection overhead cost more than the boxing they
save.

**The winning approach is the last row: 8.1x faster decode, 2.75x less heap.** The design
below therefore mandates a *hand-written* streaming decoder and explicitly forbids
`msgpack.Unmarshal` into structs.

### 3.3 Shared memory / Arrow / zero-copy input — NOT WORTH IT

The entire cgo boundary cost (including copying the 33 MB payload into Go) is ~0.14s,
~4.5% of runtime. Removing the copy entirely would save a fraction of that. Not a lever.

---

## 4. Constraint: Python-side encoding must stay at C speed

`msgspec.json.encode` encodes 1.5M cells in 0.056s because it is a C extension.

**Any wire format we choose must still be encodable by msgspec from Python.** Hand-rolling
a binary encoder in a Python loop would be 10–50x slower than msgspec and would destroy
the entire win. This rules out bespoke `struct.pack` framing.

This is why the choice is msgpack (`msgspec.msgpack.encode`, measured 0.022s) and not a
custom binary format.

Measured payload sizes, 1.5M cells, 4 distinct styles:

| Wire format | Python encode | Payload | Bytes/cell |
| --- | --- | --- | --- |
| JSON + style names (**current**) | 0.056s | 35.9 MB | 23.9 |
| JSON + int style IDs | 0.036s | 11.9 MB | 8.0 |
| **msgpack + int style IDs** | **0.022s** | **7.1 MB** | **4.7** |

Interning style names to integer IDs alone shrinks the payload 3x — because today every
cell redundantly carries a string like `"black_fill_style"` (18 bytes).

---

## 5. Stage A — internal rework, zero API change

### A1. Fix the data race (do this first, independently of the perf work)

- Delete the package-level `var styleMap` in `core/writer.go:17`; make it a field on
  `ExcelWriter` (which is already a struct — this is nearly free).
- `createCell` takes the style table as a parameter instead of reading a global.
- Python: make `StyleManager`'s registries **per-instance** instead of class-level.
  `set_custom_style()` continues to write into a process-level *default* registry which
  each new `Workbook` copies at construction — this preserves the existing documented
  pattern of registering styles at import time (see `example.py`).
- **Remove `reset_style_configs()` from the save path.** Wiping global style state is a
  side effect of saving, not a feature.
- **Regression test:** a Go `-race` test calling `core.WriteExcel` from N goroutines.
  It reproduces the bug today. Add `-race` to CI.

### A2. Wire format v2 (internal; users never see it)

- Python interns style names to `uint32` IDs at encode time (`_style_map` is already
  insertion-ordered, so its index is the ID).
- Cell on the wire becomes `[value, style_id]` instead of `[value, "style_name"]`.
- Encode with `msgspec.msgpack.encode`.
- Go decodes the cell grid with a **hand-written streaming decoder**
  (`msgpack.NewDecoder` + explicit `DecodeArrayLen` / `DecodeInterfaceLoose` /
  `DecodeInt`) that writes **directly into `[]excelize.Cell`**, resolving `StyleID` via a
  `[]int` slice index — no map, no per-cell string hash.
- **Explicitly forbidden:** reflection-based `msgpack.Unmarshal` into structs (§3.2 —
  measured 2x slower than the status quo).
- **Metadata (styles, file_props, charts, pivots, panes, validation) stays on the existing
  JSON + `marshmallow` path.** It is small and its decode cost is noise. Do not touch it.
- Add a `PYFASTEXCEL_WIRE=json` escape hatch that falls back to the JSON encoder, so the
  payload is still human-readable when debugging.

Measured effect: decode 0.902s -> 0.111s, Go heap 204 MB -> 74 MB.

### A3. Python fast path in `row_append` (signature unchanged)

Three specific defects in the current per-cell path:

1. `self.style_key = f'{style}{kwargs}'` is computed on **every cell**, but is only used
   when `kwargs` is non-empty. Guard it behind `if kwargs`.
2. `_handle_string_style` does `self._collections_list + list(self.style.REGISTERED_STYLES)`
   — **constructing a new list on every single cell** — purely to do a membership test.
   Cache it as a `set`, invalidated on style registration.
3. `validate_and_format_value` is a separate call doing two `isinstance` checks. Inline
   the common scalar/str case.

Measured: **1.8–3.5x faster**, output verified byte-identical to the current
implementation.

### A4. Output path

- `Export` returns **pointer + length** instead of a base64 C string; Python reads it with
  `ctypes.string_at(ptr, n)`. This removes the base64 encode, the 1.33x size inflation,
  the NUL-scan, and `b64decode` in one change.
- New `ExportToFile(payload, path)` calling `excelize.SaveAs` directly. `wb.save('x.xlsx')`
  routes here, so **the xlsx bytes never cross the FFI boundary at all**.
- Wrap the ctypes call in `try/finally` so `FreeCPointer` still runs if Python raises
  in between — today that path leaks the C allocation.
- Fix the panic path: on a Go panic, `Export` currently returns nil, and Python then calls
  `base64.b64decode(None)` and raises a baffling `TypeError`. Return a status code and an
  error message instead.

### Stage A projected result

| | Current | After Stage A |
| --- | --- | --- |
| Python | 0.97s | ~0.45s |
| Encode | 0.056s | 0.022s |
| Go | 2.08s | ~1.35s |
| **Total** | **3.1s** | **~1.85s (~1.7x)** |
| **Peak RSS** | **617 MB** | **~300 MB (~2x)** |
| Thread-safe | No | **Yes** |

### The ceiling, stated honestly

After Stage A, excelize's own `SetRow` (0.53s) + XML/zip (0.66s) = 1.19s is **~63% of the
new runtime** and is immovable without changing backend. And per §3.1, the obvious
alternative backend is *slower*. **~1.7x is close to the ceiling for this architecture.**

---

## 6. Stage B — new opt-in APIs (still no breaking changes)

### B1. Batch row API

`ws.append_row(values, style=...)` / `append_rows(rows)` — one Python call per row instead
of one per cell. Measured **2.5x faster** than the `row_append` loop.

### B2. Columnar / DataFrame ingestion

`ws.write_frame(df)` reading straight from numpy/arrow buffers, **never materializing a
Python tuple per cell**. This removes the 108 MB list-of-tuples from peak memory and is
the real "fast lane" for bulk data.

### B3. Handle-based streaming (the memory story)

```
h = NewWorkbook()             # Go owns the workbook; returns an opaque uint64 handle
AddSheet(h, name, meta)
WriteRows(h, sheet, chunk)    # msgpack chunk; Go SetRow's it and frees it
Save(h, path) / Close(h)
```

- Handle registry in Go: `map[uint64]*Workbook` behind a mutex. **Never pass a Go pointer
  to C** (cgo rule); the handle is an opaque integer.
- Peak memory becomes **O(chunk)** instead of O(entire workbook) — this is what makes
  multi-million-row exports viable.
- The existing `Workbook` / `StreamWriter` API is unchanged, but `save()` can chunk
  `self._data` internally, so **existing users get a bounded Go heap without changing a
  line of code**.

### B4. Faster deflate in Go (promoted from stretch to "evaluate seriously")

XML+zip is the single largest remaining cost (0.66s). Since Go's `compress/flate` already
beats rust_xlsxwriter's compressor (§3.1), the lever here is `klauspost/compress`, a
drop-in faster deflate for Go (typically 1.5–2x).

**Not yet verified:** whether excelize exposes a way to substitute its zip writer. This
needs a feasibility spike before any number is promised.

---

## 7. Testing strategy

- **Go `-race` concurrency test in CI.** It reproduces the §2 bug today; it must stay green.
- **Golden-file corpus.** For a set of workbooks covering styles, charts, pivot tables,
  tables, data validation, comments, panes, merges and grouping: assert that wire format
  v2 produces a **byte-identical xlsx** to v1. This is the safety net for the whole
  refactor.
- Extend the benchmark harness to report **peak RSS**, not just wall time. Memory is half
  the point of this work and is currently unmeasured.
- Property test: round-trip every supported cell value type and style through the msgpack
  decoder.

## 8. Risks

1. **msgpack loses JSON's readability.** Mitigated by the `PYFASTEXCEL_WIRE=json` escape
   hatch (A2).
2. **The hand-written decoder is more code** than `marshmallow.Unmarshal` and must handle
   every cell type. Mitigated by the golden-file corpus.
3. **De-globalizing `StyleManager` is the most likely source of subtle behaviour change.**
   Registering styles at import time via `set_custom_style()` is an existing public
   pattern (`example.py`) and must keep working.
4. New Go dependency: `vmihailenco/msgpack`.

## 9. Open questions (unresolved — decide before implementation)

1. **Does de-globalizing `StyleManager` break usage not visible in this repo?** The
   import-time `set_custom_style()` pattern will be preserved, but if styles are shared
   across workbooks deliberately anywhere, that needs to be known first.
2. **Is Stage B3 (handle-based streaming) in scope now, or after Stage A lands?** Stage A
   is self-contained and delivers the bug fix plus ~1.7x; B3 is the larger change and the
   only thing that truly bounds memory.

## 10. Reproducing these numbers

The benchmarks used for this document were throwaway scripts, not committed. To
reproduce: build the workload as 50,000 x 30 cells via `StreamWriter`, and time the
phases separately.

**One measurement trap worth recording:** the first version of the phase profile ran with
`tracemalloc` active during the Python build loop. `tracemalloc` intercepts every
allocation and inflated the Python phase from ~1.0s to ~3.8s, which made Python look like
62% of runtime instead of 31%. **Never profile time with `tracemalloc` enabled.** Measure
memory in a separate pass.
