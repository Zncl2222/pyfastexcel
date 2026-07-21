# Independent reproduction — 2026-07-20

A second, from-scratch measurement of the perf rework, run to corroborate the
committed Stage A reports in the parent directory. **Nothing here reuses the
committed JSON** — every number was measured fresh:

- **baseline (v1)** — old code checked out at `ab23864` (the commit before the
  perf rework) in a throwaway git worktree, its native `.so` rebuilt, then
  measured with [`benchmark/baseline_probe.py`](../../baseline_probe.py).
- **new (v2)** — current build, measured with `benchmark/perf_memory.py`
  (`--wire json`, `--wire msgpack`, and `--destination file`).

Workload is identical throughout: 50000 rows × 30 cols, 4 styles, `StreamWriter`,
3 fresh-subprocess samples.

## Numbers (mean)

| series | build (s) | export (s) | total (s) | peak RSS (MiB) |
| --- | --: | --: | --: | --: |
| baseline v1 (old, bytes) | 1.253 | 1.236 | 2.489 | 576 |
| new v2 json (bytes)      | 0.969 | 1.174 | 2.143 | 551 |
| new v2 pfx2 (bytes)      | 0.955 | 1.074 | 2.029 | 295 |
| new v2 pfx2 (file mode)  | 0.913 | 1.038 | 1.952 | 289 |

## Findings

- **Peak RSS reproduces almost exactly** vs the committed reports (baseline 575,
  json 552, pfx2 294 MiB). RSS is largely load-independent, so this is the
  hardest evidence: the msgpack path genuinely roughly halves peak RSS.
- **Time improvements reproduce in direction but smaller in magnitude.** The
  committed report claims total −27% for pfx2; here it was −18%. The gap is most
  likely environment (this run: Python 3.14.0 under a loaded WSL2; committed:
  3.14.6 in a clean devcontainer). Re-measure `build`/`total` on a quiet machine
  before publishing the larger headline figures.
- **The new "file mode" (direct-to-file `ExportToFileV2`) is marginal for this
  workload**: ~4% faster and ~6 MiB lower RSS than returning bytes. The returned
  xlsx (~5.5 MB) is tiny next to the ~290 MiB peak, which is dominated by
  Go-side build structures rather than the final bytes. File mode's memory
  benefit grows with output size.

See `version-compare.png` (old vs new) and `mode-compare.png` (write-mode
comparison).
