---
description: "xlspy performance highlights: XLSB and XLSX read/write speed, the native Rust extension, and how to measure your own workload."
---

# Performance

**2.5 million cells in under a second, as XLSB.** The native Rust extension
accelerates XLSB read and write, the XLSX read and write paths, column widths and
`load_workbook()`.

## Headline numbers

50,000 × 50 cells (2.5 million cells), medians of three runs:

| Machine | xlspy XLSB | xlspy XLSX | xlsxwriter | openpyxl |
| :--- | ---: | ---: | ---: | ---: |
| Windows 11, Ryzen 7 7840HS, Python 3.15.2 — write | **0.67 s** | **1.16 s** | 11.71 s | 20.39 s |
| Windows 11, Ryzen 7 7840HS, Python 3.15.2 — read | **0.22 s** | **0.98 s** | — | 14.35 s |
| macOS 27.2, Apple M4, Python 3.14.10 — write | **0.44 s** | **0.71 s** | 5.98 s | 9.71 s |
| macOS 27.2, Apple M4, Python 3.14.10 — read | **0.07 s** | **0.44 s** | — | 6.72 s |

The native module writes XLSB with inline strings 4.2–4.4× faster than the
pure-Python fallback and reads XLSB 30–56× faster. Cross-platform comparisons are
indicative, not a controlled hardware benchmark.

## What is fast

| Path | Why |
| :--- | :--- |
| `write_sheet()` to XLSB | rows are encoded and written as they are produced |
| `write_sheet()` to XLSX | rows spool, then the sheet is assembled once |
| `ExcelReader.get_rows()` | streams values, creates no `Cell` objects |
| `load_workbook()` | keeps only a style index per cell; styles are built lazily |
| Column autofit | native glyph metrics, 0.03 µs per cell |

## What costs time

- Shared strings grow with the number of distinct strings. Use
  `use_shared_strings=False` for high-cardinality exports.
- Styled cells and per-cell `Style` objects cost more than plain rows, but each
  distinct style combination is resolved once.
- `autofit_rows=None` measures every row and adds about 30% to an XLSB write.
- Loading a workbook builds a model; `ExcelReader` is faster for values only.

## Measure your own workload

```bash
python benchmarks/bench_io.py --out before.json
python benchmarks/bench_io.py --compare before.json --max-slowdown 1.03
```

The full methodology, hardware and raw results are in
[Benchmarks](benchmarks.md). To disable the native module:

```bash
XLSPY_ACCEL=python python your_script.py
```
