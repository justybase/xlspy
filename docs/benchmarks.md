# Benchmarks

[Home](../README.md) · [Usage guide](usage.md) · [Validation](validation.md)

`xlspy` is designed for high performance. Since version 0.1.0, the library includes a **C extension** (`_c_core`) that accelerates XLSB read and write as well as the plain XLSX write path. The C extension is **enabled by default** (compiled automatically on install). Set `XLSPY_DISABLE_C_EXT=1` to force the pure Python fallback.

All benchmarks: **50000 × 50** dataset (2.5M cells); times below are the **median of three runs**. Tests performed on **Windows** (AMD64, AMD Ryzen 7 7840HS, Python 3.14.6; xlspy 0.7.0, openpyxl 3.1.5, xlsxwriter 3.2.9).

## Write

| Library | Format | Time | Size |
|---------|--------|------|------|
| **xlspy (C_EXT)** | XLSB | **1.46 s** | 7.25 MB |
| xlspy (Python) | XLSB | 3.21 s | 7.25 MB |
| xlspy | XLSX | 4.56 s | 6.43 MB |
| [xlsxwriter](https://pypi.org/project/xlsxwriter/) | XLSX | 9.85 s | 11.57 MB |
| [openpyxl](https://pypi.org/project/openpyxl/) | XLSX | 17.77 s | 11.71 MB |

## Read

| Library | Format | Time | Notes |
|---------|--------|------|-------|
| **xlspy (C_EXT)** | XLSB | **0.60 s** | default, compiled C |
| xlspy | XLSX | 4.96 s | uses expat XML parser (C) |
| xlspy (Python) | XLSB | 6.33 s | pure Python fallback |
| [openpyxl](https://pypi.org/project/openpyxl/) | XLSX | 11.06 s | read-only mode |


## Analysis

The **10.6× read speedup** comes from two factors:
- **~60–70%** — native C compilation, no interpreter overhead per record
- **~30–40%** — algorithm simplification: flat array indexed by `col − first_col` instead of `Dict[int, Any]`, no `isinstance` per cell, no `BiffReader.read_worksheet()` method call per record

## Run the benchmarks

From the repository root, run:

```bash
python examples/performance_test.py
```

See the [benchmark scripts](../examples/README.md#benchmarks) for the available
comparisons. Results depend on the dataset, hardware, Python version, and
library versions; these measurements describe the environment above.

For correctness checks and platform test results, see [Validation](validation.md).
