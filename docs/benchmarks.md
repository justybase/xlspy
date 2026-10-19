# Benchmarks

[Home](../README.md) · [Usage guide](usage.md) · [Validation](validation.md)

`xlspy` includes a **C extension** (`_c_core`) that accelerates XLSB read and write and the XLSX write path. It is enabled by default. Set `XLSPY_DISABLE_C_EXT=1` to force the pure Python fallback.

This page reports the write and read benchmarks on three machines. All were measured on **2026-10-19** with the same scripts and parameters. Dataset: **50,000 × 50** (2.5 million cells), `mixed` layout, median of **three runs** per measurement.

## Environments

| | Windows | macOS | WSL2 |
| :--- | :--- | :--- | :--- |
| CPU and memory | AMD Ryzen 7 7840HS, 16 logical CPUs, 31 GiB RAM | Apple M4, 10 CPU cores, 16 GiB RAM | AMD Ryzen 7 7840HS (host), 16 logical CPUs, 20 GiB visible to WSL2 |
| OS | Windows 11 Pro, build 26300 | macOS 27.0 (arm64) | Ubuntu, kernel 6.18.54.1-microsoft-standard-WSL2 (x86_64) |
| Python | 3.14.7 (`py -3.14`) | 3.14.7 (Homebrew) | 3.14.8 (uv-managed build) |
| C compiler | MSVC | Apple clang 21 | GCC 15.2.0 |
| openpyxl / xlsxwriter | 3.1.5 / 3.2.9 | 3.1.5 / 3.2.9 | 3.1.5 / 3.2.9 |

The Python patch version differs on WSL2 (3.14.8 against 3.14.7). The difference is treated as negligible. The C extension was built with the default `setup.py` flags on each system.

## Results

### Write time

Median wall time, 50,000 × 50, generator input, file output. Rows with the C extension on unless marked.

| Writer | Windows | macOS | WSL2 |
| :--- | ---: | ---: | ---: |
| xlspy XLSB, inline strings | 1.11 s | 0.72 s | 0.87 s |
| xlspy XLSB, shared strings | 1.31 s | 0.80 s | 1.09 s |
| xlspy XLSX, inline strings | 1.66 s | 1.14 s | 1.40 s |
| xlspy XLSX, shared strings | 1.93 s | 1.33 s | 1.65 s |
| xlspy XLSB, inline, C extension off | 2.77 s | 1.76 s | 2.19 s |
| xlspy XLSB, shared, C extension off | 2.91 s | 1.83 s | 2.37 s |
| xlspy XLSX, inline, C extension off | 3.14 s | 2.01 s | 2.55 s |
| xlspy XLSX, shared, C extension off | 3.50 s | 2.15 s | 2.78 s |
| openpyxl (`write_only`) | 16.59 s | 10.69 s | 14.30 s |
| xlsxwriter (`constant_memory`) | 9.97 s | 6.59 s | 8.14 s |

### Peak memory and file size

Peak memory is the worker's peak resident memory, including the interpreter. File sizes are the output workbook.

| Writer | Windows peak | macOS peak | WSL2 peak | Windows size | macOS size | WSL2 size |
| :--- | ---: | ---: | ---: | ---: | ---: | ---: |
| xlspy XLSB, inline | 35.6 MB | 37.8 MB | not reported ¹ | 6.67 MiB | 6.96 MiB | 6.96 MiB |
| xlspy XLSB, shared | 97.4 MB | 110.6 MB | not reported ¹ | 7.25 MiB | 7.52 MiB | 7.52 MiB |
| xlspy XLSX, inline | 51.2 MB | 50.1 MB | not reported ¹ | 5.60 MiB | 5.77 MiB | 5.77 MiB |
| xlspy XLSX, shared | 91.9 MB | 103.9 MB | not reported ¹ | 6.43 MiB | 6.71 MiB | 6.71 MiB |
| openpyxl (`write_only`) | 42.6 MB | 42.7 MB | not reported ¹ | 11.71 MiB | 11.48 MiB | 11.48 MiB |
| xlsxwriter (`constant_memory`) | 29.2 MB | 30.8 MB | not reported ¹ | 11.55 MiB | 11.33 MiB | 11.33 MiB |

¹ The WSL2 peak is not reported. Four writers show the same 47.9 MB, which points to a measurement artifact in the Linux peak reading (`ru_maxrss`), not to real memory use. The WSL2 peak needs re-measuring once the harness reads the value correctly.

### Observations

- With the C extension on, xlspy is **1.6–2.5× faster** than with it off, on all three systems. The largest gain is for XLSB inline.
- xlspy XLSX (shared strings) is **8–9× faster** than openpyxl `write_only` on the same machine.
- The Mac mini M4 is about **1.5× faster** than the Windows machine on every writer.
- WSL2 is 14–22% faster than Windows, but the two runs write to different filesystems (ext4 in WSL2, NTFS on Windows, APFS on macOS). This comparison does not isolate the operating system.

### Read time

Median wall time to read the 50,000 × 50 workbook back. Each library reads the file written in the same run, right after writing.

| Reader | Windows | macOS | WSL2 |
| :--- | ---: | ---: | ---: |
| xlspy XLSB, C extension | 0.58 s | 0.27 s | 0.40 s |
| xlspy XLSX | 4.77 s | 2.38 s | 3.59 s |
| xlspy XLSB, C extension off | 6.26 s | 3.46 s | 4.50 s |
| openpyxl (`read_only`) | 10.53 s | 6.02 s | 8.64 s |

- The C extension reads XLSB **11–13× faster** than the pure Python path on each system.
- xlspy XLSX reads **2.2–2.5× faster** than openpyxl `read_only`.
- The Mac mini M4 reads **1.8–2.2× faster** than the Windows machine, a larger gap than for writes.
- WSL2 reads 18–31% faster than Windows, with the same filesystem caveat as for writes.

## Method

- Every measurement runs in its own worker process. The writers are interleaved across three runs, and the medians are reported.
- The data is the `mixed` dataset (the `performance_test.py` layout) streamed from a generator. Generation is part of every time, so every writer pays the same cost.
- Time runs from opening the workbook to closing it. Peak is the peak working set on Windows and `ru_maxrss` on macOS and Linux.
- Output goes to a new temporary directory on each system.
- Read times come from `examples/performance_test.py`. It writes a file, then reads it back with each library, so the OS page cache is warm for every read.

Known limitations:

- The generation cost is included in every write time. No generation-only baseline was measured, so the net writer time is not reported.
- openpyxl and xlsxwriter outputs were not compared with xlspy outputs cell by cell.
- The filesystems differ between systems (see Observations).
- Read times are measured with a warm page cache, not after a cold start.

### Result files

The write data for each system, with the environment, every run, and the medians, is in JSON. The read data is the full text output of three runs of `performance_test.py` per system:

- [`windows-2026-10-19.json`](benchmark-results/windows-2026-10-19.json) (Windows, Python 3.14.7, write)
- [`mac-m4-2026-10-19.json`](benchmark-results/mac-m4-2026-10-19.json) (macOS, Python 3.14.7, write)
- [`wsl2-2026-10-19.json`](benchmark-results/wsl2-2026-10-19.json) (WSL2, Python 3.14.8, write)
- [`read-windows-2026-10-19.log`](benchmark-results/read-windows-2026-10-19.log) (Windows, Python 3.14.7, read)
- [`read-mac-m4-2026-10-19.log`](benchmark-results/read-mac-m4-2026-10-19.log) (macOS, Python 3.14.7, read)
- [`read-wsl2-2026-10-19.log`](benchmark-results/read-wsl2-2026-10-19.log) (WSL2, Python 3.14.8, read)

## Run the benchmarks

Install the test extras and check that C acceleration is loaded:

```bash
python -m pip install -e '.[test]'
python -c "from xlspy._accel import HAVE_C_EXT; assert HAVE_C_EXT"
```

[`examples/write_benchmark.py`](../examples/write_benchmark.py) produced the tables above. It runs every measurement in a fresh worker process, repeats each writer three times by default, interleaves the writers, and reports medians:

```bash
# The full comparison: all writers, C extension on and off, one tracemalloc run per writer
python examples/write_benchmark.py --rows 50000 --cols 50 --iterations 3 --writers all --c-ext both --tracemalloc --output bench.json --csv bench.csv

# Only some writers
python examples/write_benchmark.py --writers xlsx-inline,xlsb-inline --c-ext both

# Another source tree, for before/after runs
python examples/write_benchmark.py --src ../other-checkout/src --csv other.csv
```

- **Writers**: `xlsx-shared`, `xlsx-inline`, `xlsb-shared`, `xlsb-inline` (xlspy, `--c-ext on|off|both`), `openpyxl` (`write_only=True`), and `xlsxwriter` (`constant_memory=True`). openpyxl and xlsxwriter are optional; missing ones are skipped.
- **Data**: `--dataset mixed` (default) is the eight-type layout of `performance_test.py` and is deterministic. `--dataset random --seed N` is a seeded, text-heavy alternative. With `--input stream` (default), every writer consumes the same generator. With `--input list`, the rows are built before the clock starts.
- **Metrics**: wall time, peak resident memory, output size, and, with `--tracemalloc`, a separate untimed run's Python allocation peak (memory the C extension allocates directly is not traced).
- **Targets**: `--target file` (default) or `--target memory` (an `io.BytesIO`, xlspy writers only).
- **Output**: `--output` writes JSON with the environment, every run, and the medians. The date is recorded as a day only. `--csv` writes one row per run. A failed run is recorded with its error and does not stop the series.

The harness needs no third-party packages except the optional openpyxl and xlsxwriter. It is not part of the test suite.

### Read benchmark

The read numbers come from `examples/performance_test.py`. It runs the write and the read in one pass and prints a table. Run it three times from a scratch directory, because it writes its test files to the current directory, and take the median of each read row:

```bash
python examples/performance_test.py --rows 50000 --cols 50
```

For correctness checks and platform test results, see [Validation](validation.md).
