# Benchmarks

[Home](index.md) · [Usage guide](usage.md) · [Validation](validation.md)

`xlspy` includes a **native extension** (`_rs_core`, Rust/PyO3) that accelerates XLSB read and write, the XLSX read and write paths and workbook-model loading. It is enabled by default. Set `XLSPY_ACCEL=python` to force the pure-Python fallback.

The [home page](index.md#built-for-speed) and the results below now include fresh WSL2 and Mac measurements for the 0.14.0 release candidate. The older tables remain as historical baselines: Windows results use commit `203721c` (Python 3.14.2), while the earlier [macOS and WSL2 results](#macos-and-wsl2-xlspy-0110) use release 0.11.0. The [historical results](#historical-results-c-extension-2026-10-19-to-2026-10-30) predate the Rust port.

## Release candidate 0.14.0 (2026-11-05)

The code was taken from local HEAD `a9f60f0` and built natively on each
machine. The temporary source snapshots set `pyproject.toml` to the 0.14.0
release-candidate version; benchmark code matches that commit. WSL2 ran Ubuntu
26.04.1 on an AMD Ryzen 7 7840HS with Python 3.14.4. The Mac mini ran macOS
27.2 on Apple M4 with Python 3.14.10. Both used the same deterministic input,
native Rust backend, and three measured runs. WSL2's benchmark files were on
the Linux home filesystem; Mac files were on APFS. The different CPUs, Python
patch versions and filesystems mean the cross-platform comparison is
indicative, not a controlled attribution of speed to the OS.

### Read, write and model (50,000 × 20)

`benchmarks/bench_io.py` ran each case in a fresh process. Times are medians
of three runs; the raw JSON includes each run and its interquartile range.

| Case | WSL2, seconds | Mac, seconds |
| :--- | ---: | ---: |
| Stream write XLSX | 0.391 | 0.249 |
| Stream write XLSB | 0.277 | 0.167 |
| `ExcelReader` read XLSX | 0.220 | 0.151 |
| `ExcelReader` read XLSB | 0.053 | 0.025 |
| `load_workbook()` XLSX | 0.411 | 0.221 |
| `load_workbook()` XLSB | 0.256 | 0.100 |
| Load styled XLSX | 0.438 | 0.216 |
| Load styled XLSB | 0.245 | 0.101 |
| Save loaded styled XLSX | 0.904 | 0.521 |
| Save loaded styled XLSB | 0.366 | 0.201 |
| Build and save model XLSX | 1.008 | 0.698 |
| Build and save model XLSB | 0.853 | 0.615 |
| Formula/hyperlink XLSX write | 2.336 | 1.402 |

### Large writer comparison (50,000 × 50)

`examples/write_benchmark.py` used the mixed eight-type generator, streamed
rows, and three interleaved runs per writer. xlspy results include both native
and pure-Python modes; the README shows the native inline-string XLSB/XLSX
results. Times are seconds, medians:

| Writer | WSL2 | Mac |
| :--- | ---: | ---: |
| xlspy XLSB inline, native | 0.65 | 0.44 |
| xlspy XLSB inline, Python | 2.73 | 1.95 |
| xlspy XLSX inline, native | 1.00 | 0.71 |
| xlspy XLSX inline, Python | 3.25 | 2.38 |
| xlspy XLSB shared, native | 0.86 | 0.51 |
| xlspy XLSX shared, native | 1.32 | 0.91 |
| xlsxwriter (`constant_memory`) | 8.54 | 5.98 |
| openpyxl (`write_only`) | 14.21 | 9.71 |
| PyExcelerate | 6.81 | 4.94 |

The separate end-to-end reader check (`examples/performance_test.py`) wrote
and read 50,000 × 50 files, three times per platform. Read times were measured
with a warm page cache; each reader consumed its own writer's file:

| Reader | WSL2, seconds | Mac, seconds |
| :--- | ---: | ---: |
| xlspy XLSB, native | 0.16 | 0.07 |
| xlspy XLSB, pure Python | 4.88 | 3.95 |
| xlspy XLSX, native | 0.59 | 0.44 |
| openpyxl XLSX, `read_only` | 9.14 | 6.72 |

Raw data: [WSL2 read/model JSON](benchmark-results/wsl2-2026-11-05-r1-io.json),
[Mac read/model JSON](benchmark-results/mac-m4-2026-11-05-r1-io.json),
[WSL2 writer JSON](benchmark-results/wsl2-2026-11-05-r1-write.json),
[Mac writer JSON](benchmark-results/mac-m4-2026-11-05-r1-write.json), and the
[WSL2](benchmark-results/read-wsl2-2026-11-05-r1.log),
[Mac run 1](benchmark-results/read-mac-m4-2026-11-05-r1.log),
[Mac run 2](benchmark-results/read-mac-m4-2026-11-05-r2.log) and
[Mac run 3](benchmark-results/read-mac-m4-2026-11-05-r3.log) reader logs.
The WSL2 logs are [run 1](benchmark-results/read-wsl2-2026-11-05-r1.log),
[run 2](benchmark-results/read-wsl2-2026-11-05-r2.log) and
[run 3](benchmark-results/read-wsl2-2026-11-05-r3.log).

## Rust benchmark baseline (2026-11-02)

Windows 11 Pro (build 26300), AMD Ryzen 7 7840HS (16 logical CPUs, 31 GiB RAM), Python 3.14.2, xlspy at commit `203721c` (Q7 of the quality round), openpyxl 3.1.5, xlsxwriter 3.3.1, PyExcelerate 0.13.0. Output goes to NTFS.

### Native write-path experiments (2026-11-05)

Two candidates were measured before being adopted; neither passed its gate, so the
write path is unchanged.

- **Hyperlink trailer in Rust** (`<hyperlinks>`, the worksheet `.rels` part): the
  Python trailer on a 50,000 × 20 sheet of 250,000 unique links took 1.03 s; with
  the native encoder 0.99 s, inside the run-to-run spread. The native function
  creates the same 250,000 Python relationship tuples, which offsets the faster
  XML formatting. Not adopted.
- **Native deflate** (`zlib-rs`, `libdeflater`) on the uncompressed 24.8 MB
  worksheet of a 50,000 × 20 XLSX, level 3: zlib 0.073 s (341 MB/s), `zlib-rs`
  0.061 s (409 MB/s, 1.20×), `libdeflate` 0.055 s (453 MB/s, 1.33×); at level 6
  `zlib-rs` is 1.08× zlib, at level 9 1.09×. Compression is about 0.07 s of a
  0.42 s write at level 3, so the best case saves ~3% of the whole write (below the
  8% gate), and the writer already compresses on a worker thread. Not adopted.

### Write

50,000 × 50 (2.5 million cells), `mixed` dataset from a generator to a file, median of three interleaved runs, each in its own process ([`windows-2026-11-02-rust-write.json`](benchmark-results/windows-2026-11-02-rust-write.json)). Peak is the worker's peak working set; traced is a separate run's peak of Python allocations.

| Writer | Format | Time | Peak | Traced | File |
| :--- | :--- | ---: | ---: | ---: | ---: |
| xlspy, inline strings | XLSB | **0.81 s** | 44.4 MB | 4.4 MB | 7.32 MiB |
| xlspy, shared strings | XLSB | **1.01 s** | 115.0 MB | 63.0 MB | 8.10 MiB |
| xlspy, inline strings | XLSX | **1.27 s** | 53.6 MB | 5.1 MB | 6.13 MiB |
| xlspy, shared strings | XLSX | **1.45 s** | 107.3 MB | 57.0 MB | 7.11 MiB |
| xlspy, inline strings, `XLSPY_ACCEL=python` | XLSB | 2.98 s | 43.9 MB | 2.6 MB | 7.32 MiB |
| xlspy, shared strings, `XLSPY_ACCEL=python` | XLSB | 3.21 s | 100.7 MB | 56.2 MB | 8.10 MiB |
| xlspy, inline strings, `XLSPY_ACCEL=python` | XLSX | 3.35 s | 46.4 MB | 4.8 MB | 6.13 MiB |
| xlspy, shared strings, `XLSPY_ACCEL=python` | XLSX | 3.59 s | 99.2 MB | 57.0 MB | 7.11 MiB |
| PyExcelerate | XLSX | 8.42 s | 281.2 MB | 218.5 MB | 12.31 MiB |
| xlsxwriter (`constant_memory`) | XLSX | 10.17 s | 28.9 MB | 1.1 MB | 11.55 MiB |
| openpyxl (`write_only`) | XLSX | 17.06 s | 40.5 MB | 1.0 MB | 11.71 MiB |

- The native module writes **2.5–3.7×** faster than the pure-Python path; the gain is largest for XLSB with inline strings.
- xlspy XLSB (inline) is **12.6×** faster than xlsxwriter and **21×** faster than openpyxl; xlspy XLSX (shared strings) is **7.0×** and **11.8×** faster.
- Shared strings cost memory in proportion to the distinct texts (about 57–63 MB of Python allocations here); inline strings keep the writer near 45–55 MB whatever the row count.
- The write times match the last C-extension run on the same machine (2026-10-30, below) within 0.03 s (0.81 s, 0.98 s, 1.28 s and 1.46 s then).

### Read

The 50,000 × 50 file each writer produced, read back right after writing (warm page cache), median of three runs of `examples/performance_test.py` ([`read-windows-2026-11-02-rust.log`](benchmark-results/read-windows-2026-11-02-rust.log); the log's note on the XLSX row was corrected afterwards: XLSX rows go through the native scanner).

| Reader | Format | Time |
| :--- | :--- | ---: |
| xlspy `ExcelReader`, native | XLSB | **0.14 s** |
| xlspy `ExcelReader`, native | XLSX | **0.73 s** |
| xlspy `ExcelReader`, `XLSPY_ACCEL=python` | XLSB | 6.51 s |
| openpyxl (`read_only`) | XLSX | 10.75 s |

- Native XLSB reads are **46×** faster than the pure-Python path and **77×** faster than openpyxl reading the XLSX file. Native XLSX reads are **15×** faster than openpyxl.
- Before the Rust port, the C extension read the same XLSB in 0.58 s and xlspy read the XLSX in 4.77 s (2026-10-19, below); the XLSB read path and the native XLSX scanner were reworked between those runs and this one.

### Read modes and the workbook model

`benchmarks/bench_io.py`, 50,000 × 20 (1 million cells), median of three runs, each in its own process ([`windows-2026-11-04-r3-io.json`](benchmark-results/windows-2026-11-04-r3-io.json)):

| Case | XLSX | XLSB |
| :--- | ---: | ---: |
| `ExcelReader.get_rows()` (values, no styles) | 0.27 s | 0.05 s |
| `load_workbook()` | 0.33 s | 0.16 s |
| `load_workbook()`, every 10th row bold | 0.38 s | 0.23 s |
| Save a loaded styled workbook after 1,000 value and style edits | 1.69 s | 0.33 s |
| Build the model and save it | 1.03 s | 0.90 s |
| Write rows with `XlsxWriter` / `XlsbWriter` | 0.40 s | 0.32 s |

- `load_workbook()` builds its `Cell` objects natively. It costs **1.3–1.4×** an `ExcelReader` pass for XLSX and **3.2–4.5×** for XLSB, whose plain read is much faster; `ExcelReader` creates no `Cell` objects, so use it when you only need values.
- A loaded workbook is saved by patching the worksheet in place: the time depends mostly on the size of the edited sheet, not on the number of edits. XLSX costs more because the sheet's XML is parsed once to check the edit.
- Compared with the code just before this work, measured in turns on one machine: saving a loaded XLSB 2.87 s → 0.33 s, building and saving the model 1.51 s → 1.03 s (XLSX) and 1.40 s → 0.92 s (XLSB), peak memory of that case 323 MB → 212 MB. The XLSX save after edits had no row here because it took hours at this size: 86 s at 4,000 rows, growing with the square of the rows.

### macOS and WSL2 (xlspy 0.11.0)

The same three benchmarks, run the same day on the release commit of 0.11.0 (`bf8c6ba`), from a fresh clone with the package and its Rust module built in release mode. The Windows column repeats the tables above, measured at `203721c`; a check of the Windows load rows at `bf8c6ba` gave the same times within 0.03 s.

| | Windows | macOS | WSL2 |
| :--- | :--- | :--- | :--- |
| CPU and memory | AMD Ryzen 7 7840HS, 16 logical CPUs, 31 GiB RAM | Apple M4 (Mac mini), 10 CPU cores, 16 GiB RAM | AMD Ryzen 7 7840HS (the Windows host), 16 logical CPUs, 20 GiB visible to WSL2 |
| OS | Windows 11 Pro, build 26300 | macOS 27.2 (arm64) | Ubuntu 26.04.1 LTS, kernel 6.18.54.1-microsoft-standard-WSL2 (x86_64) |
| Python | 3.14.2 | 3.14.7 | 3.14.8 |
| Filesystem | NTFS | APFS | ext4 (the WSL2 virtual disk, not `/mnt/c`) |

**Write**, 50,000 × 50, `mixed`, median of three runs ([`mac-m4-2026-11-02-rust-write.json`](benchmark-results/mac-m4-2026-11-02-rust-write.json), [`wsl2-2026-11-02-rust-write.json`](benchmark-results/wsl2-2026-11-02-rust-write.json)). Peak is `ru_maxrss` of the worker process:

| Writer | Format | Windows | macOS | WSL2 | macOS peak | WSL2 peak |
| :--- | :--- | ---: | ---: | ---: | ---: | ---: |
| xlspy, inline strings | XLSB | **0.81 s** | **0.42 s** | **0.59 s** | 46.9 MB | 39.9 MB |
| xlspy, shared strings | XLSB | **1.01 s** | **0.49 s** | **0.75 s** | 125.9 MB | 117.3 MB |
| xlspy, inline strings | XLSX | **1.27 s** | **0.69 s** | **0.89 s** | 60.3 MB | 45.5 MB |
| xlspy, shared strings | XLSX | **1.45 s** | **0.75 s** | **1.07 s** | 121.6 MB | 99.3 MB |
| xlspy, inline strings, `XLSPY_ACCEL=python` | XLSB | 2.98 s | 1.58 s | 2.24 s | 46.0 MB | 40.0 MB |
| xlspy, shared strings, `XLSPY_ACCEL=python` | XLSB | 3.21 s | 1.65 s | 2.37 s | 106.1 MB | 97.4 MB |
| xlspy, inline strings, `XLSPY_ACCEL=python` | XLSX | 3.35 s | 1.82 s | 2.45 s | 48.1 MB | 40.3 MB |
| xlspy, shared strings, `XLSPY_ACCEL=python` | XLSX | 3.59 s | 1.93 s | 2.67 s | 109.5 MB | 97.4 MB |
| PyExcelerate | XLSX | 8.42 s | 4.87 s | 6.55 s | 304.1 MB | 281.4 MB |
| xlsxwriter (`constant_memory`) | XLSX | 10.17 s | 5.59 s | 7.77 s | 32.7 MB | 35.1 MB |
| openpyxl (`write_only`) | XLSX | 17.06 s | 9.26 s | 13.05 s | 44.3 MB | 48.9 MB |

**Read**, the files each writer produced, median of three runs of `examples/performance_test.py` ([`read-mac-m4-2026-11-02-rust.log`](benchmark-results/read-mac-m4-2026-11-02-rust.log), [`read-wsl2-2026-11-02-rust.log`](benchmark-results/read-wsl2-2026-11-02-rust.log)):

| Reader | Format | Windows | macOS | WSL2 |
| :--- | :--- | ---: | ---: | ---: |
| xlspy `ExcelReader`, native | XLSB | **0.14 s** | **0.07 s** | **0.14 s** |
| xlspy `ExcelReader`, native | XLSX | **0.73 s** | **0.44 s** | **0.55 s** |
| xlspy `ExcelReader`, `XLSPY_ACCEL=python` | XLSB | 6.51 s | 3.62 s | 4.47 s |
| openpyxl (`read_only`) | XLSX | 10.75 s | 6.15 s | 7.99 s |

**Read modes and the workbook model**, `benchmarks/bench_io.py`, 50,000 × 20, median of three runs ([`mac-m4-2026-11-02-rust-io.json`](benchmark-results/mac-m4-2026-11-02-rust-io.json), [`wsl2-2026-11-02-rust-io.json`](benchmark-results/wsl2-2026-11-02-rust-io.json)):

| Case | Windows XLSX | macOS XLSX | WSL2 XLSX | Windows XLSB | macOS XLSB | WSL2 XLSB |
| :--- | ---: | ---: | ---: | ---: | ---: | ---: |
| `ExcelReader.get_rows()` (values, no styles) | 0.27 s | 0.15 s | 0.20 s | 0.05 s | 0.03 s | 0.05 s |
| `load_workbook()` | 0.37 s | 0.23 s | 0.39 s | 0.23 s | 0.12 s | 0.29 s |
| `load_workbook()`, every 10th row bold | 0.42 s | 0.49 s | 0.82 s | 0.27 s | 0.12 s | 0.26 s |
| Save a loaded styled workbook after 1,000 value and style edits | — | — | — | 2.57 s | 1.66 s | 2.77 s |
| Build the model and save it | 1.48 s | 1.03 s | 1.76 s | 1.41 s | 0.94 s | 1.60 s |
| Write rows with `XlsxWriter` / `XlsbWriter` | 0.40 s | 0.24 s | 0.31 s | 0.31 s | 0.16 s | 0.25 s |

- The ratios hold on all three systems: native XLSB (inline) writes **12.6–13.4×** faster than xlsxwriter and **21–22×** faster than openpyxl; the native module writes **2.5–3.8×** faster than the pure-Python path; native XLSB reads are **57–88×** faster than openpyxl reading XLSX.
- The Apple M4 is the fastest of the three: **1.7–2.1×** Windows for writes and reads. WSL2 runs on the same CPU as Windows and is **1.3–1.4×** faster for writes and the XLSX read, and equal for the native XLSB read; the filesystems differ (ext4 against NTFS), so this does not isolate the operating system.
- One row does not follow: `load_workbook()` of the styled XLSX costs about 2× the unstyled load on macOS and WSL2, but 1.1× on Windows. A profile on WSL2 puts the time in the native row loader, not in a Python fallback; the cause is still open.
- Peaks on macOS and WSL2 are `ru_maxrss` and are not comparable with the Windows peak working set. File sizes differ slightly from Windows (7.02 against 7.32 MiB for XLSB inline) because the zlib builds differ; the data is the same.

## Historical results (C extension, 2026-10-19 to 2026-10-30)

These runs used the C extension (`_c_core`, removed in commit `68d4f67`). The rows are labelled as they were recorded: "C extension on" is the native path, "C extension off" the pure-Python path. The port's own comparison against C is in [Validation](validation.md#rust-port-of-the-native-extension-2026-10-31).

### Environments

| | Windows | macOS | WSL2 |
| :--- | :--- | :--- | :--- |
| CPU and memory | AMD Ryzen 7 7840HS, 16 logical CPUs, 31 GiB RAM | Apple M4, 10 CPU cores, 16 GiB RAM | AMD Ryzen 7 7840HS (host), 16 logical CPUs, 20 GiB visible to WSL2 |
| OS | Windows 11 Pro, build 26300 | macOS 27.0 (arm64) | Ubuntu, kernel 6.18.54.1-microsoft-standard-WSL2 (x86_64) |
| Python | 3.14.7 (`py -3.14`) | 3.14.7 (Homebrew) | 3.14.8 (uv-managed build) |
| C compiler | MSVC | Apple clang 21 | GCC 15.2.0 |
| openpyxl / xlsxwriter | 3.1.5 / 3.2.9 | 3.1.5 / 3.2.9 | 3.1.5 / 3.2.9 |

The Python patch version differs on WSL2 (3.14.8 against 3.14.7). The difference is treated as negligible. The C extension was built with the default `setup.py` flags on each system. Both `setup.py` and the C extension have since been removed; the package is built with maturin.

### Results

#### Write time

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

#### Peak memory and file size

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

#### Observations

- With the C extension on, xlspy is **1.6–2.5× faster** than with it off, on all three systems. The largest gain is for XLSB inline.
- xlspy XLSX (shared strings) is **8–9× faster** than openpyxl `write_only` on the same machine.
- The Mac mini M4 is about **1.5× faster** than the Windows machine on every writer.
- WSL2 is 14–22% faster than Windows, but the two runs write to different filesystems (ext4 in WSL2, NTFS on Windows, APFS on macOS). This comparison does not isolate the operating system.

#### Read time

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
- These read runs predate the XLSB read changes of 2026-10-30 (see the comparison
  below). The XLSB rows were not measured again on macOS or WSL2, so the XLSB numbers
  in this table are older than the Windows results of the comparison.

#### Read modes and the workbook model

`benchmarks/bench_io.py` compares the read modes described in
[Read modes](guide/reading.md#read-modes-which-one-to-use) on a 50,000 × 20 workbook
(Windows, Python 3.14, C extension on, median of three runs, 2026-10-29,
[`results-2026-10-29-sprint.json`](https://github.com/justybase/xlspy/blob/main/benchmarks/history/results-2026-10-29-sprint.json);
before the 0.11 sprint in brackets,
[`results-write-final.json`](https://github.com/justybase/xlspy/blob/main/benchmarks/history/results-write-final.json)):

| Case | XLSX | XLSB |
| :--- | ---: | ---: |
| `ExcelReader.get_rows()` (values, no styles) | 0.28 s | 0.19 s |
| `load_workbook()` | 0.40 s (2.30 s) | 0.55 s (2.56 s) |
| `load_workbook()`, every 10th row bold | 0.82 s (2.69 s) | 0.90 s (2.93 s) |
| Save a loaded styled workbook after 1,000 value and style edits | — | 2.43 s |
| Build the model and save it | 1.91 s | 1.77 s |

- The C extension builds the model's cells itself (0.11): `load_workbook()` is
  now only **1.4–3×** slower than `ExcelReader`, which creates no `Cell`
  objects; use `ExcelReader` when you only need values.
  Files saved by Excel itself load in 0.47 s (XLSX) and 0.71 s (XLSB).
- `load_workbook()` keeps only a style index per cell, and style objects are
  built only for the cells whose style you read.
- Against the pre-sprint baseline, no other case changed by more than 5%
  (`--compare ... --max-slowdown 1.05`).
- A loaded XLSB is saved by patching the binary worksheet in place: the time
  depends mostly on the size of the edited sheet, not on the number of edits.
- The XLSB read numbers in this section predate the 2026-10-30 changes to the C
  extension's XLSB read path. They were not measured again.

**Regression check** against `results-write-final.json` (`bench_io.py`, 50,000 × 20, 2026-10-27):
every case ran at 0.90–0.99× of the baseline time, so no case is slower than before
(full results: [`results-2026-10-27.json`](https://github.com/justybase/xlspy/blob/main/benchmarks/history/results-2026-10-27.json)).

#### Comparison with other Python libraries (2026-10-30)

Re-run of the writers and readers on one machine, with the Python libraries that
read or write the same formats. Windows 11, AMD Ryzen 7 7840HS, Python 3.14.7,
the xlspy development tree at commit `75828a1`. The benchmark scripts and the XLSB read path of
the C extension (`src/xlspy/c_ext/_c_core.c`, removed in `68d4f67`) had uncommitted changes; the read
numbers below are for the changed code.
Library versions: openpyxl 3.1.5, xlsxwriter 3.2.9, PyExcelerate 0.13.0,
python-calamine 0.9.0, fastexcel 0.21.0 with pyarrow 26.0.0, pyxlsb 1.0.10.
Raw data: write, median of three runs, 50,000 × 50 (2.5 million cells):
[`results-write-2026-10-30.json`](https://github.com/justybase/xlspy/blob/main/benchmarks/history/results-write-2026-10-30.json),
[`results-write-2026-10-30.csv`](https://github.com/justybase/xlspy/blob/main/benchmarks/history/results-write-2026-10-30.csv);
read, median of five runs:
[`results-read-compare-2026-10-30-v2.json`](https://github.com/justybase/xlspy/blob/main/benchmarks/history/results-read-compare-2026-10-30-v2.json).

**Write**, time in seconds. The `mixed` dataset (the eight-type layout of
`performance_test.py`) goes from a generator to a file. Peak is the writer's peak
working set. The write path was not changed after this run, so these numbers
still apply to the current code.

| Writer | Format | Time | Peak | File |
| :--- | :--- | ---: | ---: | ---: |
| xlspy, inline strings | XLSB | 0.81 s | 43.5 MB | 7.32 MiB |
| xlspy, shared strings | XLSB | 0.98 s | 104.3 MB | 8.10 MiB |
| xlspy, inline strings | XLSX | 1.28 s | 53.2 MB | 6.31 MiB |
| xlspy, shared strings | XLSX | 1.46 s | 105.4 MB | 7.25 MiB |
| PyExcelerate | XLSX | 8.47 s | 271.8 MB | 12.31 MiB |
| xlsxwriter (`constant_memory`) | XLSX | 10.15 s | 29.4 MB | 11.55 MiB |
| openpyxl (`write_only`) | XLSX | 17.30 s | 32.4 MB | 11.71 MiB |

With the C extension off, the four xlspy writers take 2.6–3.2 s.

- xlspy XLSB with inline strings is **12.5×** faster than xlsxwriter and **21×**
  faster than openpyxl. XLSX with shared strings is **7.0×** faster than xlsxwriter
  and **11.8×** faster than openpyxl.
- PyExcelerate is the fastest of the other XLSX writers, but it holds the whole
  sheet in memory (272 MB peak, about 219 MB of Python allocations) and writes no
  styles, so dates are stored as serial numbers and its file is the largest.
- No other library in this comparison writes XLSB, so the XLSB rows have no
  competitor here.

**Read**, time in seconds, on the `bench_io.py` dataset (five cell types). The timed
pass consumes every row. The row and cell counts, which check the readers against
each other, come from a second pass over the same file that is not timed. Every
reader counted 50,000 rows and 2,500,000 non-empty cells.

| Reader | File | Time |
| :--- | :--- | ---: |
| xlspy `ExcelReader` | XLSB | 0.10 s |
| python-calamine | XLSB | 0.32 s |
| fastexcel | XLSB | 0.55 s |
| xlspy `load_workbook` (model) | XLSB | 1.14 s |
| pyxlsb | XLSB | 6.63 s |
| xlspy `ExcelReader` | XLSX (xlspy) | 0.60 s |
| xlspy `ExcelReader` | XLSX (xlsxwriter) | 0.64 s |
| python-calamine | XLSX (xlspy) | 0.89 s |
| python-calamine | XLSX (xlsxwriter) | 1.08 s |
| fastexcel | XLSX (xlspy) | 1.00 s |
| xlspy `load_workbook` (model) | XLSX (xlspy) | 1.80 s |
| openpyxl `read_only=True` | XLSX (xlspy) | 4.82 s |
| openpyxl `read_only=True` | XLSX (xlsxwriter) | 8.03 s |
| openpyxl `load_workbook` | XLSX (xlspy) | 12.15 s |
| openpyxl `load_workbook` | XLSX (xlsxwriter) | 15.61 s |

- XLSB: xlspy `ExcelReader` is the fastest measured reader. It is 3.2× faster than
  python-calamine, 5.5× faster than fastexcel and about 66× faster than pyxlsb.
  This favours xlspy: python-calamine returns the whole sheet as lists that stay in
  memory, and `ExcelReader` drops each row after it is read. When both keep every row
  (xlspy `list(get_rows())` 0.16 s, python-calamine `to_python()` 0.32 s), xlspy is
  still about 2× faster.
- XLSX: xlspy `ExcelReader` is the fastest measured XLSX reader. It is 1.5× faster
  than python-calamine on the xlspy file, 1.7× faster than fastexcel and 8× faster
  than openpyxl `read_only=True`.
- python-calamine and fastexcel build the whole sheet before they return it (a list
  of lists, and an Arrow batch in which mixed-type columns come back as strings).
  `ExcelReader` yields rows as it goes.
- `load_workbook` (the model) is 3.0× slower than `ExcelReader` on XLSX and 11×
  slower on XLSB, because it creates a `Cell` object for every value.

**XLSB read, before and after the C changes** (same file, same machine, the same
harness pass, median of five runs; the file is written by `XlsbWriter` from the
`bench_io.py` generator):

| Case | Before (`_c_core.c` at `75828a1`) | After | Change |
| :--- | ---: | ---: | ---: |
| `ExcelReader`, consume rows | 0.453 s | 0.093 s | −80 % |
| `load_workbook` (model) | 2.448 s | 1.307 s | −47 % |

What changed in `_c_core.c`:

- A date cell no longer looks up its number format in a dict and a set, and no
  longer builds a `timedelta` for each value. The date flag of every cell style is
  computed once per sheet, and the conversion reuses the XLSX serial-to-datetime
  function.
- The cells of a row are kept in a buffer that grows geometrically, instead of one
  `realloc` per cell.

Both changes keep the results identical. A differential test compared 45,045 serial
values in both date systems, including the microsecond rounding and the out-of-range
errors (exception type and message), with the baseline and with the new build. A
second test compared 1,000 random worksheets (gaps, unordered and repeated columns,
empty and wide rows, every cell kind, error cells); both builds returned the same
rows. The full test suite passes with the C extension (1883 passed, 183 skipped
because they need Microsoft Excel) and without it (1709 passed, 357 skipped because
they need the extension).

Method changes in this run:

- The read harness imports each reader's library before the timer starts. Earlier
  runs timed the import as well.
- The timed pass only consumes the rows. Earlier runs also counted the non-empty
  cells inside the timed region, which added about 0.12 s to the xlspy XLSB case and
  made the comparison with the other readers less even.
- pyxlsb cannot read a file written by `XlsbWriter` as it is. The writer stores an
  `A1:A1` placeholder in the `BrtWsDim` record, pyxlsb trusts it, and reading a
  row wider than one column raises `IndexError`. The pyxlsb row therefore reads a
  copy in which only that record holds the real bounds. The cells are the same.
  xlspy, python-calamine and fastexcel read the original file.
- Read memory is not reported. The `peak_rss_delta_mb` field in the read JSON is
  the change in resident memory at the end of the read, not the peak.

Reproduce (from the repository root, with the libraries installed:
`python -m pip install openpyxl==3.1.5 xlsxwriter==3.2.9 PyExcelerate python-calamine fastexcel pyarrow pyxlsb psutil`):

```bash
python examples/write_benchmark.py --rows 50000 --cols 50 --iterations 3 --writers all --c-ext both --tracemalloc --output benchmarks/results-write-2026-10-30.json --csv benchmarks/results-write-2026-10-30.csv
python benchmarks/bench_read_compare.py --rows 50000 --cols 50 --repeat 5 --out benchmarks/results-read-compare-2026-10-30-v2.json
```

**Not measured in this run:** a full-model write of openpyxl (`Workbook().save()`)
against `Workbook.save()` in xlspy. pandas and polars call these readers and
writers underneath and add their own overhead, so they were not measured
separately.

## Feature comparison

The table lists the API that each library exposes, checked against the installed
versions. It is not a functional test suite. For the details and the list of gaps,
see [openpyxl parity](openpyxl-parity.md).

| Capability | xlspy | openpyxl 3.1.5 | xlsxwriter 3.3.1 |
| :--- | :---: | :---: | :---: |
| Read XLSX | yes | yes | no |
| Read XLSB | yes | no | no |
| Write XLSX | yes | yes | yes |
| Write XLSB | yes | no | no |
| Streaming write | yes | yes (`write_only`) | yes (`constant_memory`) |
| Streaming read | yes (`ExcelReader`) | yes (`read_only`) | n/a |
| Modify and save an existing file | yes (XLSX and XLSB ¹) | yes | no |
| Cell styles (font, fill, border, number format) | yes | yes | yes |
| Named styles | yes | yes | no |
| Merged cells | yes | yes | yes |
| Hyperlinks | yes | yes | yes |
| Data validation | yes | yes | yes |
| Conditional formatting | yes | yes | yes |
| Tables | yes | yes | yes |
| Charts | yes (14 types) | yes | yes |
| Images | yes | yes | yes |
| Comments | yes | yes | yes |
| Rich text in cells | yes | yes | yes |
| Defined names | yes | yes | yes |
| Freeze panes, autofilter, print area | yes | yes | yes |
| Sheet protection | yes | yes | yes |
| Sparklines | yes | no | yes |
| Array formulas | yes (XLSX, `Formula(array_ref=...)`) | yes | yes |
| Create `.xlsm` / VBA project | no (`ValueError`) | keeps `vba_archive` | yes (`add_vba_project`) |
| Pivot tables | written; loaded ones preserved | preserved only | no |
| Chartsheets | yes | yes | yes |

¹ A loaded XLSB saves the edits listed in [Usage](guide/model.md#loaded-xlsx-and-xlsb-workbooks); the few it refuses raise an error at `save()`.
Some style features are not written to XLSB, so check that page before you rely on a
particular style in XLSB output.

### Python 3.15 (the figures of the README)

The README tables come from this run: commit `08ab0f1`, Python 3.15 (3.15.2 on Windows, 3.15.0 on macOS and WSL2), a fresh virtual environment with the package built in release mode, openpyxl 3.1.5 and xlsxwriter 3.3.1. The write times are `examples/write_benchmark.py --rows 50000 --cols 50 --iterations 3 --native both` (inline strings), the read times the median of three runs of `examples/performance_test.py`, the machines as in the table above.

| 50,000 × 50 | Windows | WSL2 | macOS |
| :--- | ---: | ---: | ---: |
| xlspy XLSB write, native | **0.67 s** | **0.57 s** | **0.41 s** |
| xlspy XLSB write, `XLSPY_ACCEL=python` | 2.78 s | 2.64 s | 1.81 s |
| xlspy XLSX write, native | **1.16 s** | **0.92 s** | **0.69 s** |
| xlspy XLSX write, `XLSPY_ACCEL=python` | 3.73 s | 3.12 s | 2.18 s |
| xlsxwriter (`constant_memory`) write | 11.71 s | 7.89 s | 5.42 s |
| openpyxl (`write_only`) write | 20.39 s | 13.72 s | 9.23 s |
| xlspy XLSB read, native | **0.22 s** | **0.14 s** | **0.07 s** |
| xlspy XLSB read, `XLSPY_ACCEL=python` | 6.41 s | 4.74 s | 3.82 s |
| xlspy XLSX read, native | **0.98 s** | **0.58 s** | **0.44 s** |
| openpyxl (`read_only`) read | 14.35 s | 8.13 s | 6.14 s |

Files: [`python315-windows-write.json`](benchmark-results/python315-windows-write.json), [`python315-windows-read.log`](benchmark-results/python315-windows-read.log), [`python315-wsl2-write.json`](benchmark-results/python315-wsl2-write.json), [`python315-wsl2-read.log`](benchmark-results/python315-wsl2-read.log), [`python315-mac-m4-write.json`](benchmark-results/python315-mac-m4-write.json), [`python315-mac-m4-read.log`](benchmark-results/python315-mac-m4-read.log).

- The Windows times are higher than in the Python 3.14 run above for every library, the third-party ones included (xlsxwriter 10.17 s → 11.71 s, openpyxl 17.06 s → 20.39 s), so the cause is not in xlspy; the Python version and the load of the machine on the day were not separated. The ratios hold.
- The native module writes **3.1–4.6×** faster than the pure-Python path and reads XLSB **29–55×** faster.

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

Current results (Rust module):

- Windows: [`windows-2026-11-02-rust-write.json`](benchmark-results/windows-2026-11-02-rust-write.json) (write), [`read-windows-2026-11-02-rust.log`](benchmark-results/read-windows-2026-11-02-rust.log) (read) and [`windows-2026-11-02-rust-io.json`](benchmark-results/windows-2026-11-02-rust-io.json) (read modes and model; since then [`windows-2026-11-04-r3-io.json`](benchmark-results/windows-2026-11-04-r3-io.json))
- macOS (Apple M4): [`mac-m4-2026-11-02-rust-write.json`](benchmark-results/mac-m4-2026-11-02-rust-write.json), [`read-mac-m4-2026-11-02-rust.log`](benchmark-results/read-mac-m4-2026-11-02-rust.log) and [`mac-m4-2026-11-02-rust-io.json`](benchmark-results/mac-m4-2026-11-02-rust-io.json)
- WSL2: [`wsl2-2026-11-02-rust-write.json`](benchmark-results/wsl2-2026-11-02-rust-write.json), [`read-wsl2-2026-11-02-rust.log`](benchmark-results/read-wsl2-2026-11-02-rust.log) and [`wsl2-2026-11-02-rust-io.json`](benchmark-results/wsl2-2026-11-02-rust-io.json)

Historical results (C extension):

The write data for each system, with the environment, every run, and the medians, is in JSON. The read data is the full text output of three runs of `performance_test.py` per system:

- [`windows-2026-10-19.json`](benchmark-results/windows-2026-10-19.json) (Windows, Python 3.14.7, write)
- [`mac-m4-2026-10-19.json`](benchmark-results/mac-m4-2026-10-19.json) (macOS, Python 3.14.7, write)
- [`wsl2-2026-10-19.json`](benchmark-results/wsl2-2026-10-19.json) (WSL2, Python 3.14.8, write)
- [`read-windows-2026-10-19.log`](benchmark-results/read-windows-2026-10-19.log) (Windows, Python 3.14.7, read)
- [`read-mac-m4-2026-10-19.log`](benchmark-results/read-mac-m4-2026-10-19.log) (macOS, Python 3.14.7, read)
- [`read-wsl2-2026-10-19.log`](benchmark-results/read-wsl2-2026-10-19.log) (WSL2, Python 3.14.8, read)

## Run the benchmarks

Install the test extras and check that native acceleration is loaded:

```bash
python -m pip install -e '.[test]'
python -c "from xlspy._accel import HAVE_NATIVE; assert HAVE_NATIVE"
```

[`examples/write_benchmark.py`](https://github.com/justybase/xlspy/blob/main/examples/write_benchmark.py) produced the tables above. It runs every measurement in a fresh worker process, repeats each writer three times by default, interleaves the writers, and reports medians:

```bash
# The full comparison: all writers, native module on and off, one tracemalloc run per writer
python examples/write_benchmark.py --rows 50000 --cols 50 --iterations 3 --writers all --native both --tracemalloc --output bench.json --csv bench.csv

# Only some writers
python examples/write_benchmark.py --writers xlsx-inline,xlsb-inline --native both

# Another source tree, for before/after runs
python examples/write_benchmark.py --src ../other-checkout/src --csv other.csv
```

- **Writers**: `xlsx-shared`, `xlsx-inline`, `xlsb-shared`, `xlsb-inline` (xlspy, `--native on|off|both`; `off` sets `XLSPY_ACCEL=python`), `openpyxl` (`write_only=True`), `xlsxwriter` (`constant_memory=True`) and `pyexcelerate`. The other libraries are optional; missing ones are skipped. The JSON field is `native`.
- **Data**: `--dataset mixed` (default) is the eight-type layout of `performance_test.py` and is deterministic. `--dataset random --seed N` is a seeded, text-heavy alternative. With `--input stream` (default), every writer consumes the same generator. With `--input list`, the rows are built before the clock starts.
- **Metrics**: wall time, peak resident memory, output size, and, with `--tracemalloc`, a separate untimed run's Python allocation peak (memory the native extension allocates directly is not traced).
- **Targets**: `--target file` (default) or `--target memory` (an `io.BytesIO`, xlspy writers only).
- **Output**: `--output` writes JSON with the environment, every run, and the medians. The date is recorded as a day only. `--csv` writes one row per run. A failed run is recorded with its error and does not stop the series.

The harness needs no third-party packages except the optional openpyxl and xlsxwriter. It is not part of the test suite.

### Read benchmark

The read numbers come from `examples/performance_test.py`. It runs the write and the read in one pass and prints a table. Run it three times from a scratch directory, because it writes its test files to the current directory, and take the median of each read row:

```bash
python examples/performance_test.py --rows 50000 --cols 50
```

### Read modes and model benchmark

`benchmarks/bench_io.py` runs each case in a fresh process and writes JSON.
`--compare` checks a run against a saved one and exits non-zero when a case is
slower than `--max-slowdown` times its baseline:

```bash
py -3.14 benchmarks/bench_io.py --out results.json
py -3.14 benchmarks/bench_io.py --compare docs/benchmark-results/windows-2026-11-04-r3-io.json --max-slowdown 1.03
```

For correctness checks and platform test results, see [Validation](validation.md).
