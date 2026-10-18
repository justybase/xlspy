# Benchmarks

[Home](../README.md) · [Usage guide](usage.md) · [Validation](validation.md)

`xlspy` is designed for high performance. Since version 0.1.0, the library includes a **C extension** (`_c_core`) that accelerates XLSB read and write as well as the plain XLSX write path. The C extension is **enabled by default** (compiled automatically on install). Set `XLSPY_DISABLE_C_EXT=1` to force the pure Python fallback.

Dataset: **50,000 × 50** (2.5 million cells). Results are the **median of three
runs** per environment. Windows figures are the previously recorded benchmark
results ([original Windows report](https://github.com/justybase/xlspy/blob/ed7e618/README.md#performance)); WSL2 figures were measured on **2026-10-18**.

## Environments

| | Windows | WSL2 |
| :--- | :--- | :--- |
| Platform | Windows AMD64 | Linux x86_64, WSL2 kernel 6.18.54.1 |
| CPU | AMD Ryzen 7 7840HS | AMD Ryzen 7 7840HS, 16 logical CPUs |
| Python | 3.14.6 | 3.14.4 |
| xlspy | 0.7.0 | 0.7.0 |
| openpyxl / xlsxwriter | 3.1.5 / 3.2.9 | 3.1.5 / 3.2.9 |
| Benchmark storage | Not recorded | WSL2 ext4 filesystem |

The WSL2 extension was built with GCC 15.2.0 and `-O2`; openpyxl ran without
optional lxml acceleration. Equivalent build and lxml details were not
recorded for Windows. These are results for each environment, with different
Python versions and configurations; they do not isolate the effect of the OS.

## Write

| Library | Format | Windows | WSL2 | Windows size | WSL2 size |
| :--- | :--- | ---: | ---: | ---: | ---: |
| **xlspy (C_EXT)** | XLSB | 1.46 s | 1.17 s | 7.25 MiB | 7.52 MiB |
| xlspy (Python) | XLSB | 3.21 s | 2.41 s | 7.25 MiB | 7.52 MiB |
| xlspy | XLSX | 4.56 s | 1.96 s | 6.43 MiB | 6.71 MiB |
| openpyxl (write_only) | XLSX | 17.77 s | 14.58 s | 11.71 MiB | 11.48 MiB |
| xlsxwriter (constant_memory) | XLSX | 9.85 s | 7.90 s | 11.57 MiB | 11.35 MiB |

## Read

| Library | Format | Windows | WSL2 |
| :--- | :--- | ---: | ---: |
| **xlspy (C_EXT)** | XLSB | 0.60 s | 0.43 s |
| xlspy | XLSX | 4.96 s | 3.87 s |
| xlspy (Python) | XLSB | 6.33 s | 4.64 s |
| openpyxl (read_only) | XLSX | 11.06 s | 9.23 s |

The [WSL2 measurement data](benchmark-results/wsl2-2026-10-18.json) contains
all three runs, the environment, and the calculated medians.

## Methodology

The dataset mixes strings, integers, floats, booleans, dates, datetimes,
Decimals, and large integers. The 50,000 rows include a header row. Write
timings include data generation, serialization, compression, and closing the
file. Reads consume all rows; each library reads its own output, with both
xlspy XLSB readers using the C writer's file.

The Python fallback runs in a separate process with `XLSPY_DISABLE_C_EXT=1`.
xlsxwriter uses `constant_memory=True`; openpyxl uses `write_only=True` for
writes and `read_only=True` for reads. Files are read after writing, without
flushing OS caches; timings do not measure forced durable writes to disk.

## Run the benchmarks

Install the benchmark dependencies and check that C acceleration is loaded:

```bash
python -m pip install -e '.[test]'
python -c "from xlspy._accel import HAVE_C_EXT; assert HAVE_C_EXT"
```

From the repository root, run the following command **three times**, then
report the median time for each operation:

```bash
python examples/performance_test.py --rows 50000 --cols 50
```

See the [benchmark scripts](../examples/README.md#benchmarks) for the available
comparisons. Results depend on the dataset, hardware, Python version, and
library versions; these measurements describe the environment above.

For correctness checks and platform test results, see [Validation](validation.md).
