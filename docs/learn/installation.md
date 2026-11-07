---
description: Install xlspy, verify the native extension, and choose the optional pandas or polars extras.
---

# Installation

xlspy needs **Python 3.11 or newer**. Microsoft Excel is not required.

## Install from PyPI

```bash
pip install xlspy
```

Wheels include the native Rust extension. If you install from the source
distribution, the build needs Rust 1.88 or newer:
[rustup](https://rustup.rs).

## Verify the install

```bash
python -c "import xlspy; print(xlspy.__version__)"
```

To see which backend is active:

```bash
python -c "import xlspy._accel as a; print(a.HAVE_NATIVE, a.ACCEL_BACKEND)"
```

`HAVE_NATIVE=True` means the Rust module is loaded. Set `XLSPY_ACCEL=python` to
force the pure-Python fallback; everything still works, only slower.

## Optional extras

| Extra | Install | Adds |
| :--- | :--- | :--- |
| pandas | `pip install "xlspy[pandas]"` | `to_xlsx()`, `to_xlsb()` for pandas DataFrames |
| polars | `pip install "xlspy[polars]"` | the same API for polars DataFrames |
| docs | `pip install -e ".[docs]"` | Zensical and mkdocstrings, for building this site |

The core package has **no required dependencies**. pandas and polars are imported
lazily, only when you pass a DataFrame.

## Choose a format

| | XLSX | XLSB |
| :--- | :--- | :--- |
| Excel support | every Excel version | Excel 2007+ |
| Read/write speed | fast | faster |
| File size | larger | smaller |
| Third-party readers | most tools | fewer tools |
| Use it for | interchange, reports | large exports, pipelines |

Both formats use the same API: `XlsxWriter` and `XlsbWriter` take the same rows,
styles, charts and tables.

!!! tip "Next"
    Continue with [Your first workbook](first-workbook.md).
