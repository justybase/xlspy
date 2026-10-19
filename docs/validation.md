# Validation and test results

[Home](../README.md) · [Benchmarks](benchmarks.md) · [Running the test suite](../tests/README.md)

## Windows validation (2026-10-19)

Run on Windows 11 build 26300 with Python 3.14.6, pytest 9.1.1, and Microsoft
Excel 16.0, in the same environment as the [benchmarks](benchmarks.md#results).
The C extension was built with MSVC without warnings.

- Full suite with the C extension (`python -m pytest tests`): **632 passed,
  13 skipped**. The skips are the optional `nzpy_extended` dependency and the
  opt-in Excel COM tests.
- Full suite, pure-Python fallback (`XLSPY_DISABLE_C_EXT=1`): **518 passed,
  127 skipped**. The extra skips are the C-only tests and the C-encoder
  variants of the cell-text tests.
- Excel COM tests (`XLSPY_RUN_EXCEL_COM=1 python -m pytest
  tests/test_excel_com_features.py tests/test_excel_com_updaters.py
  tests/test_comprehensive.py`): **26 passed**. Excel opens each file and the
  tests read back what it sees:
  - The combined report from `tests/test_xlsx_integration.py`: computed formulas
    (including cross-sheet ones), the merge, hyperlinks, both validations, the
    table, conditional formatting, three charts, the picture, print area,
    repeated rows, orientation and paper size, and freeze panes.
  - Freeze panes in XLSX and XLSB (rows and columns, columns only, and the
    default frozen header).
  - Tables whose names come from their header cells, including escaped text and
    a header row 150 rows down.
  - Workbooks written to an `io.BytesIO`.
  - Cell text in all four modes: control characters, CR/LF, `_xHHHH_` look-alike
    text and U+FFFE come back exactly, and at the limit Excel's `LEN` returns
    32,767 (16,383 astral characters plus one ASCII character are 16,384 code
    points but 32,767 UTF-16 units).

  The in-process COM tests in `test_comprehensive.py` print
  `Windows fatal exception: code 0x80010108` during Excel/pywin32 shutdown on
  this host; the exit code is 0 and every assertion passed.
- Excel's behaviour behind two of the fixes, checked with COM probes:
  - Cells over 32,767 units, written with the new check switched off: Excel
    refuses to open the XLSX (shared and inline) and inline-XLSB files. The
    shared-string XLSB file opens with the text cut to 32,767 units, because its
    string table writer truncates silently, as 0.7.0 did.
  - Tables: Excel refuses the workbook when a header cell differs from its
    column name, is empty, is numeric, or repeats another name, and when the
    table name is a cell reference such as `T1`. A difference in case only is
    accepted. 0.7.0 wrote `Column1`… names for tables without `columns`, so such
    workbooks did not open.
- Differential fuzzing (1,000 seeded random sheets × XLSX/XLSB × shared/inline):
  the C and Python encoders produced byte-identical archives for all 4,000
  pairs, and refused the same 864 identically.
- Mutation checks: re-introducing each fixed defect in a copy of the package
  (no length check, no XML escaping, publishing on failure, writing straight
  into the caller's stream, C without the shared-string rollback and with
  `str()` of unknown objects) makes the new tests fail in every case.

### Sanitizers (WSL2)

The C extension was built with `-fsanitize=address,undefined` (GCC 15.2.0,
`-O1 -g`) in a copy of the working tree under WSL2 (Python 3.14.4), and
loaded into the uninstrumented interpreter with `LD_PRELOAD`.
`PYTHONMALLOC=malloc` routed every `PyMem` allocation through malloc, so
ASan saw the extension's buffers. Leak detection was off, because CPython
does not free everything at exit; the suite's own refcount, RSS and
tracemalloc leak tests ran instead.

- The first run found undefined behaviour: the XLSB encoder left-shifted
  negative integers when writing RK numbers (`_c_core.c`, reported by UBSan
  on `-536870912`). It is fixed, along with a related 32-bit shift in the
  reader that only overflows where `long` is 32 bits (Windows).
- Final run: the full suite passed (**620 passed, 25 skipped**; the skips are
  Windows-only tests), the 1,000-seed differential fuzz passed, and ASan and
  UBSan reported nothing.
- Valgrind is not installed in this environment and was not used.

### Known limitations

- **Older Excel versions were not tested.** Only Excel 16.0 was available, so
  compatibility with other Excel releases is unverified. Dynamic-array spill
  formulas are not supported.
- **Copying into a caller's stream cannot be undone.** A failed write leaves a
  file object untouched, but if the stream itself fails while the finished
  workbook is being copied in, it holds part of the workbook (see the usage
  guide's *Failure handling*).
- **Formula text length** (Excel allows 8,192 characters) is not checked.
- **Updaters** (`XlsxUpdater`, `XlsbUpdater`) were not changed: they do not
  apply the new cell-text checks or escapes, or the staged output.
