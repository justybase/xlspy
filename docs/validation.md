# Validation and test results

[Home](../README.md) · [Benchmarks](benchmarks.md) · [Running the test suite](../tests/README.md)

## Windows validation

Validation reported for the Windows/Python 3.14.6 environment used in the
[benchmarks](benchmarks.md):

- C-extension test run: **268 passed, 2 skipped**; the three direct Excel COM
  tests were run separately and passed. The skips were the optional real-data
  dependency and the opt-in COM updater test.
- Pure-Python fallback test run (`XLSPY_DISABLE_C_EXT=1`): **242 passed,
  28 skipped**; C-only tests are expected to skip in this mode.
- Cross-validation against **openpyxl** and **xlsxwriter**: **11 passed**.
- Excel COM validation: direct XLSB/XLSX open/readback tests **3 passed**;
  PowerShell validation of updated XLSB/XLSX files **1 passed**.
- Additional differential gates compared C and Python writer payloads and
  reader grids, including seeded mixed-value cases, a 50000 × 50 workbook,
  malformed BIFF records, date-system edge cases, and shared-string truncation
  and corruption cases.

The legacy in-process COM tests pass but emit `0x80010108` shutdown diagnostics
from Excel/pywin32 on this host; the separate PowerShell COM validation passed
without that diagnostic.

## Local test run (WSL2)

Run on **2026-10-18** in WSL2 (Linux x86_64, Python 3.14.4, pytest 9.1.1):

- With the C extension built: **257 passed, 13 skipped**.
- Pure-Python fallback (without the C extension): **239 passed, 31 skipped**.
- The skips were for optional `nzpy_extended`/`psutil` dependencies and
  Windows Excel COM tests, which cannot run in this WSL2 environment.

These results are for WSL2/Linux; the CI workflow separately includes native
Windows runners in its operating-system matrix.
