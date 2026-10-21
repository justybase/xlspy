# Test suite

For recorded platform runs and compatibility checks, see
[Validation and test results](../docs/validation.md).

## Install and run

The test dependencies are optional and do not become runtime dependencies:

```bash
python -m pip install -e '.[test]'
python -m pytest
```

The suite is configured to treat warnings as errors and uses `pytest-randomly`
to vary test order. Reproduce a particular order with a fixed seed:

```bash
python -m pytest --randomly-seed=12345
```

To measure coverage (the configured branch-coverage threshold is 84%):

```bash
python -m pytest --cov
```

The notes/template regression suites are `test_comments_and_editing.py` and
`test_xlsx_reader_streaming.py`. For native Excel validation of notes, dimensions,
named ranges, protection and preserved form controls:

```powershell
$env:XLSPY_RUN_EXCEL_COM = "1"
python -m pytest tests/test_excel_com_metadata.py
python -m pytest tests/test_xlsb_metadata_com.py
python -m pytest tests/test_excel_com_review_fixes.py
```

## Optional C extension

Most XLSB tests run through the pure-Python reader by default when the
extension is absent. A few regressions explicitly cover both implementations.
Build the local extension and run those tests with:

```bash
gcc -shared -fPIC -O2 \
  -I"$(python -c 'import sysconfig; print(sysconfig.get_path("include"))')" \
  src/xlspy/c_ext/_c_core.c \
  -o "src/xlspy/_c_core$(python3-config --extension-suffix)"
python -m pytest tests/test_reader.py tests/test_excel_dates.py tests/test_xlsb_unicode.py
```

Remove the generated, git-ignored `.so` afterward to restore the pure-Python
test environment.

## Cross-validation goldens

`tests/data/golden_*.xlsx` are written by independent XLSX writers (openpyxl
and xlsxwriter); their JSON sidecars are expected values. The tests validate
the sidecars with both xlspy and openpyxl, which is the reference reader for
XLSX. XLSB goldens can be added when an independent Excel-produced fixture is
available; openpyxl does not read XLSB.

Regenerate the XLSX fixtures and sidecars with:

```bash
python tools/make_goldens.py
```

The date suite also exercises the 1900 leap-year compatibility boundary and
the 1904 date system. Hypothesis property tests are deterministic smoke tests
(50 examples per property), not extended fuzzing or nightly jobs.

## Optional integrations

Tests needing pandas, polars, openpyxl, or xlsxwriter skip when the relevant
package is not installed. The DataFrame feature itself remains optional and
is tested for both pandas and polars when available. Excel COM tests require
Windows, Microsoft Excel, and `XLSPY_RUN_EXCEL_COM=1`.
