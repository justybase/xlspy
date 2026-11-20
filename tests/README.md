# Test suite

For recorded platform runs and compatibility checks, see
[Validation and test results](../docs/validation.md).

## Install and run

xlspy needs Python 3.11 or later (`requires-python = ">=3.11"`); CI runs
3.11–3.15. Where the default `python` is older, name the version, e.g.
`py -3.14` on Windows. Building from source needs a Rust toolchain
(`rustup`) for the native extension, `rust/xlspy_core` (built by maturin; see
`src/xlspy/_accel.py`, where `XLSPY_ACCEL=rust|python` picks the backend).

The test dependencies are optional and do not become runtime dependencies:

```bash
python -m pip install -e '.[test]'
python -m pytest
```

Without the native extension (the pure-Python fallback that CI also runs):

```bash
XLSPY_ACCEL=python python -m pytest        # PowerShell: $env:XLSPY_ACCEL="python"
```

Tests that drive Microsoft Excel through COM (Windows, `pywin32`) are skipped
unless `XLSPY_RUN_EXCEL_COM=1` is set; run them with `-p no:faulthandler`.

The conformance suite (`test_excel_com_conformance.py`, driven by the data in
`com_cases.py` and the Excel helpers in `com_spec.py`) writes eight cases as XLSX
and XLSB with xlspy and reads them back in Excel, writes the same cases in Excel
and reads them with xlspy, and replaces a template's body with an updater. It
needs no PowerShell 7 and no external `comparison/` workspace. Ignored errors and
the updater's formulas have their own COM modules
(`test_excel_com_ignored_errors.py`, `test_xlsb_updater_formulas.py`).

## Lint and type checks

CI runs both. mypy checks every module in `src/xlspy` with
`disallow_untyped_defs`; the modules listed under `[[tool.mypy.overrides]]` in
`pyproject.toml` may still have unannotated defs (their bodies are checked). A new
module must be fully annotated:

```bash
python -m pip install ruff mypy
ruff check src tests tools benchmarks examples
mypy
```

The suite is configured to treat warnings as errors and uses `pytest-randomly`
to vary test order. Reproduce a particular order with a fixed seed:

```bash
python -m pytest --randomly-seed=12345
```

To measure coverage the way CI does, over both backends (the branch-coverage
threshold, `fail_under` in `pyproject.toml`, is 90%; only explicit
`pragma: no cover` lines are excluded):

```bash
python -m pytest --cov --cov-report=
XLSPY_ACCEL=python python -m pytest --cov --cov-append --cov-report=
python -m coverage report
```

## openpyxl compatibility harness

The parity harness compares xlspy with openpyxl-authored goldens. Regenerate
the goldens after changing the generator, then run the harness:

```bash
python tools/make_openpyxl_goldens.py
python -m pytest tests/openpyxl_compat -q
python tools/compare_openpyxl_model.py LEFT.xlsx RIGHT.xlsx
```

The harness uses openpyxl 3.1.5 (pinned in the `test` extra). A missing feature is
reported as `xfail` with an ID from
`tests/openpyxl_compat/expected_gaps.toml`. That manifest is the source of truth
for the expected gaps; `test_gap_manifest.py` keeps it in sync with
[`docs/openpyxl-parity.md`](../docs/openpyxl-parity.md), which records the
current implementation status and the remaining phases.

The notes/template regression suites are `test_comments_and_editing.py` and
`test_xlsx_reader_streaming.py`. For native Excel validation of notes, dimensions,
named ranges, protection and preserved form controls:

```powershell
$env:XLSPY_RUN_EXCEL_COM = "1"
python -m pytest tests/test_excel_com_metadata.py
python -m pytest tests/test_xlsb_metadata_com.py
python -m pytest tests/test_excel_com_review_fixes.py
python -m pytest tests/test_excel_com_conformance.py tests/test_excel_com_ignored_errors.py
```

## Excel-written corpus and local Excel validation

`tests/data/excel_365/` holds `features.xlsx` and `features.xlsb` written by
desktop Excel (Microsoft 365 reports version 16.0) plus `manifest.json` with the
build and file hashes. CI reads the committed corpus without Excel; regenerate
it on Windows with office installed:

```bash
python tools/make_excel_fixtures.py
python -m pytest tests/test_excel_fixtures.py
```

The COM and printer tests are opt-in because GitHub Actions has no Excel. Before
a release, record a local run as an artifact and check it:

```powershell
python tools/excel_validation.py
python tools/excel_validation.py --check
```

`--include-printer` adds the PageSetup tests. The artifact records the Excel
version/build, commit and JUnit counts in `validation/`; see
[`docs/validation.md`](../docs/validation.md#local-excel-validation).

LibreOffice is a secondary consumer, not an Excel replacement. Its smoke tests
convert xlspy output with headless LibreOffice and are skipped when `soffice` is
not installed; CI runs them as a non-blocking job:

```bash
python -m pytest tests/test_libreoffice_xlsb.py -q -m libreoffice
```

## Native extension

Most tests run through whichever backend `_accel.py` loaded; the
`test_native_*` modules and a few regressions pin both the native code and
the pure-Python path to the same results. After changing the Rust code,
rebuild it with `python -m pip install -e '.[test]'` (or `maturin develop
--release`), then run the Rust checks too:

```bash
cd rust/xlspy_core
cargo fmt --check && cargo clippy --release --locked --all-targets -- -D warnings
cargo test --release --locked
cd ../..
python -m pytest tests/test_native_*.py tests/test_xlsb_native_dates.py
```

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
