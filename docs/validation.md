---
description: "Platform test results, Excel COM conformance and known limitations."
---

# Validation and test results

[Home](index.md) · [Benchmarks](benchmarks.md) · [Running the test suite](https://github.com/justybase/xlspy/blob/main/tests/README.md)

## Known limitations

- **Older Excel versions were not tested.** Only Excel 16.0 was available, so
  compatibility with other Excel releases is unverified.
- **Copying into a caller's stream cannot be undone.** A failed write leaves a
  file object untouched, but if the stream itself fails while the finished
  workbook is being copied in, it holds part of the workbook (see
  [Failure handling](guide/writing.md#failure-handling)).

## Local Excel validation

GitHub Actions has no Microsoft Excel, so the `com` tests (and the `printer`
tests, which need a default printer) are opt-in. Before a release, run them on a
Windows machine with desktop Excel and commit the artifact:

```powershell
python tools/excel_validation.py                    # pytest -m com
python tools/excel_validation.py --include-printer  # also PageSetup tests
python tools/excel_validation.py --check            # before tagging
```

The tool records the Excel version and build, the commit, the Python version and
the JUnit counts in `validation/excel-validation-<date>-<sha>.json`, and prints
the table row below. `--check` refuses an artifact that was recorded from a dirty
tree, is older than 90 days, or whose commit is no longer an ancestor of `HEAD`
(changes confined to `docs/`, `validation/` and `CHANGELOG.md` are allowed).
Excel 2016 and Microsoft 365 both report version 16.0; the build number
distinguishes them. CI cannot replace this step, but the committed artifact
keeps the result reviewable.

| Date | Excel | Commit | Result |
| :--- | :--- | :--- | :--- |
| 2026-11-20 | 16.0 (20527.0) | `4343d33` | 316 passed, 0 failed, 22 skipped |

The first artifact is
`validation/excel-validation-2026-11-20-4343d33.json`. The skips are
environment limits (no default printer, display scaling, VBA trust); the Excel
conformance tests themselves all passed.

The LibreOffice tests (`pytest -m libreoffice`) are a secondary check: headless
LibreOffice can convert the output and read basic cells, but its XLSX/XLSB
support is not Excel. A LibreOffice failure is informative, not release-blocking,
unless it shows a package that Excel also refuses.

## Quality-signal refresh (2026-11-20)

Windows 11, Python 3.14.2. Coverage now runs without the broad
`raise NotImplementedError/UnsupportedFeatureError` exclusion: the combined
native and pure-Python run measures **91%** branch coverage (the floor is 90%,
and only explicit `pragma: no cover` lines are excluded). The native run passed
**10,109 tests, 343 skipped, 1 xfailed** in 4:55; the fallback run passed
**9,003 tests, 1,450 skipped** in 3:20. The openpyxl harness reads its expected
gaps from `tests/openpyxl_compat/expected_gaps.toml` and passes **93** tests;
`cargo test` passes **68** unit tests; the LibreOffice smoke suite passes **31**.
The Excel 365 corpus in `tests/data/excel_365/` loads in both formats and
exposed Excel's `indexed="81"` colors, which the model now accepts. The local
COM run recorded above passed 316 tests with Excel 16.0 (build 20527.0).

## Updater convergence (2026-11-20)

Windows 11, Python 3.14.2. `tests/test_updater_differential.py` (58 tests) runs
the same replacement, formula, note, link, name and protection edits through
`XlsxUpdater` and `XlsbUpdater` and compares the loaded model; array and spill
ranges are compared in the saved package (dimension and auto-filter), because
XLSB stores the member cells and XLSX does not. The harness found and fixed
three differences: XLSX did not widen `<dimension>`/auto-filter for array and
spill ranges, a `None` header was not written as an empty cell in XLSX, and the
pure-Python row encoder accepted values past column XFD. The full native suite
passed **10,164 tests, 343 skipped, 1 xfailed** in 5:10; the pure-Python pass
took **9,060 tests, 1,450 skipped** in 2:33. Combined coverage is **91%** branch
coverage over 27,214 statements (the floor is 90%). The LibreOffice smoke suite
passes **31** tests.

## Updater option parity (2026-11-20)

Windows 11, Python 3.14.2. The differential harness grew to **65 tests** and
now also compares `update_cells(styles=...)`, `set_merged_cells()` and
`set_data_validations()` across both formats, including clearing the last two
and the overlap/type errors. It found that XLSX dropped `allow_blank=False`
through a round trip because the writer omitted `allowBlank`; the writer now
emits `allowBlank="0"`, matching the reader's default. The full native suite
passed **10,174 tests, 343 skipped, 1 xfailed** in 5:05; the pure-Python pass
took **9,068 tests, 1,450 skipped** in 2:47. Combined coverage is **91%**
branch coverage over 27,263 statements (the floor is 90%). The LibreOffice
smoke suite passes **31** tests.

## Coverage hardening and nightly stress (2026-11-20)

Windows 11, Python 3.14.2. Added direct unit suites for the rules that the
end-to-end tests only touch incidentally: `tests/test_sort_state_validation.py`
(24 tests), `tests/test_table_totals_units.py` (38),
`tests/test_rich_text_units.py` (27), `tests/test_formula_tokenizer_edges.py`
(7) and `tests/test_loaded_styles_unit.py` (18). The five lowest modules rose
from 77% to 94% combined (`_sort_state` and `_table_totals` 100%, tokenizer and
`rich_text` 99%, `workbook/_styles.py` 89%). The full native suite passed
**10,311 tests, 343 skipped, 1 xfailed** in 5:03; the pure-Python pass took
**9,205 tests, 1,450 skipped** in 2:47. Combined coverage is **92%** branch
coverage over 27,240 statements.

`.github/workflows/nightly.yml` adds two off-PR checks: a long randomized-edit
and native-fuzz run (`XLSPY_FUZZ_ITERATIONS=2000`) across seeds 7, 101 and
2024, and a non-blocking job on the latest `openpyxl`, `pandas` and `polars`.
The regular CI matrix keeps the short deterministic fuzz budget.

The writer and the loaded style table now share `_style_xml.py` for color and
border fragments; `tests/test_style_xml_shared.py` compares both callers and
pins the `tint` spelling. `chart_parts.column_letters()` was corrected to end
at `XFD` after 16,384 entries instead of continuing with invalid labels.
`updater_utils.py` was split into `_zip_package.py` and `_xml_text.py`;
`tests/test_updater_utils_split.py` keeps the old names bound to the same
objects.

## XLSX sheet management (2026-11-20)

`XlsxUpdater` now has the sheet metadata API that `XlsbUpdater` already had:
`get_sheet_states()`, `get_active_sheet()`, `set_active_sheet()`,
`set_sheet_state()` and `reorder_sheets()`. The XLSX side edits `<sheet>`
attributes and the workbook view in `xl/workbook.xml`; a reorder carries the
`localSheetId` of defined names and the `firstSheet` tab-bar position with the
sheets, and falls back to the first visible sheet when the active one becomes
hidden. Chartsheets are included in the order.

`tests/test_updater_differential.py` grew from 65 to 74 tests and compares both
formats for visibility, active tab, local names and the error cases.
`tests/test_xlsx_sheets_unit.py` adds 6 XML-level tests (missing `<bookViews>`,
invalid `activeTab`, unknown state, tab-bar remapping, hidden active sheet).
The full native suite passed **10,326 tests, 343 skipped, 1 xfailed** in 4:16;
the pure-Python pass took **9,220 tests, 1,450 skipped** in 2:47. Combined
coverage is **92%** branch coverage over 27,353 statements (the floor is 90%).
`ruff` and `mypy` (146 files) are clean, `cargo test` passes 68, and the
LibreOffice smoke suite passes 31. The opt-in Excel COM test, extended to
reorder and hide tabs in both formats, passes locally; a LibreOffice conversion
of a reordered XLSX kept the order, visibility and active sheet (LibreOffice
reads `veryHidden` back as `hidden`).

## Sheet structure and edit fuzzing (2026-11-21)

Windows 11, Python 3.14.2. The XLSX sheet API now matches the XLSB one end to
end. `tests/test_updater_differential.py` grew to **88 tests** and drives
`add_sheet()`, `remove_sheet()` and `rename_sheet()`, as well as visibility,
active tab and reorder, through both formats; it compares the loaded model and
the error types (duplicate or invalid names, removing the last sheet, pivot
tables and slicers). `tests/test_xlsx_sheets_unit.py` (14 tests),
`tests/test_sheet_calc_chain.py` (2), `tests/test_sheet_structural_fuzz.py` (2),
`tests/test_cell_utils_edges.py` (11) and `tests/test_hyperlink_validation.py`
(13) cover the XML-level edge cases, `calcChain` invalidation and the value
helpers directly.

The long edit fuzz (`XLSPY_FUZZ_ITERATIONS=2000`) passes on seeds **7, 101 and
2024** (**17 passed, 1 skipped** per seed; the skip needs Excel COM). It found
and fixed five differences:

- assigning a hyperlink to a blank cell changed its value without marking the
  cell dirty, so a save dropped the text;
- merging a range left the hidden cells' old values in the package;
- neither updater allowed clearing a hidden cell of a merged range, although
  the model clears it when the merge is made;
- XLSX saved cell updates before new merges and XLSB before removed merges, so
  an edit made in the same save could be judged against the old merge set;
- the differential byte comparison depended on a process-wide XML namespace
  registration, so the native and pure-Python subprocesses could serialize
  relationships differently.

The full native suite passed **10,376 tests, 343 skipped, 1 xfailed** in 4:59;
the pure-Python pass took **9,270 tests, 1,450 skipped** in 3:08. Combined
coverage is **92%** branch coverage over 27,664 statements (the floor is 90%).
`ruff` and `mypy` (146 source files) are clean, `cargo test` passes 68, and the
LibreOffice smoke suite passes 31.

## Release and nightly hardening (2026-11-21)

Windows 11, Python 3.14.2. The nightly stress job now runs
`tests/test_sheet_structural_fuzz.py` and `tests/test_updater_differential.py`
alongside the long edit and native fuzz for seeds 7, 101 and 2024, and the
workflow declares read-only `contents` permissions. The release test job runs
the suite with `XLSPY_FUZZ_ITERATIONS=250`, so a tag gets a wider fuzz sample
than a pull request without stretching the job.

`tools/wheel_smoke.py` now adds a sheet, renames another and removes a third
through the updater in both formats, then reads the data back. The smoke test
runs on every wheel cibuildwheel produces (Linux x86-64 and aarch64, Windows
x64 and ARM64, macOS arm64 and x86-64) and in the sdist job. The built
`xlspy-0.15.0-cp314-cp314-win_amd64.whl` installed into a clean Python 3.14
environment passed it. The opt-in Excel COM check gained updater `add_sheet`,
`rename_sheet` and `remove_sheet` cases for both formats and passes locally
with Excel 16.0: the workbook opens with the expected tab order and values.

## Quality baseline and module boundaries (2026-11-21)

Windows 11, Python 3.14.2. The first refactoring pass (quality gates and
low-risk deduplication) is in. `ruff format` now owns formatting: 354 Python
files were reformatted in a mechanical commit listed in
`.git-blame-ignore-revs`, and both the lint job and pre-commit check
`ruff format --check`. Ruff also selects `B`, `SIM` and `C901`; the complexity
threshold is 20, with the known functions listed in per-file ignores until they
are split. Four import-linter contracts guard the layers: the core model does
not directly import the public I/O facades, internal modules do not import the
`xlspy` facade, readers/writers/updaters do not import `xlspy.workbook`, and
XLSB modules do not import XLSX modules.

The low-risk deduplication removed three copies of the tab-bar remap, moved the
shared sheet states and next-part logic to `_sheet_common.py`, moved the shared
edit transaction and defined-name resolution to `_edit_base.py`, moved row
layout to `_rows.py`, and moved XML serialization and regex helpers to
`_metadata.py`. `XlsbEditMixin` no longer inherits `XlsxEditMixin`; the nine
XML-only methods it never used are gone from its MRO. The duplicated COM
fixtures in eight test modules were replaced by the `conftest.py` fixtures.

The full native suite passed **10,376 tests, 345 skipped, 1 xfailed** in 4:58;
the pure-Python pass took **9,270 tests, 1,452 skipped** in 2:51. Combined
coverage is **92%** branch coverage over 27,651 statements. `ruff check`,
`ruff format --check`, `mypy` (149 source files) and `lint-imports` (4
contracts) are clean; the COM subset rerun locally after the fixture cleanup
passed **103 tests, 6 skipped** with Excel 16.0.

## Compatibility refresh (2026-11-17)

Windows 11, Python 3.14.2, openpyxl 3.1.5. The full suite passes with the native
backend: **10,087 passed, 343 skipped, 1 xfailed** in 4:37. The single xfail is
the known rich-text inline-string native gap. The openpyxl compatibility harness
passes **91**; the new `tests/test_openpyxl_model_parity.py` compares the added
behaviour with openpyxl itself. Ruff and mypy (142 source files) pass.

## Release candidate qualification (2026-11-05)

Windows, Python 3.14.2. The focused feature tests passed (**116 passed, 5
skipped**). Full-suite results:

| Backend | Passed | Skipped | Xfailed | Time |
| :--- | ---: | ---: | ---: | ---: |
| Rust | 9,983 | 324 | 2 | 4:24 |
| Pure Python | 8,890 | 1,418 | 1 | 2:42 |

The xfails are the known rich-text inline-string native gap and, without Pillow,
the openpyxl image-count comparison. The fallback run only has the latter.
Excel COM checks for the current formulas, tables, ignored errors, protection,
sort state and conditional-format changes passed (**88 passed, 4 skipped** in
2:20). The four skips are cases not authored by Excel or cases with no cells.

Ruff, mypy for Python 3.11 and 3.14 (140 source files), `stubtest`, Rustfmt,
Clippy, all **51 Rust tests**, the strict MkDocs build and the repository link
check (**24 passed**) also passed. MkDocs reported its existing links to
repository files outside the published docs site. The CPython 3.14 Windows
wheel passed its installed-package smoke test and **206 tests, 17 skipped** in
a clean environment. The sdist built and installed in another clean environment
and passed the same smoke test.

## Mixed loaded-workbook round trip (2026-11-05)

Windows, Python 3.11.9, Rust 1.99.1. A loaded XLSX table edit reused its
original part name but appended a second matching `<Override>` to
`[Content_Types].xml`; xlspy and openpyxl could read the package, but Excel
refused to open it. `_Package.override` now updates an existing declaration
instead of duplicating it. A regression test checks that a mixed workbook
remains idempotent after a second save and that Excel opens both XLSX and XLSB
round trips containing a renamed table/column, a structured-reference formula,
dynamic-array spill, LET formula, totals row, ignored errors and sheet
protection.

The feature matrix passed on both backends (**127 passed, 4 skipped** each).
The selected Excel COM suite passed **45 tests**. The full suite passed with
the native backend (**9981 passed, 322 skipped, 1 xfailed**) and pure Python
(**8888 passed, 1416 skipped**). Ruff, mypy (140 source files), `cargo fmt`,
Clippy, all **51 Rust tests**, and `mkdocs build --strict` passed. The strict
documentation build reports existing links to repository README/examples and
benchmark history that are not included in the docs site.

## XLSB conditional-format extension preservation (2026-11-05)

Windows, Python 3.11.9, Rust 1.99.1, Excel 16. Conditional-format rules loaded
from an XLSB can carry opaque Excel 2010 x14/FRT extension records for settings
such as negative data-bar colours and axis configuration. A priority-only rule
edit now keeps the original extension records and their GUID links. Changes to
extension-backed fields or the rule/group range are rejected with
`UnsupportedFeatureError` before writing, rather than silently dropping the
settings. Removing a rule removes its extension.

The rule model signature now tracks conditional-format group ranges, so changing
only `sqref` is recognized as an edit. Binary and round-trip tests check extension
payloads, GUID links, priority, group ranges and no-output-on-refusal. Excel COM
confirmed that a priority edit keeps the fixture's axis colour and position:
**32 passed**. The targeted CF/model tests passed with Rust (**94 passed**) and
pure Python (**94 passed**). Full suite: Rust **9985 passed, 323 skipped,
1 xfailed**; Python **8892 passed, 1417 skipped**. Ruff, mypy (140 source
files), stubtest, `cargo fmt`, Clippy, **51 Rust tests** and strict docs build
passed. The docs build continues to report pre-existing links to repository
README/examples and benchmark history outside the docs site.

## COM conformance and the XLSB write path (2026-11-05)

Windows 11, Python 3.14, native module built with `maturin develop --release`.
The full suite passes (the two expectations that changed with the new XLSB write
path were updated: `replace_sheet_data` now writes scalar formulas, and
`XlsbWriter` takes `ignored_errors`).

Modern formulas: a capture of Excel's own XLSX and XLSB files showed that reading
LET/LAMBDA/spill already worked and a no-op save is byte-identical, while writing
needed the `_xlpm.` parameter names, the XLSX `xl/metadata.xml` and the XLSB
`xl/metadata.bin` plus a `BrtCellMeta` record and an array formula group. Excel
16 recalculates LET and lambdas passed to `MAP`/`REDUCE`, and reports
`HasSpill`/`SpillingToRange` for `SEQUENCE` in both formats. A lambda called in
place (`LAMBDA(...)(...)`) is written to XLSX only.

Hashed sheet protection: an Excel-saved XLSB (`BrtSheetProtectionIso`) keeps its
hash when a permission flag changes, and Excel still unprotects with the old
password. Changing the password or its spin count is refused because the hash
cannot be recomputed without the password.

Sort state: writing an XLSX with `<autoFilter><sortState>` and saving it as XLSB
through COM showed the block Excel writes (`BrtBeginSortState` 0x212,
`BrtBeginSortCond` 0x214); the updater had looked for 0x1BB. It is now read,
saved and checked in Excel (`AutoFilter.Sort.SortFields(1).Order`).

The rewritten COM conformance suite (`tests/test_excel_com_conformance.py`, data
in `tests/com_cases.py`, Excel helpers in `tests/com_spec.py`) — eight cases in
XLSX and XLSB, written by xlspy and read by Excel, written by Excel and read by
xlspy, plus an updater round — gave **46 passed, 4 skipped** in 6.8 s. The touched
COM modules (`test_excel_com_conformance.py`,
`test_excel_com_ignored_errors.py`, `test_xlsb_updater_formulas.py`,
`test_excel_com_table_totals.py`) gave **65 passed, 4 skipped** in 76 s.

Checked in Excel 16.0: the XLSB ignored-error block (ranges and each flag read
back through `Range.Errors.Item(i).Ignore`), a loaded XLSB keeping and extending
them, and `replace_sheet_data` formulas recalculating on load (the cached values
written on purpose were wrong; Excel produced the right ones). The old conformance
module depended on PowerShell 7 and a `comparison/com` workspace that was never in
the repository; it skipped itself unconditionally.

## Edit and write rounds (2026-11-04)

Full suite at commit `37df71f`, native module / `XLSPY_ACCEL=python`:

| System | Python | Native | Pure Python |
| :--- | :--- | ---: | ---: |
| Windows 11 | 3.14 | 9813 passed, 328 skipped, 2 xfailed | 8730 passed, 1423 skipped, 1 xfailed |
| Ubuntu 26.04 on WSL2 | 3.15 | 9790 passed, 271 skipped, 2 xfailed | 8699 passed, 1363 skipped, 1 xfailed |
| macOS 27 (Apple M4) | 3.15 | 9790 passed, 271 skipped, 2 xfailed | 8699 passed, 1363 skipped, 1 xfailed |

The previous commit (`02fe4ad`) also passed on Python 3.11, 3.14 and 3.15 on WSL2 and on 3.11 and 3.14 on macOS, and `cargo test` passed on both. The two xfails are the openpyxl harness's image count without Pillow and the known gap of rich text with inline strings (`tests/test_native_write_coverage.py`). The Excel COM subset for the new table edits and defined-name comments passes with Excel 16.0 (`XLSPY_RUN_EXCEL_COM=1`). What changed in these rounds:

- Edit-sequence fuzzing (`tests/test_edit_sequences_fuzz.py`) compares a loaded workbook's saved edits with what was edited, in both formats; it found six cases where data was lost without an error (see the changelog).
- The XLSB write paths are counted (`tests/test_native_write_coverage.py`, `benchmarks/bench_xlsb_paths.py`): typical writes keep every row in the native encoder.
- Tables of a loaded XLSB can be changed and removed, and `DefinedName.comment` is written and read; both are checked against XLSX and in Excel.
- Not covered by a test: macOS and Linux have no Excel, so the COM checks run on Windows only.

## Quality round (2026-11-02)

Full suite on Windows 11 (Python 3.14, release build of the Rust module):
**2233 passed, 289 skipped, 1 xfailed**; with `XLSPY_ACCEL=python`, **1993
passed, 529 skipped, 1 xfailed**. The xfail is the openpyxl harness's image
count when Pillow is not installed (see [openpyxl parity](openpyxl-parity.md)).
`ruff`, `mypy`, `stubtest`, `cargo fmt`, `clippy -D warnings`, `cargo test` and
`mkdocs build --strict` pass. Refactors were checked by comparing the ZIP entries
of a fixed set of written files before and after (byte for byte), and the
read/write benchmarks stayed within 3% (see [benchmarks](benchmarks.md)).

## Review fixes (2026-11-01)

Full suite on Windows 11 with the project's `.venv` and the Rust module built:
**2138 passed, 192 skipped, 1 xfailed**. The new regression tests were in
`tests/test_p0_regressions.py` (since moved into the topic test files) and
`tests/test_safe_xml.py`. `ruff check src` and
`mypy` (the configuration in `pyproject.toml`) pass. At the time, mypy excluded 19
modules listed in `[tool.mypy.overrides]`; it now checks all of them.

The test count in the 2026-10-31 section below is from that date and is not
updated.

## Rust port of the native extension (2026-10-31)

The C extension (`xlspy._c_core`) was replaced by a Rust module
(`xlspy._rs_core`, PyO3 0.30), one group of functions at a time, with both
built side by side until the last step. Validated on Windows 11 with Python
3.11, 3.14 and 3.15:

- Differential tests, C against Rust, for every entry point (1,238 cases:
  fixed inputs, Hypothesis-generated rows, sheets and corrupted parts):
  identical results, exceptions and shared-string table state. They were
  removed with the C code; the comparison that stays is native against the
  pure-Python path (`tests/test_native_*.py`).
- Full suite with Rust as the default backend and, in those runs, with the C backend
  (`XLSPY_ACCEL=c`, since removed), on 3.11, 3.14 and 3.15: **3165 passed** in each run; the pure-Python fallback:
  **2987 passed**. After the C code was removed (and the 1,238 differential
  cases with it): **1927 passed** on 3.11 and 3.15, **1930 passed** on 3.14
  (with `pywin32`), **1752 passed** for the fallback. The one failure in every
  run, `test_golden_object_counts[images.xlsx]`, needs Pillow, which was not
  installed.
- Excel COM checks on the Rust build (`XLSPY_RUN_EXCEL_COM=1 pytest -m com`,
  Excel 16.0): **91 passed**, 9 skipped (printer, VBA project access,
  `nzpy_extended`).
- A wheel built by maturin from the source distribution for Python 3.11
  passed `tools/wheel_smoke.py` in a clean virtual environment.
- A mutation fuzzer (5,000 iterations per phase) through the Rust backend:
  seeds 1–5 and 7–10 held every check. Seed 1 first stopped at an XLSX cell
  with `r="A"` (no row number): the regular-expression path read it as
  column A, while ElementTree raised `ValueError`. The C extension stopped at
  the same case, so the port was not the cause; the read paths now share one
  rule for cell references (`refs.cell_position`), and the seed passes.
  Seed 6 generates a row 463,132 cells wide repeated over many gap rows: with
  rows held as full cell lists, it takes hours in either backend, so it is a
  slow input, not a failure. The unbounded tool has been removed. A bounded
  mutation test, `tests/test_native_fuzz.py` (300 seeded cases per format, each
  required to finish within 5 seconds), now runs in the ordinary suite; the
  encoder failure-path checks stay in `tests/test_native_encoders.py`.
- Speed, Rust time / C time on the same machine (below 1 is faster): XLSB
  row encoders 0.83–1.02, shared strings 0.88–0.99, column widths 1.00,
  formatted-cell pre-scan 0.79, XLSX row encoder 0.81–0.87, XLSB reader
  about 0.9, XLSX scanner 0.94–0.98, XLSB model loader 0.99–1.02.
- Behaviour differences: style and xf ids above 2**32 - 1 raise
  `OverflowError` on every platform (C raised on Windows and truncated on
  Linux and macOS); row and record sizes use 64-bit arithmetic on every
  platform (C's `long` is 32-bit on Windows); the model loader checks that
  the slot descriptors belong to the `Cell` class.

## Review fixes, read-only model and Strict reading (2026-10-21)

Validated on Windows with Python 3.14 and Excel 16.0:

- Full suite: **907 passed, 57 skipped** with the C extension and **774 passed,
  190 skipped** with `XLSPY_DISABLE_C_EXT=1`. Skips are optional database
  support and opt-in COM tests.
- Excel COM suites: **56 passed**, including 15 checks in
  `tests/test_excel_com_review_fixes.py`. Excel opened every file and confirmed:
  notes added to and removed from XLSX/XLSB workbooks with an Excel header
  picture, which stays intact; removing a sheet through `load_workbook()` keeps
  print areas and local/global names; `load_workbook()` edits of Excel-created
  files (unchanged save, values, fonts, added and renamed sheets); style edits of
  cells without `r` attributes; XLSB local names defined after use, legacy array
  formulas with date members and row options on array rows; `auto_filter_ref`;
  worksheet-local names and sheet order next to a chartsheet.
- Excel-created XLSB and XLSX files were compared with the streaming, formula
  and read-only readers. A workbook saved by Excel as Strict Open XML was read
  with values, dates, formula text, notes and defined names, and editing it was
  refused.
- Without the namespace-preserving serializer, Excel refused to open an
  Excel-created workbook after a font edit through `load_workbook()`.

Commands (`XLSPY_DISABLE_C_EXT` was the switch of the C extension; today
`XLSPY_ACCEL=python` selects the pure-Python code):

```powershell
py -3.14 -m pytest -q
$env:XLSPY_DISABLE_C_EXT = "1"; py -3.14 -m pytest -q; Remove-Item Env:XLSPY_DISABLE_C_EXT
$env:XLSPY_RUN_EXCEL_COM = "1"
py -3.14 -m pytest -q -p no:faulthandler tests/test_excel_com_review_fixes.py tests/test_excel_com_metadata.py tests/test_xlsb_metadata_com.py tests/test_excel_com_features.py tests/test_excel_com_updaters.py
```
## XLSB formula and metadata validation (2026-10-21)

Validated on Windows with Python 3.14.7 and Excel 16.0:

- Full suite: **884 passed, 42 skipped**. Skips are optional database support
  and opt-in COM tests.
- Excel COM suite: **25 passed**. Excel opened files generated by xlspy and
  files created independently by Excel. It confirmed comments, preserved form
  controls, sheet/workbook protection, defined names, recalculated scalar and
  array formulas, cached values, and sparse edits.
- Formula checks cover IF branches, IFERROR, CHOOSE, SUMIFS, COUNTIFS, XLOOKUP,
  sheet/name references, arrays, and stale caches that Excel must recalculate.
- `tests/test_xlsb_metadata.py`: **42 passed** with the C extension and **42
  passed** with `XLSPY_DISABLE_C_EXT=1`.
- Bulk-path timing used 50,000 × 20 numeric cells, one warmup, and nine
  measurements per implementation in one process, alternating baseline and
  working-tree runs. C writes were within 1%; C reads were 14–17% faster in
  this run. Python writes and reads were within 1% of baseline in both shared
  and inline string modes. These are one-machine timings, not performance
  guarantees.

Commands:

```powershell
py -3.14 -m pytest -q -p no:faulthandler
$env:XLSPY_RUN_EXCEL_COM = "1"
py -3.14 -m pytest tests/test_xlsb_metadata_com.py -q -p no:faulthandler
$env:XLSPY_DISABLE_C_EXT = "1"
py -3.14 -m pytest tests/test_xlsb_metadata.py -q
```

## Comments and template metadata validation (2026-10-20)

Validated on Windows with Python 3.14.7:

- Full suite: **842 passed, 17 skipped**. The skips are the optional
  `nzpy_extended` dependency and opt-in Excel COM checks.
- Statement/branch coverage: **88.34%**, above the configured 84% threshold.
- New Excel COM checks: **2 passed**. Excel opened the generated files with
  its normal loader and confirmed note text/authors, note dimensions, notes on
  empty cells, named ranges, sheet/workbook protection, and formula results.
  A form button created and saved by Excel survived subsequent note deletion,
  addition and sparse cell edits.
- Interoperability tests cover notes created by openpyxl and read by xlspy,
  and notes created/updated by xlspy and read by openpyxl. They also cover
  package cleanup, rollback, VML controls, opaque VBA parts, preserved chart
  and image payloads, threaded-comment preservation, prefixed XML, the 1900
  and 1904 date systems, and interleaving sparse/streamed updates.
- The incremental reader test consumes 30,000 rows with peak traced Python
  allocations below 4 MiB using inline strings. A separate test checks that the
  first row is yielded before the worksheet reaches EOF. Shared strings and
  styles are still loaded separately.

Commands used:

```powershell
py -3.14 -m pytest -q -p no:faulthandler --cov=xlspy
$env:XLSPY_RUN_EXCEL_COM = "1"
py -3.14 -m pytest tests/test_excel_com_metadata.py -q -p no:faulthandler
```

`-p no:faulthandler` suppresses the known host-specific COM shutdown diagnostic
described below; assertions and test exit codes remain enabled.

Tests marked `printer` read or set Excel's `PageSetup` (orientation, paper,
headers, page breaks). Excel does that through the default printer and, when
it cannot reach one, waits, shows a printer prompt and fails, so these tests
also need `XLSPY_RUN_EXCEL_PRINT=1`. Without it they are skipped; the other
COM tests check print areas through the sheet's `Print_Area` name instead. Coverage tools
were installed into an isolated, ignored `tmp/coverage-tools` directory and made
available through `PYTHONPATH`; the library gained no runtime dependencies.

Supported scope: classic XLSX/XLSB notes, XLSX/XLSM/XLSB template updates,
scalar formula edits, user-defined names and legacy sheet/workbook protection.
Threaded-comment editing, grouped-formula edits and reading note dimensions are
not implemented. See the [usage guide](guide/objects.md#comments-and-template-metadata-xlsxxlsmxlsb)
for the public APIs and limits.

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

### Known limitations at the time

Kept as recorded. The formula length and the updaters' text checks have
since been added; [Known limitations](#known-limitations) lists the current ones.

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
