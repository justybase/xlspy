# openpyxl parity matrix

This document tracks the practical compatibility target described in the
implementation plan.  Status levels:

| Level | Meaning |
| --- | --- |
| L0 | not supported |
| L1 | write/create only |
| L2 | read + preserve, no editing |
| L3 | read + edit + save |
| L4 | openpyxl-equivalent behaviour, including edge cases |

The harness in `tests/openpyxl_compat/` is the executable version of this
matrix.  A missing attribute is reported as `xfail` with the ID below; when a
gap is implemented the same test starts passing without changes.

## Implementation status — 2026-10-27

### Completed

**Phase 0 — compatibility harness and goldens**
- `tools/make_openpyxl_goldens.py` generates six openpyxl-authored workbooks
  (`metadata`, `objects`, `charts`, `images`, `chartsheet`, `long_tail`) plus
  `manifest.json`.
- `tests/openpyxl_compat/` compares values, styles, metadata, object counts and
  no-op ZIP preservation; `tools/compare_openpyxl_model.py` produces a semantic
  JSON diff.
- CI job `compat` pins `openpyxl==3.1.5`.
- Current results: the compatibility harness is **85 passed, 0 xfailed**; the
  full suite is **1452 passed, 132 skipped**. All planned harness checks pass.

**Phase 1 — public API, utilities and formulas**
- New `xlspy.utils.cell`, `xlspy.utils.datetime`, `xlspy.utils.escape`,
  `xlspy.utils.units`, `xlspy.utils.exceptions`.
- `styles.py` is now a package; `xlspy.styles.numbers` implements the openpyxl
  format helpers.
- `xlspy.formula.Tokenizer` and `Translator` match openpyxl on the tested
  grammar, including function names, arrays, scientific notation and shared
  references.
- Model additions: `WriteOnlyCell`, `ReadOnlyCell.number_format`,
  `Cell.is_date`/`encoding`/`base_date`, `append(dict)`, `copy_worksheet`
  hyperlinks, richer `CellRange`/`MultiCellRange`, `ExcelReader(errors=...)`,
  workbook metadata defaults, and a hybrid `DefinedNameList`/`DefinedNameDict`.
- 99 focused compatibility tests in `tests/test_compat_utils.py`.

**Phase 2 — worksheet/workbook metadata read + write**
- Read and write: `SheetFormatProperties`, sheet properties/`outlinePr`,
  full `SheetView`/`Pane`/`Selection`, full `PageSetup`, `PrintOptions`,
  odd/even/first `HeaderFooter`, page breaks, full `SheetProtection`,
  `WorkbookProtection`/security, `CalcProperties`, `BookView`, custom document
  properties and `lastPrinted`.
- AutoFilter filter columns and sort conditions, row/column `bestFit`/`style`
  and dimension attributes are read and written.
- No-op save preserves every untouched ZIP part for all goldens; metadata
  edits are verified through openpyxl and xlspy round trips.
- Fixes found during this phase: a failed `load_workbook()` now closes its
  temporary updater; theme colours accept the full `0..11` range (hyperlink
  styles use theme 10); freeze-pane changes update the model `Pane`/selection.

**Phase 3 — tables, validations, conditional formatting, named styles (XLSX/XLSM)**
- Tables are parsed from `tableParts` into `Table`/`TableColumn`/`TableStyleInfo`
  and exposed through a name-keyed `TableList` (`ws.tables["Name"]`).
- Data validations are parsed into `DataValidationList`; camelCase aliases
  (`allowBlank`, `showDropDown`, `errorStyle`, `imeMode`, ...) are available.
  Text is preserved even when its display flag is off, matching openpyxl.
- Conditional formatting is parsed into `ConditionalFormatting` groups; all
  rule families used by the goldens (cellIs, expression, colorScale, dataBar,
  iconSet, top10, aboveAverage, text, duplicates, time periods, blanks) are
  reconstructed with their differential styles.
- Named styles are parsed from `cellStyles`/`cellStyleXfs` into `NamedStyle`
  objects; `Workbook.named_styles` exposes them and `add_named_style()` writes a
  new style to both containers. `Cell.style = "Name"` applies its components.
- Editing/removing loaded tables, validations and conditional formats is
  supported. `sync_objects()` rebuilds the affected sheet elements and table
  parts when a loaded object changes or is removed; pure additions still use the
  existing `add_objects()` path so existing XML prefixes are preserved.
- The generic `Rule` class and `rule_type` aliases are exported.
- Added regression tests in `tests/openpyxl_compat/test_objects_edit.py` for
  table style edit/remove, validation edit/add/remove, CF edit/add/remove and
  named-style write/application.
- XLSB note: `XlsbWriter`/`XlsbUpdater` keep their existing table, validation
  and CF support, but loading an XLSB into the editable workbook model (and thus
  named-style editing in XLSB) belongs to Phases 5–6.

**Phase 4 — charts, images, chartsheets (read/preserve)**
- `LoadedChart` and `chart_reader.py` read DrawingML charts, including all
  fifteen chart types in the golden; `CHART_TYPES` now lists the practical
  openpyxl families (radar, bubble, stock, surface, 3D variants, projected pie).
- Images are read from drawing anchors with their data and alternative text.
- Chartsheets are read into `Workbook.chartsheets`; they remain preserved on
  save.
- The compatibility harness now has **0 xfails**.

**Phase 5 — XLSB model read (core)**
- `load_workbook(.xlsb)` no longer raises. `_xlsb_reader.py` builds a
  `Workbook`/`Worksheet`/`Cell` graph from `XlsbUpdater` and
  `_xlsb_read.iter_indexed_rows`.
- Values, formulas, dates, errors, merged ranges, hyperlinks, classic notes and
  defined names are loaded.
- No-op save is byte-identical; editing a cell and saving works through
  `XlsbUpdater.update_cells`.
- `read_only=True` keeps the existing streaming `ExcelReader` path.

**XLSB model parity — 2026-10-27**
- `styles.bin` is loaded into `XlsbStyleTable` (number formats, fonts, fills
  including gradients, borders, cell XFs, cell style XFs, named-style records
  and DXFs). Cells keep only their XF index; style objects are built on first
  access, so plain loads stay at the previous speed.
- Loaded into the model from XLSB, matching XLSX for the same workbook: row and
  column dimensions, sheet properties and tab colour, sheet views, panes and
  selections, sheet protection, page setup, margins, print options,
  header/footer, page breaks, auto filter, data validations, conditional
  formatting (all rule families except colour scales, data bars and icon
  sets), tables, charts, images and document/custom properties.
- Saved into a loaded XLSB: cell values and formulas, cell styles (new fonts,
  fills, borders, number formats and XFs are appended to `styles.bin`),
  merged cells, hyperlinks (internal and external, with their relationships),
  data validations, comments and defined names. Every other edit raises
  `NotImplementedError` at `save()` rather than being dropped.
- Excel COM tests (`tests/test_excel_com_xlsb_loaded.py`) open the saved files
  and check styles, merges, hyperlinks and validations.

### Remaining

**Phase 3 — remaining XLSB parity**
- Editing sheet layout, print setup, CF, tables, charts and images inside a
  loaded XLSB; named styles in XLSB; loading XLSB colour scales, data bars and
  icon sets.
- `IconSetRule` keeps its existing `type` field for the cfvo type and exposes
  the rule type as `rule_type`; openpyxl uses `.type == "iconSet"`.
- A conditional-formatting edit rebuilds the sheet's CF blocks and appends the
  corresponding `dxf` records; old orphaned `dxf` entries may accumulate across
  repeated edits (harmless, but a future cleanup can compact them).

**Phase 4 — remaining edit/removal**
- Loaded charts are preserved, not yet editable field-by-field; loaded images
  cannot yet be removed. Adding new charts/images already works.
- Chart options beyond the existing builder (combined charts, secondary axis,
  error bars, trendlines, data tables) and chartsheet creation remain.
- XLSB drawings are loaded; XLSB chartsheets are not loaded yet.

**Phase 6 — XLSB edit core**
- Cell values, styles, merges, hyperlinks, validations, comments and defined
  names are written through `XlsbUpdater` (`update_cells(..., styles=...)`,
  `set_merged_cells`, `set_hyperlinks`, `set_data_validations`,
  `set_comments`, `set_defined_name`/`remove_defined_name`); no-op saves remain
  byte-identical.
- Still missing: sheet add/remove/rename/reorder, protection/print/views, and
  table/CF/chart/image editing in XLSB.

**Phase 7 — long tail**
- External links, pivots, OLE/controls, templates, phonetics, custom
  `app.xml`/`custom.xml` writing, sheet backgrounds, ignored errors and
  custom/built-in table styles remain to be implemented.

## P0/P1 gaps

| ID | Area | xlspy 0.10.1 | target | Phase |
| --- | --- | --- | --- | --- |
| A1 | XLSB workbook model | `load_workbook(.xlsb)` raises for `read_only=False` → **done** (see E1–E6) | L3 | 5–6 |
| A2 | Loaded worksheet objects | tables, validations, CF, charts, images are invisible after load | L3 | 3–4 |
| A3 | Workbook metadata | no `security`, `calculation`, `views`, `epoch`, `template`, `vba_archive` | L3 | 2 |
| A4 | Defined names | read-only list; `DefinedName(name, formula, scope)` | L3 | 1 |
| A5 | New `.xlsm` / templates | `ValueError` | L2/L3 | 7 |
| A6 | External links | preservation only | L2 | 7 |
| A7 | Pivot tables | preservation/retarget only | L2 | 7 |
| A8 | Chartsheets | preserved but not loaded into the model | L2/L3 | 4 |
| A9 | OLE / controls / shapes | preservation only | L1/L2 | 7 |
| A10 | read-only / write-only | no `ReadOnlyCell.number_format`, no `WriteOnlyCell` | L3 | 1 |
| A11 | Custom document properties | writer-only; model does not read/write `custom.xml` | L3 | 2 |
| A12 | Document properties | no `lastPrinted`; model ignores `app.xml`/custom | L3 | 2 |
| B1 | Sheet views | 6 fields; no `Pane`/`Selection`/split panes | L3 | 1–2 |
| B2 | Sheet format | no `SheetFormatProperties` | L3 | 2 |
| B3 | Worksheet properties | no `outlinePr` and most flags | L3 | 2 |
| B4 | AutoFilter | only `ref`; filters/sort raise `NotImplementedError` | L3 | 2 |
| B5 | PageSetup | 5 fields | L3 | 2 |
| B6 | PrintOptions | centring only | L3 | 2 |
| B7 | HeaderFooter | odd header/footer only | L3 | 2 |
| B8 | Page breaks | writer only; no model API | L3 | 2 |
| B9 | Print area/titles | string only; no `print_titles` | L3 | 1–2 |
| B10 | Sheet protection | 3 flags + password | L3 | 2 |
| B11 | Workbook protection/security | 2 fields | L3 | 2 |
| B12 | Ignored errors/background/scenarios/smart tags | absent | L2 | 7 |
| B13 | Row/column dimensions | missing `min`, `max`, `index`, `autoSize` | L3 | 2 |
| B14 | `copy_worksheet` | does not copy hyperlinks | L3 for cells/metadata | 1 |
| C1 | Tables | write/add only; `ExcelTable` subset | L3 | 3 |
| C2 | Data validation | write/add only; no `imeMode`/camelCase aliases | L3 | 3 |
| C3 | Conditional formatting | write/add only; no read/generic `Rule`/dxf read | L3 | 3 |
| C4 | Named styles | `NotImplementedError` | L3 | 3 |
| C5 | Table styles | only `style_name` | L2/L3 | 7 |
| C6 | Charts | 7 types; no read/preserve/edit; no combined/3D/error bars | L2/L3 practical | 4 |
| C7 | Images | 4 formats; cell anchor + pixels only; no read/edit | L3 | 4 |
| C8 | Chartsheets | absent | L2/L3 | 4 |
| C9 | Drawings/shapes/connectors | preservation only | L1/L2 preserve | 7 |
| D1 | `styles.numbers` | absent | L3 | 1 |
| D2 | `utils.*` | internal helpers only | L3 | 1 |
| D3 | Formula API | internal normalizer/translator only | L3 | 1 |
| D4 | CellRange / MultiCellRange | minimal `str` subclass | L3 | 1 |
| D5 | NamedStyle / StyleProxy / StyleableObject | absent | L3 | 3 |
| D6 | Rich text phonetics | absent | L2 | 7 |
| D7 | Error representation | `ExcelReader` uses `error:#DIV/0!`; model uses `#DIV/0!` | L3 consistent | 1 |
| D8 | Date/time behaviour | mostly covered; public helpers missing | L3 | 1–2 |
| E1 | XLSB model read | absent → **done** (values, styles, metadata, objects; not CF colour scales/data bars/icon sets, chartsheets) | L2 | 5 |
| E2 | XLSB model edit | sparse `update_cells` only → values, styles, merges, hyperlinks, DV, comments, names; other edits raise | L3 | 6 |
| E3 | XLSB styles | mostly number formats → **done** (read, lazy; new styles appended on save) | L3 | 5–6 |
| E4 | XLSB objects read | absent → **done** (DV, CF, tables, charts, images) | L2 | 5 |
| E5 | XLSB objects edit | writer can create; updater cannot edit → DV edits saved; tables/CF/charts/images refused | L3 | 6 |
| E6 | XLSB sheet management | no add/remove/rename/reorder in updater | L3 | 6 |

## Long tail

| ID | Area | target | Phase |
| --- | --- | --- | --- |
| F1 | External links | L2 | 7 |
| F2 | Pivot tables | L2 | 7 |
| F3 | OLE / controls | L1/L2 | 7 |
| F4 | Templates / new `.xlsm` | L2/L3 | 7 |
| F5 | Themes / custom themes | L1/L2 | 7 |
| F6 | Encryption | out of scope (openpyxl has none) | — |
| F7 | Legacy `.xls` | out of scope (openpyxl has none) | — |

## Behavioural differences

| ID | Difference | Decision |
| --- | --- | --- |
| G1 | `ExcelReader` errors use `error:#DIV/0!` | opt-in `errors="openpyxl"`; keep legacy default |
| G2 | `Worksheet.append()` rejects `dict` | accept letter/index keys |
| G3 | OOXML `showDropDown` is inverted | openpyxl-compatible alias inverts it |
| G4 | `freeze_panes` is a string, not `Pane` | add `SheetView.pane`; accept both |
| G5 | `print_area` is a string | accept string or list |
| G6 | `Cell.style` rejects names | add `NamedStyle` in phase 3 |
| G7 | `Workbook.defined_names` is a list | hybrid list/dict keyed by name |
| G8 | `Worksheet.tables` is a list | hybrid list/dict keyed by name |
| G9 | no `Worksheet.sheet_format` | add in phase 2 |
| G10 | partial `Worksheet.page_setup` | add in phase 2 |

## Verification

```bash
python tools/make_openpyxl_goldens.py
python -m pytest tests/openpyxl_compat -q
python tools/compare_openpyxl_model.py LEFT RIGHT
python -m pytest tests/ -q
XLSPY_DISABLE_C_EXT=1 python -m pytest tests/ -q
```

Excel COM is an optional hybrid gate:

```powershell
$env:XLSPY_RUN_EXCEL_COM = "1"
python -m pytest tests/ -q -p no:faulthandler
```
