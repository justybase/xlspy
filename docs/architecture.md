# Architecture: the reader, updater and writer paths

xlspy grew in three stages. Each stage still has code in the package, and the
public API reaches more than one of them. This page maps which module does what,
so changes go to the right place. It describes the code as of 2026-11-02.

## The three generations

| Stage | Reads | Edits an existing file | Writes a new file |
|---|---|---|---|
| 1. Streaming (first) | `reader.py` (`ExcelReader`), `_xlsx_read.py`, `_xlsb_read.py`, `biff_reader.py` | `xlsx_updater.py` (`XlsxUpdater`, `XlsmUpdater`), `xlsb_updater.py` (`XlsbUpdater`), `_updater_base.py`, `updater_utils.py` | `xlsx_writer.py` (`XlsxWriter`), `writer.py` (`XlsbWriter`) |
| 2. Edit mixins | | `_xlsx_edit.py`, `_xlsb_edit.py`, `_xlsb_sheets.py` (mixed into the updaters) | |
| 3. Object model | `workbook/_reader.py` (`load_workbook`), `workbook/_fast_sheet.py`, `workbook/_xlsb_reader.py` | `workbook/_loaded.py`, `workbook/_xlsb_loaded.py`, `workbook/_xlsb_layout.py` | `workbook/_writer.py`, `workbook/_styles.py` |

Stage 3 is not independent of the earlier stages:

- `load_workbook` opens the package through `XlsxUpdater` (`workbook/_reader.py`,
  the XLSX branch) and through `XlsbUpdater` (`workbook/_xlsb_reader.py`). The
  updater holds the zip package and the shared-string table. The model reaches
  them only through `_package_handle.PackageHandle` (see below).
- `Workbook.save` writes loaded sheets with `workbook/_loaded.save_loaded`. In
  the save path `workbook/_writer.py` still constructs `XlsxUpdater`.
- `dataframe.py` and the sheet writers use `XlsxWriter` and `XlsbWriter` directly.

## Module layout after the split

Four large modules are facades. Each one keeps its public names and re-exports
everything it moved, so `from xlspy.<facade> import X` still works and gives the
same object as the module that defines it. `tests/test_module_split_identity.py`
pins the names.

| Facade | Defined in |
|---|---|
| `styles/core.py` | `styles/_attrs.py` (colour, font, fill, border, alignment, protection), `_style.py` (`Style`, `DifferentialStyle`), `_cf_rules.py`, `_tables.py`, `_sheet_checks.py`, `_keys.py` (style keys), `_ranges.py`, `_cell_xf.py` |
| `workbook/_model.py` | `workbook/_coords.py`, `_properties.py`, `_tables.py`, `_dimensions.py`, `_cell.py`, `_worksheet.py` (`Worksheet`, `WriteOnlyWorksheet`, `Chartsheet`), `_workbook.py` (`Workbook`) |
| `xlsx_writer.py` (`XlsxWriter`) | the constructor and state methods stay here; the rest is in mixins: `_xlsx_writer_sheet.py`, `_xlsx_writer_styles.py`, `_xlsx_writer_parts.py`, `_xlsx_writer_sheet_xml.py`; module helpers in `_xlsx_writer_xml.py` |
| `writer.py` (`XlsbWriter`) | the constructor and state methods stay here; the rest is in mixins: `_xlsb_writer_sheet.py`, `_xlsb_writer_bin.py`, `_xlsb_writer_views.py`, `_xlsb_writer_strings.py`, `_xlsb_writer_styles.py`, `_xlsb_writer_parts.py`; constants in `_xlsb_writer_consts.py` |

The writer mixins read attributes that the writer class sets and call each
other's methods. Each writer has one host class that declares both with their
real types: `_xlsx_writer_host._XlsxWriterHost` and
`_xlsb_writer_host._XlsbWriterHost`. Every mixin of that writer derives from
it. The declarations sit under `if TYPE_CHECKING:`, so the host is empty at
runtime. mypy checks the writer's `__init__` against the attribute types, and
it reports an override error when a method's implementation no longer matches
its stub in the host. A new attribute or cross-mixin method needs a line in
the host.

## What has been consolidated

- **Formula reference shifting.** One function, `refs.shift_references(text, rows, cols)`,
  serves both shared-formula expansion on read (`_xlsx_read.translate_formula`)
  and moved cells (`workbook/_shift.translate_formula`). Before, there were two
  implementations. The read version shifted whole-column and whole-row references
  and the move version did not. Both now match `openpyxl.formula.translate.Translator`
  (`tests/test_formula_shift.py`).
- **XML tag helper.** `_metadata.local_name` replaces five identical private copies.
- **Name clash.** The sheet-level `<sheetProtection>` serializer in
  `workbook/_loaded.py` is now `_sheet_protection_xml`. The cell-style
  `<protection locked hidden/>` serializer is `_protection_xml` in `_xlsx_writer_xml.py`.
  They are different functions with the same name before this change.
- **Hardened XML entry point.** All package parts are parsed through
  `_safe_xml.fromstring` and `_safe_xml.iterparse`.

## Phase C: decisions and what is left

Phase C started with 69 top-level names defined in more than one module (22
with identical bodies, 47 with different ones). Decisions taken:

- **Package ownership.** `ZipPackage` owns the write-back of pending parts
  (`apply_parts`: one transaction, rolled back on failure). `XlsxEditMixin._apply_parts`
  and the object model's `_Package.commit` both go through it, so the object model
  writes without going through the updater. `XlsxUpdater` and `XlsbUpdater` keep
  their public methods and stay the handle the object model uses for the sheet map
  and the date system.
- **One sheet locator.** `_sheet_data.py` finds `<sheetData>` (with or without a
  prefix, empty or not) and replaces its body. `XlsxUpdater._patch_sheet_xml`,
  `_xlsx_edit` and `workbook/_loaded.py` use it.
- **Duplicate names.** Identical copies have one owner module. Equivalent helpers are
  one function. Different meanings under one name are renamed, so each name means one
  thing. `tests/test_no_duplicate_top_level_names.py` keeps it that way, with no
  exceptions left (the model classes became `WorksheetProtection`,
  `WorksheetPageMargins`, and the writer's outline index `OutlineIndex`).

## Layers and the updater boundary (Q4)

**Layers.** Writers, readers and updaters are the lower layer; the object model
(`xlspy.workbook`) is built on them. No lower-layer module imports
`xlspy.workbook`, and `import xlspy` or opening an updater leaves the model
unloaded. Imports sit at module level; a function-level import needs a reason
in a comment (optional pandas/polars/numpy, a real cycle or a rare heavy path).
`tests/test_import_layers.py` checks all three and caps the function-level
imports (21, down from 150).

**PackageHandle.** `_package_handle.PackageHandle(updater)` is the model's only
view of an updater's private state: the staged `package`, the `sheets` map, the
`date_1904` flag, `chartsheets`, `apply_parts`, `recalculate`, `sheet_xml` /
`sheet_bin`, `part_or_none`, `tag_prefix` and `rewrite_sheet_list`. An AST test
(`tests/test_package_handle.py`) fails when a module under `workbook/` reads
`updater._<name>`; the count went from 73 to 0. A new need goes into the handle
rather than into the model.

**XLSX worksheet edits.** `_sheet_data.py` holds the shared walks over a worksheet
part: `sheet_data_match` / `replace_sheet_data_body` (the locator),
`patch_rows` (each `<row>`, with numbers for rows without `r`), `insert_rows`,
`patch_cells` (each `<c>`, with its reference) and `set_dimension_ref`, which
keeps the existing `<dimension>` element and adds a missing one in schema order.
`_xlsx_edit.update_cells`, `workbook/_loaded._patch_xml_rows` and
`XlsxUpdater._patch_sheet_xml` use them. Two rewriters stay separate on purpose:
`XlsxUpdater._patch_sheet_file` works on the byte stream, without parsing the
part, and `workbook/_xlsb_layout._patch_xlsb_rows` edits binary records.

**XLSB: who does what.**

| Concern | Updater side | Model side |
|---|---|---|
| Sheet list in `workbook.bin` (add, remove, rename, reorder) | `_xlsb_sheets.py` | calls it through `PackageHandle.rewrite_sheet_list` |
| Cell values written into a sheet part | `xlsb_updater.py` | `workbook/_xlsb_loaded.py` (save) |
| Sheet metadata records: read | `_xlsb_metadata.py` (codecs) | `workbook/_xlsb_sheet_meta.py` (into the model) |
| Sheet metadata records: write | `_xlsb_metadata.py`, `_xlsb_print.py` (shared encoders: `pane_record`, `selection_record`, `sheet_protection_record`, `margins_record`, `breaks_block`, `header_footer_block`) | `workbook/_xlsb_layout.py` (re-encodes only the changed groups) |
| Template features (CF, validations, tables) | `_xlsb_edit.py` (opt-in, never in bulk loops) | read in `workbook/_xlsb_reader.py` |

`XlsbWriter` uses the same encoders, so a record has one encoder whichever path
writes it.

**Writers.** Both `write_sheet` methods check their options through
`_writer_common.WriterFrontEnd._check_sheet_options`, which returns a
`SheetSpec` before anything is recorded; `_claim_sheet_options` then takes the
table and slicer names and the conditional-format priorities. A rejected call
leaves the writer usable on both formats.

**Corrupt input.** `exceptions.reading(path)` turns the decoders' own errors
into `InvalidFileError` naming the file (`ExcelReader.open` and its row
readers, `load_workbook`, the updater constructors), and so does a Rust panic:
`_accel` registers the native module's `PanicException` with it. Those callers
check their arguments before the block, so a `ValueError` inside it comes from
the file. Low-level decoders raise `InvalidFileError` directly.
`tests/test_native_fuzz.py` fails when the cause of an error is a panic. A `KeyError` for an unknown sheet name is left as
it is. A record cut short is an error on every XLSB read path; `ExcelReader`
stops reading at the end of the cell data, so a cut after it is not seen there.

**Long functions.** Functions over about 150 lines were split into named steps
where that helps: the XLSB worksheet encoder, both `_write_sheet` methods, the
XLSX worksheet reader (`_read_core_sheet` now reads elements through
`_ELEMENT_READERS`), `load_workbook_xlsb`, and the Python row readers outside
their per-cell loops. Each split kept the output byte for byte and the benchmarks
within 3%. `formula/tokenizer._parse` stays one state machine.

## The native module (Rust)

`rust/xlspy_core` is a PyO3 crate built by maturin into `xlspy._rs_core`
(`[tool.maturin]` in `pyproject.toml`). It replaced the C extension and keeps its
behavior; the pure-Python code stays the reference and the fallback.

**Dispatch.** Only `_accel.py` imports the module. `XLSPY_ACCEL` picks the backend
(`auto` and `rust` load the module when it is built, `python` never does). At import
`_accel` hands the module the xlspy classes it must recognize or create (`CellError`,
`Style`, `Formula`, `Hyperlink`, `InvalidFileError`) and registers its
`PanicException`. Each wrapper returns None when the native function is missing or
cannot handle the input (an over-long string, a cell format not known yet), and the
caller then runs the Python code. `ACCEL_BACKEND` records which functions are native.

**Cell hook.** The row encoders write every type they know (text, numbers, booleans,
dates, `Decimal`, XLSX formulas and links) themselves. For any other cell they call
the writer's `cell_hook(cell, column, xf, row)` (`call_hook` in `cells.rs`) instead of
handing the whole row back: the hook returns the value the Python encoder would
write (a numpy scalar's `.item()`, None for pandas' `NaT` and `NA`, `str()` of
anything else), the encoded cell (an XLSB formula, compiled in Python), a
`(text, xf)` pair (a link in a styled cell), or `NotImplemented`, and then the row
goes to the Python encoder after all. Only rich text, `(Hyperlink, Style)` cells,
array formulas and the first row that needs a new style variant (a format overlay, a
styled date or link) still do.

| Rust module | Does | Python counterpart |
|---|---|---|
| `xlsb_write.rs`, `cells.rs`, `sst.rs` | XLSB row records, shared strings | `_xlsb_writer_bin.py` |
| `xlsx_write.rs`, `xml_text.rs` | XLSX `<row>` XML | `_xlsx_writer_sheet.py` |
| `widths.rs` | column widths, the formatted-cell pre-scan | `_widths.py` |
| `xlsb_read.rs` | `ExcelReader` XLSB rows | `biff_reader.py`, `reader.py` |
| `xlsx_read.rs`, `xlsx_sst.rs` | `ExcelReader` XLSX rows and shared strings; `load_workbook` fast rows | `_xlsx_read.py`, `workbook/_fast_sheet.py` |
| `model.rs` | `load_workbook` XLSB cells, created natively | `workbook/_xlsb_reader.py` |
| `bytes.rs`, `numfmt.rs`, `dates.rs`, `text.rs`, `pyseq.rs`, `types.rs` | shared helpers | |

**Parity.** The `test_native_*` tests run the same inputs through both backends and
require equal results, and CI runs the whole suite a second time with
`XLSPY_ACCEL=python`. Combined coverage counts both runs.

**Errors.** Bad data in a file raises `InvalidFileError`. A Rust panic unwinds
(`panic = "unwind"` in the release profile) into a `PanicException`, which the read
entry points report as `InvalidFileError`; the fuzz tests fail if one ever escapes.

**Unsafe code.** The module uses CPython's C API where PyO3 has no equivalent that
is as fast. Every `unsafe` block carries a `// SAFETY:` comment, which the crate's
lints enforce (`undocumented_unsafe_blocks`, `unsafe_op_in_unsafe_fn`).

**The stub.** `src/xlspy/_rs_core.pyi` types the module. CI checks it against the
built module with `python -m mypy.stubtest xlspy._rs_core`, and
`tests/typing/stub_usage.py` checks the argument types it accepts and rejects.

**Why not abi3.** One stable-ABI wheel per platform would replace one wheel per
Python version, but the module relies on three things outside the limited API:
`model.rs` creates cells with `tp_alloc` and writes their `__slots__` through member
descriptor offsets; `text.rs` reads PEP 393 string data in place, so that strings
with lone surrogates need no decoding; and the row walk uses the
`PySequence_Fast` macros. Replacing them would cost speed on the hot paths, so the
module stays version-specific until that cost is measured and accepted.

## Where to make a change

- A new read feature: stage 3 (`workbook/_reader.py`) and, if it is a fast path,
  `_xlsx_read.py`. Keep the two results equal (`tests/test_fast_sheet_load.py`).
- A change to a loaded workbook's save: `workbook/_loaded.py`, then check
  `workbook/_writer.py`.
- A change to a new file: the writer methods are in the mixins listed above
  (`xlsx_writer.py` and `writer.py` hold the constructor and state), and the object
  model only when it is needed (`workbook/_writer.py`).
- Formula shifting: `refs.shift_references` only.
