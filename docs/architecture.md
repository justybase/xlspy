# Architecture: the reader, updater and writer paths

xlspy grew in three stages. Each stage still has code in the package, and the
public API reaches more than one of them. This page maps which module does what,
so changes go to the right place. It describes the code as of 2026-11-01.

## The three generations

| Stage | Reads | Edits an existing file | Writes a new file |
|---|---|---|---|
| 1. Streaming (first) | `reader.py` (`ExcelReader`), `_xlsx_read.py`, `_xlsb_read.py`, `biff_reader.py` | `xlsx_updater.py` (`XlsxUpdater`, `XlsmUpdater`), `xlsb_updater.py` (`XlsbUpdater`), `_updater_base.py`, `updater_utils.py` | `xlsx_writer.py` (`XlsxWriter`), `writer.py` (`XlsbWriter`) |
| 2. Edit mixins | | `_xlsx_edit.py`, `_xlsb_edit.py`, `_xlsb_sheets.py` (mixed into the updaters) | |
| 3. Object model | `workbook/_reader.py` (`load_workbook`), `workbook/_fast_sheet.py`, `workbook/_xlsb_reader.py` | `workbook/_loaded.py`, `workbook/_xlsb_loaded.py`, `workbook/_xlsb_layout.py` | `workbook/_writer.py`, `workbook/_styles.py` |

Stage 3 is not independent of the earlier stages:

- `load_workbook` opens the package through `XlsxUpdater` (`workbook/_reader.py`,
  the XLSX branch) and through `XlsbUpdater` (`workbook/_xlsb_reader.py`). The
  updater holds the zip package and the shared-string table.
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
  thing. `tests/test_no_duplicate_top_level_names.py` keeps it that way. Three public
  names are defined by two models on purpose (`SheetProtection`, `PageMargins`,
  `Outline`); renaming them changes the API, so they stay.

Still not consolidated:

1. **Row-level sheet rewriters.** The locator is shared, but the code that edits the
   rows inside `<sheetData>` is still per caller: `_xlsx_edit.update_cells` (cell
   values), `workbook/_loaded._patch_xml_rows` (row attributes),
   `XlsxUpdater._patch_sheet_file` (streaming, bytes) and
   `workbook/_xlsb_layout._patch_xlsb_rows` (binary records). They edit different
   things, so merging them is a design step. The dimension and autoFilter rewriting
   is also still in two places (`XlsxUpdater._patch_sheet_xml` and `replace_element`
   in `_xlsx_edit` / `_loaded`).
2. **XLSB pair.** `workbook/_xlsb_loaded.py` and `_xlsb_edit.py` / `_xlsb_sheets.py`
   still overlap in responsibilities. Phase C only removed the duplicate names.
3. **Updater as holder.** The object model still reads `updater._sheets`,
   `updater._date_1904` and calls `updater._apply_parts` for writes outside the
   package. Moving those into a model object is the next ownership step.
## Where to make a change

- A new read feature: stage 3 (`workbook/_reader.py`) and, if it is a fast path,
  `_xlsx_read.py`. Keep the two results equal (`tests/test_fast_sheet_load.py`).
- A change to a loaded workbook's save: `workbook/_loaded.py`, then check
  `workbook/_writer.py`.
- A change to a new file: the writer methods are in the mixins listed above
  (`xlsx_writer.py` and `writer.py` hold the constructor and state), and the object
  model only when it is needed (`workbook/_writer.py`).
- Formula shifting: `refs.shift_references` only.
