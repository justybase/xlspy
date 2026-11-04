# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- A formula in a loaded XLSB can refer to a table or column renamed in the same save, by its new name,
  and to a table added to a later sheet from an earlier one. The old name is refused with
  `UnsupportedFeatureError` or `ValueError`.
- A table removed in a save frees its name for a table added in the same save. A formula that still
  names the removed table is refused.
- `ExcelTable.totals_row_functions` selects each column's totals function (`sum`, `average`, `min`,
  ...), or `label`; without it the cell decides: text is a label, a number a sum, a formula a custom
  total.

### Fixed

- Row keyword combinations that Excel refuses are refused when a structured reference is written:
  `#All` or `#This Row` with another keyword, and `#Headers` with `#Totals` without `#Data`. Repeated
  keywords are dropped, as Excel drops them.
- An XLSB table with a totals row opens in Excel: the table part sets `fShownTotalRow`, writes every
  column's totals function or label, keeps the totals row out of its filter, and omits the filter for
  a headerless table. Moving the range of a loaded totals table recomputes that metadata.
- An XLSX table with a totals row keeps it when Excel opens the file: `totalsRowCount` and the
  per-column totals attributes are written instead of `totalsRowShown` alone.

## [0.13.0] - 2026-11-04

### Breaking changes

Migration from 0.12:

- `ExcelReader.get_rows()` on XLSB returns rows of the XLSX shape (see Changed). Code that relied on the
  old shape, rows padded from the first column or a wider row widening the rows after it, should read
  each value by its position in the row.
- `Worksheet.max_row` of a loaded XLSX no longer counts `<row>` elements without cells or the stored
  `<dimension>`. Code that sized a loop by `max_row` may see fewer rows. `append()` now writes to the
  same row in XLSX and XLSB.
- `XlsbUpdater.set_hyperlinks` sets the cells it is given and keeps the others; pass `None` for each
  link to clear a sheet (see Changed).
- `DefinedName(localSheetId=n)` and `scope=` on a loaded XLSB name a sheet, where they made a global
  name before; `hidden=True` is kept.
- `load_workbook` and `ExcelReader` raise `InvalidFileError` for a row record past the last row or a
  cell past column XFD, in both backends; they dropped or padded such data before. A damaged value
  cell is skipped. A valid file reads as before.

### Added

- Tables of a loaded XLSB can be changed and removed, as in XLSX: the range, the name, the style and
  its banding, the filter range and the column names (which must still equal the header cells), and
  `del ws.tables["Name"]` / `ws.tables.remove(...)`. The part of a changed table keeps every record
  the model cannot express (totals, column formulas, formats, filter criteria); the parts of other
  tables keep their bytes. A table with another number of columns, and a rename or removal in a
  workbook with pivot tables or slicers, raise `UnsupportedFeatureError` before anything is written.
  Until now any such edit was refused.

- Formulas that refer to tables are read from an XLSB: `SUM(Table1[Qty])`,
  `Table1[[#Totals],[Qty]]`, `Table1[[#Headers],[#Data],[Qty]]`, column ranges and `[@Col]` (as
  `Table1[[#This Row],[Col]]`, the text an XLSX file holds), with the escapes Excel uses in a column
  name. This includes the totals row and the calculated columns of every table that Excel saves.
  Removing a table that a formula refers to is refused; renaming one is safe, because a binary
  workbook names the table by its id.

- Formulas that refer to a table are written to an XLSB: `SUM(Table1[Qty])`, `Table1[[#Totals],[Qty]]`,
  `Table1[[#Data],[Qty]]`, column ranges and `[@Col]`, as the PtgList tokens Excel writes. Excel
  calculates them, and a formula that is loaded back reads as it was written. In a new workbook a
  formula can refer to a table of its own sheet or of an earlier sheet; a table of a later sheet, or
  one whose header row is not written yet, raises `UnsupportedFeatureError`. In a loaded workbook a
  formula can also refer to a table added in the same save. A defined name cannot refer to a table,
  as in Excel.

- `XlsxUpdater` and `XlsbUpdater` work as context managers: leaving the `with` block closes the
  updater, and changes are kept only by a `save()` made inside it.

- `tools/rust_coverage.py`: line and region coverage of the native module (`cargo llvm-cov`),
  from the Rust unit tests and the pytest suite run against an instrumented build.

### Changed

- `ExcelReader.get_rows()` on XLSB returns rows of the same shape as on XLSX, by every path (bulk,
  `streaming=True`, `data_only=False`): each value at its own position, and rows padded to the
  width of the first row with cells. Rows before that one are now `[]` (they were padded), a wider
  row no longer widens the rows after it, and a sheet of empty rows gives `[[], []]`, not
  `[[None], [None]]`.

- `write_dataframe` and `dataframe_to_rows` decide the plain cells of a row (text, numbers, booleans,
  missing values, `pd.Timestamp`) in Rust and leave any other cell to the Python normaliser, so the
  rows are the same. Converting 100,000 rows of five columns: 0.36 s -> 0.25 s; the rest is pandas
  itself (`itertuples`, iterating Arrow-backed strings).
- A row with rich text (`CellRichText`) is written by the native XLSB encoder, not cell by cell in
  Python: the cell is an index into the shared-string table, as before. 100,000 rows with a rich cell:
  about 1.0 s -> 0.62 s, the bytes written are the same. With `use_shared_strings=False` the row
  still goes through Python, because rich text is always a shared string.
- The XLSB writer encodes numpy integers, booleans and float16/32 scalars in Rust instead of calling
  into Python for each cell (the value they hold, as before). 100,000 rows of four numpy columns:
  0.88 s -> 0.49 s. The bytes written are the same.
- `XlsxUpdater.update_cells`, which the save of a loaded XLSX uses, finds what it needs in a sheet
  (the bounds of its cells, the cells being edited, the formulas that forbid an edit) with a
  native scan instead of a parsed element tree and a loop over every cell. On a sheet of 1,000,000
  cells, 1,000 edits take 0.40 s instead of 1.23 s, and the save of a loaded workbook 0.60 s
  instead of 1.69 s; the bytes written are the same. Only the plain markup that Excel and the
  writers produce is scanned that way (unprefixed names, double quotes, references in their normal
  form); any other part goes through ElementTree as before. A part that is not well formed, has a
  document type declaration or is not a worksheet is refused as before, before the edits are read.

- The workbook model is faster and smaller, with the same bytes in the files (50,000 x 20, measured
  in turns against the code before): saving a loaded XLSB after edits 2.87 s -> 0.33 s, because one
  native pass over the sheet's records replaces a Python object per record and the save no longer
  visits every cell; building and saving a model 1.51 s -> 1.03 s (XLSX) and 1.40 s -> 0.92 s (XLSB);
  `load_workbook` 7-19% faster. `Worksheet.append` and the rows handed to the writers run natively
  for plain values, and leave any other row or value to the Python code they stand in for. Peak
  memory while building a model falls from 323 MB to 212 MB: a new workbook no longer records which
  cells changed, which only the save of a loaded workbook reads, and `Cell` lost a flag nothing read.
- **Breaking:** `XlsbUpdater.set_hyperlinks` works as in `XlsxUpdater`: it sets the given cells,
  `None` removes a cell's link, and every other link is kept. It used to replace all of the
  sheet's links and refuse `None`. To clear a sheet, pass `None` for each link. Saving a loaded
  XLSB now rewrites only the links that changed, so one over a range survives an edit of another.
- `workbook.bin`'s sheet list and date system have one parser (`biff_utils`), shared by
  `ExcelReader` and `XlsbUpdater`. A sheet record whose strings overrun it is now refused by
  `ExcelReader` too, instead of giving a cut-off name.

### Fixed

- An 8-byte BrtFmlaError (a formula cell with an error, shorter than its 9-byte header) is not a cell,
  as an 8-byte BrtCellError already was: the native reader stored it as a blank value and the Python
  reader read past the end of the record.
- `load_workbook()` and `ExcelReader.get_rows(data_only=False)` failed on a binary workbook from Excel
  with a table that has a totals row or a calculated column (`unsupported XLSB formula token 0x18`),
  and so on any formula that uses a table by name. `data_only=True` was not affected.
- `DefinedName.comment` was dropped by every XLSX and XLSB writer, updater and reader: the attribute
  existed but was neither written nor read. It is now saved as the `comment` attribute of
  `<definedName>` and as the comment string of `BrtName` (Excel shows it in the Name Manager and
  reads what xlspy writes, in both formats), and a changed comment or `hidden` flag of a name of a
  loaded XLSB is saved by the model. Names without a comment are written as before, byte for byte.
- `ExcelReader.get_rows()` on XLSB lost data without an error: cells left of the first cell of the
  first row were dropped and the others moved. With a title in B1 over a table from A3, the first
  column of the table was gone, and a sheet starting at E6 read as if it started at A1.
  `load_workbook` and all XLSX readers were right.
- `XlsbUpdater.replace_sheet_data` on a sheet without rows put them after the end of its sheet data,
  where nothing reads them.
- `XlsxUpdater.replace_sheet_data` rounded numbers to 15 digits: 0.30000000000000004 became 0.3 and
  1234567890123456 became 1234567890123460. It now writes the digits that read back as the same
  number, like the writers and the XLSB updater. Dates with a time of day are exact too.
- Assigning a value to a cell removed its link in memory only: `ws["A1"].hyperlink = link` followed
  by `ws["A1"].value = 1` lost the link, while the same assignment on a loaded link kept it in the
  file and reported `None`. A value now keeps the link, as openpyxl does; `hyperlink = None` removes
  it.
- `Worksheet.max_row` of a loaded XLSX counted `<row>` elements without cells and the `<dimension>`
  of the sheet, which XLSB did not, so `append()` wrote to a different row in the two formats. Only
  cells and merged ranges count now, and a saved sheet's `<dimension>` is the range of its cells.
- XLSX writers dropped the spaces at the end of a text, and a tab or a line break inside one: the
  `<t>` had no `xml:space="preserve"` (only a leading space was kept), and Excel trims such a text.
  `"trail "` read back as `"trail"` in Excel. Both backends and both string tables now mark every
  text that needs it, as the updaters always did.
- Saving a loaded XLSX after cell edits took time in the square of the sheet's size: 86 s after
  1,000 edits of a 4,000 x 20 sheet, and hours at 50,000 rows. The same save takes 0.33 s now, with
  the same bytes in the file. The pattern for an XML element (`element_pattern`) ran to the end of
  the part for every element of a sheet without a namespace prefix (a backreference to a prefix that
  was not there), and `update_cells` went through every pending edit for every row. Notes, links
  and defined names use that pattern too. `update_cells` and the style edit also parsed the sheet up
  to three times, and walked all its cells to find the edited ones. The rows of `<sheetData>` are
  now walked a tag at a time.
- Other edits also took time in the square of their number, with the same bytes in the file now:
  4,000 notes on an XLSX in one `set_comments` took 9 s and take 0.2 s (each note copied the whole
  part); 4,000 `merge_cells` took 13.5 s and take 0.25 s (each compared itself with every range,
  and `cell()` went through them too); `set_hyperlinks` on an XLSX, from 0.9 s to 0.1 s (each link
  looked for a free id from 1). Checks of edits and notes against merged ranges now use an index.
  Still quadratic: adding defined names to a loaded workbook, which rewrites the workbook part at
  each `add()` so that a bad name is refused there (500 names take 1-2 s, 2,000 take 18-32 s).
- `XlsbUpdater.update_cells` on a damaged sheet (no sheet data, or a row, cell, dimension or merge
  record too short for what it holds) raises `InvalidFileError` instead of `StopIteration` or
  `struct.error`.
- A centred header under an autofilter or table drop-down button (`Alignment(horizontal="center")`)
  got room for the button on one side only, so its text ran into the button. Excel's *AutoFit*
  keeps centred text clear of the button on both sides (32 pixels, not 16: measured at 100%
  scaling, on one line and on two); the estimate now does the same, for a centred header style
  and for a `(value, Style)` header cell. Right-aligned and distributed headers keep the
  16 pixels (Excel adds none and lets the button cover the end of the text).
- Damaged XLSB worksheets, found by a new differential fuzzer that reads each corrupted case through
  the native and the pure-Python reader (`tests/test_native_fuzz_differential.py`):
  - A cell past column XFD made the Python reader (`XLSPY_ACCEL=python`, or no native module) pad
    its row up to that column: minutes and gigabytes for one cell in column 2,147,483,647. The
    native reader dropped the cell. `ExcelReader` now refuses it in both, with the
    `InvalidFileError` `load_workbook` already raised ("XLSB column exceeds Excel limits").
  - A row record shorter than 4 bytes made the native reader take its row index from the bytes
    of the next record, so a readable sheet could be refused as having a row past the last.
  - The Python reader read cells after `BrtEndSheetData` (the native one stops there), and lost
    or misplaced cells whose records came out of column order. Its row decoder is now the
    twin of the native one, so both read a damaged sheet the same way.
  - A record the cell reader ignores (an overrunning string, say) before the first row made the
    native reader start a phantom empty row.
  - The Python shared-strings reader sized its list by the count in `BrtBeginSst`: a damaged
    count allocated gigabytes. The count is not trusted any more, and a cut-short `BrtBeginSst`
    is refused as in the native reader.
  - `load_workbook` on a damaged XLSB: the Python loader refused a cell record too short for its
    value, and read any byte but 0 of a boolean as `True`; the native loader dropped the cells of
    a row record past the last row. Both backends now skip a short value cell, read only 1 as
    `True`, and refuse a row past the last. The fuzzer compares `load_workbook` too.

- `XlsxUpdater` and `XlsbUpdater` meant different things by the same call:
  - `DefinedName(..., localSheetId=n)` became a global name in XLSB and could add a second entry
    in XLSX; `hidden=True` was lost in XLSB. Either `scope` or `localSheetId` now names the sheet
    in both formats (`KeyError` for a sheet that is not there, `ValueError` when the two disagree),
    and both report both.
  - Changing or removing a name left stale formula results in XLSX, where it did not ask for a
    recalculation on load; XLSB did. Removing a name that does not exist rewrote `workbook.bin`.
  - `XlsbUpdater.set_comments` refused to remove a classic note on a cell with a threaded comment;
    as in XLSX, only adding one there is refused.
  - `XlsxUpdater.set_hyperlinks` with an empty mapping accepted a sheet that does not exist.
- XLSB notes got a random GUID, so two saves of the same workbook differed. The GUID now comes from
  the comments part and the cell (checked in Excel).

## [0.12.0] - 2026-11-03

### Fixed

- Automatic column widths ignored a cell's alignment indent (`Alignment(indent=n)`, Excel's
  *Format Cells > Indent*, typical for the labels of nested row groups) and measured tabs as
  nothing, so such labels ran past the column. Excel's *AutoFit* adds 9 pixels per indent level at
  any font size (twice for centred and distributed text) and draws a tab half an em wide; the
  estimate now does the same, in the Python and the native meter (exact on 176 cells measured at
  100% scaling, 8 to 20 pt).
- Automatic column widths measured every font as Calibri and ignored italic, so columns in a wider
  font (Arial, Verdana, Consolas: up to 20% wider than Calibri) or in italic cut their text off.
  The widths now use the glyph widths of 21 popular fonts in regular, bold, italic and bold italic
  (Calibri, Calibri Light, Aptos, Aptos Narrow, Arial, Arial Narrow, Arial Black, Times New Roman,
  Cambria, Consolas, Courier New, Lucida Console, Segoe UI, Tahoma, Verdana, Georgia, Trebuchet MS,
  Century Gothic, Lucida Sans Unicode, Microsoft Sans Serif, Comic Sans MS), as GDI draws them at
  96 DPI;
  Excel's margin follows the font's digit width and italic text gets room for its slant. On 3,696
  cells measured in Excel (8 to 40 pixels) and 560 more on other texts, every one is Excel's
  *AutoFit* plus the spare pixel (two for some italic sizes). Any other font is measured as
  Calibri 1.4 times as wide, which covers the text fonts of Windows and Office (a fixed-pitch name
  such as `Cascadia Mono` as Courier New).
  `src/xlspy/_font_metrics.py` (generated by `tools/gdi_column_widths.py`) replaces
  `_calibri_metrics.py`; `text_pixels`, `text_width` and `cell_width` take `font=` and `italic=`.
- Rich text (`CellRichText`) was measured as plain text in the cell's font. Each run is now
  measured in its own font (a `TextBlock`'s unset fields are Calibri 11, as Excel reads them), and
  italic runs get their share of the slant room, as Excel's *AutoFit* does.
- Characters the cell's font lacks were measured as a digit, so emoji (Excel draws them about 1.4
  em wide), symbols and arrows (up to 1.3 em) and the Unicode spaces (an em space is a whole em)
  came out too narrow. The widths now take the fallback fonts Excel draws them with (Segoe UI
  Emoji, and the widest of Segoe UI Symbol, Cambria Math, Segoe UI, Tahoma and Arial), and the
  fonts' own glyphs for General Punctuation, arrows, mathematical operators, box drawing and other
  symbol blocks; a soft hyphen is a hyphen.
- Header cells under an autofilter or table drop-down button were one pixel narrower than Excel's
  *AutoFit*: Excel adds 16 pixels for the button at any font size (measured at 100% scaling on 429
  headers, 11 and 16 pt), xlspy added 14. `FILTER_BUTTON` is now 16/7 characters, so a filtered
  header gets Excel's width plus the usual spare pixel (`COLUMN_NAME`: 17.57, Excel 17.43).
- The native row iterators (`XlsbReader`, `XlsxReader` in `rust/xlspy_core`) were invisible to the
  garbage collector: a reference cycle through one of them (its shared-string list, or in
  `load_workbook` the sheet and cells it builds, referring back to the iterator) was never freed.
  They now report what they hold (`__traverse__` / `__clear__`).
- Registering xlspy's classes with the native module again (a module reload) took one more
  reference to each class every time; the same class is now kept once.
- An XLSB cell with a date format whose serial is no date (NaN, infinite, past 9999-12-31) stopped
  the read: `ExcelReader` raised a bare `OverflowError` (or refused the whole sheet as
  `InvalidFileError` for NaN), in both the native and the Python reader. Such a cell now reads as
  its number, as in the XLSX readers and the workbook model (Excel shows it as `####`).
- A damaged workbook part could raise `KeyError` instead of `InvalidFileError`: a `<sheet>` or
  `<definedName>` without its `name` or `r:id` in `workbook.xml`, and in `load_workbook` an XLSB
  `workbook.bin` whose sheet list the updater and the reader read differently.
- A failed open gave its resources back only when the exception was collected:
  `load_workbook(read_only=True)` kept the file open after a damaged sheet (on Windows, locked),
  and `XlsxUpdater`/`XlsbUpdater` given the other format, or `load_workbook` on an XLSB whose
  reader failed, kept their staging directory. They are released before the exception propagates.
- `XlsbWriter.write_sheet` raised `TypeError: row must be a list` for rows given as tuples or as a
  numpy array (any row the native encoder saw); `XlsxWriter` took them. Every sequence is a row now.
- An int past the float range (`10**400`) was text in an `XlsbWriter` row the native encoder wrote
  and an `OverflowError` everywhere else; both writers now write its digits as text.
- `XlsbUpdater.replace_sheet_data` wrote `True` and `False` as the numbers 1 and 0 (a conversion of
  every `numbers.Real` to float caught `bool` first); they are booleans now, as in `XlsxUpdater`.
- pandas' `NaT` in a raw row (`write_sheet`, not `write_dataframe`) was a NaN number in XLSB, which
  Excel does not open as a date, and the text `nan` in XLSX; `NA` was the text `<NA>`. Both are
  empty cells now, as in the DataFrame path, and the column widths measure them as empty.

### Changed

- The automatic column widths are measured by the native extension (`WidthMeter` in
  `rust/xlspy_core/src/autofit.rs`), 10 to 40 times faster than in Python: every cell but numpy
  scalars, str subclasses, ints past 64 bits and a few rare number formats (`*@`, a literal with a
  line break), which still go through the Python rules. Number formats are parsed in Python once
  per style and applied in Rust; `(value, format)` and `(value, Style)` overlays, formulas,
  hyperlinks, NaN and infinities are measured in Rust as well. A cell that cannot widen its column
  is skipped (exact bounds: a number in General in a column already as wide as any number can
  need, text shorter than the column can hold of its widest character). The widths are the same
  as the Python path's (checked on 1.8 million numbers, 1.2 million numbers in 460 number formats
  and mixed styled rows). `autofit_rows=None` (every row) now adds about 30% to an XLSB file
  (50,000 x 50: 0.35 s and 0.45 s, where it added 4.5 s) and 10-20% to an XLSX file. The writers
  look the row styles up once per row, where the meter looked them up a second time. The native
  meter is freed with its writer (the garbage collector sees its callbacks). The
  pure-Python path (`XLSPY_ACCEL=python`) is unchanged.
- `XlsbWriter` writes rows with dates or datetimes in a styled column (`styles={"D:D": ...}`) with
  the native encoder, as `XlsxWriter` does: it sent every such row to the Python encoder, because
  the style's date variant was handed to the native encoder only for rows with `(value, format)`
  cells. 200,000 x 10 with column styles: 0.64 s, was 6.7 s (XLSX: 0.8 s). The bytes are the same.
- More `XlsbWriter` rows stay on the native encoder: after an array formula only the rows of its
  range go to the Python encoder (every later row of the sheet did), and a column whose cells are
  all `(value, format)` overlays on a column style registers that style, so the rows after the
  first go native (none did). Empty rows past an array formula's range get no row record.
- The native row encoders no longer hand a whole row to the Python encoder for one cell of a type
  they do not know: they ask the writer for that cell alone (`cell_hook`). numpy scalars, times,
  `timedelta`, bytes and other objects (written as `str()`, as before), pandas' `NaT`/`NA`,
  `(None, format)` blanks, XLSB formulas and links (also in styled cells) are written natively;
  100,000 x 6 with such a column: numpy ints 0.3 s (were 0.95 s), times 0.3 s (0.9 s), XLSB
  formulas 1.3 s (1.9 s; compiling the formula is the rest), XLSB links 0.5 s (1.1 s). The bytes
  are those of the Python encoder, which writes a styled blank without its `r` now, as the native
  encoder always did. Rich text, `(Hyperlink, Style)` cells and array formulas still go to the
  Python encoder.
- `XlsxUpdater` and `XlsbUpdater.replace_sheet_data` (and `replace_sheet_data_stream`) write their
  rows natively (`update_rows.rs`): text, numbers, booleans and dates there, any other cell through
  the updater's Python code, one cell at a time, so the bytes and the errors are the same.
  100,000 x 6: XLSX 0.73 s (was 1.61 s), XLSB 0.40 s (1.35 s); streamed, XLSX 1.5 s (3.0 s) and
  XLSB 1.4 s (3.8 s). The streaming shared-string index keeps the last 65,536 strings' indexes in
  memory and no longer flushes its file for every new string, and the streamed XLSB sheet's row
  headers are patched natively, without parsing the rows twice in Python.

## [0.11.1] - 2026-11-02

### Fixed

- Automatic column widths were up to 7% narrow for long text: a column Excel's *AutoFit* made 66.29
  wide came out 61.86. The character widths had been measured with Excel on a display at 175%
  scaling, where Excel measures text at that size and hinting rounds letters differently (`i` is
  0.46 of a digit there, 0.57 at 100%); column widths are defined at 96 DPI. The estimate now
  rebuilds the pixel widths GDI draws for Calibri at 96 DPI, for every font size and for bold
  (`src/xlspy/_calibri_metrics.py`, generated by `tools/gdi_column_widths.py`), and applies the
  margin Excel's AutoFit adds: `max(px + m, floor(px * 29/28) + m - 2)` pixels, with `m` = 2 at
  11 pt and 2 more for every 8 pixels of font size. Over 11,786 cells measured with Excel at 100%
  scaling (8 to 36 pt, regular and bold, text and numbers) it matches to the pixel; xlspy adds one
  spare pixel. At other scalings the widths are usually a little wider than Excel's own.
- General numbers are sized as Excel sizes them: 11 characters besides the sign
  (`-6619407.002`, was `-6619407`), rounded half up as decimals (`-9526850.491`), scientific when it
  keeps more digits (`4.99518E-06`), and room for a decimal point as wide as a digit. A number with
  11 whole digits and a fraction lost its trailing zeros (`49062358930` was measured as
  `4906235893`), and a negative number rounded to zero showed `-0`.
- Default dates (`numFmt` 14 and 22) are sized for the widest common locale (`12/31/2024`): at
  `30.09.2024` a column could show `####` for a reader with `2024-09-30`.
- The project page on PyPI: README.md linked the documentation, the examples and the banner with
  relative paths, which PyPI resolves against pypi.org (`https://pypi.org/project/xlspy/docs/usage.md`
  was a 404). The README now links to the files on GitHub, and `tests/test_docs_links.py` checks those
  links against the working tree and that the README has no relative link left. The PyPI sidebar also
  lists the documentation, the changelog and the issue tracker.

### Added

- `write_sheet(autofit_rows=..., autofit_padding=...)`, also in `SheetOptions`, `write_dataframe()`,
  `to_xlsx()` and `to_xlsb()`: how many rows the automatic widths measure (`None`: all of them, `0`:
  none) and extra room in characters for every measured column. The rows are measured as they
  stream past (the XLSB writer now writes its column records after the rows, like the dimension),
  so `None` costs time, not memory.

### Changed

- The automatic widths measure the first 1000 rows instead of 100. Measuring is about 1-2
  microseconds per cell; on 50,000 x 50 cells the XLSB file takes about 0.1 s longer.
- Bold text is measured with Calibri Bold's own widths instead of 1.03 times the regular ones.

## [0.11.0] - 2026-11-02

### Breaking changes

Migration from 0.10:

| Before | Now |
| :--- | :--- |
| `XlsxWriter(path, compressionLevel=6, useSharedStrings=False)` | `XlsxWriter(path, compression_level=6, use_shared_strings=False)` (same for `XlsbWriter`, `to_xlsx`, `to_xlsb`) |
| `XLSPY_DISABLE_C_EXT=1` | `XLSPY_ACCEL=python` |
| `xlspy._accel.HAVE_C_EXT` | `xlspy._accel.HAVE_NATIVE` |
| type of `ws.protection` / `ws.page_margins`: `SheetProtection` / `PageMargins` from `xlspy.workbook._model` | `xlspy.workbook.WorksheetProtection` / `xlspy.workbook.WorksheetPageMargins` |
| `xlspy.outline.Outline` (the writer's outline index) | `xlspy.outline.OutlineIndex` |
| XLSB formula pointing at a missing formula group: `NotImplementedError` | `InvalidFileError` |
| a corrupt part: `xml.etree.ElementTree.ParseError`, `ExpatError`, `UnicodeDecodeError`, `struct.error`, `IndexError` or a plain `ValueError`, depending on the path | `InvalidFileError` (the original error is its `__cause__`) |

- The old writer keywords raise `TypeError` that names the new one (`argument 'compressionLevel' was renamed to
  'compression_level'`). Positional arguments are unchanged.
- Each name now means one class. `xlspy.SheetProtection` and `xlspy.PageMargins` are the writer settings, as
  before; the openpyxl-shaped objects behind `ws.protection` and `ws.page_margins` were also called
  `SheetProtection` and `PageMargins`, and `xlspy.SheetProtection(select_locked_cells=True)` means the opposite of
  the openpyxl-style `selectLockedCells=True`. The model classes are now `WorksheetProtection` and
  `WorksheetPageMargins` (exported from `xlspy.workbook`); both assignments still accept the public classes.
- Corrupt files raise `InvalidFileError` from `ExcelReader.open()` and its row readers, `load_workbook()` and the
  `XlsxUpdater`/`XlsbUpdater` constructors. It is a `ValueError`, so `except ValueError` still matches; code that
  caught `ParseError` or `UnicodeDecodeError` must catch `InvalidFileError` instead. A panic in the native reader
  (`PanicException`, a `BaseException`) is reported the same way. A missing sheet name is still a `KeyError`.
- `ExcelReader` raises `InvalidFileError` for an XLSB record cut short in a sheet's cell data, on both backends. It
  used to stop there and return the rows read so far as if they were the whole sheet. `load_workbook` raised already
  and now also catches a record id cut in the middle.

### Added

- Python 3.11 and 3.15 support: `requires-python = ">=3.11"`, CI and wheels for 3.11–3.15.
- Pivot tables (`PivotTable`, `PivotValue`): new pivots in XLSX and XLSB, passed as
  `write_sheet(..., pivots=[...])` (and `write_dataframe`). One row field, at most one column field, any number of
  filters (page fields) and values (sum, count, average, min, max). The source is a table written in the same
  call; the pivot can sit on that sheet below the data or on a sheet written after it. The report is rendered
  from the streamed rows and the cache refreshes when the file opens. Compared against Excel's own layouts and
  refreshed in Excel through COM. The XLSB records were mapped from files Excel saved
  (`tools/capture_xlsb.py --pivots`). Editing existing pivots is not included.
- XLSB pivot tables cannot read dates, booleans or NaN in their fields (`ValueError`), and a pivot placed on
  another sheet needs that sheet written with `write_sheet` (an empty list is enough).
- Sparklines (`SparklineGroup`): line, column and win/loss groups with colours, markers, high/low/first/last/negative
  points, empty-cell handling and axis limits, in XLSX and XLSB. The XLSB records were mapped from files Excel
  saved, one option at a time (`tools/capture_xlsb.py`).
- Table slicers (`Slicer`): a slicer on a table column, placed on the same sheet as the table, with a caption,
  size and column count, in XLSX and XLSB (binary slicer and cache parts, as Excel writes them).
- Sheet view switches in both writers and the model: `show_row_col_headers`, `show_zero_values` and
  `right_to_left`.
- Shapes and text boxes (`Shape`): rectangles, rounded rectangles, ellipses and text boxes with fill, border, font,
  alignment and an optional `link_cell` that shows a cell's value live (KPI cards). XLSX and XLSB.
- Chart formatting: `Chart.gap_width` and `hole_size` (doughnut), and chart-area and plot-area `fill`, `border`
  and `plot_fill`, plus `font_size` and `font_color`. XLSX and XLSB.
- `examples/business_dashboard.py`: a sales dashboard with KPI cards, charts, a slicer and sparklines, written as
  XLSX or XLSB from the same code. Its figures are `SUMIFS` over the rows the slicer leaves visible.
- `tools/capture_xlsb.py`: builds sparkline and slicer variants in Excel (COM) and prints the XLSB record bytes
  that differ from a base file.
- Exception classes: `XlspyError` is the base of the errors xlspy raises. `InvalidFileError` (also a `ValueError`
  and a `zipfile.BadZipFile`) is raised for a file that is not a valid Excel package or whose data cannot be read,
  and `UnsupportedFeatureError` (also a `NotImplementedError`) for a feature xlspy cannot write. Code that caught
  the built-in types keeps working. openpyxl's `InvalidFileException` is a subclass of `InvalidFileError`.
- Wheels for Linux aarch64, Windows ARM64 and macOS Intel (x86_64), in addition to the x86_64 and arm64 builds.
  The publish workflow runs on `v*` tags, refuses a tag that does not match the version in `pyproject.toml`, and
  installs and smoke-tests the source distribution in a clean virtual environment before it uploads.

- Loaded charts (XLSX and XLSB) can be edited: `LoadedChart.title`, `x_axis_title`,
  `y_axis_title` and `legend_position` (`None` removes the title or legend). The chart part is
  edited as text, so the rest of the file stays as Excel wrote it.
- `LoadedChart.set_series_title()`, `set_series_color()` and `series_titles`.
- XLSB formulas: the Analysis ToolPak functions (`EOMONTH`, `WORKDAY`, `NETWORKDAYS`, `GCD`, `XIRR`, ...,
  98 in all) with the ids, arities and operand classes Excel writes, and the newer `_xlfn` functions
  (`BITAND`, `NUMBERVALUE`, `UNICHAR`, `RRI`, ... also prefixed in XLSX). `RANDBETWEEN` is volatile.
- `Worksheet.ignored_errors` (`IgnoredError`): the ranges whose error checks are skipped, read and
  written for XLSX (XLSB refuses to save them).
- `Workbook.external_links` (`ExternalLink`): the workbooks that formulas refer to, XLSX and XLSB,
  read only and kept as they are on save.
- Charts: combined charts (`ChartSeries.chart_type`), a secondary value axis (`secondary_axis`,
  `Chart.y2_axis`), `Trendline` and `ErrorBars`, in XLSX and XLSB.
- Charts and images can be added to a loaded XLSB workbook, including on new sheets, and so can tables.
- Named styles (`add_named_style()`) and new chartsheets (`create_chartsheet()`) can be saved into a
  loaded XLSB workbook.
- Editing conditional formats in a loaded XLSB keeps the Excel 2010 extension (negative bar colours,
  axis, gradient) of rules that were not changed; extension rules the model cannot read are kept too.

- `benchmarks/bench_io.py`: subprocess-isolated read/write/model-load benchmark
  for XLSX and XLSB with JSON output, fallback-rate and zlib-level cases, and
  `--compare` regression gating.
- `XlsbWriter.write_sheet(auto_filter_ref=...)`; model-built `.xlsb` files now
  keep an explicit auto-filter range, and loaded `.xlsb` models read it.
- `load_workbook()` reads XLSB into the same model as XLSX: cell styles (from
  `styles.bin`, resolved lazily on first access), row/column dimensions, sheet
  properties, views and panes, selections, sheet format, sheet protection,
  print setup (with print areas and titles), autofilter columns, data
  validations, conditional formatting (colour scales, data bars and icon sets
  included), tables, charts, images, sheet states, the active sheet and
  document properties.
- Saving a loaded `.xlsb` workbook writes cell style changes (new styles are
  appended to `styles.bin`), merged cells, hyperlinks, data validations,
  conditional formats (new differential styles are added once), sheet
  layout (column/row dimensions, views, freeze panes, tab colour, sheet
  format), autofilter, sheet protection, print setup, print areas and titles,
  document properties (core and custom), removed charts and images, and sheet
  management: add, remove, rename (formulas, validations and defined names
  follow), reorder, hide and the active sheet. Adding tables, charts, images
  or chartsheets still raises `NotImplementedError` at `save()` instead of
  being dropped silently.
- `XlsbUpdater` sheet management: `add_sheet()`, `remove_sheet()`,
  `rename_sheet()`, `reorder_sheets()`, `set_sheet_state()`,
  `set_active_sheet()`, `get_sheet_states()`, `get_active_sheet()`, and
  `set_conditional_formats()`.
- XLSB colour scales, data bars and icon sets are read and written with
  formula thresholds (`FormatObject("formula", ...)`), which the XLSB writer
  used to reject.
- Conditional formats accept multi-range sqrefs (`"A1:A5 C1:C5"`), as Excel
  writes them; both loaders used to drop such rules.
- Removing a chart or image from `ws._charts` / `ws._images` (as in openpyxl)
  removes it from a loaded XLSX or XLSB file, with the parts only it used.
- `Workbook.create_chartsheet()` for XLSX (new and loaded workbooks).
- New `.xlsm`, `.xltx` and `.xltm` workbooks; `load_workbook(...,
  keep_vba=True)` sets `vba_archive`, whose VBA project a new `.xlsm` takes.
- `Workbook.move_sheet()`; renaming a sheet rewrites formulas, validation
  formulas and defined names that refer to it (`xlspy.formula.sheets`).
- `SheetProtection.password` hashes a plain password as openpyxl does
  (`set_password()`, `enable()`, `disable()`).
- `py.typed`; a CI "lint" job with ruff and mypy (`pyproject.toml` lists the
  modules still waiting for type fixes).
- `XlsbUpdater.update_cells(..., styles={ref: xf})`, `set_merged_cells()`,
  `set_hyperlinks()` and `set_data_validations()`.
- Benchmark cases `load_xlsb_styled` and `save_loaded_xlsb_styled`.
- Usage guide: "Read modes" section (`ExcelReader`, `read_only=True`,
  `load_workbook()`) and an XLSX/XLSB table of what a loaded workbook reads and
  saves.
- openpyxl compatibility harness (`tests/openpyxl_compat/`), six openpyxl-authored
  goldens (`tests/data/openpyxl_parity/`), `tools/make_openpyxl_goldens.py` and
  `tools/compare_openpyxl_model.py`.
- Public `xlspy.utils` modules (`cell`, `datetime`, `escape`, `units`,
  `exceptions`), `xlspy.styles.numbers` and `xlspy.formula` (`Tokenizer`,
  `Translator`).
- openpyxl-like model metadata: sheet format, full sheet views/panes/selection,
  full page setup/print options/header-footer, page breaks, sheet/workbook
  protection, calculation properties, workbook views, custom document
  properties and `lastPrinted`.
- Read support for tables, data validations, conditional formatting and named
  styles in loaded workbooks.
- Editing and removing loaded tables, data validations and conditional
  formatting rules, with table parts, relationships and sheet elements rebuilt
  from the model on save.
- `NamedStyle` write support: `add_named_style()` appends to
  `cellStyleXfs`/`cellStyles`, and `Cell.style = "Name"` applies a loaded or
  newly added named style.
- Model conveniences: `WriteOnlyCell`, `ReadOnlyCell.number_format`,
  `Cell.is_date`/`encoding`/`base_date`, `append(dict)`, `copy_worksheet`
  hyperlinks, and richer `CellRange`/`MultiCellRange`.
- Chart/image/chartsheet reading: `LoadedChart`, `chart_reader.py`, image
  anchors/descriptions and `Workbook.chartsheets`; the practical chart-type
  catalogue now includes radar, bubble, stock, surface, 3D variants and
  projected pie.
- XLSB workbook model read: values, formulas, dates, errors, merged ranges,
  hyperlinks, classic notes and defined names, with byte-identical no-op saves.
- XLSB model edits: cell values, comments and defined-name additions/removals
  through `XlsbUpdater`; existing cell XF styles are preserved.

### Changed

- The macOS wheels need macOS 10.12 or later on Intel (x86_64) and 11.0 on Apple silicon: the
  native module is Rust now, whose standard library starts at 10.12. The 0.10 wheels claimed 10.9.
- `XlsxWriter` and `XlsbWriter` choose the default autofilter by one rule (`tests/test_autofilter_parity.py`).
  An explicit `auto_filter_ref` is written as given. Otherwise a filter is added only when data follows the
  header, and it spans the columns up to the last one holding a value and the rows down to the last data row or
  the last row of an array formula. What changes:
  - XLSX: a trailing `None`, a formatted cell without a value, or a styled empty header cell no longer widens
    the filter (`[["h1", "h2"], [1, 2, None]]` gave `A1:C2`, now `A1:B2`, as XLSB); the rows of an array
    formula that reaches past the data are now inside the filter (as XLSB).
  - XLSB: a header with no data rows below it, or a sheet with no values, no longer gets a filter (as XLSX);
    `auto_filter=True` over exactly a table's range no longer adds a second filter beside the table's own (as
    XLSX); the filter buttons of `auto_filter_ref` now widen their columns (as XLSX).
- `ExcelReader` reads workbooks with chartsheets: `get_rows`/`read_all` on a chartsheet return no rows and
  `get_comments` an empty dict, where they raised `InvalidFileError` ("no worksheet part") or `KeyError`.
  `get_chartsheet_names()` lists them. A read-only `load_workbook` leaves chartsheets out of `worksheets`, like a
  full load. A worksheet whose part is missing from the package raises `InvalidFileError` instead of `KeyError`.
- `ws.tables` also works like openpyxl's name-keyed `TableList`: `del ws.tables["Sales"]`, `get`, `keys`,
  `values` and `items()` (`(name, ref)` pairs). Iterating still gives the tables. Tables added as `ExcelTable` are
  now found by name too.
- Type annotations: every public class and function exported from `xlspy` and `xlspy.workbook` is fully
  annotated (a test guards this), so mypy and IDEs check calls into the library. Writers, `Workbook.save`,
  `save_workbook`, `to_xlsx` and `to_xlsb` accept `str`, any `os.PathLike` (such as `pathlib.Path`) or a binary
  file object; `to_xlsx`/`to_xlsb` return the target they were given. `Worksheet.iter_rows`/`iter_cols` are
  typed by `values_only` (cells or values). `ReadOnlyWorksheet.max_row`/`max_column` are `int | None`, as they
  are `None` until the sheet is sized.
- Internal: mypy requires annotations on every def by default (`disallow_untyped_defs`), and also runs with
  `warn_return_any` and `no_implicit_reexport`. The modules still to annotate are listed in `pyproject.toml`
  (61 of 134). The writer mixins are typed through one host class per writer, so attributes and methods shared
  between mixins have real types instead of `Any`.
- Internal: 18 private helpers that nothing called are removed (leftovers of the module split and older code paths).
- `__version__` is read from the installed package metadata, so `pyproject.toml` is the only place the version is
  set. A source tree that was never installed reports `0+unknown`.
- Tests that start Microsoft Excel (`test_excel_com_conformance.py`, the Excel checks in `test_comprehensive.py`) are
  marked `com` and need `XLSPY_RUN_EXCEL_COM=1`, like the other Excel tests. Each test has a 300 s time limit
  (`pytest-timeout`).
- Internal: CI lints `tests`, `tools`, `benchmarks` and `examples` with ruff, type-checks for Python 3.11 and 3.14,
  runs the pure-Python fallback on Linux, Windows and macOS with Python 3.11 and 3.14, and the release workflow runs
  the test suite before it builds wheels. The `test` extra pins `openpyxl==3.1.5` (the version the compatibility
  goldens were made with), which replaces the separate `compat` job. The Excel fixtures under `tests/data` that
  the tests need are committed.
- Internal: names that several modules defined are consolidated. Identical copies now have one owner
  module; helpers with the same behaviour are one function; different functions that shared a name are
  renamed, so each name means one thing. Only private names changed, with one exception:
  `xlspy.workbook._model._PLAIN_TYPES` (an accidental re-export) is no longer exported. The package's
  write-back of pending parts is owned by `ZipPackage.apply_parts`, and `<sheetData>` is located in one
  module (`xlspy._sheet_data`).
- Internal: the four largest modules are split. `styles/core.py` and `workbook/_model.py` are now facades
  over focused modules; `xlsx_writer.py` and `writer.py` keep `XlsxWriter` and `XlsbWriter` and take their
  methods from mixins. Every name the old modules exported still imports from the same path, and the
  object is the one its defining module holds. Stdlib names no longer leak through `xlspy.styles` and
  `xlspy.workbook._model`. The accelerator call sites moved with their methods, so code that patches
  `xlspy.xlsx_writer.accel_encode_xlsx_row` or `xlspy.writer.encode_xlsb_row` must patch the new module.
- Ragged rows: the first row present in an XLSX sheet sets the width the other rows are padded to in
  `ExcelReader.get_rows()` / `read_all()` (as documented in `docs/usage.md`). `Worksheet.iter_rows()` pads to the
  sheet width, as openpyxl does. A blank first row used to count as width 0, so the padding did not happen; now it is the same
  whether or not the sheet starts with a blank row.
- The native accelerator selects one backend, the Rust module. Its private switches were renamed (`_c_*` to
  `_native_*` in `xlspy._accel`). The public names (`HAVE_NATIVE`, `HAVE_C_EXT`, `ACCEL_MODE`, `ACCEL_BACKEND`,
  `XLSPY_ACCEL`, `XLSPY_DISABLE_C_EXT`) are unchanged.
- Mypy checks every module in `src/xlspy`: the `ignore_errors` list in `pyproject.toml` (nineteen modules at the start
  of this cycle) is gone.
- Ruff also enforces `PLE`, `PIE`, `TRY203`, `C4`, `RET`, `B010` and `B023`, which report nothing after this change.
- Pivot tables no longer resize the columns of their sheet when Excel refreshes them on open, unless the
  new `PivotTable(autofit_columns=True)` asks for it (Excel's "Autofit column widths on update", written for
  XLSX and XLSB). Before, a pivot placed on its data sheet could narrow the columns the sheet had been given.
- Plain rows given to `write_sheet` (XLSX and XLSB) write numpy integer, `bool_`, float16 and float32 scalars as the
  Python value they hold, as the DataFrame path already did. They used to be written as text (`"1"`, `"True"`,
  `"0.1"`). A float32 `0.1` is now written as `0.10000000149011612`, and a float32 `NaN` raises `ValueError`, as a
  float `NaN` does. Auto column widths measure the value that is written.
- `XlsxUpdater` and `XlsbUpdater` raise `ValueError` for NaN and infinite numbers, as `update_cells` already did.
  They used to write them into the file.
- `XlsxWriter` and `XlsbWriter` raise `ValueError` for NaN and infinite numbers. XLSX used to write them as text
  (`NaN`, `∞`) in the cell.
- A loaded XLSB whose auto filter has custom criteria (such as `> 2`, not a value list) keeps the criteria when the
  file is saved unchanged, and the filter can be removed. Changing that filter raises `NotImplementedError` on save;
  it used to drop the criteria and leave a blank-match filter in Excel.
- `add_named_style()` on a new workbook (one not loaded from a file) raises `NotImplementedError` on save, because
  the new-file writers do not write named styles. Save the workbook first, load it, add the style and save again.
  Loaded XLSX and XLSB workbooks keep named styles.
- Renaming a sheet rewrites the sparkline formulas that refer to it.
- Removing a sheet that holds table slicers from a loaded workbook (XLSX or XLSB) raises `NotImplementedError`
  before anything is written.
- The native accelerator is now a Rust module (`xlspy._rs_core`, PyO3) instead of the C extension
  (`xlspy._c_core`, removed). It covers the same paths with the same results: XLSB and XLSX reading and row
  encoding, shared strings, column widths and workbook-model loading; the reads and writes are as fast or faster.
  The build uses maturin, so installing from the source distribution needs a Rust toolchain. `XLSPY_ACCEL=python`
  selects the pure-Python fallback (`XLSPY_DISABLE_C_EXT=1` still works); `xlspy._accel.HAVE_NATIVE` reports
  the extension, and `HAVE_C_EXT` stays as an alias.
- Native style and xf ids above 2**32 - 1 raise `OverflowError` on every platform (the C extension raised on
  Windows and silently truncated them on Linux and macOS).
- Errors from the native readers for data they cannot parse (a truncated XLSB record, a row index past the last row,
  a column past Excel's limit, an XLSX worksheet with no row closed) raise `InvalidFileError` instead of a plain
  `ValueError`. The message is unchanged.
- `NotImplementedError` raised for unsupported features is now `UnsupportedFeatureError`, which is still a
  `NotImplementedError`.
- The minimum supported Rust toolchain is 1.88 (`rust-version` in `rust/xlspy_core/Cargo.toml`), and CI builds the
  crate on it.
- XLSX writing is faster for formula and hyperlink cells, which the native row encoder now writes. The sheets are the
  same: the native and Python writers are compared package by package. A 50,000 × 20 sheet of formulas went from 0.86 s
  to 0.29 s; a sheet of hyperlinks from 2.1 s to 1.3 s (its relationships are still written by Python).
- XLSX shared strings are parsed natively: about 6× faster than the regular-expression parser on a million short
  strings, and `ExcelReader` opens such a file in 0.16 s instead of 0.86 s.
- Sheets whose text has character references or line breaks no longer fall back to ElementTree. `ExcelReader` reads
  a 50,000 × 50 sheet of that kind in 0.69 s instead of 6.5 s, and `load_workbook` loads it about 8× faster.


- The mutation fuzzer is back as a bounded test, `tests/test_native_fuzz.py` (300 seeded cases per format; set
  `XLSPY_FUZZ_ITERATIONS` for more). The unbounded tool, `tools/fuzz_native.py`, stays removed.

- Much faster `load_workbook()` (50k x 20 cells: XLSX 6.5 s to 0.39 s, XLSB
  5.5 s to 0.55 s; files saved by Excel: 0.47 s and 0.71 s). The C extension
  builds the model's cells directly (`XlsxReader` model mode,
  `xlsb_load_cells`), setting their slots without Python per-cell code;
  formulas, rich text and unusual rows go back to Python. Without the C
  extension XLSX rows are matched with regular expressions
  (`workbook/_fast_sheet.py`) and XLSB sheets are walked once with
  `biff_utils.iter_record_payloads`. Differential tests compare the paths.
- `IconSetRule.type` is the rule type, `"iconSet"`, as in openpyxl; the
  threshold type is `cfvo_type`. `IconSetRule(type=...)` still works with a
  `DeprecationWarning`.
- `copy_worksheet()` gives the copy its own dimensions, view, properties,
  autofilter and print settings instead of sharing the source's objects.
- Loaded XLSX: repeated conditional format edits reuse equal `dxf` entries
  instead of appending copies to `styles.xml`.
- The writers share their argument checks and sheet bookkeeping
  (`_writer_common.WriterFrontEnd`).
- Faster cell writes on the model (`Cell.value` for plain int/float/bool/str and
  memoised number-format changes): building 1M cells 5.5 s to 3.5 s.
- A regex fast path for XLSX shared strings in `ExcelReader`, with expat as the
  fallback.
- `XlsbUpdater.update_cells()` walks the worksheet once instead of three times,
  and BIFF12 record headers are decoded inline: saving 1,000 edits into a
  50k x 20 loaded XLSB takes 2.4 s instead of 5.0 s.
- XLSB style record encoders moved from `writer.py` to `_xlsb_style_records.py`
  so the writer and the loaded-workbook saver share them.
- Faster streaming writes (C extension): exact `date`/`datetime` values are
  read through the C datetime API instead of attribute lookups, and whole
  numbers below 1e16 skip the general float-to-text routine in XLSX. Per
  million cells: XLSB datetimes 507 to 148 ms, XLSX dates 396 to 164 ms, XLSX
  integers 233 to 168 ms. Output is byte-identical to the Python encoders.
- XLSX shared strings are escaped a chunk at a time instead of string by
  string (1M unique strings: 262 to 154 ms).
- `Worksheet.append()` stores plain int/float/bool/str values, and dates with
  their default format, directly when the row is new and the sheet has no
  merged cells; saving the model reads each row once. Building and saving a
  50k x 20 model: 3.5 s to 1.9 s (XLSX), 3.3 s to 1.8 s (XLSB).
- Worksheets and shared strings are compressed on a worker thread while rows
  are encoded (`_zip_threads.py`; zlib releases the GIL). An XLSX sheet's
  rows are compressed as they are written and joined to its head and tail as
  one deflate stream, as pigz does. Streaming XLSX writes are 15% faster
  (50k x 20, level 3), more at higher levels (level 9: 25%); file sizes are
  unchanged. The members hold the same bytes; XLSB members at levels 3-9 are
  also byte-identical when compressed.
- `xlspy.styles` is now a package with the original names re-exported from
  `xlspy.styles.core`; `styles.py` remains import-compatible.
- `ExcelReader` accepts `errors="openpyxl"` for bare `#DIV/0!`-style values;
  the default remains the historical `error:#DIV/0!` form. Error values are
  a `str` subclass, so text that starts with `error:` is never taken for an
  error (by `errors="openpyxl"` or by `load_workbook()` on XLSB).
- Renaming a sheet also updates both ends of 3D references (`Jan:Mar!B2`),
  conditional-format formulas, internal hyperlinks and the series of loaded
  charts.
- Sheet and chartsheet names follow Excel's rules and are unique ignoring
  case across worksheets and chartsheets. Renaming a loaded chartsheet
  raises `NotImplementedError`.
- Saving a loaded workbook refuses (`NotImplementedError`) to rewrite a
  sheet's data validations or conditional formats while it holds rules the
  model cannot read, such as a prompt-only validation, and to remove one
  object of a grouped drawing; both would lose data.

### Removed

- **Breaking:** `XLSPY_DISABLE_C_EXT` and `xlspy._accel.HAVE_C_EXT`, kept from before the Rust port. Use
  `XLSPY_ACCEL=python` and `HAVE_NATIVE`. Setting `XLSPY_DISABLE_C_EXT` now makes `import xlspy` raise
  `ImportError`, so a script that still sets it does not silently run the native code. The benchmark scripts take
  `--native on|off|both` instead of `--c-ext`, and their JSON reports `native` instead of `c_ext`.

### Fixed

- `XlsbWriter.write_sheet` checks `auto_filter_ref` before the rows stream, as `XlsxWriter` does. Before, a
  non-string or malformed range failed after the sheet part was started, and the writer could no longer save.
  Both writers now share one option check (`WriterFrontEnd._check_sheet_options`).
- `ExcelReader.open()` (and `with ExcelReader(path)`) closes the archive when the workbook part fails to parse.
  Before, the ZIP file stayed open, and on Windows the file could not be deleted until the process ended.
- **Breaking:** an XLSB cell formula that points at a shared or array formula the sheet does not define raises
  `InvalidFileError` instead of `NotImplementedError`; a malformed reference to one raises `InvalidFileError`
  (still a `ValueError`) instead of a plain `ValueError`. The file is damaged; nothing is unimplemented.
- The openpyxl-compatible exceptions in `xlspy.utils.exceptions` (`CellCoordinatesException`, `SheetTitleException`
  and the others) derive from `XlspyError`, so `except XlspyError` catches them.
- The column-letter caches in `xlspy.utils.cell` are bounded (mixed-case names could grow them to ~140,000 entries).
- The type stub `_rs_core.pyi` lists the `write_r` parameter of `encode_xlsx_row`; `stubtest` failed without it.
- A named style whose fill is a colour string (`fill="00FF00"`) no longer raises `AttributeError` when it is written
  to XLSX or added to an XLSB workbook. A colour string is a solid fill of that colour.
- A corrupt XLSX `sharedStrings` part without an `sst` element raises `ValueError` ("sharedStrings element was not
  found") instead of `TypeError`.
- `ws.protection` and `ws.page_margins` accept the public `SheetProtection` and `PageMargins` objects and store
  them as the model's own types, so a sheet protected with `xlspy.SheetProtection(password=...)` saves and loads.
- Writing a sheet whose data is far from the origin (for example `A1` and `XFD1048576`) no longer builds every
  empty row between them. A row ends at its last cell, and empty rows with no height, flag or style are not written
  (XLSX and XLSB). The file for `A1` and `XFD1048576` is 2.7 KB instead of 2.7 MB, and it takes under a second.
- A row past the last cell that has a height, a hidden flag or an outline level is written. Before, its height was
  lost on save.
- XLSX rows carry no `r` attribute when they follow the previous row, so dense data is not made larger. The row
  number is written only after a gap. Excel, openpyxl and xlspy read both forms the same way.
- Empty rows at the end of the data are still written, so an empty frame keeps its one row.

- Files Excel would repair, or that held wrong data:
  - XLSB `replace_sheet_data` wrote the sheet dimension into the selection record (0x98) and left the real
    dimension (BrtWsDim, 0x94) stale. Both the in-memory and streaming paths are fixed.
  - Chart formulas are XML-escaped: a sheet named `Q&A` produced a chart part that did not parse (XLSX and XLSB).
  - Sheet names holding `"` are escaped in `workbook.xml`; control characters in sheet names are rejected.
  - `replace_sheet_data` drops `calcChain` when formula cells change, so Excel no longer repairs the workbook for
    entries that point at cells without formulas.
  - `replace_sheet_data` writes formula cells as formulas in XLSX (it wrote their text before). In XLSB it raises
    `UnsupportedFeatureError`; use `update_cells()` for formulas there.
  - `replace_sheet_data` (XLSX and XLSB, both paths) raises `UnsupportedFeatureError` before writing anything when a
    table's header cell would not equal its column name (ignoring case), or when the new data does not reach the
    header row. Excel repairs such tables; the package is left unchanged.
  - The pure-Python XLSB cell reader (used when the native module is not available) looped forever when a sheet
    stream ended right after a cell record. The native reader was not affected.
  - The native XLSX and XLSB row writers check the row length again before every element read.
  - Built-in number formats 41–44 (accounting) are read and kept; they used to load as `General` and be lost on save.
  - A sheet added without `write_sheet` gets its part in XLSB, as it already did in XLSX. A second `write_sheet` on
    the same sheet raises: it used to write duplicate ZIP entries.
  - Table names are compared without regard to case, as Excel compares them.
  - A loaded sheet with an empty `<sheetData/>` keeps its new rows inside it. A missing `<dimension>` is inserted
    after `<sheetPr>`, as the schema orders it.
  - XLSB autofilters and column widths cover columns past the first 100 rows, and explicit widths for those columns
    are kept.
  - Charts: radar series reject trendlines and error bars; a fractional `title_size` is written as whole hundredths
    of a point; control characters in chart, shape, image and hyperlink tooltip text raise `ValueError`;
    `set_series_color(None)` raises `TypeError` instead of writing an empty colour.
  - Formulas: the `_xlfn.` prefix is added to 30 more functions (for example `COT`, `SEC`, `FORECAST.ETS`,
    `CONFIDENCE.NORM`); `FILTER` and `SORT` take `_xlfn._xlws.`.
  - External links keep a placeholder for a reference that cannot be resolved, so later `[n]` numbers stay correct.
- XLSX cell references are read the same way on every path (native scanner, regular-expression fast path,
  ElementTree, workbook model). A reference without a row (`r="A"`), with row 0 or past row 1,048,576, or with a
  column past `XFD` now raises `ValueError`; the fast paths used to accept some of them, reading only the column
  letters, depending on which letters the sheet had used before. Leading zeros (`A01`) are accepted everywhere.
- The type annotations of `PivotTable` (`source`, and `value_specs` for the normalised `values`) pass mypy.
- XLSB tables are written with the same name and display name, as Excel does. The display name used to get a
  trailing `_`, so Excel could not resolve the table by its name as a pivot source ("cannot retrieve data from
  range").
- XLSX pivot tables with several values and a column field list the data field in the repeated column items,
  as Excel does.
- Radar charts (`Chart("radar", ...)`) are written as radar charts. They used to be written as pie charts, and so
  were `area3d`, `bar3d`, `bubble`, `line3d`, `pie3d`, `projected_pie`, `stock`, `surface` and `surface3d`; those
  types now raise `ValueError` when the chart is created.
- `to_xlsx()` and `to_xlsb()` no longer delete an existing file at `path` when the conversion fails; the file is
  left as it was.
- A loaded XLSX keeps its autofilter criteria other than value lists (such as custom filters) when another edit
  rewrites the sheet, and keeps style-only column definitions (`<col style=...>`) when a column width changes.
- A sheet added to a loaded XLSX keeps the sheet protection, ignored errors, row and column breaks and the default
  row height that were set on it.
- `Workbook()` saves tables added as openpyxl-style `Table` objects (it raised `TypeError`).
- `XlsxUpdater` and `XlsbUpdater` keep control characters in text. XLSX stores them as `_xHHHH_` escapes, as
  Excel does; XLSB stores them as they are. Text longer than Excel's 32,767-character limit raises `ValueError`
  instead of being written.
- `ReadOnlyCell` is exported from `xlspy`.
- An unclosed workbook opened by an updater no longer emits a `ResourceWarning` for its temporary directory when it
  is garbage collected; the directory is removed by `close()` or when the updater is collected.
- The native XLSX reader decodes character references (`&amp;`, `&#65;`) and carriage returns in cell values itself.
  Sheets that have them are no longer read by the slower ElementTree path, and typed values with references
  (`t="e"`, `t="d"`, `t="b"`, numbers) read the same as ElementTree does; they used to differ.


- Adding a sheet to a loaded workbook wrote `r:id` without declaring the `xmlns:r` prefix, so saving that workbook
  failed with `ParseError: unbound prefix`.

- `load_workbook()` of an XLSB file with formulas that refer to another workbook raised
  `NotImplementedError`; they now read as `='[1]Sheet'!A1`. Writing such a reference is still refused.
- `LoadedChart.title` is the chart's own title, not the title of its first axis.
- `load_workbook()` of an XLSB file with a chartsheet raised `KeyError`; chartsheets are now loaded
  (`Workbook.chartsheets`) and kept when sheets are renamed, added or reordered.

- XLSB sheet protection permissions were read inverted (the file stores
  "allowed" flags as "protected" bits); a new workbook ignored
  `ws.protection`.
- The XLSB writer used record ids 0x86/0x87 for BrtBeginColInfos and
  BrtEndColInfos instead of 0x186/0x187; loading such a file and editing
  column widths now drops the old wrappers. Excel tolerated the old ids, but
  LibreOffice dropped every cell of such a sheet (#1).
- `XlsbWriter` wrote `BrtWsDim` as `(0, 0, 0, 0)` whatever the data (#1). Excel
  ignores the record on load, but readers that size rows from it failed: pyxlsb
  raised `IndexError` on the first cell past column A, and
  `load_workbook(read_only=True)` reported no `max_row`. The record now holds
  the used range by Excel's rule, checked against the files Excel saves: rows
  with cells or a non-default row header (height, hidden, outline), columns of
  every cell record (values, formulas, styled blanks, array members); an empty
  sheet keeps zeros. Since the record precedes the rows, a sheet larger than the
  64 KB row buffer is compressed on the worker thread into a spool (memory up to
  1 MiB, then disk) and written behind its head as one deflate stream, as the
  XLSX writer does for `<dimension>`. Speed and peak memory are unchanged.
  Read-only rows are now padded to the dimension's width, as for a file Excel
  saved; files from older versions still read as unsized.
- `XlsxWriter` and `XlsbWriter` wrote a `(None, format)` or `(None, Style)` row entry as the
  text "None". It is now a formatted cell without a value, written like a blank with a range
  style: Excel, LibreOffice and openpyxl see an empty cell with the format or style.
- A sheet added through the model and saved as `.xlsb` got a dimension
  stretched to A1 (`create_sheet("New")["D5"] = 1` gave `A1:D5`, Excel writes `D5`).
- XLSB icon sets: `reverse` was written inverted, and `4Rating`/`5Rating`
  were rejected; data bars without lengths now use Excel's 10%/90%.
- XLSX dataBar `showValue`/`minLength`/`maxLength` were read from `<cfRule>`
  instead of `<dataBar>`; numeric CF thresholds are read as numbers.
- Loaded XLSX workbooks dropped edits to document properties on save.
- Loaded XLSB cells had `data_type` `"n"` whatever their value.
- The regular-expression path for XLSX shared strings kept literal CR and
  CRLF line breaks (as Excel writes multi-line text) instead of turning them
  into `\n` as the XML parser does; it now applies XML end-of-line handling.
- The XLSB C encoder could store `datetime` values with microseconds one
  unit in the last place away from `datetime_to_serial` (and from the Python
  encoder and XLSX); it now computes the serial the same way, and leaves
  datetimes beyond the exact range (after about 2185) to the Python encoder.
- Loaded XLSX conditional formats lost the formulas of `cellIs` and
  `expression` rules (the reader looked for `<f>` instead of `<formula>`), so
  editing the sheet's conditional formats rewrote those rules without formulas.
- A failed `load_workbook()` closes its temporary updater package instead of
  leaking a `ResourceWarning`.
- Theme colours accept indices `0..11` (hyperlink styles use theme 10).
- A no-op save no longer rewrites `xl/workbook.xml`; untouched ZIP parts are
  preserved.
- Changing `freeze_panes` updates the model pane and selection so the saved
  sheet reflects the new split.
- Updaters/evaluators see conditional-formatting groups and flattened rules
  correctly, so adding objects to a loaded sheet does not duplicate or reject
  existing rules.
- XLSB error cells (`#N/A`, `#DIV/0!`, ...) read as `None` through the bulk
  `ExcelReader` path, with and without the C extension; they are error values
  again, as in the streaming path.
- `ExcelReader` and `ExcelReader.open()` accept `pathlib.Path` objects, not
  only strings.

### Security

- Rows wider than Excel's 16,384 columns are refused. The native XLSB reader used to accept 1,048,576, so a crafted
  first row could make it allocate a million-cell row per record. The pure-Python XLSB reader raises the same
  `ValueError`.
- The XLSX streaming reader refuses a sheet part that grows past 64 MiB without closing a row, and it no longer
  rescans a pending row from its start on every chunk.
- Package parts declared above 2 GiB uncompressed are refused before they are read, on every archive open path.
- The Rust module is built with `panic = "unwind"`: a panic inside a reader raises `PanicException` in Python
  instead of ending the interpreter.
- Every XML parse goes through one helper (`xlspy._safe_xml`). A document type declaration or entity declaration in
  the prolog is refused with `InvalidFileError`, which blocks entity-expansion ("billion laughs") payloads.
- A package whose parts add up to more than 4 GiB uncompressed is refused, not only a part above 2 GiB.

### Internal

- The object model reaches an updater's package only through `PackageHandle`; `workbook/*` reads no `updater._*`
  attribute (73 before). Function-level imports went from 150 to 21, and no reader, writer or updater module
  imports `xlspy.workbook` (`tests/test_import_layers.py`).
- Shared XLSX worksheet walks in `_sheet_data.py` (`patch_rows`, `insert_rows`, `patch_cells`,
  `set_dimension_ref`); shared XLSB record encoders for panes, selections, protection, margins, breaks and
  headers/footers, used by both the writer and the loaded-workbook save.
- The longest functions are split into named steps (the XLSB worksheet encoder, both `write_sheet` paths, the
  XLSX and XLSB model loaders, the Python row readers). Output is byte for byte the same; benchmarks are within 3%.
- `docs/architecture.md` describes the layers, `PackageHandle`, the XLSB updater/model map and the error policy.
- Rust: every `unsafe` block states why it is sound, and the crate's lints (`undocumented_unsafe_blocks`,
  `unsafe_op_in_unsafe_fn`) keep it so. The row walk and integer conversion are shared helpers; the pure encoders
  have Rust unit tests (16 to 27).
- Tests are organized by topic instead of by review round, with shared helpers in `tests/helpers.py`. CI measures
  coverage over the native and the pure-Python run together (floor 89%), and a smoke test runs every script in
  `examples/`.
- Documentation: the usage guide is split into topic pages under `docs/guide/`; an API reference is generated from
  the docstrings with mkdocs and mkdocstrings (`docs/api/`, the `docs` extra; CI runs `mkdocs build --strict`);
  every public method and property has a docstring (a test keeps it so); `CONTRIBUTING.md` describes the
  development setup. The benchmarks were measured again on the Rust module (`docs/benchmarks.md`).

## [0.10.1] - 2026-10-24

### Changed

- The default ZIP compression level of both writers is 3 (was 4), aligned with
  the other libraries in the benchmark comparison. Pass `compressionLevel=` to
  override it.

## [0.10.0] - 2026-10-24

### Changed

- XLSX `get_rows()` / `read_all()` with the default `data_only=True` read sheets with a
  fast path over the raw XML: 3.9 s → 2.1 s on a 50,000 × 20 sheet from the Python regex
  reader, and 0.45 s with the optional C extension (`read_xlsx_worksheet`), which scans
  the worksheet stream in bounded chunks. Rows the fast path cannot read exactly (rich
  inline strings, entities in `str` cells, unusual markup) and sheets with prefixed tags
  are read by the ElementTree path as before, so results are unchanged. Formula reads
  (`data_only=False`) are unchanged, and without the extension nothing changes.

### Fixed

- XLSX `get_rows()` / `read_all()` keep worksheet row positions: rows absent from
  the XML between present rows are now yielded as empty rows. Previously every
  later row shifted up. Rows before the first row present in the XML are `[]`,
  and rows after it are padded to the width of the first present row, as before.

## [0.9.2] - 2026-10-23

### Added

- Hyperlinks on a loaded workbook: links set, removed or moved by
  `insert_rows()`, `delete_rows()` and `move_range()` are saved instead of
  raising `NotImplementedError`, and new sheets of a loaded workbook keep
  their links. `XlsxUpdater.set_hyperlinks()` edits links directly.
- `load_workbook()` reads a sheet's print area, print titles, page setup,
  fit-to-page, centring, header and footer into the model.

### Fixed

- A date in a styled cell (a range style, a `(value, Style)` cell, or a model
  cell with a font) was written with the General format and showed as a
  serial number in Excel. It now gets the default date or date-time format
  unless its style sets one; both writers and their C encoders. Assigning a
  date to a model cell sets that format, as for times.
- Clearing print settings on a loaded sheet (`ws.print_area = None`, default
  margins, no fit to page, an empty header) removed nothing from the file.
- Rich text in shared strings was escaped twice, so `\r`, control characters
  and literal `_xHHHH_` text came back as escape codes.
- `load_workbook(rich_text=True)` failed on runs with `u val="none"` or an
  out-of-range font size written by other programs.
- Saving row metadata of a loaded sheet dropped `collapsed="1"` from a
  collapsed group's summary row.
- XLSB charts whose categories were only rich-text cells got an invalid
  numeric cache.

## [0.9.1] - 2026-10-22

### Fixed

- Saving a loaded workbook no longer rewrites the worksheet part of a sheet
  whose print setup was not edited: the page setup read from the file was
  mistaken for a pending change, so untouched sheets gained a rewritten
  `pageMargins` element (and were no longer copied byte-for-byte). Printed
  settings changed through the model are still saved.

## [0.9.0] - 2026-10-22

### Added

- `show_grid_lines=False` on `write_sheet()`/`write_dataframe()` of both
  writers hides a sheet's grid lines (XLSX `showGridLines="0"`, XLSB
  `fDspGrid`). The workbook model now also saves
  `ws.sheet_view.showGridLines = False` for new XLSX/XLSB workbooks.
- `(value, Style)` cells: a per-cell style in the row data, applied on top of
  header and range styles like `(value, number_format)`. They give the same
  file as one `styles` entry per cell, 3.5-5x faster.
- Open-ended `styles` keys: whole columns (`"B:D"`), columns from a row down
  (`"B2:D"`) and whole rows (`"1:1"`), for streamed sheets of unknown length.
- Document properties: `set_properties()` on `XlsxWriter` and `XlsbWriter`
  writes title, subject, author, manager, company, category, keywords,
  comments, status, creation date and custom properties (*File → Info* in
  Excel). The workbook model now saves `wb.properties` for new workbooks.
- Charts (XLSX): `"area"` and `"doughnut"` types, and
  `grouping="stacked"`/`"percent_stacked"` for bar, horizontal bar, line and
  area charts.
- `Image` accepts image bytes or a binary file object (`io.BytesIO`) besides a
  path, and GIF and BMP pictures besides PNG and JPEG.
- `tab_color` on `write_sheet()`/`write_dataframe()` of both writers colours the
  sheet tab (an RGB string or a `Color`, theme colours included). The workbook
  model saves `ws.sheet_properties.tabColor` for new workbooks.
- `header_rows=N` on `write_sheet()` of both writers: for a header of several
  rows the automatic filter starts on the last header row and `header_style`
  covers every header row.
- openpyxl compatibility of the workbook model:
  - `ws.add_chart()`, `add_image()`, `add_table()`, `add_data_validation()` and
    `ws.conditional_formatting.add(range, rule)` for new workbooks (xlspy's
    classes; openpyxl objects get a `TypeError` naming the xlspy class).
  - Print setup: `print_title_rows`/`print_title_cols`, `print_area`,
    `page_setup`, `page_margins`, `print_options`,
    `sheet_properties.pageSetUpPr.fitToPage`, `oddHeader`/`oddFooter`; and
    `sheet_view.zoomScale` for new workbooks.
  - `insert_rows()`, `delete_rows()`, `insert_cols()`, `delete_cols()` and
    `move_range(..., translate=True)`, in new and loaded workbooks. Row heights
    and column widths move with their rows and columns. 5-6x faster than
    openpyxl on 100,000 cells.
- Rich text: `CellRichText`, `TextBlock` and `InlineFont` (openpyxl's names) in
  `XlsxWriter` (shared and inline strings), the workbook model and XLSX
  updates; `load_workbook(..., rich_text=True)` reads them back.
- Conditional formatting (XLSX): `TopBottomRule`, `AboveAverageRule`,
  `TextRule`, `DuplicateRule`, `TimePeriodRule` and `BlanksRule`.
- `zoom=` on `write_sheet()`/`write_dataframe()` of both writers.
- `Hyperlink` cells in `XlsbWriter`: external (web, mail) and internal links,
  tooltips and the link look, as in XLSX. `XlsbWriter` used to raise
  `TypeError` for them. Excel reads back the targets, sub-addresses, tooltips
  and styling.
- `XlsbWriter` now writes what `XlsxWriter` writes, with the same parameters.
  It used to raise `TypeError` or `ValueError` for these. Each one was checked
  against the same file saved as XLSX and opened in Excel 16:
  - all chart options: horizontal bar, scatter, area and doughnut charts,
    stacking, axes, legend, data labels, markers and colours;
  - `images=` (PNG, JPEG, GIF, BMP), in the same drawing as the charts;
  - `print_settings=`: orientation, paper, scale, fit to pages, margins,
    centring, headers and footers, page breaks, print area and print titles;
  - `data_validations=`: all seven types, lists from text or ranges (other
    sheets too), prompts and error messages;
  - conditional formats with a `DifferentialStyle`: `CellIsRule`,
    `FormulaRule`, `TopBottomRule`, `AboveAverageRule`, `TextRule`,
    `DuplicateRule`, `TimePeriodRule` and `BlanksRule`. Gradient fills,
    diagonal borders, vertical alignment and rare font options still raise
    `TypeError` in XLSB.

  The workbook model saves images, data validation and print setup to `.xlsb`.
- Loaded workbooks save added objects: `add_chart()`, `add_image()`,
  `add_table()`, `add_data_validation()`, `conditional_formatting.add()` and
  print setup on a sheet of a workbook from `load_workbook()`. They are added
  next to the file's own objects, which are kept; checked in Excel on files
  written by xlspy and on files saved by Excel. Saving used to raise
  `NotImplementedError`.
- Rich text (`CellRichText`) in `XlsbWriter`, as Excel stores it: a shared
  string with font runs, whose fonts join the cell fonts in `styles.bin`.
  Excel shows the same fonts per character as for the XLSX file. It used to
  raise `TypeError`.
- XLSB formulas: the `EDATE` function, and the volatile-function marker that
  Excel expects before `TODAY`, `NOW`, `RAND`, `OFFSET`, `INDIRECT`, `CELL`
  and `INFO`.

### Changed

- The same workbook gives the same bytes on every save: ZIP members carry
  Excel's fixed 1980-01-01 timestamp instead of the current time, and the
  revision ids in XLSB table parts are derived from the table instead of
  being random.
- `XlsxWriter` writes `date` and `datetime` cells without a format as
  `yyyy-mm-dd` and `yyyy-mm-dd hh:mm`, like `XlsbWriter`. It used Excel's
  built-in formats 14 and 22, which change with the reader's language
  (`m/d/yyyy` in English Excel, `dd.mm.yyyy` in Polish), so the same data
  looked different in XLSX and XLSB. Every XLSX `styles.xml` now has these
  two number formats, and user formats are numbered from 166.
- Chart caches: an empty cell is a gap in the series (no point) instead of a
  zero, as Excel stores it.

- Faster styled writes, with byte-identical output (measured with the new
  `tools/bench_styles.py`, 10,000 rows × 12 columns):
  - `styles` ranges are found per row by a sweep instead of a scan over all
    ranges, so many ranges no longer cost rows × ranges. One range per cell
    (60,000 ranges) went from 21 s to 0.18 s, and the workbook model's
    per-cell styles from 21 s to 0.1 s.
  - `(value, format)` cells inside styled ranges reuse their merged style:
    2x faster in Python, 6x (XLSX) / 11x (XLSB) with the C extension.
  - The C row encoders now also write `(value, format)` and `(value, Style)`
    cells, styled cells, and rows with a height, hidden flag or outline level.
  - Header and range styles in XLSX use the C encoder: 2x faster.
  - A1 range parsing is about twice as fast.

- Automatic column widths follow what Excel shows. Measured against Excel's
  AutoFit on 87 sample cells, the old estimate was too narrow for 11 (XLSX) /
  8 (XLSB) of them (by up to 15 characters); the new one for none, at about
  one character of extra room on average:
  - per-character widths measured in Excel instead of 1.25 per character, so
    capitals, `W`/`M`/`@` and CJK text no longer get cut;
  - numbers and dates as formatted (`#,##0.00`, `0%`, `dddd, d mmmm yyyy`,
    General's 11 characters), so they no longer show `####`;
  - bold, font size and wrapped text from the cell's style, and two extra
    characters for a filter or table drop-down button on the header;
  - the 5-pixel cell padding Excel adds to a file's column width is now
    included, and XLSB no longer truncates widths to whole characters;
  - XLSX and XLSB share one algorithm. XLSB columns are no longer narrower
    than Excel's default, and columns after empty cells are measured too.
- `XlsxWriter` keeps 1 MiB of encoded rows in memory before spilling to disk
  (was 8 MiB, `spool_max_bytes`). Writing 50,000 × 50 cells with
  `useSharedStrings=False` takes the same time and 13 MB less memory (peak
  40 MB instead of 52 MB). With shared strings the peak is set by the string
  table and barely changes.

### Fixed

- The workbook model no longer makes first-row text bold when saving a new
  workbook (it inherited the writers' default header look).
- Table headers written as `(value, Style)` or `(value, format)` cells were
  rejected as "not text".
- Saving a loaded workbook kept the old height of a row whose height was
  cleared (or moved away by `insert_rows`).
- XLSX borders: a `Border` with a left or right edge was written with empty
  `<start/>`/`<end/>` elements, which Excel reads as those edges, so it showed
  no left or right border. `start`/`end` are now written only when set.
- Chart legends and titles overlapped the plot area: Excel draws them over the
  plot when the `<overlay>` element is missing. Legends and titles are now
  written with `overlay val="0"`, so the legend sits beside the plot and the
  title above it. This affects XLSX and XLSB charts. The XLSX chart XML of
  default charts changes too (one `overlay` element per legend and title).
- XLSB files without a chart opened in Microsoft 365 in Aptos Narrow instead
  of Calibri: without a theme part Excel uses its current default theme. The
  theme is now always written, so XLSB files look like the XLSX ones (fonts
  and theme colours).
- Excel refused XLSB files whose chart series pointed at `Formula` cells: the
  chart cache held the formula text. It now holds the formula's cached
  `value` (or nothing).
- Pie charts dropped their categories, so the legend showed 1, 2, 3 instead of
  the slice names (XLSX and XLSB).
- XLSB line charts were drawn smoothed.
- A single-series XLSB chart without a title often lost Excel's automatic
  title (the series name) when another workbook was open in Excel. Charts
  without a title are now saved as Excel saves them: an empty automatic title
  for a single series and `autoTitleDeleted val="0"`, in both formats. This
  changes the chart XML of untitled XLSX charts.
- A conditional format fill given as `PatternFill("solid", start_color=...)`
  showed no colour in Excel, which paints a conditional solid fill with the
  background colour. The background colour is now set to the start colour
  when `end_color` is missing (XLSX and XLSB). This changes the XLSX
  `styles.xml` of such rules.
- XLSB icon sets now serialize threshold types and show/reverse flags according
  to the file format, and explicitly reject formula thresholds that need
  unsupported formula-token encoding.
- Loading a workbook now preserves its existing page margins, including when
  only one margin is changed. `rich_text=True` also retains rich runs in inline
  strings.
- Translating moved array formulas now moves their `array_ref` with the formula.

### Documentation

- `docs/usage.md`: what clicking a hyperlink does in Excel (an internal link
  jumps to its sheet and cell, the `Use CTRL+Click to follow hyperlink` option,
  and how to select a linked cell without following it), which readers return
  hyperlinks, and the XLSB hyperlink support. The feature support table and the
  `HYPERLINK` formula section now point to it.

## [0.8.1] - 2026-10-21

### Fixed

- `load_workbook(..., read_only=True)` no longer truncates XLSB files written by
  `XlsbWriter`: their A1:A1 dimension placeholder is treated as unsized.
  XLSB `update_cells()` now includes existing cells in the stored dimension.

## [0.8.0] - 2026-10-21

### Added

- Classic XLSX/XLSB cell notes: `Comment`, `comments=` on sheet/DataFrame exports,
  `ExcelReader.get_comments()` and updater `get_comments()`/`set_comments()`.
  Updates preserve unrelated note formatting and VML controls; unused note parts
  are removed when the last note is deleted. XLSM macros remain opaque.
- XLSX/XLSM/XLSB updater `update_cells()` for sparse value/formula changes preserving
  existing styles, notes and untouched cell records. Formula edits request Excel
  recalculation and remove the obsolete calculation chain.
- `DefinedName` and writer/updater methods to set, list and remove global or
  worksheet-local names; the reader can list user-defined names.
- `SheetProtection` and `WorkbookProtection` for sheet selection restrictions
  and workbook structure protection, including optional legacy password hashes.
- XLSX/XLSM/XLSB formula text reading with `get_rows(..., data_only=False)` and
  `read_all(..., data_only=False)`, including shared and legacy array formulas.
- XLSB formula compilation, cached results and legacy array formulas, including
  sheet/name references and supported built-in/newer functions. Unsupported
  grammar and tokens raise explicit errors. Optional incremental binary reading
  uses `streaming=True`; ordinary data retains the C/Python bulk paths.
- Binary comment, protection and defined-name editing, preserving unrelated
  records and VML controls. Deleted name slots remain hidden to preserve binary
  formula indexes. Sparse edits invalidate binary indexes and calculation chains.
- Excel COM interoperability checks against both generated XLSB files and
  independently Excel-created binary templates.
- An openpyxl-like workbook model: `Workbook`, `load_workbook()`, `Worksheet`,
  `Cell` and related dimension/view classes for random cell access, styles,
  merges, notes and patching edits back into loaded XLSX/XLSM packages. Read-only
  mode follows openpyxl's semantics (worksheet coordinates, empty rows for
  missing rows, stored dimensions, `ReadOnlyCell` objects) for XLSX and XLSB.
- Strict Open XML workbooks can be read with `ExcelReader` and read-only
  `load_workbook()`; editing them raises `NotImplementedError`.

- Row and column groups (Excel's outline `+`/`−` buttons) in `XlsxWriter` and
  `XlsbWriter`: `row_groups` and `column_groups` take `Group(first, last,
  collapsed=False)` ranges on `write_sheet`, `write_dataframe`, `to_xlsx` and
  `to_xlsb`. Nesting sets the level (up to seven), summaries go below and to
  the right, and collapsed groups are hidden. Sheets without groups are written
  as before.
- XLSX formulas: `Formula(text, value=..., array_ref=...)` writes formula
  cells with optional cached results and legacy array formulas. Post-2007
  functions get the `_xlfn.` prefix the file format requires; `LET` and the
  `LAMBDA` family are refused.
- XLSX data validation: `DataValidation` for list (literal or range source),
  whole, decimal, date, time, text length and custom rules, with operators,
  prompts, error messages and dropdown visibility.
- XLSX hyperlinks: `Hyperlink` as a cell value, for external web and mail
  addresses or `#Sheet!A1` locations, with text, tooltip and a built-in link style.
  Unsafe or unsupported schemes are refused.
- XLSX images: `Image` for PNG and JPEG files, placed at an anchor cell with
  pixel sizes read from the file when not given.
- XLSX print settings: `PrintSettings` and `PageMargins` for orientation, paper
  size, margins, print area, repeating rows and columns, fit-to-page or scale,
  centering, headers and footers, and manual page breaks.
- XLSX chart types and options: horizontal bar (`"barh"`) and scatter XY
  (`"scatter"`, with `x_values`) charts; `ChartAxis`, `ChartLegend`,
  `DataLabels`, `ChartMarker`, series colours, `title_size`, and the
  `x_axis`/`y_axis` settings. The original `Chart("bar")` remains a column chart.
- Bounded-memory streaming. XLSX `write_sheet` consumes its input once and
  encodes rows into a spool (in memory up to `spool_max_bytes`, then on disk).
  `XlsbWriter` gains `useSharedStrings=False` (inline text, so no shared-string
  table grows) and chart caches now keep only the cells the charts read.
- `XlsbWriter` rejects features it cannot write (formulas, hyperlinks, and the
  new chart options) with an explicit error rather than writing them as text.
- `xlspy.refs`, `xlspy.charts`, `xlspy.formulas`, `xlspy.data_validation`,
  `xlspy.hyperlinks`, `xlspy.images` and `xlspy.printing` modules; the new
  public classes are exported from `xlspy`.

- XLSB inline text (`useSharedStrings=False`) is encoded by the C extension
  (`encode_xlsb_row_inline`). Rows it cannot encode fall back to the Python
  encoder.
- `examples/write_benchmark.py`: a reproducible write benchmark (isolated
  worker processes, medians, peak memory, optional tracemalloc, JSON and CSV
  output with environment metadata) for the xlspy writers with and without the
  C extension, openpyxl and xlsxwriter.

### Changed

- `sharedStrings.xml` is written in chunks of 4,096 entries instead of one
  growing string. The bytes are unchanged. On Windows, 50,000 × 50 XLSX went
  from 3.73 s to 1.98 s, and 200,000 × 50 from 41.12 s / 327 MB to 13.48 s /
  260 MB peak.
- Cell text longer than Excel's limit of 32,767 UTF-16 code units raises
  `ValueError` in every writer path: XLSX and XLSB, shared and inline strings,
  the C and Python encoders, styled cells, DataFrames, hyperlink text and cached
  formula results. Previously XLSX wrote such text unchanged, which Excel 16.0
  refuses to open, and shared-string XLSB silently cut it to 32,767 units.
- The C row encoders hand any row holding an object of another type (such as a
  `Formula` or `Hyperlink`) to the Python encoder, instead of writing its
  `str()`. This replaces a per-row type scan in both writers. The C XLSX encoder
  also copies text with no characters to escape in runs instead of byte by byte.
  Together these made 200,000 × 50 writes 5–12% faster on Windows (XLSX shared
  13.69 → 12.04 s, XLSX inline 11.56 → 10.79 s, XLSB shared 5.44 → 5.17 s, XLSB
  inline 4.76 → 4.40 s).
- `XlsxWriter` and `XlsbWriter` share one implementation of the open, save and
  close lifecycle (`xlspy._output.StagedWriter`).
- XLSB chart caches keep one list per referenced column instead of a dict keyed
  by cell. On 200,000 × 3 with charts, the Python allocation peak fell from
  119.5 MB to 105.7 MB.
- A workbook is published only after a complete, successful save. `XlsxWriter`
  and `XlsbWriter` build the archive in a staging file: a temporary file next to
  a path target, moved into place with `os.replace`, or a spool (memory, then
  disk) copied into a file-object target. If `write_sheet` raises, saving fails,
  or the `with` block exits with an exception, the target is left as it was: no
  new file, an existing file unchanged, a stream untouched. Previously the
  partial workbook was saved on exit, and a stream received the written members
  and a ZIP end record.
- After a failed write, `write_sheet` and `save()` raise `RuntimeError`; `save()`
  also refuses to run twice. The file is published when the writer closes, so a
  writer closed without a successful `save()` publishes nothing.
- Chart and cell-reference classes moved to `xlspy.charts` and `xlspy.refs`.
  `xlspy.styles` still exports `Chart` and `ChartSeries`.
- The `_FilterDatabase` defined name quotes its sheet name, so sheets with
  spaces get a valid filter range.
- Explicit `column_widths` for columns beyond the data are written.

### Fixed

- XLSB `get_rows(..., streaming=True)` and `data_only=False` now return the same
  row layout as the default bulk reader (missing rows, first column and widths).
- Adding classic notes no longer writes into a header/footer picture VML part,
  and deleting the last note keeps that part's relationship.
- XLSB formulas may refer to a sheet-qualified local name defined later, legacy
  array formulas accept date values in member cells, and array rows past the data
  keep row heights, hidden rows and outline levels.
- Saving a loaded workbook after removing or reordering sheets removes the sheet
  relationship, remaps worksheet-local names and drops the calculation chain.
  Rewritten `workbook.xml` and `styles.xml` keep Excel's namespace prefixes; before,
  Excel refused workbooks it had created once they were edited and saved.
- Style edits of cells without an `r` attribute no longer clear their values.
- `load_workbook()` maps worksheet-local names correctly when chartsheets exist, and
  saving keeps chartsheet positions and the active tab.
- An invalid `XlsxWriter` `auto_filter_ref` is rejected before rows are written.

- XLSX row reading now parses incrementally and releases completed rows instead
  of loading the entire worksheet before yielding. Rows wider than the first
  are no longer truncated, and inline rich text retains all displayed fragments.
- Workbook and relationship parsing accepts XML namespace prefixes introduced
  by other spreadsheet tools and metadata updates.

- `write_sheet` no longer materializes the whole sheet in memory (XLSX and XLSB).
- XLSB inline text counted code points instead of UTF-16 code units, so text
  containing characters outside the Basic Multilingual Plane (such as emoji)
  produced a record whose character count did not match its payload.
- XLSB inline text containing a lone surrogate was written as invalid UTF-16. It
  is now refused with `UnicodeEncodeError`, as the shared-string path already did.
- The C extension sized XLSB column widths from UTF-8 bytes instead of
  characters, so non-ASCII headers got different widths from the Python writer.
- XLSX text with control characters (such as `"\x07"` or `"\x00"`), carriage
  returns, or U+FFFE/U+FFFF produced ill-formed XML in all four string paths. Such
  characters are now written as OOXML `_xHHHH_` escapes, literal `_xHHHH_` text
  is protected with `_x005F_`, and the XLSX reader decodes both. Excel 16.0 reads
  the original text back.
- The C XLSB encoder dropped a cell whose `str()` raised and wrote the row without
  it; the error now reaches the caller, as in the Python encoder. When the C
  shared-string XLSB encoder hands a row back to Python, the strings it had
  already added to the table are removed again, so string indexes stay in step.
- Undefined behaviour found with UBSan: the C XLSB encoder left-shifted negative
  integers when writing RK numbers, and the C reader built 32-bit values by
  shifting into a 32-bit `long` (Windows). Both now use unsigned arithmetic; the
  bytes written and the values read are unchanged.
- Tables (XLSX and XLSB) without explicit `columns` were written with the
  names `Column1`, `Column2`, … while the header cells held other text; Excel
  16.0 refuses to open such a workbook. Names now come from the header row
  (captured while the rows stream, at any depth). Header rows Excel would refuse
  (an empty or non-text header cell, explicit `columns` that differ from the
  cells other than in case, duplicate names) raise `ValueError` and nothing is
  published. Table names that read as cell references (`T1`, `R1C1`, `C`) are
  refused when the `ExcelTable` is created. This also affected the
  `xlsx_report_features.py` example and was present in 0.7.0.
- An error raised while the C shared-string encoders looked up a string (for
  example by a `str` subclass whose `__hash__` raises) was printed as "Exception
  ignored" and then reported as `MemoryError`. The original error now reaches
  the caller, from both the XLSX and the XLSB encoder.
- Row groups (XLSX and XLSB) set outline state only on rows that held data, so
  a collapsed group ending on the last data row lost the `collapsed` flag on its
  summary row, and a group reaching past the data lost its level there. These
  rows are now written empty with their level, hidden and collapsed flags. A
  group ending on the sheet's last row or column gets no summary past the sheet.
- `Formula(value=...)` accepted `inf`, `-inf` and non-finite `Decimal` cached
  values and wrote `<v>inf</v>`, which is not a valid number in OOXML. Only
  finite cached numbers are accepted now.
- An XLSX workbook whose drawings held only images contained `theme1.xml`
  without a workbook relationship to it. The theme is now linked whenever it
  is written.
- List data validations refused any unquoted source containing a comma, which
  ruled out dynamic lists such as `OFFSET($B$1,0,0,5,1)`. Only commas outside
  function arguments, strings and quoted sheet names are refused now.
- Internal hyperlinks were checked against sheet names after stripping every
  apostrophe and comparing case-sensitively, so `#'''x'!A1` (sheet `'x`) and
  `#sheet1!A1` (sheet `Sheet1`) were refused. Only the enclosing quotes are
  removed, and names are compared ignoring case, as Excel does.
- A `write_sheet` call rejected by validation could leave state behind: XLSB
  kept the sheet's charts, table ids and conditional-format priorities, and both
  writers kept table names, so a retry produced orphaned parts or refused a
  valid name. Every option is now validated before the writer records anything.
- Table header rows are checked as soon as they stream past instead of after
  the whole sheet has been encoded, so a bad header fails before millions of
  rows are read. A `Hyperlink` header cell names its column by its text.
- `hidden_columns` with a negative index was accepted by `XlsxWriter`. Both
  writers now share one validation for hidden rows and columns.
- The XLSB readers (C and Python) padded the gap before a row record with
  empty rows up to whatever index it held, so a corrupt index such as
  530,579,456 produced hundreds of millions of rows and never finished (this
  hung `tests/test_c_fuzz.py`). A row index past 1,048,575 now raises
  `ValueError`; row `0xFFFFFFFF` still reads as -1 on every platform.
- `add_sheet` accepted sheet names Excel refuses to open: empty or longer than
  31 characters, containing `: \ / ? * [ ]`, starting or ending with an
  apostrophe, the reserved `History`, or repeating another name ignoring case.
  Both writers now raise `ValueError` (`TypeError` for a non-string).
- `Formula` named `bool` as an allowed cached value in the error raised for a
  `bool`; the message now says a bool is refused.

## [0.7.0] - 2026-10-18

### Added

- `tests/test_c_memleaks.py`: RSS / tracemalloc / GC-population regression
  tests over all C entry points, including error paths (note: refcount
  assertions are avoided -- `sys.getrefcount` reads one less than expected
  for GC-tracked objects under pytest on Python 3.14).
- `tests/test_c_read_correctness.py`: C-vs-Python XLSB reader regressions for
  left-ragged cells, unordered first-row columns, short/overrunning records,
  1904 date overflow parity, and wide-layout fallback.
- XLSB tables (`ExcelTable`) and conditional formatting for color scales,
  data bars, and icon sets; formula and cell-is rules remain XLSX-only.
- Structured cell styling for XLSX: `Font`, `PatternFill`, `GradientFill`
  (linear and path, with `Stop` colors/positions), `Border`/`Side`,
  `Alignment`, and `Protection`, usable via `Style(font=..., fill=...,
  border=..., alignment=..., protection=...)` alongside the existing flat
  properties.
- Differential styles (`DifferentialStyle`) as the `dxf` of conditional
  formatting rules.
- `IconSetRule` conditional formatting.
- XLSB cell protection: `Protection(locked=..., hidden=...)` is encoded in the
  BrtXF record (`fLocked`/`fHidden` bits plus the `applyProtection` flag).
- XLSB text alignment: `Alignment` (`horizontal`, `vertical`, `wrap_text`,
  `shrink_to_fit`, `indent`, `text_rotation`, `justify_last_line`,
  `reading_order`) is encoded in the BrtXF record with the `applyAlignment`
  flag. Rotation follows the OOXML clockwise convention; BIFF12 stores it
  counter-clockwise.
- XLSB borders: structured `Border`/`Side` styles are emitted as BrtBorder
  records (top/bottom/left/right/diagonal, diagonal direction flags) and linked
  from BrtXF via `ixBorder` with the `applyBorder` flag.
- XLSB structured fonts and fills: `Font` (name, size, bold, italic, strike,
  underline incl. accounting variants, super/subscript, color, charset, family,
  scheme, outline/shadow/condense/extend) is emitted as BrtFont records;
  `PatternFill` (all 19 pattern types) and linear/path `GradientFill` (with
  per-stop colors and positions) are emitted as BrtFill records, linked from
  BrtXF with the `applyFont`/`applyFill` flags.
- `XlsbWriter.write_sheet` now accepts `merges`, `column_widths`,
  `hidden_columns`, `row_heights`, `hidden_rows`, `freeze_panes`, and
  `auto_filter`, with the same semantics and validation as `XlsxWriter`.
  Encoded as MERGECELLS, BrtColInfo, BrtRowHdr, BrtPane/BrtSel records;
  `freeze_panes=""` removes the default frozen top row, `auto_filter=False`
  omits the autofilter. `write_dataframe` and `dataframe.to_xlsb` pass the new
  arguments through.
- Charts for XLSX: `Chart`/`ChartSeries` (`bar`/`line`/`pie`, title, legend
  labels from a literal or a cell, categories, anchor cell, size in cm, style)
  via `write_sheet(..., charts=[...])`, with drawing and chart parts,
  relationships, and content types.
- Charts for XLSB: the same `Chart`/`ChartSeries` model via
  `write_sheet(..., charts=[...])`, `write_dataframe(..., charts=...)` and
  `dataframe.to_xlsb(..., charts=...)` on `XlsbWriter`. Sheets link to the
  DrawingML parts with a `BrtDrawing` record ([MS-XLSB] 2.4.351); the chart and
  drawing XML is shared with the XLSX writer. Like Excel itself, XLSB charts
  carry cached values (`strCache`/`numCache`), text categories referenced via
  `strRef`, and an explicit `varyColors` flag.
- Chart workbooks bundle the standard Office theme (`xl/theme/theme1.xml`,
  with a workbook relationship and content-type override) so both backends
  resolve the same default chart styles and accent colors.
- Demo workbook generator (`tools/make_styling_demo.py`, XLSX + XLSB) with a
  per-sheet checklist (`tmp/xlspy-test-files/README.md`) and a Windows Excel
  COM validation script (`tools/validate_excel_com.ps1`).

### Changed

- Faster XLSB writes: shared-strings table serialization and the
  formatted-row pre-scan moved to the C extension (50k×50 write
  1.92 s → 1.43 s with byte-identical output).
- Faster XLSX writes: the plain (unstyled) row encoder moved to the C
  extension, with byte-identical output including shared-strings dedup,
  float formatting, datetimes, and XML escaping (50k×50 write
  5.18 s → 4.12 s; rows with `(value, format)` tuples still use the
  Python path, which owns format registration).
- `xl/styles.bin` is now built section by section (Fmts, Fonts, Fills, Borders,
  CellStyleXFs, CellXFs) with correct record framing instead of patching byte
  offsets in a template; multi-byte VLQ record headers are handled properly.
- README documents the XLSB style parity, charts for both formats, and the
  remaining format limitations.

### Fixed

- XLSB writer C extension now matches the pure-Python reference exactly:
  integers beyond the int64 range are stored as doubles (were strings),
  `int`/`float`/`str` subclasses take the numeric/string path, and column
  autofit unwraps `(value, format)` tuples and measures `Decimal` as
  `float` (were `str(tuple)` / `str(Decimal)`).
- Fixed a latent heap buffer overrun in the C row encoder: the per-cell
  buffer reservation (16 bytes) was smaller than the largest cell record
  (`BrtCellReal`, 18 bytes).
- Fixed a per-call reference leak in the C XLSB row encoder: the result
  bytes leaked one reference per encoded row.
- XLSB reader now parses the shared-strings table in the C extension with
  behavior identical to the pure-Python reference, quirks included
  (50k×50 read 1.32 s → 0.74 s).
- C XLSB worksheet reader now matches the reference on first-row column
  ordering, left-of-layout cells, short/overrunning cell records, and date
  conversion errors. Unusually wide layouts delegate to the Python reader
  instead of using a C-specific memory limit. Hostile SST lengths no longer
  produce a `SystemError` from a negative Unicode allocation size.
- A font that names its typeface no longer claims a theme font scheme, so
  Excel renders the requested typeface instead of substituting the theme font.
- Conditional formatting priorities no longer collide between rules that
  pre-reserve a priority and rules with automatically assigned priorities.
- Freeze panes in XLSB now mirror the template's pane flags (`0x03`) and active
  pane instead of the spec-pure `0x02`: Excel hid the sheet tabs and the
  row/column headers with `0x02`.
- Worksheet record sequence is now well-formed: the sheet-data section
  contained a duplicated `BrtBeginSheetData` record with a stray `BrtACEnd`,
  which Excel tolerated but which violates the Worksheet ABNF.
- XLSX sheets with charts now carry the `<drawing r:id="..."/>` element
  (after `<pageMargins>`, before `<tableParts>`), matching openpyxl output —
  without it Excel would not display the charts. The element was missing
  because openpyxl resolves charts through the sheet rels, so round-trip
  tests passed anyway. Plain files stay byte-clean: no drawing parts, no theme
  and no relationship changes unless charts are used.
- XLSB charts no longer render transposed: Excel's XLSB loader apparently
  rebuilds chart series from the sheet when `varyColors` is absent and text
  categories are referenced through `numRef`, which split a single-series line
  over a vertical range into one series per row (legend showing the category
  names). The rendered result now matches the XLSX output.
- XLSB chart axes are now explicit (`delete="0"`, `axPos`, `tickLblPos`), like
  real Excel-generated XLSB files. Without those attributes Excel's XLSB
  loader did not draw the category/value axis tick labels (Q1–Q4, 0–180),
  although the XLSX loader applies the defaults and showed them.
- XLSB charts now spell numeric styles out through the modern `c14:style`
  extension (`mc:AlternateContent` with `c14:style` = style + 100 and the
  legacy element as fallback, [MS-ODRAWXML] 2.7.3.5), like real
  Excel-generated XLSB files. Excel's XLSB loader did not apply the bare
  legacy `style` element the way the XLSX loader does, so a styled chart
  (e.g. `style=10`) rendered with the default style in XLSB while XLSX
  showed the styled one. Charts without an explicit style are unchanged.
- XLSB charts now pin the classic Office 2007 accent palette (`4F81BD`,
  `C0504D`, `9BBB59`, …) and the legacy 18 pt bold title, so they render
  like the XLSX output. The XLSB loader only knows the modern style engine:
  it painted series with the theme accents (Office 2013 palette) and used
  14 pt titles, while the XLSX loader draws legacy styles with their
  built-in classic colours and larger titles.

## 0.1.0 – 0.6.0 (2026-05-25 – 2026-10-16)

Releases before 0.7.0 have no changelog entries; their history is in the git log.

[Unreleased]: https://github.com/justybase/xlspy/compare/v0.13.0...HEAD
[0.13.0]: https://github.com/justybase/xlspy/compare/v0.12.0...v0.13.0
[0.12.0]: https://github.com/justybase/xlspy/compare/v0.11.1...v0.12.0
[0.11.1]: https://github.com/justybase/xlspy/compare/v0.11.0...v0.11.1
[0.11.0]: https://github.com/justybase/xlspy/compare/v0.10.1...v0.11.0
[0.10.1]: https://github.com/justybase/xlspy/compare/v0.10.0...v0.10.1
[0.10.0]: https://github.com/justybase/xlspy/compare/96e1efd47e29882578e74965014a3ea0fdf3a759...v0.10.0
[0.9.2]: https://github.com/justybase/xlspy/compare/v0.9.1...96e1efd47e29882578e74965014a3ea0fdf3a759
[0.9.1]: https://github.com/justybase/xlspy/compare/v0.9.0...v0.9.1
[0.9.0]: https://github.com/justybase/xlspy/compare/v0.8.1...v0.9.0
[0.8.1]: https://github.com/justybase/xlspy/compare/v0.8.0...v0.8.1
[0.8.0]: https://github.com/justybase/xlspy/compare/v0.7.0...v0.8.0
[0.7.0]: https://github.com/justybase/xlspy/releases/tag/v0.7.0
