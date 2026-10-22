# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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

### Changed

- Faster XLSB writes: shared-strings table serialization and the
  formatted-row pre-scan moved to the C extension (50k×50 write
  1.92 s → 1.43 s with byte-identical output).
- Faster XLSX writes: the plain (unstyled) row encoder moved to the C
  extension, with byte-identical output including shared-strings dedup,
  float formatting, datetimes, and XML escaping (50k×50 write
  5.18 s → 4.12 s; rows with `(value, format)` tuples still use the
  Python path, which owns format registration).

### Added

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

- `xl/styles.bin` is now built section by section (Fmts, Fonts, Fills, Borders,
  CellStyleXFs, CellXFs) with correct record framing instead of patching byte
  offsets in a template; multi-byte VLQ record headers are handled properly.
- README documents the XLSB style parity, charts for both formats, and the
  remaining format limitations.

### Fixed

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
