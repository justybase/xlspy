# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.7.0] - 2026-10-18

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
