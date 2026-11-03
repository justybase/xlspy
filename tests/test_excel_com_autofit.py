"""Excel as the oracle for the automatic column widths: xlspy's estimate is
Excel's AutoFit plus one spare pixel, for text and numbers in any font size;
dates take the widest locale rendering, so they only must not be narrower.
Column widths are defined at 96 DPI, and Excel measures AutoFit at the
display's scaling, so this runs only on a display at 100%."""

import ctypes
import datetime
import sys

import pytest

from test_column_widths import EXCEL_AUTOFIT_100, EXCEL_AUTOFIT_SAMPLES, SPARE
from xlspy import Alignment, Font, Group, Style, XlsbWriter, XlsxWriter
from xlspy._widths import cell_width
from xlspy.refs import column_label

pytestmark = pytest.mark.com  # skipped unless XLSPY_RUN_EXCEL_COM=1 (conftest.py)


def _system_dpi():
    if sys.platform != "win32":
        return None
    user32 = ctypes.windll.user32
    return user32.GetDpiForSystem() if hasattr(user32, "GetDpiForSystem") else None


def test_widths_match_excel_autofit(tmp_path, excel, open_workbook):
    if _system_dpi() != 96:
        pytest.skip("Excel's AutoFit depends on the display scaling; needs 100% (96 DPI)")
    cells = [(text, None) for text, _r, _b in EXCEL_AUTOFIT_100]
    cells += [(text, Style(bold=True)) for text, _r, _b in EXCEL_AUTOFIT_100]
    for value, number_format, size, bold, _width in EXCEL_AUTOFIT_SAMPLES:
        cells.append((value, Style(number_format=number_format,
                                   font=Font(name="Calibri", size=size or 11, bold=bold))))
    cells += [
        (1234567.891, Style(number_format="#,##0.00")), (-98765.4321, None), (0.4567, Style(number_format="0.00%")),
        (49062358930.42197, None), (-9.623203e-05, None), (True, None), ("Środa, 30 października", None),
        (datetime.date(2024, 9, 30), None), (datetime.datetime(2024, 12, 31, 23, 45), None),
    ]
    path = tmp_path / "autofit.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Probe")
        row = [value if style is None else (value, style) for value, style in cells]
        # Row 1 stays empty: the writer bolds it.
        writer.write_sheet([[None] * len(row), row], auto_filter=False,
                           column_widths=dict.fromkeys(range(len(row)), 8.43))
    sheet = open_workbook(path).Worksheets(1)
    sheet.Range(f"A1:{column_label(len(cells) - 1)}2").Columns.AutoFit()
    for i, (value, style) in enumerate(cells):
        excel_width = sheet.Columns(i + 1).ColumnWidth
        font = style.font if style is not None else None
        ours = cell_width(value, style.number_format if style is not None else None,
                          bold=bool(style is not None and (style.bold or (font is not None and font.bold))),
                          size=font.size if font is not None else None)
        shown = sheet.Cells(2, i + 1).Text
        if isinstance(value, (datetime.date, bool)):
            assert ours >= excel_width, (value, ours, excel_width, shown)
        else:
            assert ours == pytest.approx(excel_width + SPARE, abs=0.006), (value, ours, excel_width, shown)


def test_filter_button_widths_match_excel_autofit(tmp_path, excel, open_workbook):
    """Headers under an autofilter button: Excel's AutoFit plus one spare pixel."""
    if _system_dpi() != 96:
        pytest.skip("Excel's AutoFit depends on the display scaling; needs 100% (96 DPI)")
    headers = ["COLUMN_NAME", "WWWWWWWW", "Customer name", "Kwota brutto", "Zażółć gęślą jaźń", "ID_0123456789"]
    path = tmp_path / "filtered.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Probe")
        writer.write_sheet([headers, [1] * len(headers), [2] * len(headers), [3] * len(headers)], auto_filter=True)
    sheet = open_workbook(path).Worksheets(1)
    ours = [sheet.Columns(i + 1).ColumnWidth for i in range(len(headers))]
    assert sheet.AutoFilterMode
    sheet.Range(f"A1:{column_label(len(headers) - 1)}4").Columns.AutoFit()
    for i, header in enumerate(headers):
        # Excel reports both widths rounded to 0.01: compare whole pixels.
        assert round(ours[i] * 7) == round(sheet.Columns(i + 1).ColumnWidth * 7) + 1, header


@pytest.mark.parametrize("writer_cls, ext", [(XlsxWriter, "xlsx"), (XlsbWriter, "xlsb")])
def test_indented_outline_widths_match_excel_autofit(tmp_path, excel, open_workbook, writer_cls, ext):
    """Three nested row groups whose labels step in by an alignment indent (or
    tabs) per level: Excel draws the indent as blank room before the text,
    and AutoFit makes the column that much wider."""
    if _system_dpi() != 96:
        pytest.skip("Excel's AutoFit depends on the display scaling; needs 100% (96 DPI)")
    labels = [("Total", 0), ("REGION 01 - Lidia Badach", 1), ("Okręg Mazowsze Północ", 2),
              ("0577 Warszawa, ul. Gen. Bora Komorowskiego Tadeusza 37", 3), ("0889 Warszawa", 3)]
    styles = [Style(alignment=Alignment(indent=level)) for level in range(4)]
    centred = [Style(alignment=Alignment(horizontal="center", indent=level), font=Font(size=14))
               for level in range(4)]
    rows = [[None] * 4]
    for text, level in labels:
        rows.append([(text, styles[level]), "\t" * level + text, (text, centred[level]),
                     (12345.5 * level, Style(number_format="#,##0.00", alignment=Alignment(indent=level)))])
    path = tmp_path / f"outline.{ext}"
    with writer_cls(path) as writer:
        writer.add_sheet("Probe")
        writer.write_sheet(rows, auto_filter=False, row_groups=[Group(1, 5), Group(2, 5), Group(3, 5)])
    sheet = open_workbook(path).Worksheets(1)
    ours = [sheet.Columns(i + 1).ColumnWidth for i in range(4)]
    sheet.Range("A1:D6").Columns.AutoFit()
    for i in range(4):
        assert round(ours[i] * 7) == round(sheet.Columns(i + 1).ColumnWidth * 7) + 1, i


def _autofit_pixels_of(sheet, count):
    """(widths before, widths after Excel's AutoFit) of the first ``count`` columns, in pixels."""
    before = [round(sheet.Columns(i + 1).Width * 4 / 3) for i in range(count)]
    sheet.Rows(1).ClearContents()
    sheet.Range(sheet.Cells(1, 1), sheet.Cells(2, count)).Columns.AutoFit()
    return before, [round(sheet.Columns(i + 1).Width * 4 / 3) for i in range(count)]


def test_font_widths_match_excel_autofit(tmp_path, excel, open_workbook):
    """Every font with metrics, in four styles: Excel's AutoFit plus the spare
    pixel (one more where italic rounding differs); other fonts, rich text and
    characters from fallback fonts are never narrower than Excel's."""
    if _system_dpi() != 96:
        pytest.skip("Excel's AutoFit depends on the display scaling; needs 100% (96 DPI)")
    from xlspy._fonts import _metrics
    from xlspy.rich_text import CellRichText, InlineFont, TextBlock

    text = "Kraków, al. Jana Pawła II 12/3"
    exact = [(text, Style(font=Font(name=name, size=size, bold=bold, italic=italic)))
             for name, _packed in _metrics.FAMILIES.values()
             for bold in (False, True) for italic in (False, True) for size in (9, 11, 16)]
    at_least = [(text, Style(font=Font(name=name, bold=bold)))
                for name in ("Book Antiqua", "Garamond", "Century", "Bookman Old Style", "Cascadia Mono")
                for bold in (False, True)]
    at_least += [
        CellRichText("ab ", TextBlock(InlineFont(sz=20), text)),
        CellRichText("Kwota: ", TextBlock(InlineFont(b=True, sz=14, rFont="Verdana"), "1 234 567,89 zł")),
        (CellRichText(TextBlock(InlineFont(b=True), text)), Style(font=Font(size=16))),
        (CellRichText("plain ", TextBlock(InlineFont(i=True), text)), Style(font=Font(name="Georgia", italic=True))),
        "  " + text, "😀👍 " + text, "➜ ✈ ☎ ↔ ⇔ ═║", "∑∫√∞≈≠≤≥ № ℃ ⅓", "soft\xadhyphen" * 3,
    ]
    path = tmp_path / "fonts.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Probe")
        writer.write_sheet([[None] * (len(exact) + len(at_least)), exact + at_least], auto_filter=False)
    sheet = open_workbook(path).Worksheets(1)
    ours, theirs = _autofit_pixels_of(sheet, len(exact) + len(at_least))
    for i, cell in enumerate(exact):
        assert 1 <= ours[i] - theirs[i] <= 2, (cell, ours[i], theirs[i])
    for i, cell in enumerate(at_least, len(exact)):
        assert ours[i] >= theirs[i] or theirs[i] <= 64, (cell, ours[i], theirs[i])
