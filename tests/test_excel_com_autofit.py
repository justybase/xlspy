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
from xlspy import Font, Style, XlsxWriter
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
