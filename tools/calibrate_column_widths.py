"""Measure Excel's AutoFit widths (COM, Windows + Excel) for the width estimator.

Writes probe values with xlspy (so the workbook has xlspy's default font,
Calibri 11), lets Excel AutoFit every column and records ``ColumnWidth``.

    python tools/calibrate_column_widths.py chars  out.json
        Per-character widths: each printable ASCII character and a few
        non-ASCII ones repeated 10 and 40 times; slope = width per character,
        intercept = padding. Regular and bold.
    python tools/calibrate_column_widths.py corpus out.json
        Realistic cells (words, sentences, codes, numbers with formats, dates)
        with their AutoFit width, to score estimators against.

``tools/score_column_widths.py`` compares an estimator with a corpus file.
"""

import datetime
import json
import os
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.normpath(os.path.join(HERE, "..", "src")))

from xlspy import Style, XlsxWriter  # noqa: E402
from xlspy.refs import column_label  # noqa: E402

EXTRA_CHARS = "ąćęłńóśźżĄĆĘŁŃÓŚŹŻéüßÉÜЖжΩ中日한€£°§±×–—…„”"


def _autofit(cells):
    """``cells``: list of (value, Style | None) -> AutoFit width per cell."""
    import win32com.client
    path = os.path.join(tempfile.mkdtemp(prefix="xlspy-width-"), "probe.xlsx")
    row = [value if style is None else (value, style) for value, style in cells]
    with XlsxWriter(path) as writer:
        writer.add_sheet("Probe")
        # A blank first row: the writer bolds row 1 and filters it.
        writer.write_sheet([[None] * len(row), row], auto_filter=False,
                           column_widths={i: 8.43 for i in range(len(row))})
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    try:
        book = app.Workbooks.Open(path)
        sheet = book.Worksheets(1)
        last = column_label(len(row) - 1)
        sheet.Range(f"A1:{last}2").Columns.AutoFit()
        widths = [sheet.Columns(i + 1).ColumnWidth for i in range(len(row))]
        texts = [sheet.Cells(2, i + 1).Text for i in range(len(row))]
        book.Close(False)
    finally:
        app.Quit()
    return widths, texts


def measure_chars():
    chars = [chr(c) for c in range(33, 127)] + [" "] + list(EXTRA_CHARS)
    bold = Style(bold=True)
    cells = []
    for style in (None, bold):
        for char in chars:
            # A space alone is trimmed by nothing, but keep it inside letters.
            text10 = ("x" + char * 10 + "x") if char == " " else char * 10
            text40 = ("x" + char * 40 + "x") if char == " " else char * 40
            cells += [(text10, style), (text40, style)]
    widths, _texts = _autofit(cells)
    result = {"regular": {}, "bold": {}}
    i = 0
    for name in ("regular", "bold"):
        for char in chars:
            w10, w40 = widths[i], widths[i + 1]
            slope = (w40 - w10) / 30
            result[name][char] = {"slope": round(slope, 4), "w10": w10, "w40": w40}
            i += 2
    return result


def corpus_cells():
    """Realistic cells: (value, Style | None, label)."""
    words = [
        "Product", "Quantity", "Comment", "Customer name", "ID", "Total amount",
        "WWWWWW", "Maximum", "illicit", "Łódź", "Zażółć gęślą jaźń",
        "The quick brown fox jumps over the lazy dog",
        "ORDER-2024-000123", "john.smith@example.com", "https://example.com/a/b?c=1",
        "Wrocław, ul. Świdnicka 12/4", "中文字符测试", "MIXED Case Header",
        "A", "mmmmmmmmmm", "iiiiiiiiii", "1234567890", "first line\nsecond longer line here",
        "Tab\tseparated", "   padded", "ÉCOLE ÉLÉMENTAIRE", "€ 1 234,50",
    ]
    numbers = [
        (0, None), (7, None), (42, None), (1234, None), (123456789, None),
        (1234567890123, None), (3.14159, None), (0.1 + 0.2, None), (1 / 3, None),
        (-98765.4321, None), (1e-7, None), (2.5e15, None),
        (1234567.891, "#,##0.00"), (1234567.891, "#,##0"), (0.4567, "0%"),
        (0.4567, "0.00%"), (1234.5, '#,##0.00 "zł"'), (1234.5, "#,##0.00 €"),
        (12345, "000000000"), (1234.5678, "0.00E+00"), (-1234.5, "#,##0.00;[Red]-#,##0.00"),
        (True, None), (False, None),
    ]
    dates = [
        (datetime.date(2024, 9, 30), None), (datetime.datetime(2024, 9, 30, 13, 45, 12), None),
        (datetime.date(2024, 9, 30), "dd.mm.yyyy"), (datetime.date(2024, 9, 30), "d mmmm yyyy"),
        (datetime.date(2024, 9, 30), "dddd, d mmmm yyyy"), (datetime.date(2024, 9, 30), "yyyy-mm-dd"),
        (datetime.date(2024, 9, 30), "mmmm yyyy"), (datetime.datetime(2024, 9, 30, 13, 45, 12), "dd.mm.yyyy hh:mm"),
        (datetime.datetime(2024, 9, 30, 13, 45, 12), "yyyy-mm-dd\"T\"hh:mm:ss"),
        (datetime.datetime(2024, 9, 30, 13, 45, 12), "h:mm AM/PM"),
    ]
    cells = []
    for text in words:
        cells.append((text, None, "text"))
        cells.append((text, Style(bold=True), "text-bold"))
    for value, fmt in numbers:
        cells.append((value, None if fmt is None else Style(number_format=fmt), "number"))
    for value, fmt in dates:
        cells.append((value, None if fmt is None else Style(number_format=fmt), "date"))
    return cells


def measure_corpus():
    cells = corpus_cells()
    widths, texts = _autofit([(value, style) for value, style, _ in cells])
    records = []
    for (value, style, kind), width, text in zip(cells, widths, texts):
        records.append({
            "kind": kind,
            "value": value.isoformat() if isinstance(value, (datetime.date, datetime.datetime)) else value,
            "python_type": type(value).__name__,
            "number_format": style.number_format if style is not None else None,
            "bold": bool(style is not None and style.bold),
            "excel_width": width,
            "excel_text": text,
        })
    return records


def main():
    if len(sys.argv) != 3 or sys.argv[1] not in ("chars", "corpus"):
        sys.exit(__doc__)
    result = measure_chars() if sys.argv[1] == "chars" else measure_corpus()
    with open(sys.argv[2], "w", encoding="utf-8") as handle:
        json.dump(result, handle, ensure_ascii=False, indent=1)
    print(f"wrote {sys.argv[2]}")


if __name__ == "__main__":
    main()
