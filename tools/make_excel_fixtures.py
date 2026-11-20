"""Generate small workbooks with desktop Excel for the CI read-only corpus.

Excel writes both files (XLSX and XLSB); CI then reads them without Excel.
The manifest records the Excel version and build, so a fixture refresh is
visible in the diff.  Run on Windows with Excel installed:

    python tools/make_excel_fixtures.py

Microsoft 365 reports version 16.0, like Excel 2016+; the build number
distinguishes the installation.
"""

from __future__ import annotations

import datetime as dt
import hashlib
import json
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parents[1]
OUT = ROOT / "tests" / "data" / "excel_365"

# Excel's enum values (win32com constants are not populated without EnsureDispatch).
XL_OPENXML_WORKBOOK = 51
XL_EXCEL12 = 50
XL_SRC_RANGE = 1
XL_YES = 1


def _fill_sheet(sheet) -> None:
    sheet.Name = "Data"
    sheet.Range("A1").Value = "name"
    sheet.Range("B1").Value = "value"
    sheet.Range("C1").Value = "date"
    sheet.Range("D1").Value = "link"
    sheet.Range("A2").Value = "alpha"
    sheet.Range("B2").Value = 12500
    # Value2 is the raw serial, so the fixture does not depend on the machine's
    # timezone (COM converts a Python datetime through UTC).
    sheet.Range("C2").Value2 = 45292  # 2024-01-01
    sheet.Range("A3").Value = "beta"
    sheet.Range("B3").Value = 8400
    sheet.Range("C3").Value2 = 45328  # 2024-02-06
    sheet.Range("A4").Value = "total"
    sheet.Range("B4").Formula = "=SUM(B2:B3)"

    header = sheet.Range("A1:D1")
    header.Font.Bold = True
    header.Interior.Color = 0x523C18  # BGR of RGB(24, 60, 82)
    header.Font.Color = 0xFFFFFF
    sheet.Range("B2:B3").NumberFormat = "#,##0"
    sheet.Range("C2:C3").NumberFormat = "yyyy-mm-dd"

    sheet.Hyperlinks.Add(
        Anchor=sheet.Range("D2"),
        Address="https://example.com/excel-fixture",
        TextToDisplay="fixture",
    )
    sheet.Range("A1").AddComment("Written by Excel for the read-only corpus")

    sheet.Range("E1").Value = "merged"
    sheet.Range("E1:F1").Merge()

    # A real table: Excel writes the table part and structured references.
    sheet.ListObjects.Add(XL_SRC_RANGE, sheet.Range("A1:C3"), None, XL_YES)


def main() -> int:
    if sys.platform != "win32":
        print("Excel fixtures need Windows and a desktop Excel installation", file=sys.stderr)
        return 1
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError:
        print("pywin32 is required: pip install pywin32", file=sys.stderr)
        return 1

    OUT.mkdir(parents=True, exist_ok=True)
    app = win32com.client.DispatchEx("Excel.Application")
    try:
        app.Visible = False
        app.DisplayAlerts = False
        app.EnableEvents = False
        version, build = str(app.Version), str(app.Build)
        workbook = app.Workbooks.Add()
        try:
            _fill_sheet(workbook.Worksheets(1))
            workbook.SaveAs(str(OUT / "features.xlsx"), FileFormat=XL_OPENXML_WORKBOOK)
            workbook.SaveAs(str(OUT / "features.xlsb"), FileFormat=XL_EXCEL12)
        finally:
            workbook.Close(SaveChanges=False)
    finally:
        app.Quit()

    manifest = {
        "excel": {
            "version": version,
            "build": build,
        },
        "generated": dt.date.today().isoformat(),
        "files": {
            name: hashlib.sha256((OUT / name).read_bytes()).hexdigest()
            for name in ("features.xlsx", "features.xlsb")
        },
    }
    (OUT / "manifest.json").write_text(
        json.dumps(manifest, indent=2) + "\n", encoding="utf-8"
    )
    print(f"wrote {OUT} with Excel {version} ({build})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
