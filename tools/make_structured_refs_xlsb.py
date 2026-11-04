"""Make ``tests/data/structured_refs.xlsb`` with Microsoft Excel, and the formulas it holds.

The workbook has tables, a totals row, calculated columns and formulas that refer to
tables ([MS-XLSB] PtgList) in every form Excel writes. ``structured_refs.json`` holds, per
sheet and cell, the formula as Excel reports it, written the way a file stores it
(``Tbl[[#This Row],[Col]]``, not the display form ``Tbl[@Col]``). The test reads the
workbook with xlspy and compares.

Windows with Excel and ``pywin32`` only:

    python tools/make_structured_refs_xlsb.py tests/data/structured_refs.xlsb
"""

from __future__ import annotations

import json
import re
import sys

import win32com.client

XL_XLSB = 50
NAMES = ["plain", "two words", "a,b", "a:b", "a[b]", "a#b", "it's", "a@b", 'say "x"', " lead", "trail ", "é"]


def file_form(text: str) -> str:
    """Excel shows ``[@Col]``; a file stores ``[[#This Row],[Col]]``."""
    text = re.sub(r"\[@\]", "[#This Row]", text)
    text = re.sub(r"\[@\[", "[[#This Row],[", text)
    return re.sub(r"\[@((?:[^\[\]']|'.)+)\]", r"[[#This Row],[\1]]", text)


def qualified(text: str, table: str) -> str:
    """Inside a table Excel leaves the table's name out; a file always has it."""
    out = []
    depth = 0
    index = 0
    while index < len(text):
        char = text[index]
        if char == "'" and depth:  # an escaped character inside brackets
            out.append(text[index:index + 2])
            index += 2
            continue
        if char == "[":
            if depth == 0:
                out.append(table)
            depth += 1
        elif char == "]":
            depth -= 1
        out.append(char)
        index += 1
    return "".join(out)


def main(path: str) -> None:
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    expected: dict[str, dict[str, str]] = {}
    try:
        book = excel.Workbooks.Add()
        sheet = book.Worksheets(1)
        sheet.Name = "Data"
        # A table that is removed, so that table ids and positions in the file differ.
        sheet.Range("X1:Y1").Value = ("a", "b")
        sheet.Range("X2:Y2").Value = (1, 2)
        junk = sheet.ListObjects.Add(1, sheet.Range("X1:Y2"), None, 1)
        junk.Name = "Junk"
        sheet.Range("A1:D1").Value = ("name", "qty", "my col", "z")
        for row in range(2, 6):
            sheet.Range(f"A{row}:D{row}").Value = (f"r{row}", row, row * 2, 1)
        tbl = sheet.ListObjects.Add(1, sheet.Range("A1:D5"), None, 1)
        tbl.Name = "Tbl"
        tbl.ShowTotals = True
        tbl.ListColumns(2).TotalsCalculation = 1  # sum
        tbl.ListColumns(3).TotalsCalculation = 2  # average
        junk.Delete()
        sheet.Range("H1:I1").Value = ("k", "v")
        sheet.Range("H2:I2").Value = ("x", 2)
        other = sheet.ListObjects.Add(1, sheet.Range("H1:I2"), None, 1)
        other.Name = "Other"
        # A column inserted before others: positions and column ids differ from here on.
        tbl.ListColumns(4).Delete()
        tbl.ListColumns.Add(2)
        tbl.ListColumns(2).Name = "ins"
        tbl.ListColumns.Add()
        tbl.ListColumns(5).Name = "calc"
        tbl.ListColumns(5).DataBodyRange.Formula = "=[@qty]*2+[@[my col]]"

        rows = [(), ("#All",), ("#Data",), ("#Headers",), ("#Totals",), ("#Headers", "#Data"), ("#Data", "#Totals")]
        columns = [(), ("qty",), ("ins",), ("qty", "my col"), ("ins", "my col")]
        texts = []
        for combo in rows:
            for column in columns:
                parts = list(combo)
                if len(column) == 1:
                    parts.append(column[0])
                elif len(column) == 2:
                    parts.append(f"{column[0]}]:[{column[1]}")
                if not parts:
                    texts.append("Tbl")
                elif len(parts) == 1 and not (len(column) == 2):
                    texts.append(f"Tbl[{parts[0]}]")
                else:
                    texts.append("Tbl[" + ",".join(f"[{p}]" for p in parts) + "]")
        texts += ["Tbl[[#This Row],[qty]]", "Tbl[[#This Row],[qty]:[my col]]", "Tbl[#This Row]", "Other[[#Totals],[v]]",
                  "Other[k]"]
        cells: dict[str, str] = {}
        cell_row = 20
        for text in dict.fromkeys(texts):
            formula = f"={text}" if "This Row" in text else f"=SUM({text})"
            address = f"AD{cell_row}"
            try:
                sheet.Range(address).Formula = formula
                cells[address] = file_form(sheet.Range(address).Formula)
                cell_row += 1
            except Exception:  # noqa: BLE001 - Excel refuses a form that makes no sense
                continue
        sheet.Range("AF20").Formula = "=Tbl[[#Totals],[qty]]+Other[[#Totals],[v]]*2"
        cells["AF20"] = file_form(sheet.Range("AF20").Formula)
        sheet.Range("AF21").Formula = "=INDEX(Tbl[[#Data],[name]],2)&Other[[#Headers],[k]]"
        cells["AF21"] = file_form(sheet.Range("AF21").Formula)
        # The formulas Excel itself wrote: the totals row and the calculated column.
        for cell in sheet.UsedRange.Cells:
            if cell.HasFormula and cell.Address.replace("$", "") not in cells:
                text = file_form(cell.Formula)
                if cell.ListObject is not None:
                    text = qualified(text, cell.ListObject.Name)
                cells[cell.Address.replace("$", "")] = text
        expected["Data"] = dict(sorted(cells.items()))

        # Column names with characters that are escaped or bracketed.
        special = book.Worksheets.Add(After=sheet)
        special.Name = "Names"
        for index, name in enumerate(NAMES):
            special.Cells(1, index + 1).Value = name
            special.Cells(2, index + 1).Value = index
        table = special.ListObjects.Add(1, special.Range(special.Cells(1, 1), special.Cells(2, len(NAMES))), None, 1)
        table.Name = "Sp"
        special_cells: dict[str, str] = {}
        for index, name in enumerate(NAMES):
            escaped = re.sub(r"(['#\[\]@])", r"'\1", name)
            address = f"A{5 + index}"
            try:
                special.Range(address).Formula = f"=SUM(Sp[[{escaped}]])"
                special_cells[address] = file_form(special.Range(address).Formula)
                special.Range(f"B{5 + index}").Formula = f"=COUNT(Sp[[#Headers],[{escaped}]])"
                special_cells[f"B{5 + index}"] = file_form(special.Range(f"B{5 + index}").Formula)
            except Exception:  # noqa: BLE001
                print("Excel refused a column name:", repr(name))
        expected["Names"] = special_cells
        book.SaveAs(path, XL_XLSB)
        book.Close(False)
    finally:
        excel.Quit()
    with open(path.rsplit(".", 1)[0] + ".json", "w", encoding="utf-8") as stream:
        json.dump(expected, stream, indent=1, ensure_ascii=False)
        stream.write("\n")
    print({sheet: len(cells) for sheet, cells in expected.items()}, "formulas")


if __name__ == "__main__":
    main(sys.argv[1])
