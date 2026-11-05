"""Capture modern formulas and diagnose what xlspy already does with them.

Builds one workbook per formula in Microsoft Excel (COM), saves it as XLSX and
XLSB, and prints the parts xlspy has to understand:

- XLSX: the worksheet XML (`cm`/`vm`, `<f>`), `xl/metadata.xml`, defined names;
- XLSB: the worksheet records (formula cells, `BrtCellMeta` 0x31, `BrtValueMeta`
  0x32), the metadata part and the `BrtName` entries (`_xlfn.*`, `_xlpm.*`).

It then runs the xlspy paths that have to keep the file intact: `ExcelReader`,
`load_workbook(data_only=False)`, a no-op save (compared part by part), and an
`update_cells` outside the formula's cell.

Usage:
    py -3.14 tools/capture_formulas.py OUT_DIR [name ...]
    names: let_simple let_two lambda_call map_lambda reduce filter sort unique sequence
"""

from __future__ import annotations

import argparse
import os
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "../src")))

import win32com.client  # noqa: E402

from xlspy import ExcelReader, XlsbUpdater, XlsxUpdater, load_workbook  # noqa: E402
from xlspy._xlsb_metadata import read_wide  # noqa: E402
from xlspy.biff_utils import iter_records  # noqa: E402

XL_XLSB, XL_XLSX = 50, 51

FORMULAS = {
    "let_simple": "=LET(x,1,x+1)",
    "let_two": "=LET(x,1,y,2,x+y)",
    "lambda_call": "=LAMBDA(x,x*2)(3)",
    "map_lambda": "=MAP(A1:A3,LAMBDA(v,v*2))",
    "reduce": "=REDUCE(0,A1:A3,LAMBDA(a,b,a+b))",
    "filter": "=FILTER(A1:B4,B1:B4>2)",
    "sort": "=SORT(A1:B4)",
    "unique": "=UNIQUE(A1:A4)",
    "sequence": "=SEQUENCE(3,2)",
}

DATA = [[1, 10], [2, 20], [3, 30], [4, 40]]
FORMULA_CELL = "D1"
FORMULA_RECORDS = set(range(0x08, 0x13))


def build(excel, path, formula):
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Name = "Data"
        for row_number, row in enumerate(DATA, start=1):
            for column_number, value in enumerate(row, start=1):
                sheet.Cells(row_number, column_number).Value = value
        target = sheet.Range(FORMULA_CELL)
        try:
            target.Formula2 = formula
        except Exception:
            target.Formula = formula
        book.SaveAs(str(path), XL_XLSB if path.suffix == ".xlsb" else XL_XLSX)
    finally:
        book.Close(False)


def dump_xlsx(path):
    with zipfile.ZipFile(path) as archive:
        names = sorted(archive.namelist())
        print("  parts:", names)
        if "xl/metadata.xml" in names:
            print("  metadata.xml:", archive.read("xl/metadata.xml").decode("utf-8"))
        sheet = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")
        print("  sheet1.xml:", sheet)
        workbook = archive.read("xl/workbook.xml").decode("utf-8")
        print("  workbook.xml:", workbook)


def dump_xlsb(path):
    with zipfile.ZipFile(path) as archive:
        names = sorted(archive.namelist())
        print("  parts:", names)
        for name in names:
            if "metadata" in name:
                data = archive.read(name)
                print(f"  {name}: {data.hex()}")
                for record in iter_records(data):
                    payload = data[record.data_start:record.data_end]
                    print(f"    {record.record_id:#06x} len={record.length} {payload.hex()}")
        data = archive.read("xl/worksheets/sheet1.bin")
        for record in iter_records(data):
            payload = data[record.data_start:record.data_end]
            if record.record_id in FORMULA_RECORDS or record.record_id in (0x31, 0x32):
                print(f"  record {record.record_id:#06x} len={record.length} {payload.hex()}")
        workbook = archive.read("xl/workbook.bin")
        for record in iter_records(workbook):
            if record.record_id != 0x27:
                continue
            payload = workbook[record.data_start:record.data_end]
            flags, _reserved, scope = (
                int.from_bytes(payload[0:4], "little"),
                int.from_bytes(payload[4:5], "little"),
                int.from_bytes(payload[5:9], "little"),
            )
            text, _ = read_wide(payload, 9)
            print(f"  name {text!r} flags={flags:#x} scope={scope:#x}")


def _parts(path):
    with zipfile.ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def diagnose(path, suffix):
    print("  -- xlspy --")
    try:
        with ExcelReader(path) as reader:
            rows = reader.read_all("Data", data_only=False)
            print("  reader D1/E1:", rows[0][3:5] if rows and len(rows[0]) > 4 else rows[:1])
    except Exception as error:
        print("  reader failed:", type(error).__name__, error)

    book = None
    try:
        book = load_workbook(path, data_only=False)
        cell = book["Data"][FORMULA_CELL]
        print("  model D1:", repr(cell.value), "data_type:", cell.data_type)
        noop = path.with_name(path.stem + "-noop" + path.suffix)
        book.save(noop)
        before, after = _parts(path), _parts(noop)
        changed = [name for name in sorted(set(before) | set(after)) if before.get(name) != after.get(name)]
        print("  no-op changed parts:", changed or "none")
    except Exception as error:
        print("  load/save failed:", type(error).__name__, error)
    finally:
        if book is not None:
            book.close()

    updater_class = XlsxUpdater if suffix == "xlsx" else XlsbUpdater
    target = path.with_name(path.stem + "-edit" + path.suffix)
    try:
        updater = updater_class(str(path))
        try:
            updater.update_cells("Data", {"C10": 1})
            updater.save(str(target))
        finally:
            updater.close()
        with ExcelReader(target) as reader:
            print("  update_cells outside: ok, C10 =", reader.read_all("Data")[9][2])
    except Exception as error:
        print("  update_cells outside failed:", type(error).__name__, error)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("out_dir")
    parser.add_argument("names", nargs="*", default=sorted(FORMULAS))
    args = parser.parse_args()
    out_dir = Path(os.path.abspath(args.out_dir))
    out_dir.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    try:
        for name in args.names:
            formula = FORMULAS[name]
            print(f"=== {name}: {formula}")
            for suffix in ("xlsx", "xlsb"):
                path = out_dir / f"{name}.{suffix}"
                build(excel, path, formula)
                print(f"  file: {path} ({os.path.getsize(path)} bytes)")
                if suffix == "xlsx":
                    dump_xlsx(Path(path))
                else:
                    dump_xlsb(Path(path))
                diagnose(Path(path), suffix)
    finally:
        excel.Quit()


if __name__ == "__main__":
    main()
