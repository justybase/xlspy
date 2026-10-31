"""Read benchmark: xlspy vs the fastest Python readers, with JSON output.

Each file is written by the library that owns the format (xlspy for its XLSX and
XLSB, xlsxwriter for its XLSX). Every read runs in a fresh child process. The
reader's library is imported before the clock starts. The timed pass only consumes
the rows; the non-empty cell count is taken in a second, untimed pass over the same
file. Every case reports the number of rows and non-empty cells it saw, and the
parent warns when two readers disagree, so the readers can be checked against each other.

Readers:
  xlspy ExcelReader, load_workbook   XLSX and XLSB (this package)
  python-calamine                    XLSX and XLSB, Rust (calamine); the whole sheet as lists
  fastexcel                          XLSX and XLSB, Rust (calamine); the whole sheet as an Arrow batch
  pyxlsb                             XLSB, pure Python
  openpyxl                           XLSX, pure Python (read_only and full load)

xlsxwriter and PyExcelerate cannot read files, so they appear only as writers of
the inputs. xlspy is the only library in this comparison that writes XLSB.

pyxlsb trusts the BrtWsDim record, and XlsbWriter writes an A1:A1 placeholder
there. The pyxlsb case therefore reads a copy of the xlspy file with the real
bounds written into that record. The cells are the same.

Usage (from the repository root):
    py -3.14 benchmarks/bench_read_compare.py [--rows 50000] [--cols 50]
        [--repeat 3] [--out benchmarks/results-read-compare.json]

Requires openpyxl, xlsxwriter, python-calamine, fastexcel (with pyarrow), pyxlsb and psutil.
"""
from __future__ import annotations

import argparse
import datetime
import importlib.metadata
import json
import os
import platform
import statistics
import struct
import subprocess
import sys
import tempfile
import time
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "src"))

SHEET = "S"
SHEET_PART = "xl/worksheets/sheet1.bin"
# BrtWsDim with the A1:A1 placeholder: record id 0x94 (two-byte BIFF12 id), size 16, 16 zero bytes.
WSDIM_PLACEHOLDER = b"\x94\x01\x10" + bytes(16)

CASES = [
    # (label, file key, reader)
    ("xlspy ExcelReader, XLSX", "xlspy_xlsx", "xlspy_reader"),
    ("xlspy ExcelReader, XLSB", "xlspy_xlsb", "xlspy_reader"),
    ("xlspy ExcelReader, xlsxwriter file", "xlsxwriter_xlsx", "xlspy_reader"),
    ("xlspy load_workbook, XLSX", "xlspy_xlsx", "xlspy_load"),
    ("xlspy load_workbook, XLSB", "xlspy_xlsb", "xlspy_load"),
    ("python-calamine, XLSX", "xlspy_xlsx", "calamine"),
    ("python-calamine, XLSB", "xlspy_xlsb", "calamine"),
    ("python-calamine, xlsxwriter file", "xlsxwriter_xlsx", "calamine"),
    ("fastexcel, XLSX", "xlspy_xlsx", "fastexcel"),
    ("fastexcel, XLSB", "xlspy_xlsb", "fastexcel"),
    ("pyxlsb, XLSB", "xlspy_xlsb_fixed", "pyxlsb"),
    ("openpyxl read_only, XLSX", "xlspy_xlsx", "openpyxl_readonly"),
    ("openpyxl read_only, xlsxwriter file", "xlsxwriter_xlsx", "openpyxl_readonly"),
    ("openpyxl load_workbook, XLSX", "xlspy_xlsx", "openpyxl_load"),
    ("openpyxl load_workbook, xlsxwriter file", "xlsxwriter_xlsx", "openpyxl_load"),
]

FILES = {
    "xlspy_xlsx": "xlspy.xlsx",
    "xlspy_xlsb": "xlspy.xlsb",
    "xlspy_xlsb_fixed": "xlspy-dimension.xlsb",
    "xlsxwriter_xlsx": "xlsxwriter.xlsx",
}


def _rss_mb():
    import psutil
    return psutil.Process().memory_info().rss / (1024 * 1024)


def _walk(rows, count_cells=True):
    """Consume every row; return (row count, non-empty cell count).

    With ``count_cells=False`` the cells are not inspected and the count is None,
    so the timed pass does not include the cost of counting.
    python-calamine reports empty cells as '', the other readers as None.
    """
    row_count = 0
    filled = 0
    for row in rows:
        row_count += 1
        if count_cells:
            filled += sum(1 for value in row if value is not None and value != "")
    return row_count, (filled if count_cells else None)


def make_reader(name):
    """Import the reader's library and return ``read(path, count_cells=True)``.

    Returns ``(rows, filled)``. Runs before the clock starts, so import time is
    not part of the result.
    """
    if name == "xlspy_reader":
        from xlspy import ExcelReader

        def read(path, count_cells=True):
            with ExcelReader(path) as source:
                return _walk(source.get_rows(SHEET), count_cells)
        return read
    if name == "xlspy_load":
        from xlspy import load_workbook

        def read(path, count_cells=True):
            workbook = load_workbook(path)
            try:
                return _walk(workbook[SHEET].iter_rows(values_only=True), count_cells)
            finally:
                workbook.close()
        return read
    if name == "calamine":
        from python_calamine import CalamineWorkbook

        def read(path, count_cells=True):
            # to_python() builds the whole sheet as Python lists before the walk.
            workbook = CalamineWorkbook.from_path(path)
            return _walk(workbook.get_sheet_by_name(SHEET).to_python(), count_cells)
        return read
    if name == "fastexcel":
        import fastexcel

        def read(path, count_cells=True):
            # to_arrow() builds the whole sheet as one Arrow batch before any count.
            sheet = fastexcel.read_excel(path).load_sheet_by_name(SHEET, header_row=None)
            batch = sheet.to_arrow()
            if not count_cells:
                return batch.num_rows, None
            filled = sum(len(column) - column.null_count for column in batch.columns)
            return batch.num_rows, filled
        return read
    if name == "pyxlsb":
        from pyxlsb import open_workbook

        def read(path, count_cells=True):
            with open_workbook(path) as workbook, workbook.get_sheet(SHEET) as worksheet:
                return _walk(([cell.v for cell in row] for row in worksheet.rows(sparse=False)),
                             count_cells)
        return read
    if name == "openpyxl_readonly":
        import openpyxl

        def read(path, count_cells=True):
            workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
            try:
                return _walk(workbook[SHEET].iter_rows(values_only=True), count_cells)
            finally:
                workbook.close()
        return read
    if name == "openpyxl_load":
        import openpyxl

        def read(path, count_cells=True):
            workbook = openpyxl.load_workbook(path, data_only=True)
            try:
                return _walk(workbook[SHEET].iter_rows(values_only=True), count_cells)
            finally:
                workbook.close()
        return read
    raise SystemExit(f"unknown reader {name}")


def child(args):
    """Executed in the child process; prints one JSON line with the metrics.

    The timed pass only consumes the rows. The row and cell counts come from a
    second pass over the same file, which is not timed.
    """
    read = make_reader(args.reader)
    before = _rss_mb()
    start = time.perf_counter()
    read(args.path, count_cells=False)
    seconds = time.perf_counter() - start
    peak_rss_delta = max(_rss_mb() - before, 0.0)
    row_count, filled = read(args.path, count_cells=True)
    print(json.dumps({
        "seconds": seconds,
        "rows": row_count,
        "non_empty_cells": filled,
        "peak_rss_delta_mb": peak_rss_delta,
    }))


def _with_dimension(source, target, rows, cols):
    """Copy an XLSB file, writing the real bounds into its BrtWsDim record.

    Only that record changes, so the cells are the same as in ``source``.
    """
    replacement = b"\x94\x01\x10" + struct.pack("<IIII", 0, rows - 1, 0, cols - 1)
    with zipfile.ZipFile(source) as zin, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == SHEET_PART:
                if data.count(WSDIM_PLACEHOLDER) != 1:
                    raise RuntimeError("BrtWsDim placeholder not found exactly once")
                data = data.replace(WSDIM_PLACEHOLDER, replacement, 1)
            zout.writestr(item, data)


def make_files(directory, rows, cols):
    """Write the input files once; all of them hold the same values."""
    from bench_io import rows as generate

    import xlspy
    import xlsxwriter
    xlspy_xlsx = os.path.join(directory, FILES["xlspy_xlsx"])
    xlspy_xlsb = os.path.join(directory, FILES["xlspy_xlsb"])
    xlspy_xlsb_fixed = os.path.join(directory, FILES["xlspy_xlsb_fixed"])
    xlsxwriter_path = os.path.join(directory, FILES["xlsxwriter_xlsx"])

    with xlspy.XlsxWriter(xlspy_xlsx) as writer:
        writer.add_sheet(SHEET)
        writer.write_sheet(generate(rows, cols))
    with xlspy.XlsbWriter(xlspy_xlsb) as writer:
        writer.add_sheet(SHEET)
        writer.write_sheet(generate(rows, cols))
    _with_dimension(xlspy_xlsb, xlspy_xlsb_fixed, rows, cols)

    book = xlsxwriter.Workbook(xlsxwriter_path, {"constant_memory": True})
    sheet = book.add_worksheet(SHEET)
    date_format = book.add_format({"num_format": "yyyy-mm-dd"})
    for index, row in enumerate(generate(rows, cols)):
        for column, value in enumerate(row):
            if isinstance(value, datetime.date):
                sheet.write_datetime(index, column, datetime.datetime.combine(value, datetime.time()), date_format)
            elif value is not None:
                sheet.write(index, column, value)
    book.close()
    return {"xlspy_xlsx": xlspy_xlsx, "xlspy_xlsb": xlspy_xlsb,
            "xlspy_xlsb_fixed": xlspy_xlsb_fixed, "xlsxwriter_xlsx": xlsxwriter_path}


def _versions():
    import xlspy
    versions = {"python": platform.python_version(), "xlspy": getattr(xlspy, "__version__", None)}
    for dist in ("openpyxl", "xlsxwriter", "python-calamine", "fastexcel", "pyarrow", "pyxlsb", "psutil"):
        try:
            versions[dist] = importlib.metadata.version(dist)
        except importlib.metadata.PackageNotFoundError:
            versions[dist] = None
    return versions


def parent(args):
    # Fail before writing any file if a benchmark dependency is missing.
    for module in ("xlsxwriter", "openpyxl", "python_calamine", "fastexcel", "pyarrow", "pyxlsb", "psutil"):
        __import__(module)

    results = {}
    with tempfile.TemporaryDirectory(prefix="xlspy-read-bench-") as directory:
        paths = make_files(directory, args.rows, args.cols)
        for label, file_key, reader in CASES:
            runs = []
            for _ in range(args.repeat):
                completed = subprocess.run(
                    [sys.executable, os.path.abspath(__file__), "--child", reader,
                     "--path", paths[file_key]],
                    capture_output=True, text=True, env=dict(os.environ),
                )
                if completed.returncode:
                    runs.append({"error": completed.stderr.strip().splitlines()[-1:]})
                    break
                runs.append(json.loads(completed.stdout.strip().splitlines()[-1]))
            good = [r for r in runs if "error" not in r]
            if not good:
                results[label] = runs[0]
                print(f"{label:42s} {results[label]}", flush=True)
                continue
            merged = dict(good[0])
            merged["seconds"] = statistics.median(r["seconds"] for r in good)
            merged["peak_rss_delta_mb"] = statistics.median(r["peak_rss_delta_mb"] for r in good)
            merged["runs"] = [round(r["seconds"], 3) for r in good]
            results[label] = merged
            print(f"{label:42s} {merged['seconds']:7.2f} s  {merged['rows']} rows  "
                  f"{merged['non_empty_cells']} cells", flush=True)

    counts = {(r["rows"], r["non_empty_cells"]) for r in results.values() if "error" not in r}
    if len(counts) > 1:
        print(f"WARNING: readers disagree on (rows, non-empty cells): {sorted(counts)}", file=sys.stderr)

    document = {
        "python": sys.version.split()[0],
        "rows": args.rows,
        "cols": args.cols,
        "c_ext": not os.environ.get("XLSPY_DISABLE_C_EXT"),
        "versions": _versions(),
        "results": results,
    }
    if args.out:
        with open(args.out, "w", encoding="utf-8") as handle:
            json.dump(document, handle, indent=2)
    return 0


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--rows", type=int, default=50000)
    parser.add_argument("--cols", type=int, default=50)
    parser.add_argument("--repeat", type=int, default=3)
    parser.add_argument("--out")
    parser.add_argument("--child", dest="reader")
    parser.add_argument("--path")
    args = parser.parse_args()
    if args.reader:
        child(args)
        return 0
    return parent(args)


if __name__ == "__main__":
    raise SystemExit(main())
