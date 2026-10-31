"""Read benchmark: xlspy vs openpyxl on the same data, with JSON output.

Each file is written by the library that owns the format (xlspy for its XLSX and
XLSB, xlsxwriter and openpyxl for their own XLSX). Every read runs in a fresh
child process and walks every row with ``values_only``. The count of non-empty
cells is reported for each case so that the readers can be checked against each
other.

xlsxwriter cannot read files, so it appears only as a writer of the input that
the readers consume. XLSB has no openpyxl reader, so XLSB is read by xlspy only.

Usage (from the repository root):
    py -3.14 benchmarks/bench_read_compare.py [--rows 50000] [--cols 50]
        [--repeat 3] [--out benchmarks/results-read-compare.json]

Requires openpyxl and xlsxwriter (optional for xlspy-only runs, which still work).
"""
from __future__ import annotations

import argparse
import datetime
import json
import os
import statistics
import subprocess
import sys
import tempfile
import time

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "src"))

CASES = [
    # (case, file key, reader)
    ("xlspy ExcelReader, XLSX", "xlspy_xlsx", "xlspy_reader"),
    ("xlspy ExcelReader, XLSB", "xlspy_xlsb", "xlspy_reader"),
    ("xlspy ExcelReader, xlsxwriter file", "xlsxwriter_xlsx", "xlspy_reader"),
    ("xlspy load_workbook, XLSX", "xlspy_xlsx", "xlspy_load"),
    ("xlspy load_workbook, XLSB", "xlspy_xlsb", "xlspy_load"),
    ("openpyxl read_only, XLSX", "xlspy_xlsx", "openpyxl_readonly"),
    ("openpyxl read_only, xlsxwriter file", "xlsxwriter_xlsx", "openpyxl_readonly"),
    ("openpyxl load_workbook, XLSX", "xlspy_xlsx", "openpyxl_load"),
    ("openpyxl load_workbook, xlsxwriter file", "xlsxwriter_xlsx", "openpyxl_load"),
]

FILES = {
    "xlspy_xlsx": "xlspy.xlsx",
    "xlspy_xlsb": "xlspy.xlsb",
    "xlsxwriter_xlsx": "xlsxwriter.xlsx",
}


def _rss_mb():
    import psutil
    return psutil.Process().memory_info().rss / (1024 * 1024)


def _walk(rows):
    """Return (row count, non-empty cell count) after consuming every row."""
    row_count = 0
    filled = 0
    for row in rows:
        row_count += 1
        filled += sum(1 for value in row if value is not None)
    return row_count, filled


def run_read(reader, path, sheet):
    if reader == "xlspy_reader":
        from xlspy import ExcelReader
        with ExcelReader(path) as source:
            return _walk(source.get_rows(sheet))
    if reader == "xlspy_load":
        from xlspy import load_workbook
        workbook = load_workbook(path)
        try:
            return _walk(workbook[sheet].iter_rows(values_only=True))
        finally:
            workbook.close()
    if reader == "openpyxl_readonly":
        import openpyxl
        workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
        try:
            return _walk(workbook[sheet].iter_rows(values_only=True))
        finally:
            workbook.close()
    if reader == "openpyxl_load":
        import openpyxl
        workbook = openpyxl.load_workbook(path, data_only=True)
        try:
            return _walk(workbook[sheet].iter_rows(values_only=True))
        finally:
            workbook.close()
    raise SystemExit(f"unknown reader {reader}")


def child(args):
    """Executed in the child process; prints one JSON line with the metrics."""
    before = _rss_mb()
    start = time.perf_counter()
    row_count, filled = run_read(args.reader, args.path, "S")
    seconds = time.perf_counter() - start
    print(json.dumps({
        "seconds": seconds,
        "rows": row_count,
        "non_empty_cells": filled,
        "peak_rss_delta_mb": max(_rss_mb() - before, 0.0),
    }))


def make_files(directory, rows, cols):
    """Write the three input files once; all of them hold the same values."""
    from bench_io import rows as generate

    import xlspy
    import xlsxwriter
    xlspy_xlsx = os.path.join(directory, FILES["xlspy_xlsx"])
    xlspy_xlsb = os.path.join(directory, FILES["xlspy_xlsb"])
    xlsxwriter_path = os.path.join(directory, FILES["xlsxwriter_xlsx"])

    with xlspy.XlsxWriter(xlspy_xlsx) as writer:
        writer.add_sheet("S")
        writer.write_sheet(generate(rows, cols))
    with xlspy.XlsbWriter(xlspy_xlsb) as writer:
        writer.add_sheet("S")
        writer.write_sheet(generate(rows, cols))

    book = xlsxwriter.Workbook(xlsxwriter_path, {"constant_memory": True})
    sheet = book.add_worksheet("S")
    date_format = book.add_format({"num_format": "yyyy-mm-dd"})
    for index, row in enumerate(generate(rows, cols)):
        for column, value in enumerate(row):
            if isinstance(value, datetime.date):
                sheet.write_datetime(index, column, datetime.datetime.combine(value, datetime.time()), date_format)
            elif value is not None:
                sheet.write(index, column, value)
    book.close()
    return {"xlspy_xlsx": xlspy_xlsx, "xlspy_xlsb": xlspy_xlsb, "xlsxwriter_xlsx": xlsxwriter_path}


def parent(args):
    import xlsxwriter  # noqa: F401  (fail early if the optional writer is missing)
    import openpyxl  # noqa: F401

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
                print(f"{label:40s} {results[label]}", flush=True)
                continue
            merged = dict(good[0])
            merged["seconds"] = statistics.median(r["seconds"] for r in good)
            merged["peak_rss_delta_mb"] = statistics.median(r["peak_rss_delta_mb"] for r in good)
            merged["runs"] = [round(r["seconds"], 3) for r in good]
            results[label] = merged
            print(f"{label:40s} {merged['seconds']:7.2f} s  {merged['non_empty_cells']} cells", flush=True)

    document = {
        "python": sys.version.split()[0],
        "rows": args.rows,
        "cols": args.cols,
        "c_ext": not os.environ.get("XLSPY_DISABLE_C_EXT"),
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
