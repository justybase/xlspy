"""Read/write benchmark for XLSX and XLSB, with JSON output.

Cases: streaming write, streaming read (``ExcelReader``), full-model
``load_workbook`` (plain and styled), peak RSS delta, XLSX reader fallback rate,
and a zlib-level sweep for XLSX writing.

Usage (from the repository root):
    py -3.14 benchmarks/bench_io.py [--rows 50000] [--cols 20] [--repeat 3]
        [--out benchmarks/results.json] [--compare baseline.json --max-slowdown 1.5]

Each case runs in a fresh subprocess so peak memory and import state are isolated.
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

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, os.path.join(ROOT, "src"))

CASES = [
    "write_xlsx", "write_xlsb",
    "read_xlsx", "read_xlsb",
    "load_xlsx", "load_xlsb",
    "load_xlsx_styled", "load_xlsb_styled", "save_loaded_xlsb_styled",
    "write_model_xlsx", "write_model_xlsb",
    "xlsx_fallback_rate",
]


def rows(count, cols):
    base = datetime.date(2023, 1, 1)
    yield [f"Header_{j}" for j in range(cols)]
    for i in range(1, count):
        out = []
        for j in range(cols):
            kind = j % 5
            if kind == 0:
                out.append(f"String_{i % 5000}_{j}")  # repeated strings -> shared table
            elif kind == 1:
                out.append(i * j)
            elif kind == 2:
                out.append(i / (j + 1))
            elif kind == 3:
                out.append(base + datetime.timedelta(days=i % 3000))
            else:
                out.append((i + j) % 2 == 0)
        yield out


def rss_mb():
    import psutil
    return psutil.Process().memory_info().rss / (1024 * 1024)


def _write(path, kind, count, cols, level=None):
    from xlspy import XlsbWriter, XlsxWriter
    start = time.perf_counter()
    if kind == "xlsx":
        kwargs = {} if level is None else {"compressionLevel": level}
        with XlsxWriter(path, **kwargs) as writer:
            writer.add_sheet("S")
            writer.write_sheet(rows(count, cols))
    else:
        with XlsbWriter(path) as writer:
            writer.add_sheet("S")
            writer.write_sheet(rows(count, cols))
    return time.perf_counter() - start


def _model_write(path, count, cols):
    from xlspy import Workbook
    start = time.perf_counter()
    workbook = Workbook()
    sheet = workbook.active
    for row in rows(count, cols):
        sheet.append(row)
    workbook.save(path)
    return time.perf_counter() - start


def _styled(path, count, cols):
    from xlspy import Font, Workbook
    workbook = Workbook()
    sheet = workbook.active
    bold = Font(bold=True)
    for index, row in enumerate(rows(count, cols)):
        sheet.append(row)
        if index % 10 == 0:
            sheet.cell(row=index + 1, column=1).font = bold
    workbook.save(path)


def run_case(case, path, count, cols):
    """Executed in the child process; returns a dict of metrics."""
    before = rss_mb()
    result = {}
    if case in ("write_xlsx", "write_xlsb"):
        result["seconds"] = _write(path, case[6:], count, cols)
        result["size_mb"] = os.path.getsize(path) / 1048576
    elif case == "write_model_xlsx" or case == "write_model_xlsb":
        result["seconds"] = _model_write(path, count, cols)
        result["size_mb"] = os.path.getsize(path) / 1048576
    elif case in ("read_xlsx", "read_xlsb"):
        from xlspy import ExcelReader
        _write(path, case[5:], count, cols)
        before = rss_mb()
        start = time.perf_counter()
        total = 0
        with ExcelReader(path) as reader:
            for row in reader.get_rows("S"):
                total += len(row)
        result["seconds"] = time.perf_counter() - start
        result["cells"] = total
    elif case == "save_loaded_xlsb_styled":
        # Restyle every 50th row of a loaded styled XLSB, then time the in-place save.
        from xlspy import Font, load_workbook
        _styled(path, count, cols)
        workbook = load_workbook(path)
        italic = Font(italic=True)
        sheet = workbook.active
        for row in range(2, count + 1, 50):
            sheet.cell(row=row, column=2).font = italic
            sheet.cell(row=row, column=2).value = -row
        before = rss_mb()
        start = time.perf_counter()
        workbook.save(path[:-5] + "-out.xlsb")
        result["seconds"] = time.perf_counter() - start
        workbook.close()
    elif case in ("load_xlsx", "load_xlsb", "load_xlsx_styled", "load_xlsb_styled"):
        from xlspy import load_workbook
        if case.endswith("_styled"):
            _styled(path, count, cols)
        else:
            _write(path, case[5:], count, cols)
        before = rss_mb()
        start = time.perf_counter()
        workbook = load_workbook(path)
        result["seconds"] = time.perf_counter() - start
        workbook.close()
    elif case == "xlsx_fallback_rate":
        from xlspy import ExcelReader, _xlsx_read
        _write(path, "xlsx", count, cols)
        calls = {"n": 0}
        original = _xlsx_read._row_from_element

        def counting(*args, **kwargs):
            calls["n"] += 1
            return original(*args, **kwargs)

        _xlsx_read._row_from_element = counting
        total = 0
        with ExcelReader(path) as reader:
            for _ in reader.get_rows("S"):
                total += 1
        result["rows"] = total
        result["fallback_rows"] = calls["n"]
        result["fallback_pct"] = 100.0 * calls["n"] / max(total, 1)
    elif case.startswith("zlib"):
        level = int(case[4:])
        result["seconds"] = _write(path, "xlsx", count, cols, level)
        result["size_mb"] = os.path.getsize(path) / 1048576
    else:
        raise SystemExit(f"unknown case {case}")
    result["peak_rss_delta_mb"] = max(rss_mb() - before, 0.0)
    return result


def child(args):
    print(json.dumps(run_case(args.case, args.path, args.rows, args.cols)))


def parent(args):
    results = {}
    cases = list(CASES) + [f"zlib{level}" for level in (1, 3, 6, 9)]
    with tempfile.TemporaryDirectory(prefix="xlspy-bench-") as directory:
        for case in cases:
            runs = []
            for _ in range(1 if case == "xlsx_fallback_rate" else args.repeat):
                path = os.path.join(directory, case + (".xlsb" if "xlsb" in case else ".xlsx"))
                completed = subprocess.run(
                    [sys.executable, os.path.abspath(__file__), "--child", case,
                     "--path", path, "--rows", str(args.rows), "--cols", str(args.cols)],
                    capture_output=True, text=True, env=dict(os.environ),
                )
                if completed.returncode:
                    runs.append({"error": completed.stderr.strip().splitlines()[-1:]})
                    break
                runs.append(json.loads(completed.stdout.strip().splitlines()[-1]))
            good = [r for r in runs if "error" not in r]
            if not good:
                results[case] = runs[0]
            else:
                merged = dict(good[0])
                if "seconds" in merged:
                    merged["seconds"] = statistics.median(r["seconds"] for r in good)
                results[case] = merged
            print(f"{case:22s} {results[case]}", flush=True)
    document = {
        "python": sys.version.split()[0], "rows": args.rows, "cols": args.cols,
        "c_ext": not os.environ.get("XLSPY_DISABLE_C_EXT"), "results": results,
    }
    if args.out:
        with open(args.out, "w", encoding="utf-8") as handle:
            json.dump(document, handle, indent=2)
    if args.compare:
        return compare(document, args.compare, args.max_slowdown)
    return 0


def compare(document, baseline_path, limit):
    with open(baseline_path, encoding="utf-8") as handle:
        baseline = json.load(handle)["results"]
    failed = False
    for case, now in document["results"].items():
        before = baseline.get(case, {})
        if "seconds" in now and "seconds" in before:
            ratio = now["seconds"] / before["seconds"]
            flag = "REGRESSION" if ratio > limit else "ok"
            failed |= ratio > limit
            print(f"{case:22s} {ratio:5.2f}x  {flag}")
    return 1 if failed else 0


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--rows", type=int, default=50000)
    parser.add_argument("--cols", type=int, default=20)
    parser.add_argument("--repeat", type=int, default=3)
    parser.add_argument("--out")
    parser.add_argument("--compare")
    parser.add_argument("--max-slowdown", type=float, default=1.5)
    parser.add_argument("--child")
    parser.add_argument("--path")
    args = parser.parse_args()
    if args.child:
        args.case = args.child
        child(args)
        return 0
    return parent(args)


if __name__ == "__main__":
    raise SystemExit(main())
