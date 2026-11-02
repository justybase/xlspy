"""Read/write benchmark for XLSX and XLSB, with JSON output.

Cases: streaming write, streaming read (``ExcelReader``), full-model
``load_workbook`` (plain and styled), peak RSS delta, XLSX reader and model
fallback rates (plain file and a file whose cells need XML decoding), a write
of Formula and Hyperlink cells, and a zlib-level sweep for XLSX writing.

Usage (from the repository root):
    py -3.14 benchmarks/bench_io.py [--rows 50000] [--cols 20] [--repeat 7]
        [--only read_xlsx,load_xlsx] [--out benchmarks/results.json]
        [--compare baseline.json --max-slowdown 1.5]

Each case runs in a fresh subprocess so peak memory and import state are isolated.
Timed cases report the median of ``--repeat`` runs and the interquartile range
(IQR) of those runs, so the noise of a run is visible next to its result.
"""
from __future__ import annotations

import argparse
import datetime
import json
import os
import platform
import statistics
import subprocess
import sys
import tempfile
import time
import zipfile

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, os.path.join(ROOT, "src"))

CASES = [
    "write_xlsx", "write_xlsb",
    "read_xlsx", "read_xlsb",
    "load_xlsx", "load_xlsb",
    "load_xlsx_styled", "load_xlsb_styled", "save_loaded_xlsb_styled",
    "write_model_xlsx", "write_model_xlsb",
    "write_formula_hyperlink_xlsx",
    "read_xlsx_entities", "load_xlsx_entities",
    "xlsx_fallback_rate", "xlsx_fallback_rate_model",
    "xlsx_entities_fallback_rate", "xlsx_entities_fallback_rate_model",
]

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
SHEET_PART = "xl/worksheets/sheet1.xml"


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


def formula_hyperlink_rows(count, cols):
    """Every fourth cell is a Formula, every fourth a Hyperlink; the rest are numbers and text."""
    from xlspy import Formula, Hyperlink
    yield [f"Header_{j}" for j in range(cols)]
    for i in range(1, count):
        out = []
        for j in range(cols):
            kind = j % 4
            if kind == 0:
                out.append(Formula(f"{i}+{j}", value=i + j))
            elif kind == 1:
                out.append(Hyperlink(f"https://example.com/item/{i}/{j}", text=f"Item {i}"))
            elif kind == 2:
                out.append(i * j)
            else:
                out.append(f"String_{i % 5000}_{j}")
        yield out


def _entity_cell(ref, j, i):
    """One cell whose text needs XML decoding (entities, CR/LF, character references)."""
    kind = j % 5
    if kind == 0:
        return f'<c r="{ref}" t="inlineStr"><is><t>a &amp; b &lt; c {i}</t></is></c>'
    if kind == 1:
        return f'<c r="{ref}"><v>{i * j}</v></c>'
    if kind == 2:
        # A literal CR LF (normalized to LF by the XML parser) and a kept carriage return.
        return f'<c r="{ref}" t="inlineStr"><is><t>line {i}\r\nnext &#13;x</t></is></c>'
    if kind == 3:
        return f'<c r="{ref}" t="str"><v>x &lt; {i} &amp; y</v></c>'
    return f'<c r="{ref}" t="d"><v>2023-01-{i % 28 + 1:02d}T00:00:00</v></c>'


def write_entity_xlsx(path, count, cols):
    """An XLSX whose worksheet holds the cells of ``_entity_cell``.

    The package is a normal xlspy file with its worksheet part replaced, so the
    styles and the other parts are valid.
    """
    from xlspy import XlsxWriter
    from xlspy.refs import column_label

    template = path + ".template"
    with XlsxWriter(template) as writer:
        writer.add_sheet("S")
        writer.write_sheet([["x"]])
    labels = [column_label(j) for j in range(cols)]
    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<worksheet xmlns="{MAIN_NS}"><sheetData>'
    ]
    header = "".join(
        f'<c r="{labels[j]}1" t="inlineStr"><is><t>Header &amp; {j}</t></is></c>' for j in range(cols)
    )
    parts.append(f'<row r="1">{header}</row>')
    for i in range(2, count + 1):
        cells = "".join(_entity_cell(f"{labels[j]}{i}", j, i) for j in range(cols))
        parts.append(f'<row r="{i}">{cells}</row>')
    parts.append("</sheetData></worksheet>")
    sheet = "".join(parts).encode("utf-8")
    with zipfile.ZipFile(template) as zin, zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = sheet if item.filename == SHEET_PART else zin.read(item.filename)
            zout.writestr(item, data)
    os.remove(template)


def rss_mb():
    import psutil
    return psutil.Process().memory_info().rss / (1024 * 1024)


def accel_info():
    from xlspy import _accel
    return {
        "mode": _accel.ACCEL_MODE,
        "have_native": _accel.HAVE_NATIVE,
        "backends": dict(_accel.ACCEL_BACKEND),
    }


def _write_rows(path, kind, data, level=None):
    from xlspy import XlsbWriter, XlsxWriter
    start = time.perf_counter()
    if kind == "xlsx":
        kwargs = {} if level is None else {"compressionLevel": level}
        with XlsxWriter(path, **kwargs) as writer:
            writer.add_sheet("S")
            writer.write_sheet(data)
    else:
        with XlsbWriter(path) as writer:
            writer.add_sheet("S")
            writer.write_sheet(data)
    return time.perf_counter() - start


def _write(path, kind, count, cols, level=None):
    return _write_rows(path, kind, rows(count, cols), level)


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


def _fallback_rate(path, model):
    """Read ``path`` and count the rows the reader hands to ElementTree.

    ``model`` counts ``workbook._fast_sheet._load_row`` (``load_workbook``);
    otherwise ``_xlsx_read._row_from_element`` (``ExcelReader``). The counted
    function is looked up as a module global at call time, so patching the
    module attribute is enough.
    """
    from xlspy import ExcelReader, _xlsx_read, load_workbook
    from xlspy.workbook import _fast_sheet

    module, name = (_fast_sheet, "_load_row") if model else (_xlsx_read, "_row_from_element")
    original = getattr(module, name)
    calls = {"n": 0}

    def counting(*args, **kwargs):
        calls["n"] += 1
        return original(*args, **kwargs)

    setattr(module, name, counting)
    try:
        if model:
            workbook = load_workbook(path)
            try:
                total = workbook["S"].max_row
            finally:
                workbook.close()
        else:
            total = 0
            with ExcelReader(path) as reader:
                for _ in reader.get_rows("S"):
                    total += 1
    finally:
        setattr(module, name, original)
    return {"rows": total, "fallback_rows": calls["n"],
            "fallback_pct": 100.0 * calls["n"] / max(total, 1)}


def _timed_read(path, reader_kind):
    """Time a full pass over the rows; the cell count is taken in an untimed pass."""
    if reader_kind == "reader":
        from xlspy import ExcelReader
        opened = ExcelReader
    else:
        from xlspy import load_workbook
        opened = load_workbook
    before = rss_mb()
    start = time.perf_counter()
    total = 0
    if reader_kind == "reader":
        with opened(path) as reader:
            for row in reader.get_rows("S"):
                total += len(row)
    else:
        workbook = opened(path)
        try:
            for row in workbook["S"].iter_rows(values_only=True):
                total += len(row)
        finally:
            workbook.close()
    seconds = time.perf_counter() - start
    return {"seconds": seconds, "cells": total, "before_rss_mb": before}


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
    elif case == "write_formula_hyperlink_xlsx":
        result["seconds"] = _write_rows(path, "xlsx", formula_hyperlink_rows(count, cols))
        result["size_mb"] = os.path.getsize(path) / 1048576
    elif case in ("read_xlsx", "read_xlsb"):
        _write(path, case[5:], count, cols)
        timed = _timed_read(path, "reader")
        before = timed.pop("before_rss_mb")
        result.update(timed)
    elif case == "read_xlsx_entities":
        write_entity_xlsx(path, count, cols)
        timed = _timed_read(path, "reader")
        before = timed.pop("before_rss_mb")
        result.update(timed)
    elif case == "load_xlsx_entities":
        # Load only, as load_xlsx does, so the two cases are comparable.
        from xlspy import load_workbook
        write_entity_xlsx(path, count, cols)
        before = rss_mb()
        start = time.perf_counter()
        workbook = load_workbook(path)
        result["seconds"] = time.perf_counter() - start
        workbook.close()
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
    elif case in ("xlsx_fallback_rate", "xlsx_fallback_rate_model",
                  "xlsx_entities_fallback_rate", "xlsx_entities_fallback_rate_model"):
        if case.startswith("xlsx_entities"):
            write_entity_xlsx(path, count, cols)
        else:
            _write(path, "xlsx", count, cols)
        result.update(_fallback_rate(path, model=case.endswith("_model")))
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


def _iqr(values):
    if len(values) < 2:
        return 0.0
    q1, _, q3 = statistics.quantiles(values, n=4, method="inclusive")
    return q3 - q1


def parent(args):
    results = {}
    cases = list(CASES) + [f"zlib{level}" for level in (1, 3, 6, 9)]
    if args.only:
        wanted = [part.strip() for part in args.only.split(",") if part.strip()]
        cases = [case for case in cases if any(part in case for part in wanted)]
    with tempfile.TemporaryDirectory(prefix="xlspy-bench-") as directory:
        for case in cases:
            runs = []
            is_rate = "fallback_rate" in case
            for _ in range(1 if is_rate else args.repeat):
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
                    seconds = [r["seconds"] for r in good]
                    merged["seconds"] = statistics.median(seconds)
                    merged["iqr"] = _iqr(seconds)
                    merged["runs"] = [round(s, 3) for s in seconds]
                results[case] = merged
            print(f"{case:34s} {results[case]}", flush=True)
    document = {
        "python": sys.version.split()[0],
        "platform": platform.platform(),
        "rows": args.rows, "cols": args.cols, "repeat": args.repeat,
        "accel": accel_info(),
        "results": results,
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
            print(f"{case:34s} {ratio:5.2f}x  {flag}")
    return 1 if failed else 0


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--rows", type=int, default=50000)
    parser.add_argument("--cols", type=int, default=20)
    parser.add_argument("--repeat", type=int, default=7)
    parser.add_argument("--only", help="comma-separated substrings; run only matching cases")
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
