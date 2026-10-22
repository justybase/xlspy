"""Benchmark styled writes: time per scenario, writer and C-extension mode.

Every run happens in a fresh worker process (this script re-invoked with
``--worker``). Rows and styles are built before the clock starts; the timed
part is opening the writer, ``write_sheet`` (or ``Workbook.save``) and closing
it. Each scenario repeats ``--iterations`` times, interleaved, and the report
gives the median and minimum. ``--tracemalloc`` adds one untimed run per
scenario for the peak of Python allocations.

``--keep DIR`` keeps one output file per scenario, so the files of two code
versions can be compared with ``tools/compare_xlsx_parts.py``.

Usage:
    python tools/bench_styles.py --rows 10000 --output tmp/bench_styles_base.json
    python tools/bench_styles.py --compare tmp/bench_styles_base.json tmp/bench_styles_a1.json
"""

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

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_SRC = os.path.normpath(os.path.join(HERE, "..", "src"))
COLS = 12


def _rows(n):
    base = datetime.date(2024, 1, 1)
    return [
        [f"item-{r}", r, r * 1.5, base + datetime.timedelta(days=r % 3650)]
        + [r * c for c in range(4, COLS)]
        for r in range(n)
    ]


def _letter(col):
    from xlspy.refs import column_label
    return column_label(col)


def _scenarios():
    """name -> builder(rows_count) returning (kind, payload)."""
    from xlspy import (
        Alignment, Border, Font, Group, PatternFill, Side, Style,
    )
    header = Style(fill="1F4E78", font_color="FFFFFF", bold=True)
    bold = Style(font=Font(bold=True))
    boxed = Style(border=Border(left=Side(style="thin"), right=Side(style="thin")))
    centered = Style(alignment=Alignment(horizontal="center"))
    money = Style(number_format="#,##0.00")
    zebra = Style(pattern=PatternFill(fill_type="solid", start_color="DDEBF7"))

    def plain(n):
        return "sheet", (_rows(n), {})

    def header_only(n):
        return "sheet", (_rows(n), {"header_style": header})

    def column_range(n):
        return "sheet", (_rows(n), {"header_style": header,
                                    "styles": {f"B2:C{n}": money}})

    def overlapping(n):
        return "sheet", (_rows(n), {"header_style": header, "styles": {
            f"A1:L{n}": boxed,
            f"B2:D{n}": money,
            f"C2:C{n // 2}": bold,
            f"A{n // 4}:L{n // 4 + 50}": centered,
        }})

    def fmt_tuples(n):
        rows = [[r[0]] + [(v, "#,##0.00") for v in r[1:3]] + [(r[3], "yyyy-mm-dd")] + r[4:]
                for r in _rows(n)]
        return "sheet", (rows, {})

    def fmt_and_range(n):
        rows = [[r[0]] + [(v, "#,##0.00") for v in r[1:3]] + [(r[3], "yyyy-mm-dd")] + r[4:]
                for r in _rows(n)]
        return "sheet", (rows, {"header_style": header,
                                "styles": {f"A1:L{n}": boxed}})

    def per_cell(n):
        styles = {}
        for r in range(n):
            for c in range(0, COLS, 2):
                styles[f"{_letter(c)}{r + 1}"] = bold if r % 2 else boxed
        return "sheet", (_rows(n), {"styles": styles})

    def per_cell_tuples(n):
        # per_cell written with (value, Style) cells instead of 6 ranges per row.
        rows = [[(v, bold if r % 2 else boxed) if c % 2 == 0 else v
                 for c, v in enumerate(row)] for r, row in enumerate(_rows(n))]
        return "sheet", (rows, {})

    def column_open(n):
        # column_range with an open range ("B2:C") instead of "B2:C{n}".
        return "sheet", (_rows(n), {"header_style": header, "styles": {"B2:C": money}})

    def zebra_rows(n):
        styles = {f"A{r + 1}:L{r + 1}": zebra for r in range(1, n, 2)}
        return "sheet", (_rows(n), {"header_style": header, "styles": styles})

    def range_heights(n):
        return "sheet", (_rows(n), {
            "header_style": header,
            "styles": {f"A1:L{n}": boxed},
            "row_heights": {r: 18 for r in range(0, n, 3)},
            "row_groups": [Group(r, r + 4) for r in range(10, n - 10, 20)],
        })

    def heatmap(n):
        # Up to 4096 distinct fills: one style per (r, c) colour bucket.
        styles = {}
        palette = [Style(fill=f"{(i * 37) % 256:02X}{(i * 91) % 256:02X}{(i * 13) % 256:02X}")
                   for i in range(4096)]
        for r in range(min(n, 1000)):
            for c in range(COLS):
                styles[f"{_letter(c)}{r + 1}"] = palette[(r * COLS + c) % len(palette)]
        return "sheet", (_rows(n), {"styles": styles})

    def workbook_cells(n):
        from xlspy.workbook import Workbook
        wb = Workbook()
        ws = wb.active
        for row in _rows(n):
            ws.append(row)
        font = Font(bold=True)
        for r, row in enumerate(ws.iter_rows(), start=1):
            for c, cell in enumerate(row):
                if c % 2 == 0:
                    cell.font = font
                    if r % 2:
                        cell.number_format = "#,##0.00"
        return "workbook", wb

    def dataframe_formats(n):
        import pandas as pd
        df = pd.DataFrame(_rows(n), columns=[f"c{i}" for i in range(COLS)])
        return "dataframe", (df, {"formats": {"c1": "#,##0", "c2": "#,##0.00"},
                                  "header_style": header})

    return {
        "plain": plain,
        "header": header_only,
        "column_range": column_range,
        "overlapping": overlapping,
        "fmt_tuples": fmt_tuples,
        "fmt_and_range": fmt_and_range,
        "per_cell": per_cell,
        "per_cell_tuples": per_cell_tuples,
        "column_open": column_open,
        "zebra_rows": zebra_rows,
        "range_heights": range_heights,
        "heatmap": heatmap,
        "workbook_cells": workbook_cells,
        "dataframe_formats": dataframe_formats,
    }


def _write(kind, payload, path, writer_cls):
    if kind == "workbook":
        payload.save(path)
        return
    with writer_cls(path) as writer:
        writer.add_sheet("Data")
        if kind == "sheet":
            rows, options = payload
            writer.write_sheet(rows, **options)
        else:
            df, options = payload
            writer.write_dataframe(df, **options)


def worker(args):
    import tracemalloc
    from xlspy import XlsbWriter, XlsxWriter
    from xlspy import _accel
    writer_cls = XlsxWriter if args.format == "xlsx" else XlsbWriter
    kind, payload = _scenarios()[args.scenario](args.rows)
    out_dir = args.keep or tempfile.mkdtemp(prefix="xlspy-bench-")
    os.makedirs(out_dir, exist_ok=True)
    path = os.path.join(out_dir, f"{args.scenario}.{args.format}")
    if args.trace:
        tracemalloc.start()
    started = time.perf_counter()
    _write(kind, payload, path, writer_cls)
    elapsed = time.perf_counter() - started
    peak = tracemalloc.get_traced_memory()[1] if args.trace else None
    result = {
        "elapsed_s": elapsed,
        "tracemalloc_mb": peak / 2**20 if peak is not None else None,
        "file_bytes": os.path.getsize(path),
        "c_ext": _accel.HAVE_C_EXT,
    }
    if not args.keep:
        os.remove(path)
    print(json.dumps(result))


def _run_worker(args, scenario, fmt, c_ext, trace=False, keep=None):
    env = os.environ.copy()
    env["PYTHONPATH"] = args.src + os.pathsep + env.get("PYTHONPATH", "")
    env["XLSPY_DISABLE_C_EXT"] = "0" if c_ext else "1"
    cmd = [sys.executable, os.path.abspath(__file__), "--worker",
           "--scenario", scenario, "--format", fmt, "--rows", str(args.rows)]
    if trace:
        cmd.append("--trace")
    if keep:
        cmd += ["--keep", keep]
    done = subprocess.run(cmd, env=env, capture_output=True, text=True)
    if done.returncode != 0:
        return {"error": done.stderr.strip().splitlines()[-1] if done.stderr else "failed"}
    return json.loads(done.stdout.strip().splitlines()[-1])


def orchestrate(args):
    scenarios = args.scenarios.split(",") if args.scenarios else list(_scenarios_names())
    formats = args.formats.split(",")
    modes = {"on": [True], "off": [False], "both": [True, False]}[args.c_ext]
    runs = {}
    for iteration in range(args.iterations):
        for scenario in scenarios:
            for fmt in formats:
                for c_ext in modes:
                    key = f"{scenario}/{fmt}/{'c' if c_ext else 'py'}"
                    keep = None
                    if args.keep and iteration == 0:
                        keep = os.path.join(args.keep, "c" if c_ext else "py")
                    result = _run_worker(args, scenario, fmt, c_ext, keep=keep)
                    runs.setdefault(key, []).append(result)
                    if "error" in result:
                        print(f"{key}: ERROR {result['error']}", file=sys.stderr)
    summary = {}
    for key, results in runs.items():
        times = [r["elapsed_s"] for r in results if "elapsed_s" in r]
        if not times:
            summary[key] = {"error": results[0].get("error")}
            continue
        summary[key] = {
            "median_s": statistics.median(times),
            "min_s": min(times),
            "file_bytes": results[0]["file_bytes"],
        }
    if args.tracemalloc:
        for scenario in scenarios:
            for fmt in formats:
                for c_ext in modes:
                    key = f"{scenario}/{fmt}/{'c' if c_ext else 'py'}"
                    result = _run_worker(args, scenario, fmt, c_ext, trace=True)
                    if "tracemalloc_mb" in result and key in summary:
                        summary[key]["tracemalloc_mb"] = result["tracemalloc_mb"]
    report = {
        "label": args.label,
        "rows": args.rows,
        "cols": COLS,
        "iterations": args.iterations,
        "python": sys.version,
        "platform": platform.platform(),
        "summary": summary,
        "runs": runs,
    }
    for key, value in summary.items():
        if "error" in value:
            print(f"{key:36s} ERROR {value['error']}")
        else:
            mem = value.get("tracemalloc_mb")
            mem_text = f"  {mem:7.1f} MB" if mem is not None else ""
            print(f"{key:36s} {value['median_s']:8.4f}s (min {value['min_s']:.4f})"
                  f"  {value['file_bytes']:>10,d} B{mem_text}")
    if args.output:
        os.makedirs(os.path.dirname(os.path.abspath(args.output)), exist_ok=True)
        with open(args.output, "w", encoding="utf-8") as handle:
            json.dump(report, handle, indent=1)


def _scenarios_names():
    sys.path.insert(0, DEFAULT_SRC)
    return _scenarios().keys()


def compare(before_path, after_path):
    with open(before_path, encoding="utf-8") as handle:
        before = json.load(handle)["summary"]
    with open(after_path, encoding="utf-8") as handle:
        after = json.load(handle)["summary"]
    print(f"{'scenario':36s} {'before':>9s} {'after':>9s} {'delta':>8s} {'speedup':>8s}  size")
    for key in before:
        b, a = before[key], after.get(key)
        if a is None or "median_s" not in b or "median_s" not in a:
            print(f"{key:36s} (missing or failed)")
            continue
        delta = (a["median_s"] - b["median_s"]) / b["median_s"] * 100
        size = "same" if a["file_bytes"] == b["file_bytes"] else f"{a['file_bytes'] - b['file_bytes']:+,d} B"
        print(f"{key:36s} {b['median_s']:9.4f} {a['median_s']:9.4f} {delta:+7.1f}% "
              f"{b['median_s'] / a['median_s']:7.2f}x  {size}")


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--worker", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--scenario", help=argparse.SUPPRESS)
    parser.add_argument("--format", help=argparse.SUPPRESS)
    parser.add_argument("--trace", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--rows", type=int, default=10000)
    parser.add_argument("--iterations", type=int, default=5)
    parser.add_argument("--scenarios", help="comma-separated subset")
    parser.add_argument("--formats", default="xlsx,xlsb")
    parser.add_argument("--c-ext", choices=("on", "off", "both"), default="on")
    parser.add_argument("--tracemalloc", action="store_true")
    parser.add_argument("--keep", help="directory for one output file per scenario")
    parser.add_argument("--label", default="")
    parser.add_argument("--output")
    parser.add_argument("--src", default=DEFAULT_SRC)
    parser.add_argument("--compare", nargs=2, metavar=("BEFORE", "AFTER"))
    args = parser.parse_args()
    if args.compare:
        compare(*args.compare)
    elif args.worker:
        worker(args)
    else:
        orchestrate(args)


if __name__ == "__main__":
    main()
