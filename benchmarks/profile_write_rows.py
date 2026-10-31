"""Where XLSX row writing spends its time, to decide whether batching rows pays (C4).

Reports the share of ``XlsxWriter.write_sheet`` spent in the native row encoder,
in the per-row Python glue around it (row attributes, style ids, the loop), and in
the rest (streams, shared strings, the zip). The shares come from a cProfile run,
whose overhead is larger for small calls, so the unprofiled time is reported
next to them.

Usage (from the repository root):
    py -3.14 benchmarks/profile_write_rows.py [--rows 50000] [--cols 20] [--repeat 3]
"""
from __future__ import annotations

import argparse
import cProfile
import json
import os
import pstats
import statistics
import sys
import tempfile
import time

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, os.path.join(ROOT, "src"))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Glue: the Python work done once per row, outside the native encoder and the cell writer.
GLUE = ("_row_attrs", "styles_for_row", "_style_ids_for_row")


VARIANTS = ("plain", "formula", "formula_as_values", "hyperlink", "hyperlink_as_text")


def variant_rows(count, cols, variant):
    """Rows of the same shape for every variant: every fourth cell is special.

    ``formula`` and ``hyperlink`` put a Formula or a Hyperlink there; the ``*_as_*``
    variants put the plain value the special cell stands for, so the difference
    between a variant and its plain twin is the cost of the special cell's row path.
    """
    from xlspy import Formula, Hyperlink
    yield [f"Header_{j}" for j in range(cols)]
    for i in range(1, count):
        out = []
        for j in range(cols):
            kind = j % 4
            if kind == 0 and variant == "formula":
                out.append(Formula(f"{i}+{j}", value=i + j))
            elif kind == 0 and variant == "formula_as_values":
                out.append(i + j)
            elif kind == 1 and variant == "hyperlink":
                out.append(Hyperlink(f"https://example.com/item/{i}/{j}", text=f"Item {i}"))
            elif kind == 1 and variant == "hyperlink_as_text":
                out.append(f"Item {i}")
            elif kind == 2:
                out.append(i * j)
            else:
                out.append(f"String_{i % 5000}_{j}")
        yield out


def _write(path, data):
    from xlspy import XlsxWriter
    with XlsxWriter(path) as writer:
        writer.add_sheet("S")
        writer.write_sheet(data)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--rows", type=int, default=50000)
    parser.add_argument("--cols", type=int, default=20)
    parser.add_argument("--repeat", type=int, default=3)
    parser.add_argument("--data", choices=VARIANTS, default="plain",
                        help="a row shape from variant_rows; 'plain' is bench_io's mixed rows")
    parser.add_argument("--variants", action="store_true",
                        help="time every variant without a profiler instead of profiling one")
    parser.add_argument("--no-native-rows", action="store_true",
                        help="write every row with the Python encoder (the pre-native baseline "
                             "for rows that hold Formula or Hyperlink cells)")
    parser.add_argument("--out")
    args = parser.parse_args()

    from bench_io import rows
    from xlspy import _accel, xlsx_writer

    if args.no_native_rows:
        xlsx_writer.accel_encode_xlsx_row = lambda *a, **k: None

    def data():
        if args.data == "plain":
            return rows(args.rows, args.cols)
        return variant_rows(args.rows, args.cols, args.data)

    # The rows are built before any timing or profiling: building Formula and
    # Hyperlink objects (Formula normalizes its text) is not part of writing.
    if args.variants:
        timings = {}
        with tempfile.TemporaryDirectory(prefix="xlspy-profile-") as directory:
            path = os.path.join(directory, "v.xlsx")
            for variant in VARIANTS:
                built = list(variant_rows(args.rows, args.cols, variant))
                runs = []
                for _ in range(args.repeat):
                    start = time.perf_counter()
                    _write(path, built)
                    runs.append(time.perf_counter() - start)
                timings[variant] = {"median_seconds": statistics.median(runs),
                                    "runs": [round(r, 3) for r in runs]}
                print(f"{variant:20s} {timings[variant]['median_seconds']:7.3f} s", flush=True)
        if args.out:
            with open(args.out, "w", encoding="utf-8") as handle:
                json.dump({"rows": args.rows, "cols": args.cols, "repeat": args.repeat,
                           "accel_backends": dict(_accel.ACCEL_BACKEND), "variants": timings},
                          handle, indent=2)
        return 0

    built = list(data())
    with tempfile.TemporaryDirectory(prefix="xlspy-profile-") as directory:
        path = os.path.join(directory, "p.xlsx")
        runs = []
        for _ in range(args.repeat):
            start = time.perf_counter()
            _write(path, built)
            runs.append(time.perf_counter() - start)
        unprofiled = statistics.median(runs)

        profile = cProfile.Profile()
        profile.enable()
        _write(path, built)
        profile.disable()

    stats = pstats.Stats(profile).stats
    total = None
    body = {"tottime": 0.0, "cumtime": 0.0}
    native = 0.0
    glue = {name: 0.0 for name in GLUE}
    python_cells = 0.0  # _write_row_cells: the Python cell writer, for rows the native encoder hands back
    formula_cells = 0.0
    for (filename, _line, name), (_cc, _nc, tottime, cumtime, _callers) in stats.items():
        if name == "write_sheet":
            total = cumtime
        if name == "_write_body" and filename.endswith("xlsx_writer.py"):
            body["tottime"] += tottime
            body["cumtime"] += cumtime
        if name == "encode_xlsx_row":
            native += tottime
        if name in glue and filename.endswith("xlsx_writer.py"):
            glue[name] += cumtime
        if name == "_write_row_cells" and filename.endswith("xlsx_writer.py"):
            python_cells += cumtime
        if name == "_write_formula_cell" and filename.endswith("xlsx_writer.py"):
            formula_cells += cumtime
    if total is None:
        raise SystemExit("write_sheet not found in the profile")

    glue_total = body["tottime"] + sum(glue.values())
    result = {
        "data": args.data,
        "native_rows": not args.no_native_rows,
        "rows": args.rows, "cols": args.cols, "repeat": args.repeat,
        "accel_backends": dict(_accel.ACCEL_BACKEND),
        "unprofiled_seconds": unprofiled,
        "profiled_write_sheet_seconds": total,
        "native_row_encoder_seconds": native,
        "body_loop_self_seconds": body["tottime"],
        "row_glue_seconds": glue,
        "glue_share_of_write": glue_total / total,
        "native_share_of_write": native / total,
        "python_cell_writer_share_of_write": python_cells / total,
        "formula_cell_writer_share_of_write": formula_cells / total,
    }
    print(json.dumps(result, indent=2))
    if args.out:
        with open(args.out, "w", encoding="utf-8") as handle:
            json.dump(result, handle, indent=2)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
