"""Shared-strings micro-benchmark: how fast the sharedStrings part is parsed.

One million short strings, three measurements:

  regex   ``reader._parse_shared_strings_fast`` on the decoded part (plain markup).
  reader  ``ExcelReader`` open on a file with those strings, the whole cost a user pays.
  expat   ``ExcelReader`` open on a file where one string holds a character
          reference, which sends the whole part to the XML parser.

When the native parser is built, ``native`` reports ``_accel.read_xlsx_shared_strings``
on the same bytes, and ``reader`` then uses it too.

Usage (from the repository root):
    py -3.14 benchmarks/bench_sst.py [--strings 1000000] [--repeat 7] [--out benchmarks/sst.json]
"""
from __future__ import annotations

import argparse
import json
import os
import platform
import statistics
import sys
import tempfile
import time
import zipfile

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, os.path.join(ROOT, "src"))

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
SST_PART = "xl/sharedStrings.xml"


def sst_xml(count, entity=False):
    """The sharedStrings part: ``count`` short plain strings, or with one character reference."""
    items = [f"<si><t>s{i}</t></si>" for i in range(count)]
    if entity:
        items[0] = "<si><t>&#65;</t></si>"
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<sst xmlns="{MAIN_NS}" count="{count}" uniqueCount="{count}">' + "".join(items) + "</sst>")


def make_book(path, count, entity):
    """A workbook whose sharedStrings part is replaced by the generated one."""
    from xlspy import XlsxWriter

    template = path + ".template"
    with XlsxWriter(template) as writer:
        writer.add_sheet("S")
        writer.write_sheet([["x"]])
    data = sst_xml(count, entity).encode("utf-8")
    with zipfile.ZipFile(template) as zin, zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            zout.writestr(item, data if item.filename == SST_PART else zin.read(item.filename))
    os.remove(template)


def _iqr(values):
    if len(values) < 2:
        return 0.0
    q1, _, q3 = statistics.quantiles(values, n=4, method="inclusive")
    return q3 - q1


def timed(function, repeat):
    runs = []
    for _ in range(repeat):
        start = time.perf_counter()
        function()
        runs.append(time.perf_counter() - start)
    return {"seconds": statistics.median(runs), "iqr": _iqr(runs), "runs": [round(r, 4) for r in runs]}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--strings", type=int, default=1_000_000)
    parser.add_argument("--repeat", type=int, default=7)
    parser.add_argument("--out")
    args = parser.parse_args()

    from xlspy import ExcelReader, _accel, reader as reader_module

    content = sst_xml(args.strings)
    results = {}
    parsed = reader_module._parse_shared_strings_fast(content)
    if parsed is None or len(parsed) != args.strings:
        raise SystemExit("regex parser did not take the plain part; the benchmark is invalid")
    results["regex"] = timed(lambda: reader_module._parse_shared_strings_fast(content), args.repeat)
    native = getattr(_accel, "read_xlsx_shared_strings", None)
    if native is not None:
        raw = content.encode("utf-8")
        check = native(raw)
        if check is None or len(check) != args.strings:
            raise SystemExit("native parser did not take the plain part; the benchmark is invalid")
        results["native"] = timed(lambda: native(raw), args.repeat)

    with tempfile.TemporaryDirectory(prefix="xlspy-sst-bench-") as directory:
        plain = os.path.join(directory, "plain.xlsx")
        expat = os.path.join(directory, "expat.xlsx")
        make_book(plain, args.strings, entity=False)
        make_book(expat, args.strings, entity=True)

        def open_plain():
            with ExcelReader(plain):
                pass

        def open_expat():
            with ExcelReader(expat):
                pass

        results["reader"] = timed(open_plain, args.repeat)
        results["expat"] = timed(open_expat, args.repeat)

    for name, result in results.items():
        print(f"{name:8s} {result['seconds']:8.3f} s  (iqr {result['iqr']:.3f})", flush=True)
    document = {
        "python": sys.version.split()[0],
        "platform": platform.platform(),
        "strings": args.strings,
        "repeat": args.repeat,
        "accel": {"mode": _accel.ACCEL_MODE, "have_native": _accel.HAVE_NATIVE,
                  "backends": dict(_accel.ACCEL_BACKEND)},
        "results": results,
    }
    if args.out:
        with open(args.out, "w", encoding="utf-8") as handle:
            json.dump(document, handle, indent=2)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
