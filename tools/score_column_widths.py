"""Score column-width estimators against Excel AutoFit widths.

    python tools/score_column_widths.py tmp/widths/corpus.json

For each estimator: how many cells come out narrower than Excel's AutoFit
(the text would be cut or show ####), the worst shortfall, and the mean
excess width. ``old`` is the 0.8.1 XLSX formula, ``old_xlsb`` the XLSB one.
"""

import datetime
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.normpath(os.path.join(HERE, "..", "src")))

from xlspy._widths import cell_text, cell_width  # noqa: E402


def _value(record):
    value, kind = record["value"], record["python_type"]
    if kind == "datetime":
        return datetime.datetime.fromisoformat(value)
    if kind == "date":
        return datetime.date.fromisoformat(value)
    return value


def old(value, fmt, bold):
    if isinstance(value, datetime.datetime):
        return 18.0
    if isinstance(value, datetime.date):
        return 10.14
    return max(8.43, len(str(value)) * 1.25 + 2)


def old_xlsb(value, fmt, bold):
    if isinstance(value, datetime.datetime):
        return 18.0
    return int(1.25 * len(str(value).split("\n")[0]) + 4)


def new(value, fmt, bold):
    return max(8.43, cell_width(value, fmt, bold=bold))


def new_raw(value, fmt, bold):
    return cell_width(value, fmt, bold=bold)


def main(path):
    records = json.load(open(path, encoding="utf-8"))
    print(f"{'estimator':10s} {'too narrow':>11s} {'worst':>7s} {'mean excess':>12s}")
    for name, estimator in (("old", old), ("old_xlsb", old_xlsb), ("new", new), ("new_raw", new_raw)):
        short, worst, excess = 0, 0.0, []
        for record in records:
            width = estimator(_value(record), record["number_format"], record["bold"])
            gap = width - record["excel_width"]
            if gap < -0.05:
                short += 1
                worst = min(worst, gap)
            excess.append(gap)
        print(f"{name:10s} {short:5d}/{len(records):<5d} {worst:7.2f} {sum(excess) / len(excess):12.2f}")
    if "-v" in sys.argv:
        for record in records:
            value = _value(record)
            est = new_raw(value, record["number_format"], record["bold"])
            flag = "SHORT" if est < record["excel_width"] - 0.05 else ""
            print(f"{record['excel_width']:6.2f} {est:6.2f} {flag:5s} "
                  f"{cell_text(value, record['number_format'])!r:40.40s} excel={record['excel_text']!r}")


if __name__ == "__main__":
    main(sys.argv[1])
