"""Compare an xlspy-loaded workbook with an openpyxl-loaded one semantically.

Usage:
    python tools/compare_openpyxl_model.py LEFT RIGHT [--json OUT] [--allow-missing]

LEFT is loaded with ``xlspy.load_workbook``; RIGHT with
``openpyxl.load_workbook``.  The tool compares values, number formats, style
components, merges and the object counts that the model exposes.  It exits 1
when a difference is found; missing xlspy attributes are reported separately
and do not fail the run with ``--allow-missing``.
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

import openpyxl  # noqa: E402

from xlspy import load_workbook  # noqa: E402


def _norm(value):
    if isinstance(value, bool) or value is None:
        return value
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, (dt.datetime, dt.date, dt.time)):
        return value.isoformat()
    if isinstance(value, dt.timedelta):
        return value.total_seconds()
    return value


def _color(value):
    if value is None or isinstance(value, str):
        return value
    return getattr(value, "rgb", None) or getattr(value, "theme", None) or str(value)


def _style_facts(cell):
    facts = {"number_format": getattr(cell, "number_format", None)}
    font = getattr(cell, "font", None)
    if font is not None:
        facts["font"] = {
            "name": getattr(font, "name", None),
            "size": getattr(font, "size", None),
            "bold": getattr(font, "bold", None),
            "italic": getattr(font, "italic", None),
            "underline": getattr(font, "underline", None),
            "color": _color(getattr(font, "color", None)),
        }
    fill = getattr(cell, "fill", None)
    if fill is not None:
        facts["fill"] = {
            "type": getattr(fill, "fill_type", None) or getattr(fill, "patternType", None),
            "fg": _color(getattr(fill, "fgColor", None) or getattr(fill, "start_color", None)),
            "bg": _color(getattr(fill, "bgColor", None) or getattr(fill, "end_color", None)),
        }
    border = getattr(cell, "border", None)
    if border is not None:
        facts["border"] = {
            side: getattr(getattr(border, side, None), "style", None)
            for side in ("left", "right", "top", "bottom")
        }
    alignment = getattr(cell, "alignment", None)
    if alignment is not None:
        facts["alignment"] = {
            name: getattr(alignment, name, None)
            for name in ("horizontal", "vertical", "wrap_text", "shrink_to_fit",
                         "indent", "text_rotation")
        }
    protection = getattr(cell, "protection", None)
    if protection is not None:
        facts["protection"] = {
            name: getattr(protection, name, None) for name in ("locked", "hidden")
        }
    return facts


def _objects(worksheet):
    result = {}
    for name in ("tables", "data_validations", "conditional_formatting", "charts", "images"):
        if not hasattr(worksheet, name):
            result[name] = "<missing>"
            continue
        try:
            value = getattr(worksheet, name)
            if name == "data_validations":
                value = list(value)
            elif name == "conditional_formatting":
                value = list(value)
            elif name in ("charts", "images"):
                value = list(value)
            result[name] = len(value)
        except Exception as exc:  # pragma: no cover - diagnostic tool
            result[name] = f"<error: {type(exc).__name__}: {exc}>"
    return result


def compare(left_path, right_path, *, allow_missing=False):
    xwb = load_workbook(str(left_path))
    owb = openpyxl.load_workbook(str(right_path), data_only=False)
    differences = []
    missing = []

    xsheets = list(xwb.sheetnames)
    osheets = list(owb.sheetnames)
    if xsheets != osheets:
        differences.append({
            "kind": "sheet_names",
            "left": xsheets,
            "right": osheets,
        })

    for name in osheets:
        if name not in xsheets:
            continue
        xs = xwb[name]
        osheet = owb[name]
        if not hasattr(osheet, "iter_rows"):
            missing.append({"sheet": name, "kind": "chartsheet"})
            continue

        xmax_row = getattr(xs, "max_row", 0)
        xmax_col = getattr(xs, "max_column", 0)
        omax_row = osheet.max_row
        omax_col = osheet.max_column
        if (xmax_row, xmax_col) != (omax_row, omax_col):
            differences.append({
                "sheet": name,
                "kind": "dimensions",
                "left": [xmax_row, xmax_col],
                "right": [omax_row, omax_col],
            })

        for row in range(1, omax_row + 1):
            for column in range(1, omax_col + 1):
                try:
                    xcell = xs.cell(row, column)
                except Exception as exc:
                    missing.append({
                        "sheet": name, "kind": "cell",
                        "where": f"R{row}C{column}",
                        "error": f"{type(exc).__name__}: {exc}",
                    })
                    continue
                ocell = osheet.cell(row, column)
                xvalue = _norm(getattr(xcell, "value", None))
                ovalue = _norm(ocell.value)
                if xvalue != ovalue:
                    differences.append({
                        "sheet": name, "kind": "value",
                        "where": f"R{row}C{column}",
                        "left": xvalue, "right": ovalue,
                    })
                xfacts = _style_facts(xcell)
                ofacts = _style_facts(ocell)
                if xfacts != ofacts:
                    differences.append({
                        "sheet": name, "kind": "style",
                        "where": f"R{row}C{column}",
                        "left": xfacts, "right": ofacts,
                    })

        xmerge = str(getattr(xs, "merged_cells", ""))
        omerge = str(osheet.merged_cells)
        if xmerge != omerge:
            differences.append({
                "sheet": name, "kind": "merged_cells",
                "left": xmerge, "right": omerge,
            })

        xobjects = _objects(xs)
        oobjects = {
            "tables": len(osheet.tables),
            "data_validations": len(osheet.data_validations.dataValidation),
            "conditional_formatting": len(osheet.conditional_formatting),
            "charts": len(getattr(osheet, "_charts", [])),
            "images": len(getattr(osheet, "_images", [])),
        }
        for key, oval in oobjects.items():
            xval = xobjects.get(key)
            if xval == "<missing>":
                missing.append({"sheet": name, "kind": key})
            elif xval != oval:
                differences.append({
                    "sheet": name, "kind": key,
                    "left": xval, "right": oval,
                })

    owb.close()
    try:
        xwb.close()
    except Exception:
        pass

    report = {
        "left": str(left_path),
        "right": str(right_path),
        "differences": differences,
        "missing": missing,
    }
    report["ok"] = not differences and (allow_missing or not missing)
    return report


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("left")
    parser.add_argument("right")
    parser.add_argument("--json", type=Path)
    parser.add_argument("--allow-missing", action="store_true")
    args = parser.parse_args(argv)

    report = compare(args.left, args.right, allow_missing=args.allow_missing)
    text = json.dumps(report, indent=2, ensure_ascii=False)
    if args.json:
        args.json.write_text(text + "\n", encoding="utf-8")
    print(text)
    return 0 if report["ok"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
