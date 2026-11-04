"""Random sequences of edits, checked against what they should have done.

Comparing the backends with each other cannot find a rule both of them share, so
each sequence here also keeps its own record of what the sheet should hold, and a
saved file has to agree with it through every way of reading it. For each seed:

- the updaters and the workbook model, on XLSX and on XLSB, end with what the record
  says (cells, notes, links, merges and names);
- ``ExcelReader`` sees those cells by every path (bulk, ``streaming=True``,
  ``data_only=False``), and both formats return rows of the same shape;
- a refused edit changes nothing, and both formats refuse the same edits;
- the package is one Excel would open without a repair (see ``_lint``);
- the native and the Python backends write the same bytes.

``XLSPY_FUZZ_ITERATIONS`` and ``XLSPY_FUZZ_SEED`` change the budget and the seed (as
in ``test_native_fuzz.py``); a failure message names the seed of the case. A sample
of the files is also opened in Excel, see ``test_excel_opens_edited_files``.
"""

from __future__ import annotations

import gc
import hashlib
import json
import os
import pathlib
import posixpath
import random
import re
import subprocess
import sys
import tempfile
import zipfile
from xml.etree import ElementTree as ET

import pytest

from xlspy import (
    Comment,
    ExcelReader,
    Font,
    Formula,
    Hyperlink,
    XlsbUpdater,
    XlsbWriter,
    XlsxUpdater,
    XlsxWriter,
    load_workbook,
)
from xlspy._metadata import DefinedName
from xlspy.biff_utils import iter_records
from xlspy.refs import column_label, parse_cell_range

# The default run stays short; both budgets scale with XLSPY_FUZZ_ITERATIONS.
ITERATIONS = int(os.environ.get("XLSPY_FUZZ_ITERATIONS", "300")) // 5
SEED = int(os.environ.get("XLSPY_FUZZ_SEED", "7"))

FORMATS = {"xlsx": (XlsxWriter, XlsxUpdater), "xlsb": (XlsbWriter, XlsbUpdater)}
MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

VALUES = [
    0, 1, -7, 2**31, 2**53, 1.5, -0.25, 0.1 + 0.2, 1 / 3, 1234567890123456, 1e300, 5e-324, True, False,
    "text", " lead", "trail ", "a\nb", "tab\t", "<&>\"'", "żółć", "😀", "_x0041_", "x" * 300, "same", "same",
]  # fmt: skip
FORMULAS = ["1+1", "A1*2", "SUM(A1:B2)", '"a"&"b"', "IF(A1>0,1,2)"]
MERGES = ["J1:K2", "B3:C3", "E5:F7", "A10:A12", "H1:H2"]
LINKS = ["https://a.example/", "https://b.example/x?a=1&b=2"]
AUTHORS = ["Jan", "New <&>"]
SHEETS = ("Data", "Other")


def _ref(rng: random.Random, far: bool = False) -> str:
    if far and rng.random() < 0.1:
        return rng.choice(["AA100", "B200", "XFD12"])
    return f"{rng.choice('ABCDEFGH')}{rng.randint(1, 12)}"


def _number(value):
    return float(value) if isinstance(value, (int, float)) and not isinstance(value, bool) else value


class Sheet:
    """What a sheet should hold."""

    def __init__(self):
        self.cells: dict[str, object] = {}
        self.notes: dict[str, tuple[str, str]] = {}
        self.links: dict[str, str] = {}
        self.merges: set[str] = set()

    def put_rows(self, rows):
        self.cells = {
            f"{column_label(c)}{r + 1}": _number(value)
            for r, row in enumerate(rows)
            for c, value in enumerate(row)
            if value is not None
        }

    def put(self, ref, value):
        if value is None:
            self.cells.pop(ref, None)
        elif isinstance(value, Formula):
            self.cells[ref] = "=" + value.text
        else:
            self.cells[ref] = _number(value)

    def state(self):
        return {"cells": self.cells, "notes": self.notes, "links": self.links, "merges": sorted(self.merges)}


def _write_source(path, ext, seed):
    """A workbook of two sheets, and the record of what it holds."""
    rng = random.Random(f"{seed}:source")
    rows = [
        [rng.choice([*VALUES, None]) for _ in range(rng.randint(0, 6))]
        for _ in range(rng.randint(0, 10))
    ]
    options: dict = {"auto_filter": False}
    data = Sheet()
    if rng.random() < 0.4:
        data.merges = set(rng.sample(MERGES, rng.randint(1, 3)))
        options["merges"] = sorted(data.merges)
    data.put_rows(rows)
    for merge in data.merges:
        first = merge.split(":")[0]
        for ref in _inside(merge):
            if ref != first:
                data.cells.pop(ref, None)
    if rng.random() < 0.4:
        notes = {}
        for i in range(rng.randint(1, 4)):
            ref = _ref(rng)
            if not any(ref in _inside(merge) and ref != merge.split(":")[0] for merge in data.merges):
                notes[ref] = (f"note {i}", rng.choice(AUTHORS))
        if notes:
            options["comments"] = {ref: Comment(*note) for ref, note in notes.items()}
            data.notes = notes
    with FORMATS[ext][0](str(path), use_shared_strings=rng.random() < 0.7) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(_without_merged(rows, data.merges), **options)
        writer.add_sheet("Other")
        writer.write_sheet([[1, "two"]], auto_filter=False)
    other = Sheet()
    other.put_rows([[1, "two"]])
    return {"Data": data, "Other": other}


def _inside(merge):
    c1, r1, c2, r2 = parse_cell_range(merge)
    return {f"{column_label(c)}{r + 1}" for c in range(c1, c2 + 1) for r in range(r1, r2 + 1)}


def _without_merged(rows, merges):
    """A writer takes what it is given: only the first cell of a merged range holds a value."""
    hidden = {ref for merge in merges for ref in _inside(merge) if ref != merge.split(":")[0]}
    return [
        [None if f"{column_label(c)}{r + 1}" in hidden else value for c, value in enumerate(row)]
        for r, row in enumerate(rows)
    ]


# ---- the updaters ----


def _updater_edit(rng: random.Random):
    """(method, arguments) of one edit, valid or not."""
    sheet = "Data" if rng.random() < 0.9 else "Other"
    kind = rng.choices(
        ["values", "formula", "note", "note_off", "link", "link_off", "name", "name_off", "bad", "reopen", "replace"],
        [8, 2, 3, 2, 3, 2, 2, 1, 3, 1, 1],
    )[0]
    if kind == "values":
        count = rng.randint(1, 5)
        return "update_cells", (sheet, {_ref(rng, True): rng.choice([*VALUES, None]) for _ in range(count)})
    if kind == "formula":
        return "update_cells", (sheet, {_ref(rng): Formula(rng.choice(FORMULAS))})
    if kind == "note":
        notes = {_ref(rng, True): Comment(f"n{rng.randint(0, 99)}", rng.choice(AUTHORS)) for _ in range(rng.randint(1, 3))}
        return "set_comments", (sheet, notes)
    if kind == "note_off":
        return "set_comments", (sheet, {_ref(rng): None for _ in range(rng.randint(1, 3))})
    if kind == "link":
        links = {_ref(rng): Hyperlink(rng.choice(LINKS)) for _ in range(rng.randint(1, 3))}
        return "set_hyperlinks", (sheet, links)
    if kind == "link_off":
        return "set_hyperlinks", (sheet, {_ref(rng): None for _ in range(rng.randint(1, 2))})
    if kind == "name":
        return "set_defined_name", (
            rng.choice([
                DefinedName("Rate", "Data!$A$1"),
                DefinedName("Rate", "Data!$B$2"),
                DefinedName("Scoped", "Other!$A$1", "Other"),
                DefinedName("Hid", "Data!$C$3", hidden=True),
            ]),
        )
    if kind == "name_off":
        return "remove_defined_name", (rng.choice(["Rate", "Hid"]),)
    if kind == "replace":
        rows = [[rng.choice([*VALUES, None]) for _ in range(rng.randint(0, 4))] for _ in range(rng.randint(0, 5))]
        return "replace_sheet_data", (sheet, rows)
    if kind == "reopen":
        return "reopen", ()
    # Each of these has to be refused whole, the part that is fine with it.
    return rng.choice([
        ("update_cells", ("Missing", {"A1": 1})),
        ("update_cells", (sheet, {"A1": 999, "not a ref": 1})),
        ("update_cells", (sheet, {"A1": 999, "A1:B2": 1})),
        ("update_cells", (sheet, {"A1": 999, "B1": "x" * 40000})),
        ("set_comments", (sheet, {"A1": Comment("kept?", "Jan"), "B1": "text"})),
        ("set_comments", ("Missing", {"A1": None})),
        ("set_hyperlinks", (sheet, {"A1": Hyperlink("https://kept.example/"), "B1": "https://x/"})),
        ("set_hyperlinks", ("Missing", {"A1": None})),
        ("replace_sheet_data", (sheet, [[999, "x" * 40000]])),
    ])


def _record_edit(sheets, names, method, args):
    """An edit that went through, on the record."""
    if method == "update_cells":
        for ref, value in args[1].items():
            sheets[args[0]].put(ref, value)
    elif method == "set_comments":
        notes = sheets[args[0]].notes
        for ref, note in args[1].items():
            if note is None:
                notes.pop(ref, None)
            else:
                notes[ref] = (note.text, note.author)
    elif method == "set_hyperlinks":
        links = sheets[args[0]].links
        for ref, link in args[1].items():
            if link is None:
                links.pop(ref, None)
            else:
                links[ref] = link.target
    elif method == "set_defined_name":
        name = args[0]
        names[(name.name, name.scope)] = (name.formula, bool(name.hidden))
    elif method == "remove_defined_name":
        for key in [key for key in names if key[0] == args[0]]:
            del names[key]
    elif method == "replace_sheet_data":
        sheets[args[0]].put_rows(args[1])


def _parts(path):
    with zipfile.ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist() if name != "docProps/core.xml"}


def _run_updater(tmp, ext, seed, steps=8):
    """Returns the saved file, the record, what was refused, and what went wrong."""
    rng = random.Random(f"{seed}:edits")
    path = tmp / f"source-{seed}.{ext}"
    sheets = _write_source(path, ext, seed)
    names: dict = {}
    trace: list = []
    problems: list = []
    updater_type = FORMATS[ext][1]
    updater = updater_type(str(path))
    # Half of the runs save before each edit, for the file to compare with if it is
    # refused; the others keep every edit pending until the end.
    snapshots = seed % 2 == 0
    try:
        for step in range(steps):
            method, args = _updater_edit(rng)
            if method == "reopen":
                current = tmp / f"reopened-{seed}-{step}.{ext}"
                updater.save(str(current))
                updater.close()
                updater = updater_type(str(current))
                continue
            before = None
            if snapshots:
                updater.save(str(tmp / f"before.{ext}"))
                before = _parts(tmp / f"before.{ext}")
            try:
                getattr(updater, method)(*args)
            except (ValueError, TypeError, KeyError) as error:
                trace.append((step, method, type(error).__name__))
                if before is not None:
                    updater.save(str(tmp / f"after.{ext}"))
                    changed = sorted(n for n in before.keys() | _parts(tmp / f"after.{ext}").keys()
                                     if before.get(n) != _parts(tmp / f"after.{ext}").get(n))
                    if changed:
                        problems.append(("refused edit changed the file", f"{method} {changed}"))
            except Exception as error:  # noqa: BLE001 - what the test is looking for
                problems.append(("crash", f"{method}: {type(error).__name__}: {error}"))
            else:
                trace.append((step, method, "ok"))
                _record_edit(sheets, names, method, args)
        final = tmp / f"final-{seed}.{ext}"
        updater.save(str(final))
    finally:
        updater.close()
    return final, sheets, names, trace, problems


# ---- the workbook model ----


def _model_edit(rng: random.Random, wb, sheets, ext):
    """One edit of a loaded workbook, on the record too."""
    title = rng.choice(list(sheets))
    ws = wb[title]
    sheet = sheets[title]
    kind = rng.choices(
        ["value", "style", "note", "note_off", "link", "link_off", "merge", "unmerge", "append"], [8, 3, 3, 1, 3, 1, 2, 1, 2]
    )[0]
    ref = _ref(rng)
    hidden = {r for merge in sheet.merges for r in _inside(merge) if r != merge.split(":")[0]}
    if kind == "value":
        value = rng.choice([*VALUES, None])
        if ref in hidden:
            return
        ws[ref].value = value
        sheet.put(ref, value)
    elif kind == "style":
        if ref not in hidden:
            ws[ref].font = Font(bold=True)
    elif kind == "note":
        if ref not in hidden:
            note = (f"m{rng.randint(0, 99)}", rng.choice(AUTHORS))
            ws[ref].comment = Comment(*note)
            sheet.notes[ref] = note
    elif kind == "note_off":
        if ref not in hidden:
            ws[ref].comment = None
            sheet.notes.pop(ref, None)
    elif kind == "link":
        if ref not in hidden:
            target = rng.choice(LINKS)
            ws[ref].hyperlink = Hyperlink(target)
            sheet.links[ref] = target
    elif kind == "link_off":
        if ref not in hidden:
            ws[ref].hyperlink = None
            sheet.links.pop(ref, None)
    elif kind == "merge":
        merge = rng.choice(MERGES)
        if any(_inside(merge) & _inside(other) for other in sheet.merges):
            return
        ws.merge_cells(merge)
        sheet.merges.add(merge)
    elif kind == "unmerge":
        if sheet.merges:
            merge = rng.choice(sorted(sheet.merges))
            ws.unmerge_cells(merge)
            sheet.merges.discard(merge)
    elif kind == "append":
        values = [rng.choice([*VALUES, None]) for _ in range(rng.randint(1, 4))]
        if any(f"{column_label(c)}{ws.max_row + 1}" in hidden for c in range(len(values))):
            return  # Refused: a merged range keeps its other cells to itself.
        ws.append(values)
        row = ws.max_row
        for c, value in enumerate(values):
            if value is not None:
                sheet.put(f"{column_label(c)}{row}", value)


def _run_model(tmp, ext, seed, steps=10):
    rng = random.Random(f"{seed}:model")
    path = tmp / f"model-source-{seed}.{ext}"
    sheets = _write_source(path, ext, seed)
    wb = load_workbook(str(path))
    try:
        for step in range(steps):
            _model_edit(rng, wb, sheets, ext)
            if rng.random() < 0.1:
                current = tmp / f"model-{seed}-{step}.{ext}"
                wb.save(str(current))
                wb.close()
                wb = load_workbook(str(current))
        final = tmp / f"model-final-{seed}.{ext}"
        wb.save(str(final))
    finally:
        wb.close()
    return final, sheets


# ---- what a saved file holds ----


def _read_state(path):
    wb = load_workbook(str(path))
    try:
        result = {}
        for ws in wb.worksheets:
            cells = [cell for row in ws._cells.values() for cell in row.values()]
            result[ws.title] = {
                "cells": {c.coordinate: _number(c.value) for c in cells if c.value is not None},
                "notes": {c.coordinate: (c.comment.text, c.comment.author) for c in cells if c.comment},
                "links": {c.coordinate: c.hyperlink.target for c in cells if c.hyperlink},
                "merges": sorted(str(r) for r in ws.merged_cells.ranges),
            }
        return result
    finally:
        wb.close()


def _read_names(path):
    wb = load_workbook(str(path))
    try:
        # XLSB writes the sheet of a name in quotes whether it needs them or not.
        return {(name.name, name.scope): (name.formula.replace("'", ""), bool(name.hidden)) for name in wb.defined_names}
    finally:
        wb.close()


def _diff(want, got):
    keys = [key for key in want.keys() | got.keys() if want.get(key) != got.get(key)]
    return {key: (want.get(key), got.get(key)) for key in sorted(keys, key=str)[:4]}


def _rows_to_cells(rows):
    return {
        f"{column_label(c)}{r + 1}": _number(value)
        for r, row in enumerate(rows)
        for c, value in enumerate(row)
        if value is not None
    }


def _check_file(path, sheets, ext):
    """Everything a saved file should say about the record."""
    problems = []
    for problem in _lint(_parts(path), ext):
        problems.append((f"lint {ext}: {re.sub(r'[0-9]+', 'N', problem)}", problem))
    state = _read_state(path)
    for title, sheet in sheets.items():
        want = sheet.state()
        for part in ("cells", "notes", "links"):
            if state[title][part] != want[part]:
                problems.append((f"{part} differ from the record {ext}", f"{title}: {_diff(want[part], state[title][part])}"))
        if state[title]["merges"] != want["merges"]:
            problems.append((f"merges differ from the record {ext}", (title, want["merges"], state[title]["merges"])))
        formulas = {ref for ref, value in want["cells"].items() if isinstance(value, str) and value.startswith("=")}
        plain = {ref: value for ref, value in want["cells"].items() if ref not in formulas}
        with ExcelReader(str(path)) as reader:
            for label, options in (
                ("bulk", {}),
                ("streaming", {"streaming": True}),
                ("formulas", {"data_only": False}),
            ):
                got = _rows_to_cells(list(reader.get_rows(title, **options)))
                if label == "formulas":
                    expected = want["cells"]
                else:
                    expected = plain
                    got = {ref: value for ref, value in got.items() if ref not in formulas}
                if got != expected:
                    problems.append((f"reader {label} differs from the record {ext}", f"{title}: {_diff(expected, got)}"))
    return problems


def _shape(path, title):
    with ExcelReader(str(path)) as reader:
        rows = [[_number(value) for value in row] for row in reader.get_rows(title)]
    for row in rows:
        while row and row[-1] is None:
            row.pop()
    while rows and not rows[-1]:
        rows.pop()
    return rows


def _check_pair(files, sheets):
    """Both formats, for the same edits, return rows of the same shape."""
    problems = []
    for title, sheet in sheets.items():
        if any(isinstance(value, str) and value.startswith("=") for value in sheet.cells.values()):
            continue  # A formula has a cached value in one format and none in the other.
        # None at the end of a row is written as a blank cell in XLSX and not at all in XLSB.
        xlsx, xlsb = _shape(files["xlsx"], title), _shape(files["xlsb"], title)
        if xlsx != xlsb:
            problems.append(("row shape differs between the formats", f"{title}: {xlsx[:4]} {xlsb[:4]}"))
    return problems


def _lint(data, ext):
    """What makes Excel offer a repair: dangling relationships, parts without a type,
    two of what there can be one of, and rows or cells out of order."""
    bad = []
    names = set(data) | {"docProps/core.xml"}
    for name, body in data.items():
        if name.endswith((".xml", ".rels", ".vml")):
            try:
                ET.fromstring(body)
            except ET.ParseError as error:
                bad.append(f"not well-formed {name}: {error}")
    types = ET.fromstring(data["[Content_Types].xml"])
    defaults = {n.get("Extension").lower() for n in types if n.tag.endswith("Default")}
    overrides = {n.get("PartName").lstrip("/") for n in types if n.tag.endswith("Override")}
    bad += [f"type without a part: {part}" for part in overrides - names]
    bad += [
        f"part without a type: {name}"
        for name in names - {"[Content_Types].xml"}
        if name not in overrides and name.rsplit(".", 1)[-1].lower() not in defaults
    ]
    for name, body in data.items():
        if not name.endswith(".rels"):
            continue
        base = posixpath.dirname(posixpath.dirname(name))
        ids = []
        for rel in ET.fromstring(body):
            ids.append(rel.get("Id"))
            if rel.get("TargetMode") == "External":
                continue
            target = rel.get("Target")
            full = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(base, target))
            if full not in names:
                bad.append(f"dangling {name}: {target}")
        if len(ids) != len(set(ids)):
            bad.append(f"two relationships with an id in {name}")
    (_lint_xlsx_sheets if ext == "xlsx" else _lint_xlsb_sheets)(data, bad)
    return bad


def _rels_of(data, part):
    rels = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
    return {r.get("Id"): r for r in ET.fromstring(data[rels])} if rels in data else {}


SHEET_ORDER = (
    "sheetPr dimension sheetViews sheetFormatPr cols sheetData sheetCalcPr sheetProtection protectedRanges scenarios "
    "autoFilter sortState dataConsolidate customSheetViews mergeCells phoneticPr conditionalFormatting dataValidations "
    "hyperlinks printOptions pageMargins pageSetup headerFooter rowBreaks colBreaks customProperties cellWatches "
    "ignoredErrors smartTags drawing legacyDrawing legacyDrawingHF picture oleObjects controls webPublishItems "
    "tableParts extLst"
).split()


def _lint_xlsx_sheets(data, bad):
    table = ET.fromstring(data["xl/sharedStrings.xml"]) if "xl/sharedStrings.xml" in data else None
    for name, body in data.items():
        if re.fullmatch(r"xl/comments\d*\.xml", name):
            root = ET.fromstring(body)
            authors = root.findall(f"{{{MAIN}}}authors/{{{MAIN}}}author")
            notes = root.findall(f"{{{MAIN}}}commentList/{{{MAIN}}}comment")
            refs = [note.get("ref") for note in notes]
            if len(refs) != len(set(refs)):
                bad.append(f"{name}: two notes on a cell")
            if any(int(note.get("authorId", 0)) >= len(authors) for note in notes):
                bad.append(f"{name}: author out of range")
        if not re.fullmatch(r"xl/worksheets/sheet\d+\.xml", name):
            continue
        root = ET.fromstring(body)
        last_row = 0
        for row in root.findall(f"{{{MAIN}}}sheetData/{{{MAIN}}}row"):
            number = int(row.get("r", last_row + 1))
            if number <= last_row:
                bad.append(f"{name}: rows out of order at {number}")
            last_row = number
            column = -1
            for cell in row.findall(f"{{{MAIN}}}c"):
                reference = cell.get("r")
                c, r = (column + 1, number - 1) if reference is None else parse_cell_range(reference)[:2]
                if r + 1 != number:
                    bad.append(f"{name}: cell {reference} in row {number}")
                if c <= column:
                    bad.append(f"{name}: cells out of order at {reference}")
                column = c
                if cell.get("t") == "s" and (table is None or int(cell.find(f"{{{MAIN}}}v").text) >= len(table)):
                    bad.append(f"{name}: shared string past the table")
        ranks = [SHEET_ORDER.index(child.tag.split("}")[-1]) for child in root if child.tag.split("}")[-1] in SHEET_ORDER]
        if ranks != sorted(ranks):
            bad.append(f"{name}: elements out of order")
        rels = _rels_of(data, name)
        links = root.findall(f"{{{MAIN}}}hyperlinks/{{{MAIN}}}hyperlink")
        refs = [link.get("ref") for link in links]
        if len(refs) != len(set(refs)):
            bad.append(f"{name}: two links on a cell")
        for link in links:
            rid = link.get(f"{{{RNS}}}id")
            if rid is not None and rid not in rels:
                bad.append(f"{name}: link without a relationship")
        legacy = root.find(f"{{{MAIN}}}legacyDrawing")
        if legacy is not None and legacy.get(f"{{{RNS}}}id") not in rels:
            bad.append(f"{name}: legacy drawing without a relationship")
        kinds = [rel.get("Type").rsplit("/", 1)[-1] for rel in rels.values()]
        if ("comments" in kinds) != (legacy is not None):
            bad.append(f"{name}: notes and legacy drawing do not go together")
        merges = root.find(f"{{{MAIN}}}mergeCells")
        if merges is not None and merges.get("count") and int(merges.get("count")) != len(merges):
            bad.append(f"{name}: merge count")


def _lint_xlsb_sheets(data, bad):
    for name, body in data.items():
        if not re.fullmatch(r"xl/worksheets/sheet\d+\.bin", name):
            continue
        last_row, column, in_data = -1, -1, False
        for record in iter_records(body):
            kind = record.record_id
            if kind == 0x91:
                in_data = True
            elif kind == 0x92:
                in_data = False
            elif in_data and kind == 0:
                number = int.from_bytes(body[record.data_start:record.data_start + 4], "little")
                if number <= last_row:
                    bad.append(f"{name}: rows out of order at {number}")
                last_row, column = number, -1
            elif in_data and 1 <= kind <= 0x12:
                c = column + 1 if kind >= 0xC else int.from_bytes(body[record.data_start:record.data_start + 4], "little")
                if c <= column:
                    bad.append(f"{name}: cells out of order at column {c} of row {last_row}")
                column = c
            elif kind in (0x00, *range(0x01, 0x13)):
                bad.append(f"{name}: row or cell record outside the sheet data")


def check_updater_seed(seed):
    """What is wrong with a seed's edits by the updaters: (kind, detail)."""
    with tempfile.TemporaryDirectory() as folder:
        tmp = pathlib.Path(folder)
        problems: list = []
        files, traces, records = {}, {}, {}
        for ext in FORMATS:
            final, sheets, names, trace, found = _run_updater(tmp, ext, seed)
            problems += found
            problems += _check_file(final, sheets, ext)
            got = _read_names(final)
            if got != names:
                problems.append((f"names differ from the record {ext}", _diff(names, got)))
            files[ext], traces[ext], records[ext] = final, trace, sheets
        if traces["xlsx"] != traces["xlsb"]:
            first = next((a, b) for a, b in zip(traces["xlsx"], traces["xlsb"]) if a != b)
            problems.append(("the formats refuse different edits", first))
        problems += _check_pair(files, records["xlsx"])
        return problems


def check_model_seed(seed):
    """What is wrong with a seed's edits of a loaded workbook."""
    with tempfile.TemporaryDirectory() as folder:
        tmp = pathlib.Path(folder)
        problems: list = []
        files = {}
        record = None
        for ext in FORMATS:
            final, sheets = _run_model(tmp, ext, seed)
            problems += _check_file(final, sheets, ext)
            files[ext], record = final, sheets
        problems += _check_pair(files, record)
        return problems


def _digests(seeds):
    """The bytes of each seed's files, for comparing the backends."""
    result = {}
    with tempfile.TemporaryDirectory() as folder:
        tmp = pathlib.Path(folder)
        for seed in seeds:
            for ext in FORMATS:
                final = _run_updater(tmp, ext, seed)[0]
                result[f"edit-{seed}-{ext}"] = {n: hashlib.sha256(b).hexdigest() for n, b in _parts(final).items()}
                final = _run_model(tmp, ext, seed)[0]
                result[f"model-{seed}-{ext}"] = {n: hashlib.sha256(b).hexdigest() for n, b in _parts(final).items()}
    return result


def _report(findings):
    lines = [f"{seed}: {kind}: {str(detail)[:400]}" for seed, kind, detail in findings[:8]]
    return f"{len(findings)} problems, first ones:\n" + "\n".join(lines)


def test_updater_edit_sequences():
    findings = [(seed, *problem) for seed in range(SEED, SEED + ITERATIONS) for problem in check_updater_seed(seed)]
    assert not findings, _report(findings)


def test_model_edit_sequences():
    findings = [(seed, *problem) for seed in range(SEED, SEED + ITERATIONS) for problem in check_model_seed(seed)]
    assert not findings, _report(findings)


def test_backends_write_the_same_bytes():
    from xlspy import _accel

    if not _accel.HAVE_NATIVE or os.environ.get("XLSPY_ACCEL") == "python":
        pytest.skip("needs the native extension")
    seeds = range(SEED, SEED + max(ITERATIONS // 3, 10))
    env = {**os.environ, "XLSPY_ACCEL": "python"}
    # The same text as the path it runs under: this file, as a script.
    done = subprocess.run(
        [sys.executable, __file__, "--digests", str(seeds.start), str(seeds.stop)],
        env=env, capture_output=True, text=True, check=True,
    )  # fmt: skip
    theirs = json.loads(done.stdout)
    ours = _digests(seeds)
    differing = {
        case: sorted(n for n in ours[case].keys() | theirs[case].keys() if ours[case].get(n) != theirs[case].get(n))
        for case in ours
        if ours[case] != theirs[case]
    }
    assert not differing, str(dict(list(differing.items())[:8]))


@pytest.mark.com
def test_excel_opens_edited_files(excel, tmp_path):
    """Excel has the last word on what needs a repair: with alerts off it would drop
    parts of the file, so the record has to be still there after it opens it."""
    for seed in range(SEED, SEED + 12):
        for ext in FORMATS:
            for final, sheets in (_run_updater(tmp_path, ext, seed)[:2], _run_model(tmp_path, ext, seed)):
                book = excel.Workbooks.Open(str(final), UpdateLinks=0)
                try:
                    for title, sheet in sheets.items():
                        ws = book.Worksheets(title)
                        for ref, value in list(sheet.cells.items())[:30]:
                            if isinstance(value, str) and value.startswith("="):
                                assert ws.Range(ref).Formula == value, (seed, ext, final.name, title, ref)
                            elif isinstance(value, str):
                                assert ws.Range(ref).Value == value, (seed, ext, final.name, title, ref)
                            elif not isinstance(value, bool) and abs(value) < 1e15:
                                assert ws.Range(ref).Value == pytest.approx(value, rel=1e-14, abs=1e-300), (seed, ext, final.name, ref)
                        assert ws.Comments.Count == len(sheet.notes), (seed, ext, final.name, title)
                        assert ws.Hyperlinks.Count == len(sheet.links), (seed, ext, final.name, title)
                finally:
                    book.Close(SaveChanges=False)
                    del book
                    gc.collect()


if __name__ == "__main__":  # The Python backend's side of test_backends_write_the_same_bytes.
    if sys.argv[1] == "--digests":
        print(json.dumps(_digests(range(int(sys.argv[2]), int(sys.argv[3])))))
