"""Incremental XLSX rows and shared-formula expansion."""

import datetime
import re
from xml.etree import ElementTree as ET

from ._metadata import MAIN, namespace_of
from ._text import xstring_unescape
from .excel_dates import serial_to_datetime
from .refs import parse_cell_range, column_label

_TOKEN = re.compile(
    r'"(?:[^"]|"")*"|\x27(?:[^\x27]|\x27\x27)*\x27|\[[^\]]*\]'
    r'|(?<![\w.])(?P<cell>\$?[A-Za-z]{1,3}\$?[1-9]\d*)(?![\w.!(])'
    r'|(?<![\w.])(?P<cols>\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3})(?![\w.])'
    r'|(?<![\w.])(?P<rows>\$?[1-9]\d*:\$?[1-9]\d*)(?![\w.])'
)


def translate_formula(text, origin, target):
    """Translate relative A1 references in a shared formula (without evaluating)."""
    c1, r1, _, _ = parse_cell_range(origin)
    c2, r2, _, _ = parse_cell_range(target)
    dc, dr = c2 - c1, r2 - r1

    def shift_col(label):
        if label.startswith("$"):
            return label
        col, _, _, _ = parse_cell_range(label + "1")
        return column_label(col + dc) if 0 <= col + dc < 16384 else "#REF!"

    def shift_row(label):
        if label.startswith("$"):
            return label
        value = int(label) + dr
        return str(value) if 1 <= value <= 1048576 else "#REF!"

    def shift(match):
        if match["cell"]:
            col, row = re.fullmatch(r"(\$?[A-Za-z]+)(\$?\d+)", match[0]).groups()
            try:
                col = shift_col(col)
            except ValueError:
                return match[0]  # A name such as XFE1, rather than a cell.
            row = shift_row(row)
            return "#REF!" if "#REF!" in (col, row) else col + row
        if match["cols"]:
            return ":".join(shift_col(part) for part in match[0].split(":"))
        if match["rows"]:
            return ":".join(shift_row(part) for part in match[0].split(":"))
        return match[0]

    return _TOKEN.sub(shift, text)


def iter_indexed_rows(reader, sheet_name, data_only=True):
    """Yield ``(row_number, [(column, value), ...])`` with worksheet coordinates
    (1-based rows, 0-based columns). Rows absent from the XML are not yielded."""
    rid = reader._worksheet_name_to_id[sheet_name]
    path = reader._resolve_path(reader._worksheet_id_to_location[rid])
    shared = {}
    next_row = 1
    ns = MAIN
    with reader._archive.open(path) as stream:
        stack = []
        for event, node in ET.iterparse(stream, events=("start", "end")):
            if event == "start":
                if not stack:
                    ns = namespace_of(node.tag) or MAIN  # Transitional or Strict
                stack.append(node)
                continue
            if node.tag == f"{{{ns}}}row" and len(stack) >= 2 and stack[-2].tag == f"{{{ns}}}sheetData":
                row_number = int(node.get("r", next_row))
                next_row = row_number + 1
                cells = []
                next_col = 0
                for cell in node:
                    if cell.tag != f"{{{ns}}}c":
                        continue
                    ref = cell.get("r")
                    col = parse_cell_range(ref)[0] if ref else next_col
                    ref = ref or f"{column_label(col)}{row_number}"
                    next_col = col + 1
                    formula = cell.find(f"{{{ns}}}f")
                    if not data_only and formula is not None:
                        if formula.get("t") == "shared":
                            index = formula.attrib["si"]
                            if formula.text:
                                shared[index] = (ref, formula.text)
                            if index not in shared:
                                raise ValueError("shared formula precedes its master")
                            origin, text = shared[index]
                            value = "=" + translate_formula(text, origin, ref)
                        else:
                            value = "=" + (formula.text or "")
                    else:
                        value = _value(reader, cell, ns)
                    cells.append((col, value))
                stack[-2].remove(node)
                node.clear()
                stack.pop()
                yield row_number, cells
                continue
            stack.pop()


def iter_rows(reader, sheet_name, data_only=True):
    first_width = None
    for _row_number, cells in iter_indexed_rows(reader, sheet_name, data_only):
        row = [None] * (max((col for col, _ in cells), default=-1) + 1)
        for col, value in cells:
            row[col] = value
        if first_width is None:
            first_width = len(row)
        row.extend([None] * max(0, first_width - len(row)))
        yield row


def _value(reader, cell, ns=MAIN):
    kind = cell.get("t", "n")
    if kind == "inlineStr":
        inline = cell.find(f"{{{ns}}}is")
        return xstring_unescape("".join(
            n.text or "" for child in inline for n in
            ([child] if child.tag == f"{{{ns}}}t" else child.findall(f"{{{ns}}}t") if child.tag == f"{{{ns}}}r" else [])
        )) if inline is not None else ""
    value = cell.findtext(f"{{{ns}}}v", "")
    if kind == "s":
        index = int(value) if value else 0
        return reader._shared_string_array[index] if 0 <= index < len(reader._shared_string_array) else value
    if kind == "b":
        return value == "1"
    if kind == "e":
        return "error:" + value
    if kind == "str":
        return xstring_unescape(value)
    if kind == "d":
        return datetime.datetime.fromisoformat(value) if value else None
    if not value:
        return None
    style = int(cell.get("s", "0"))
    if style < len(reader._styles_cell_xfs_array):
        fmt = reader._styles_cell_xfs_array[style]["numFmtId"]
        if reader._number_formats_type_dict.get(fmt) is datetime.datetime:
            try:
                return serial_to_datetime(float(value), reader._date_1904)
            except (ValueError, OverflowError):
                return value
    return float(value) if any(ch in value for ch in ".eE") else int(value)
