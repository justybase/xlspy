"""Incremental XLSX rows and shared-formula expansion."""

import codecs
import datetime
import re
from xml.etree import ElementTree as ET

from ._metadata import MAIN, namespace_of
from ._text import xstring_unescape
from . import _accel
from .excel_dates import serial_to_datetime
from .refs import parse_cell_range, column_label

_TOKEN = re.compile(
    r'"(?:[^"]|"")*"|\x27(?:[^\x27]|\x27\x27)*\x27|\[[^\]]*\]'
    r'|(?<![\w.])(?P<cell>\$?[A-Za-z]{1,3}\$?[1-9]\d*)(?![\w.!(])'
    r'|(?<![\w.])(?P<cols>\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3})(?![\w.])'
    r'|(?<![\w.])(?P<rows>\$?[1-9]\d*:\$?[1-9]\d*)(?![\w.])'
)

# Fast path for data_only reads: regular expressions over the raw sheet XML.
# Anything they cannot handle exactly is handed to the ElementTree path, row by row.
_CHUNK = 1 << 18  # Bounds the text buffer: each chunk is decoded and concatenated into it.
_SHEET_DATA_OPEN_B = re.compile(rb"<sheetData\b[^>]*?(/?)>")
_ROOT = re.compile(r"<worksheet\b[^>]*>")
_ROW = re.compile(r"<row\b([^>]*?)(?:/>|>(.*?)</row>)", re.S)
_ROW_NUM = re.compile(r' r="(\d+)"')
_CELL = re.compile(
    r'<c((?: \w+="[^"]*")*)\s*'
    r'(?:/>|>(?:<f\b[^>]*/>|<f\b[^>]*>[^<]*</f>)?(?:<v>([^<]*)</v>|<is>(.*?)</is>)?</c>)',
    re.S,
)
_ATTR = re.compile(r' (\w+)="([^"]*)"')
_REF = re.compile(r"([A-Z]{1,3})\d+")
_INLINE_T = re.compile(r"<t(?: [^>]*)?>([^<]*)</t>")
_PREFIXED = re.compile(r"</?[A-Za-z_][\w.-]*:")
_XMLNS = re.compile(r'xmlns="([^"]+)"')
_ENCODING = re.compile(r'encoding="([^"]+)"')


class _Fallback(Exception):
    """Raised by the fast path for a row it does not handle exactly."""


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
    if data_only:
        with reader._archive.open(path) as stream:
            probe = _probe_fast(stream)
            if probe is not None:
                rows = _iter_c(reader, stream, probe)
                if rows is not None:
                    yield from rows
                    return
                yield from _iter_fast(reader, stream, probe)
                return
    yield from _iter_etree(reader, path, data_only)


def iter_rows(reader, sheet_name, data_only=True):
    """Yield one list per worksheet row. Rows absent from the XML between present
    rows are emitted as empty rows, so row positions match the sheet. The first
    row present in the XML fixes the width that other rows are padded to; rows
    before it, when they are absent, are empty lists."""
    first_width = None
    next_row = 1
    for row_number, cells in iter_indexed_rows(reader, sheet_name, data_only):
        for _ in range(row_number - next_row):
            yield [None] * (first_width or 0)
        next_row = row_number + 1
        row = [None] * max(first_width or 0, max((col for col, _ in cells), default=-1) + 1)
        for col, value in cells:
            row[col] = value
        if first_width is None:
            first_width = len(row)
        yield row


def _iter_etree(reader, path, data_only):
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
                row_number, cells = _row_from_element(reader, node, ns, next_row, data_only, shared)
                next_row = row_number + 1
                stack[-2].remove(node)
                node.clear()
                stack.pop()
                yield row_number, cells
                continue
            stack.pop()


def _row_from_element(reader, node, ns, next_row, data_only, shared):
    row_number = int(node.get("r", next_row))
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
    return row_number, cells


def _date_style_flags(reader):
    """Per cell-style flag: True when the style's number format is a date."""
    formats = reader._number_formats_type_dict
    return [formats.get(xf["numFmtId"]) is datetime.datetime
            for xf in reader._styles_cell_xfs_array]


class _FastSheet:
    """Per-sheet state of the fast path: workbook tables and a column-letter cache."""

    def __init__(self, reader, ns, root_tag):
        self.strings = reader._shared_string_array
        self.date_1904 = reader._date_1904
        self.is_date_style = _date_style_flags(reader)
        self.columns = {}
        self.reader = reader
        self.ns = ns
        self.root_tag = root_tag


def _probe_fast(stream):
    """Read up to the sheetData start tag. Return ``(text, decoder, root_tag, ns, raw)``
    when the fast path applies (UTF-8, unprefixed tags, non-empty sheetData), otherwise
    ``None``. ``text`` and ``decoder`` continue the Python reader; ``raw`` is the same
    tail of bytes for the C reader."""
    raw = b""
    match = None
    while True:
        chunk = stream.read(_CHUNK)
        if not raw and chunk[:2] in (codecs.BOM_UTF16_LE, codecs.BOM_UTF16_BE):
            return None
        raw += chunk
        match = _SHEET_DATA_OPEN_B.search(raw)
        if match or not chunk:
            break
    if match is None or match.group(1):
        return None
    head = raw[:match.end()]
    head_text = head.decode("utf-8")
    root = _ROOT.search(head_text)
    if root is None or _PREFIXED.search(head_text):
        return None
    declaration = _ENCODING.search(head_text, 0, root.start())
    if declaration and declaration.group(1).lower() not in ("utf-8", "utf8"):
        return None
    xmlns = _XMLNS.search(root.group(0))
    ns = xmlns.group(1) if xmlns else MAIN
    decoder = codecs.getincrementaldecoder("utf-8")()
    tail = raw[match.end():]
    return decoder.decode(tail, False), decoder, root.group(0), ns, tail


def _iter_c(reader, stream, probe):
    """Rows from the C scanner, or None when the extension does not provide it."""
    _text, _decoder, root_tag, ns, raw = probe
    rows = _accel.read_xlsx_worksheet(
        stream, raw, reader._shared_string_array, _date_style_flags(reader),
        reader._date_1904)
    if rows is None:
        return None
    return _c_rows(reader, rows, root_tag, ns)


def _c_rows(reader, rows, root_tag, ns):
    """Resolve the C scanner's fallback rows with the XML parser, as _iter_fast does."""
    next_row = 1
    for item in rows:
        if isinstance(item, bytes):
            root = ET.fromstring((root_tag + item.decode("utf-8") + "</worksheet>").encode("utf-8"))
            row_number, cells = _row_from_element(reader, root[0], ns, next_row, True, None)
            rows.resume(row_number)
        else:
            row_number, cells = item
        next_row = row_number + 1
        yield row_number, cells


def _iter_fast(reader, stream, probe):
    text, decoder, root_tag, ns, _raw = probe
    sheet = _FastSheet(reader, ns, root_tag)
    next_row = 1
    while True:
        chunk = stream.read(_CHUNK)
        text += decoder.decode(chunk, not chunk)
        pos = 0
        for match in _ROW.finditer(text):
            pos = match.end()
            try:
                row_number, cells = _fast_row(sheet, match.group(1), match.group(2), next_row)
            except _Fallback:
                root = ET.fromstring((root_tag + match.group(0) + "</worksheet>").encode("utf-8"))
                row_number, cells = _row_from_element(reader, root[0], ns, next_row, True, None)
            next_row = row_number + 1
            yield row_number, cells
        done = "</sheetData>" in text[pos:]
        text = text[pos:]
        if done or not chunk:
            return


def _fast_row(sheet, attrs, body, next_row):
    row_match = _ROW_NUM.search(attrs)
    if row_match:
        row_number = int(row_match.group(1))
    elif " r=" in attrs:
        raise _Fallback
    else:
        row_number = next_row
    if body is None:
        return row_number, []

    cells = []
    next_col = 0
    matched = 0
    columns = sheet.columns
    for cell in _CELL.finditer(body):
        matched += 1
        attrs_text, v, inline = cell.group(1, 2, 3)
        attributes = dict(_ATTR.findall(attrs_text)) if attrs_text else {}
        ref = attributes.get("r")
        if ref:
            # Column letters are the reference without its trailing digits; only unseen letters are validated.
            letters = ref.rstrip("0123456789")
            col = columns.get(letters)
            if col is None:
                if _REF.fullmatch(ref) is None:
                    raise _Fallback
                col = columns[letters] = _column_index(letters)
        else:
            col = next_col
        next_col = col + 1
        if v and "t" not in attributes:
            # Plain number (the common case), inlined; a date style goes through the shared converter.
            style_text = attributes.get("s")
            style = int(style_text) if style_text is not None else 0
            if style < len(sheet.is_date_style) and sheet.is_date_style[style]:
                value = _fast_value(sheet, attributes, v, inline)
            else:
                value = float(v) if ("." in v or "e" in v or "E" in v) else int(v)
        else:
            value = _fast_value(sheet, attributes, v, inline)
        cells.append((col, value))
    # Any <c that the pattern did not match, or a prefixed tag, means this row needs the XML parser.
    if matched != body.count("<c") or _PREFIXED.search(body):
        raise _Fallback
    return row_number, cells


def _fast_value(sheet, attributes, v, inline):
    kind = attributes.get("t", "n")
    if kind == "inlineStr":
        if inline is None:
            return ""
        text_match = _INLINE_T.fullmatch(inline)
        if text_match is None or "&" in inline or "\r" in inline:
            raise _Fallback
        return xstring_unescape(text_match.group(1))
    if kind == "s":
        index = int(v) if v else 0
        return sheet.strings[index] if 0 <= index < len(sheet.strings) else (v or "")
    if kind == "b":
        return (v or "") == "1"
    if kind == "e":
        return "error:" + (v or "")
    if kind == "str":
        if v and ("&" in v or "\r" in v):
            raise _Fallback
        return xstring_unescape(v or "")
    if kind == "d":
        return datetime.datetime.fromisoformat(v) if v else None
    if not v:
        return None
    style_text = attributes.get("s")
    style = int(style_text) if style_text is not None else 0
    if style < len(sheet.is_date_style) and sheet.is_date_style[style]:
        try:
            return serial_to_datetime(float(v), sheet.date_1904)
        except (ValueError, OverflowError):
            return v
    return float(v) if ("." in v or "e" in v or "E" in v) else int(v)


def _column_index(letters):
    index = 0
    for char in letters:
        index = index * 26 + (ord(char) - 64)
    return index - 1


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
