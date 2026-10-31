"""Streaming loader for the ``<sheetData>`` of an XLSX worksheet.

Rows are matched with regular expressions, like ``_xlsx_read._iter_fast``, but
keep what the editable model needs: style ids, cell types and formula text.
A row the patterns cannot reproduce exactly goes through ElementTree and the
shared ``_load_row``, so both paths build identical cells. The rest of the
worksheet part (everything outside ``<sheetData>``) is returned for the normal
element-by-element reader.
"""

from __future__ import annotations

import datetime as _datetime
from xml.etree import ElementTree as ET

from .. import _accel, _xlsx_read
from .._metadata import MAIN
from .._text import xstring_unescape
from ._model import CELL_SLOTS, Cell
from ._reader import _apply_row_attributes, _date_kind, _load_row, _number_value

_ROW_ATTRIBUTES = (" ht=", " hidden=", " outlineLevel=", " collapsed=")
_TYPED = frozenset({"s", "inlineStr", "b", "e", "str", "d"})


class _Fallback(Exception):
    pass


def read_sheet_data(sheet, package, sheet_path, style_table, *, data_only, date_1904,
                    shared_strings, rich_text):
    """Load the rows into ``sheet``. Return ``(max_row, max_column, remainder)``
    where ``remainder`` is the worksheet XML with an empty ``<sheetData/>``, or
    ``None`` (nothing loaded) when the part needs the element-by-element reader."""
    with package.open_entry(sheet_path) as stream:
        probe = _xlsx_read._probe_fast(stream, with_head=True)
        if probe is None or probe[3] != MAIN:
            return None
        text, decoder, root_tag, _ns, raw, head = probe
        if _accel._c_read_xlsx_worksheet is not None:
            return _load_rows_c(
                sheet, stream, raw, root_tag, head, style_table,
                data_only, date_1904, shared_strings, rich_text,
            )
        return _load_rows(
            sheet, stream, text, decoder, root_tag, head, style_table,
            data_only, date_1904, shared_strings, rich_text,
        )


_DATE_KINDS = {None: 0, "date": 1, "time": 2, "duration": 3}


def _load_rows_c(sheet, stream, raw, root_tag, head, style_table,
                 data_only, date_1904, shared_strings, rich_text):
    """The C scanner creates the cells of plain rows; the rows it leaves (formulas,
    row attributes, unusual markup) go through ElementTree as in ``_load_rows``."""
    kinds = [_DATE_KINDS[_date_kind(style.number_format)] for style in style_table.styles]
    rows = _accel._c_read_xlsx_worksheet(
        stream, raw, shared_strings, kinds, date_1904,
        (Cell, CELL_SLOTS, sheet, sheet._cells, data_only, rich_text),
    )
    next_row = 1
    max_row = max_column = 0
    shared: dict = {}
    for item in rows:
        root = ET.fromstring((root_tag + item.decode("utf-8") + "</worksheet>").encode("utf-8"))
        row_number, row_max, column_max = _load_row(
            sheet, root[0], next_row, style_table, data_only=data_only,
            date_1904=date_1904, shared=shared, shared_strings=shared_strings,
            rich_text=rich_text,
        )
        rows.resume(row_number)
        next_row = row_number + 1
        max_row = max(max_row, row_max)
        max_column = max(max_column, column_max)
    loaded_rows, loaded_columns = rows.bounds()
    tail = rows.tail()
    end = tail.find(b"</sheetData>")
    tail = tail[end + len(b"</sheetData>"):] if end >= 0 else b""
    tail += stream.read()
    prefix = head[:head.rfind(b"<sheetData")]
    return (
        max(max_row, loaded_rows), max(max_column, loaded_columns),
        prefix + b"<sheetData/>" + tail,
    )


def _load_rows(sheet, stream, text, decoder, root_tag, head, style_table,
               data_only, date_1904, shared_strings, rich_text):
    columns: dict[str, int] = {}
    date_kinds: dict[int, str | None] = {}
    shared: dict = {}
    strings_len = len(shared_strings)
    set_cell = sheet._set_cell_raw
    cell_pattern = _xlsx_read._CELL
    attr_pattern = _xlsx_read._ATTR
    ref_pattern = _xlsx_read._REF
    column_index = _xlsx_read._column_index
    inline_pattern = _xlsx_read._INLINE_T
    prefixed = _xlsx_read._PREFIXED
    fromisoformat = _datetime.datetime.fromisoformat
    state = {"next_row": 1, "max_row": 0, "max_column": 0}

    def date_kind_of(style_id):
        try:
            return date_kinds[style_id]
        except KeyError:
            style = style_table.style_at(style_id)
            kind = _date_kind(style.number_format if style is not None else None)
            date_kinds[style_id] = kind
            return kind

    def fast_row(attrs, body):
        row_match = _xlsx_read._ROW_NUM.search(attrs)
        if row_match:
            row_number = int(row_match.group(1))
        elif " r=" in attrs:
            raise _Fallback
        else:
            row_number = state["next_row"]
        row_attrs = None
        if attrs and any(name in attrs for name in _ROW_ATTRIBUTES):
            found = attr_pattern.findall(attrs)
            if len(found) != attrs.count("="):
                raise _Fallback
            row_attrs = dict(found)
        cells = []
        if body is not None:
            if not data_only and "<f" in body:
                raise _Fallback
            matched = 0
            next_col = 0
            for cell in cell_pattern.finditer(body):
                matched += 1
                attrs_text, v, inline = cell.group(1, 2, 3)
                attributes = dict(attr_pattern.findall(attrs_text)) if attrs_text else {}
                ref = attributes.get("r")
                if ref:
                    letters = ref.rstrip("0123456789")
                    col = columns.get(letters)
                    if col is None:
                        if ref_pattern.fullmatch(ref) is None:
                            raise _Fallback
                        col = columns[letters] = column_index(letters)
                    digits = ref[len(letters):]
                    if not (digits.isascii() and digits.isdigit()):
                        raise _Fallback
                    row = int(digits)
                    if not 1 <= row <= 1048576:
                        raise _Fallback
                else:
                    col, row = next_col, row_number
                next_col = col + 1
                style_text = attributes.get("s")
                style_id = int(style_text) if style_text else 0
                kind = attributes.get("t", "n")
                if kind not in _TYPED:
                    data_type = "n"
                    if not v:
                        value = None
                    else:
                        try:
                            number = float(v)
                        except ValueError:
                            value = v
                        else:
                            date_kind = date_kind_of(style_id)
                            if date_kind is None:
                                value = int(number) if number.is_integer() else number
                            else:
                                value = _number_value(number, date_kind, date_1904)
                elif kind == "s":
                    data_type = "s"
                    raw = v or ""
                    try:
                        index = int(raw)
                    except ValueError:
                        value = raw
                    else:
                        value = shared_strings[index] if 0 <= index < strings_len else raw
                elif kind == "inlineStr":
                    data_type = "s"
                    if rich_text:
                        raise _Fallback
                    if inline is None:
                        value = None
                    else:
                        text_match = inline_pattern.fullmatch(inline)
                        if text_match is None or "&" in inline or "\r" in inline:
                            raise _Fallback
                        value = xstring_unescape(text_match.group(1))
                elif kind == "b":
                    data_type = "b"
                    value = (v or "") == "1"
                elif kind == "e":
                    data_type = "e"
                    if v and "&" in v:
                        raise _Fallback
                    value = v or ""
                elif kind == "str":
                    data_type = "s"
                    if v and ("&" in v or "\r" in v):
                        raise _Fallback
                    value = xstring_unescape(v or "")
                else:  # "d"
                    data_type = "d"
                    if v and "&" in v:
                        raise _Fallback
                    raw = v or ""
                    try:
                        value = fromisoformat(raw) if raw else None
                    except ValueError:
                        value = raw
                cells.append((row, col + 1, value, data_type, style_id))
            if matched != body.count("<c") or prefixed.search(body):
                raise _Fallback
        # Nothing is created before the row is known to be exact.
        if row_attrs is not None:
            _apply_row_attributes(sheet, row_number, row_attrs)
        state["next_row"] = row_number + 1
        max_row = max(state["max_row"], row_number)
        max_column = state["max_column"]
        for row, column, value, data_type, style_id in cells:
            if value is None and style_id == 0:
                continue
            set_cell(row, column, value, style_id=style_id, data_type=data_type)
            if row > max_row:
                max_row = row
            if column > max_column:
                max_column = column
        state["max_row"] = max_row
        state["max_column"] = max_column

    def slow_row(row_text):
        root = ET.fromstring((root_tag + row_text + "</worksheet>").encode("utf-8"))
        row_number, row_max, column_max = _load_row(
            sheet, root[0], state["next_row"], style_table, data_only=data_only,
            date_1904=date_1904, shared=shared, shared_strings=shared_strings,
            rich_text=rich_text,
        )
        state["next_row"] = row_number + 1
        state["max_row"] = max(state["max_row"], row_max)
        state["max_column"] = max(state["max_column"], column_max)

    row_pattern = _xlsx_read._ROW
    chunk_size = _xlsx_read._CHUNK
    while True:
        chunk = stream.read(chunk_size)
        text += decoder.decode(chunk, not chunk)
        pos = 0
        for match in row_pattern.finditer(text):
            pos = match.end()
            try:
                fast_row(match.group(1), match.group(2))
            except _Fallback:
                slow_row(match.group(0))
        end = text.find("</sheetData>", pos)
        if end >= 0:
            tail = text[end + len("</sheetData>"):]
            break
        text = text[pos:]
        if not chunk:
            tail = ""
            break
    rest = stream.read()
    if rest:
        tail += decoder.decode(rest, True)
    prefix = head[:head.rfind(b"<sheetData")]
    remainder = prefix + b"<sheetData/>" + tail.encode("utf-8")
    return state["max_row"], state["max_column"], remainder
