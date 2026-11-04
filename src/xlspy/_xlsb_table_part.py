"""Read an XLSB table part (``xl/tables/tableN.bin``) into plain values.

The records are mapped in [MS-XLSB] 2.4.100 (BrtBeginList), 2.4.103
(BrtBeginListCol), 2.4.320 (BrtTableStyleClient) and 2.4.13 (BrtBeginAFilter).
The object model turns the result into a ``Table``; the updater only needs the
header. Both use this module, so the updater does not load the object model.
"""

from __future__ import annotations

import struct
from collections.abc import Callable, Iterable
from dataclasses import dataclass, field

from ._xlsb_metadata import read_wide, records, wide_string
from .biff_utils import build_record, iter_record_payloads
from .refs import column_label, parse_cell_range

_BEGIN_LIST, _BEGIN_LIST_COL, _TABLE_STYLE_CLIENT, _BEGIN_AUTO_FILTER = 0x157, 0x15B, 0x201, 0xA1
_END_LIST = 0x158
_TOTALS_FUNCTIONS = {
    1: "average", 2: "count", 3: "countNums", 4: "max", 5: "min",
    6: "sum", 7: "stdDev", 8: "var", 9: "custom",
}


def _range(r1: int, r2: int, c1: int, c2: int) -> str:
    return f"{column_label(c1)}{r1 + 1}:{column_label(c2)}{r2 + 1}"


@dataclass
class TablePartColumn:
    id: int
    name: str | None
    totals_function: str | None
    totals_label: str | None


@dataclass
class TablePartStyle:
    name: str | None
    first_column: bool
    last_column: bool
    row_stripes: bool
    column_stripes: bool


@dataclass
class TablePart:
    id: int
    name: str | None
    ref: str
    comment: str | None
    header_rows: int
    totals_rows: int | None
    auto_filter_ref: str | None = None
    columns: list[TablePartColumn] = field(default_factory=list)
    style: TablePartStyle | None = None


def parse_table_part(data: bytes) -> TablePart | None:
    """The table a part describes, or None when it has no BrtBeginList."""
    table = None
    for record_id, payload in iter_record_payloads(data):
        if record_id == _BEGIN_LIST and table is None:
            r1, r2, c1, c2, _kind, ident, header, totals, _flags = struct.unpack_from("<9I", payload, 0)
            # Six dxf ids and the connection id precede the names.
            name, offset = read_wide(payload, 36 + 24 + 4, nullable=True)
            _display, offset = read_wide(payload, offset, nullable=True)
            comment, offset = read_wide(payload, offset, nullable=True)
            table = TablePart(
                id=ident, name=name, ref=_range(r1, r2, c1, c2), comment=comment or None,
                header_rows=header, totals_rows=totals or None,
            )
        elif table is None:
            continue
        elif record_id == _BEGIN_AUTO_FILTER and table.auto_filter_ref is None:
            table.auto_filter_ref = _range(*struct.unpack_from("<4I", payload, 0))
        elif record_id == _BEGIN_LIST_COL:
            ident, totals_function = struct.unpack_from("<II", payload, 0)
            name, offset = read_wide(payload, 24, nullable=True)
            caption, offset = read_wide(payload, offset, nullable=True)
            label, offset = read_wide(payload, offset, nullable=True)
            table.columns.append(TablePartColumn(
                id=ident, name=caption or name,
                totals_function=_TOTALS_FUNCTIONS.get(totals_function),
                totals_label=label or None,
            ))
        elif record_id == _TABLE_STYLE_CLIENT:
            flags = struct.unpack_from("<H", payload, 0)[0]
            style, _ = read_wide(payload, 2, nullable=True)
            table.style = TablePartStyle(
                name=style,
                first_column=bool(flags & 0x01), last_column=bool(flags & 0x02),
                row_stripes=bool(flags & 0x04), column_stripes=bool(flags & 0x08),
            )
    return table


@dataclass
class TableEdit:
    """What a table part is to say after an edit; None keeps what the part holds."""

    ref: str
    name: str
    header_rows: int
    totals_rows: int
    column_names: list[str]
    auto_filter_ref: str | None = None
    style: TablePartStyle | None = None


def _nullable(text: str | None) -> bytes:
    return struct.pack("<I", 0xFFFFFFFF) if text is None else wide_string(text)


def _rfx(ref: str) -> bytes:
    first_col, first_row, last_col, last_row = parse_cell_range(ref)
    return struct.pack("<IIII", first_row, last_row, first_col, last_col)


def _style_payload(style: TablePartStyle, rest: bytes, old_flags: int = 0) -> bytes:
    flags = (old_flags & ~0x0F) | (
        (0x01 if style.first_column else 0) | (0x02 if style.last_column else 0)
        | (0x04 if style.row_stripes else 0) | (0x08 if style.column_stripes else 0)
    )
    return struct.pack("<H", flags) + _nullable(style.name) + rest


def patch_table_part(data: bytes, edit: TableEdit) -> bytes:
    """``data`` with the records an edit changes re-encoded and every other record kept.

    The range, names, header and totals counts of BrtBeginList, the range of the
    table's filter, the names of its columns and its style are what the object
    model can change; the rest of the part (column formulas, totals, formats,
    filter criteria, extensions) keeps its bytes. The edit must keep the table's
    number of columns.
    """
    out: list[bytes] = []
    column = 0
    seen_list = seen_filter = seen_style = False
    for record, payload in records(data):
        kind = record.record_id
        original = data[record.header_start:record.data_end]
        if kind == _BEGIN_LIST and not seen_list:
            seen_list = True
            head = struct.unpack_from("<9I", payload, 0)
            _name, after_name = read_wide(payload, 64, nullable=True)
            _display, after_display = read_wide(payload, after_name, nullable=True)
            _comment, after_comment = read_wide(payload, after_display, nullable=True)
            strings = payload[64:after_display]
            if _name != edit.name or _display != edit.name:
                strings = _nullable(edit.name) + _nullable(edit.name)
            new = (
                _rfx(edit.ref) + struct.pack("<5I", head[4], head[5], edit.header_rows, edit.totals_rows, head[8])
                + payload[36:64] + strings + payload[after_display:]
            )
            out.append(build_record(kind, new))
        elif kind == _BEGIN_AUTO_FILTER and not seen_filter:
            seen_filter = True
            if edit.auto_filter_ref is None:
                out.append(original)
            else:
                out.append(build_record(kind, _rfx(edit.auto_filter_ref) + payload[16:]))
        elif kind == _BEGIN_LIST_COL:
            name, after_name = read_wide(payload, 24, nullable=True)
            caption, after_caption = read_wide(payload, after_name, nullable=True)
            wanted = edit.column_names[column]
            column += 1
            if (caption or name) == wanted:
                out.append(original)
            elif caption:
                out.append(build_record(kind, payload[:after_name] + _nullable(wanted) + payload[after_caption:]))
            else:
                out.append(build_record(kind, payload[:24] + _nullable(wanted) + payload[after_name:]))
        elif kind == _TABLE_STYLE_CLIENT and not seen_style:
            seen_style = True
            if edit.style is None:
                out.append(original)
            else:
                _style, after_style = read_wide(payload, 2, nullable=True)
                old_flags = struct.unpack_from("<H", payload, 0)[0]
                out.append(build_record(kind, _style_payload(edit.style, payload[after_style:], old_flags)))
        elif kind == _END_LIST and not seen_style and edit.style is not None:
            seen_style = True
            out.append(build_record(_TABLE_STYLE_CLIENT, _style_payload(edit.style, b"")))
            out.append(original)
        else:
            out.append(original)
    if not seen_list or column != len(edit.column_names):
        raise ValueError("the table part does not match the edit (number of columns)")
    return b"".join(out)


def table_catalog(names: Iterable[str], read: Callable[[str], bytes | None]) -> dict[int, tuple[str, list[str]]]:
    """The tables of a package by id: ``{id: (display name, column names)}``.

    ``names`` are the package's part names and ``read(name)`` gives a part's bytes.
    A structured reference in a formula names its table by this id and its columns
    by their positions. A part that cannot be read contributes nothing.
    """
    catalog: dict[int, tuple[str, list[str]]] = {}
    for name in names:
        if not (name.startswith("xl/tables/") and name.endswith(".bin")):
            continue
        try:
            data = read(name)
            part = parse_table_part(data) if data else None
        except (struct.error, IndexError, TypeError, ValueError):
            part = None
        if part is not None and part.name:
            catalog[part.id] = (part.name, [column.name or "" for column in part.columns])
    return catalog
