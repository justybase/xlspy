"""Read an XLSB table part (``xl/tables/tableN.bin``) into plain values.

The records are mapped in [MS-XLSB] 2.4.100 (BrtBeginList), 2.4.103
(BrtBeginListCol), 2.4.320 (BrtTableStyleClient) and 2.4.13 (BrtBeginAFilter).
The object model turns the result into a ``Table``; the updater only needs the
header. Both use this module, so the updater does not load the object model.
"""

from __future__ import annotations

import struct
from dataclasses import dataclass, field

from ._xlsb_metadata import read_wide
from .biff_utils import iter_record_payloads
from .refs import column_label

_BEGIN_LIST, _BEGIN_LIST_COL, _TABLE_STYLE_CLIENT, _BEGIN_AUTO_FILTER = 0x157, 0x15B, 0x201, 0xA1
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
