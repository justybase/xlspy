"""Pivot tables as BIFF12 parts: the XLSB form of ``pivots.py``.

The layout was mapped from pivot tables Excel saved as XLSB (one option at a
time, with the cache saved without records and set to refresh on open), so the
fixed bytes below are Excel's own. Two parts per pivot:

* ``xl/pivotCache/pivotCacheDefinitionN.bin``: the source (a table), then one
  BrtBeginPCDField per source column. A field's BrtBeginPCDFAtbl holds its
  type flags and, for a numeric field, the range; an axis field lists its
  shared items after it (BrtPCDIString 0x18, BrtPCDINumber 0x15 and
  BrtPCDIMissing 0x14).
* ``xl/pivotTables/pivotTableN.bin``: BrtBeginSXView, the location, one
  BrtBeginSXVD per source column (its axis, and the items of an axis field),
  the row and column field lists, the row and column items (BrtBeginSXLI),
  the page fields (BrtBeginSXPI) and the data fields (BrtBeginSXDI).

The workbook lists each cache by id in BrtBeginPivotCacheIDs, which follows
BrtCalcProp and comes before BrtFileRecover.
"""

from __future__ import annotations

import struct
import uuid

from ._xlsb_metadata import wide_string
from .biff_utils import build_record
from .pivots import _BLANK, PivotRender, _is_number, _sort_key, _Stats
from .refs import parse_cell_range

CT_PIVOT_TABLE_BIN = "application/vnd.ms-excel.pivotTable"
CT_PIVOT_CACHE_BIN = "application/vnd.ms-excel.pivotCacheDefinition"

_AC_BEGIN, _AC_END, _UID = 0x25, 0x26, 0x0C00
_FRT_BEGIN, _FRT_END = 0x23, 0x24

# Cache definition records.
_CACHE_DEF, _END_CACHE_DEF = 0xB3, 0xB4
_FIELDS, _END_FIELDS = 0xB5, 0xB6
_FIELD, _END_FIELD = 0xB7, 0xB8
_SOURCE, _END_SOURCE = 0xB9, 0xBA
_SOURCE_RANGE, _END_SOURCE_RANGE = 0xBB, 0xBC
_FIELD_ITEMS, _END_FIELD_ITEMS = 0xBD, 0xBE
_ITEM_MISSING, _ITEM_NUMBER, _ITEM_STRING = 0x14, 0x15, 0x18

# Pivot table records.
_VIEW, _END_VIEW = 0x118, 0x13B
_LOCATION, _END_LOCATION = 0x13A, 0x139
_VDS, _END_VDS = 0x11F, 0x120
_VD, _END_VD = 0x11D, 0x11E
_VIS, _END_VIS = 0x11B, 0x11C
_VI, _END_VI = 0x11A, 0x119
_ROW_FIELDS, _END_ROW_FIELDS = 0x135, 0x136
_COL_FIELDS, _END_COL_FIELDS = 0x137, 0x138
_ROW_ITEMS, _END_ROW_ITEMS = 0x12B, 0x12C
_COL_ITEMS, _END_COL_ITEMS = 0x12D, 0x12E
_LI, _END_LI = 0x129, 0x12A
_LI_INDEXES, _END_LI_INDEXES = 0x184, 0x185
_PIS, _END_PIS = 0x123, 0x124
_PI, _END_PI = 0x121, 0x122
_DIS, _END_DIS = 0x127, 0x128
_DI, _END_DI = 0x125, 0x126
_STYLE = 0x201

# Workbook records.
_CACHE_IDS, _END_CACHE_IDS = 0x180, 0x181
_CACHE_ID, _END_CACHE_ID = 0x182, 0x183

_AC_HEADER = bytes.fromhex("010000100080")
# BrtBeginSXView up to idCache: the flags Excel writes for a new pivot (grand
# totals on, compact and outline layout, headers shown) and the version bytes.
_VIEW_PREFIX = bytes.fromhex("08414000f2650b20df02000002000803ffffffff0100000000000000")
# Excel's "Autofit column widths on update" (PivotTable.HasAutoFormat, the XLSX
# useAutoFormatting attribute). Found by saving the same pivot from Excel with the
# setting on and off: only this bit of BrtBeginSXView differs (byte 5, mask 0x01).
_AUTOFORMAT_BYTE, _AUTOFORMAT_MASK = 5, 0x01


def _view_prefix(autofit_columns: bool) -> bytes:
    prefix = bytearray(_VIEW_PREFIX)
    if autofit_columns:
        prefix[_AUTOFORMAT_BYTE] |= _AUTOFORMAT_MASK
    else:
        prefix[_AUTOFORMAT_BYTE] &= ~_AUTOFORMAT_MASK & 0xFF
    return bytes(prefix)
# BrtBeginSXVD after its axis byte, and BrtBeginPCDField before the name.
_VD_TAIL = bytes.fromhex("01001000000000" "5fa104000a000000ffffffff")
_FIELD_PREFIX = bytes.fromhex("04000000") + bytes(16)
_REFRESHED_DATE = 45000.0  # a fixed date: the pivot refreshes when the file opens
_REFRESHED_BY = "xlspy"
_CACHE_FLAGS = bytes.fromhex("08030814ffffffff")
_PAGE_ITEM_TAIL = bytes.fromhex("fe001000ffffffff00")
_STYLE_FLAGS = 0x0033

# Extension blocks Excel writes after the data fields (the hide-values-row
# setting lives in them) and after the cache fields.
_VIEW_TAIL = (
    build_record(_FRT_BEGIN, bytes.fromhex("020e0000"))
    + build_record(0x426, bytes(9) + b"\xff" * 12) + build_record(0x427)
    + build_record(_FRT_END)
    + build_record(_FRT_BEGIN, bytes.fromhex("03110000"))
    + build_record(0x1388, bytes.fromhex("010100")) + build_record(0x1389)
    + build_record(_FRT_END)
)
_CACHE_TAIL = (
    build_record(_FRT_BEGIN, bytes.fromhex("020e0000"))
    + build_record(0x42A, bytes(9)) + build_record(0x42B)
    + build_record(_FRT_END)
    + build_record(_FRT_BEGIN, bytes.fromhex("54000200"))
    + build_record(0x1421, b"\x01")
    + build_record(_FRT_END)
)

_AXIS_ROW, _AXIS_COL, _AXIS_PAGE, _AXIS_DATA = 0x01, 0x02, 0x04, 0x08
_SUBTOTALS = {"sum": 0, "count": 1, "average": 2, "max": 3, "min": 4}
_DISPLAY_GRAND = 0x0D  # BrtBeginSXLI item type of a grand total

# BrtBeginPCDFAtbl type flags (read off Excel's files).
_BASE, _NON_DATE, _STRING, _BLANK_FLAG = 0x0400, 0x0002, 0x0008, 0x0010
_SEMI_MIXED, _MIXED, _NUMBER, _INTEGER = 0x0001, 0x0020, 0x0140, 0x0080


def _uid(kind: str, number: int, name: str) -> bytes:
    """A stable revision id, so the same workbook always writes the same bytes."""
    return uuid.uuid5(uuid.NAMESPACE_URL, f"xlspy:pivot:{kind}:{number}:{name}").bytes


def _header(kind: str, number: int, name: str) -> bytes:
    return (build_record(_AC_BEGIN, _AC_HEADER) + build_record(_UID, _uid(kind, number, name))
            + build_record(_AC_END))


def _check(stats: _Stats, field: str) -> None:
    if stats.unsupported:
        raise ValueError(
            f"pivot field {field!r} holds dates, booleans or NaN, which a pivot "
            "table in XLSB cannot read; use XLSX for this pivot"
        )


def _type_flags(stats: _Stats) -> int:
    """The BrtBeginPCDFAtbl flags for what a field holds."""
    flags = _BASE | _NON_DATE
    if stats.count:
        flags |= _NUMBER | (_INTEGER if stats.integer else 0)
    if stats.text or not stats.count:
        flags |= _STRING | _SEMI_MIXED  # a field without numbers is typed as text
    if stats.text and stats.count:
        flags |= _MIXED
    if stats.blank:
        flags |= _BLANK_FLAG | _SEMI_MIXED
    return flags


def _field_items(stats: _Stats, values: list | None) -> bytes:
    """BrtBeginPCDFAtbl with the field's shared items, which only an axis field has."""
    count = len(values) if values is not None else 0
    payload = struct.pack("<HI", _type_flags(stats), count)
    if stats.count:
        payload += struct.pack("<dd", stats.low, stats.high)
    items = []
    for value in values or ():
        if value is _BLANK:
            items.append(build_record(_ITEM_MISSING))
        elif _is_number(value):
            items.append(build_record(_ITEM_NUMBER, struct.pack("<d", float(value))))
        else:
            items.append(build_record(_ITEM_STRING, wide_string(str(value))))
    return build_record(_FIELD_ITEMS, payload) + b"".join(items) + build_record(_END_FIELD_ITEMS)


def pivot_cache_bin(render: PivotRender) -> bytes:
    """The cache definition: the source table and one field per column."""
    refreshed = _CACHE_FLAGS + struct.pack("<d", _REFRESHED_DATE) + b"\x01"
    parts = [
        _header("cache", render.cache_id, render.plan.name),
        build_record(
            _CACHE_DEF,
            refreshed + struct.pack("<I", render.records) + wide_string(_REFRESHED_BY),
        ),
        build_record(_SOURCE, bytes(8)),
        build_record(_SOURCE_RANGE, b"\x01\x00\x00" + wide_string(render.plan.table.display_name)),
        build_record(_END_SOURCE_RANGE),
        build_record(_END_SOURCE),
        build_record(_FIELDS, struct.pack("<I", len(render.column_names))),
    ]
    for index, name in enumerate(render.column_names):
        values = render.shared_items.get(index)
        if values is not None:
            stats = _Stats()
            for value in values:
                stats.add(value)
        else:
            stats = render.value_stats.get(index) or _Stats()
        _check(stats, name)
        parts.append(build_record(_FIELD, _FIELD_PREFIX + wide_string(name)))
        parts.append(_field_items(stats, values))
        parts.append(build_record(_END_FIELD))
    parts.append(build_record(_END_FIELDS))
    parts.append(_CACHE_TAIL)
    parts.append(build_record(_END_CACHE_DEF))
    return b"".join(parts)


def _items_records(indexes: list[int]) -> bytes:
    """A BrtBeginSXVIs block: one BrtBeginSXVI per item and the default subtotal."""
    parts = [build_record(_VIS, struct.pack("<I", len(indexes) + 1))]
    for index in indexes:
        parts.append(build_record(_VI, b"\x00\x00\x00" + struct.pack("<I", index)))
        parts.append(build_record(_END_VI))
    parts.append(build_record(_VI, b"\x01\x00\x00" + struct.pack("<i", -1)))
    parts.append(build_record(_END_VI))
    parts.append(build_record(_END_VIS))
    return b"".join(parts)


def _line(repeat: int, kind: int, indexes: list[int], data: int = 0) -> bytes:
    """One BrtBeginSXLI (a row or column item) with its item indexes."""
    parts = [build_record(
        _LI, struct.pack("<HHII", repeat, kind, len(indexes), data),
    )]
    if indexes:
        parts.append(build_record(_LI_INDEXES, struct.pack(f"<{len(indexes)}I", *indexes)))
        parts.append(build_record(_END_LI_INDEXES))
    parts.append(build_record(_END_LI))
    return b"".join(parts)


def _row_items(render: PivotRender) -> bytes:
    lines = [_line(0, 0, [position]) for position in range(len(render.row_items))]
    lines.append(_line(0, _DISPLAY_GRAND, [0]))
    return (build_record(_ROW_ITEMS, struct.pack("<I", len(lines))) + b"".join(lines)
            + build_record(_END_ROW_ITEMS))


def _column_fields_and_items(render: PivotRender) -> bytes:
    count = len(render.value_captions)
    col_field = render.col_field
    if col_field is None:
        if count == 1:
            fields = b""
            lines = [_line(0, 0, [])]
        else:
            fields = _field_list([-2])
            lines = [_line(0, 0, [0])] + [_line(0, 0, [index], index) for index in range(1, count)]
    elif count == 1:
        fields = _field_list([col_field])
        lines = [_line(0, 0, [position]) for position in range(len(render.col_items))]
        lines.append(_line(0, _DISPLAY_GRAND, [0]))
    else:
        fields = _field_list([col_field, -2])
        lines = []
        for position in range(len(render.col_items)):
            lines.append(_line(0, 0, [position, 0]))
            lines.extend(_line(1, 0, [index], index) for index in range(1, count))
        lines.append(_line(0, _DISPLAY_GRAND, [0]))
        lines.extend(_line(0, _DISPLAY_GRAND, [0], index) for index in range(1, count))
    return (fields + build_record(_COL_ITEMS, struct.pack("<I", len(lines))) + b"".join(lines)
            + build_record(_END_COL_ITEMS))


def _field_list(fields: list[int]) -> bytes:
    return (build_record(_COL_FIELDS, struct.pack(f"<I{len(fields)}i", len(fields), *fields))
            + build_record(_END_COL_FIELDS))


def pivot_table_bin(render: PivotRender, data_field_format_ids: list[int | None]) -> bytes:
    """The pivot table definition; ``data_field_format_ids`` are the styles.bin number formats."""
    pivot = render.pivot
    first_col, first_row, last_col, last_row = parse_cell_range(render.ref)
    pages = len(render.filter_indices)
    location = struct.pack(
        "<9I", first_row, last_row, first_col, last_col,
        first_row + render.first_header_row, first_row + render.first_data_row,
        first_col + 1, pages, 1 if pages else 0,
    )
    data_fields = set(render.value_indices)
    parts = [
        _header("table", render.pivot_id, render.plan.name),
        build_record(
            _VIEW,
            _view_prefix(pivot.autofit_columns) + struct.pack("<I", render.cache_id)
            + wide_string(render.plan.name) + wide_string(pivot.data_caption),
        ),
        build_record(_LOCATION, location),
        build_record(_END_LOCATION),
        build_record(_VDS, struct.pack("<I", len(render.column_names))),
    ]
    for index in range(len(render.column_names)):
        axis = 0
        if index == render.row_field:
            axis |= _AXIS_ROW
        if index == render.col_field:
            axis |= _AXIS_COL
        if index in render.filter_indices:
            axis |= _AXIS_PAGE
        if index in data_fields:
            axis |= _AXIS_DATA
        parts.append(build_record(_VD, bytes((axis,)) + _VD_TAIL))
        values = render.shared_items.get(index)
        if values is not None:
            display = (
                render.row_display if index == render.row_field
                else render.col_display if index == render.col_field
                else sorted(values, key=_sort_key)
            )
            positions = {value: position for position, value in enumerate(values)}
            parts.append(_items_records([positions[value] for value in display]))
        parts.append(build_record(_END_VD))
    parts.append(build_record(_END_VDS))
    parts.append(build_record(_ROW_FIELDS, struct.pack("<Ii", 1, render.row_field)))
    parts.append(build_record(_END_ROW_FIELDS))
    parts.append(_row_items(render))
    parts.append(_column_fields_and_items(render))
    if render.filter_indices:
        parts.append(build_record(_PIS, struct.pack("<I", pages)))
        for index in render.filter_indices:
            parts.append(build_record(_PI, struct.pack("<I", index) + _PAGE_ITEM_TAIL))
            parts.append(build_record(_END_PI))
        parts.append(build_record(_END_PIS))
    parts.append(build_record(_DIS, struct.pack("<I", len(render.value_captions))))
    for column, spec, numfmt in zip(render.value_indices, pivot.value_specs, data_field_format_ids):
        payload = struct.pack(
            "<6I", column, _SUBTOTALS[spec.function], 0, 0, 0, numfmt or 0,
        ) + b"\x01" + wide_string(spec.caption)
        parts.append(build_record(_DI, payload))
        parts.append(build_record(_END_DI))
    parts.append(build_record(_END_DIS))
    parts.append(build_record(_STYLE, struct.pack("<H", _STYLE_FLAGS) + wide_string(pivot.style)))
    parts.append(_VIEW_TAIL)
    parts.append(build_record(_END_VIEW))
    return b"".join(parts)


def workbook_pivot_caches_bin(cache_ids: list[tuple[int, str]]) -> bytes:
    """The workbook's list of caches: ``(cache id, relationship id)`` per cache."""
    if not cache_ids:
        return b""
    body = b"".join(
        build_record(_CACHE_ID, struct.pack("<I", cache_id) + wide_string(rel_id))
        + build_record(_END_CACHE_ID)
        for cache_id, rel_id in cache_ids
    )
    return build_record(_CACHE_IDS) + body + build_record(_END_CACHE_IDS)
