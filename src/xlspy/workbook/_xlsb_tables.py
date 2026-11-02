"""Add tables to a sheet of a loaded XLSB package.

A table is a binary part (``xl/tables/tableN.bin``) written with the XLSB
writer's own encoder; the sheet lists its parts in ``BrtBeginListParts``.
"""

from __future__ import annotations
from .._package_handle import PackageHandle

import struct
from typing import TYPE_CHECKING, Any, cast

from .._xlsb_metadata import read_wide, replace_records, wide_string
from ..biff_utils import build_record, iter_record_payloads
from ..styles.core import table_column_names
from ..writer import XlsbWriter
from .._xlsx_edit import _bytes
from ._objects import _Package, _as_excel_table, _overlaps, _rels_path, pending_objects
from ._xlsb_sheet_meta import LIST_PART, read_table_part

if TYPE_CHECKING:
    from ..xlsb_updater import XlsbUpdater
    from ._worksheet import Worksheet

BEGIN_LIST_PARTS = 0x294
END_LIST_PARTS = 0x296
_TABLE_TYPE = "application/vnd.ms-excel.table"
_AFTER_LIST_PARTS = (0x25, 0x82)


def _header_row(sheet: Worksheet, table: Any) -> list:
    """The values of the table's header row, indexed by column."""
    from ..refs import parse_cell_range

    first_col, first_row, last_col, _ = parse_cell_range(table.ref)
    row: list = [None] * (last_col + 1)
    for column in range(first_col, last_col + 1):
        cell = sheet._existing_cell(first_row + 1, column + 1)
        row[column] = None if cell is None else cell.value
    return row


def add_tables(updater: XlsbUpdater, sheet: Worksheet) -> None:
    added = pending_objects(sheet)["tables"]
    if not added:
        return
    package = _Package(updater)
    handle = PackageHandle(updater)
    path = handle.sheets[sheet.title]
    rels = package.rels(path)

    existing = []
    for name in sorted(package.names):
        if name.startswith("xl/tables/") and name.endswith(".bin") and name not in package.removed:
            try:
                table = read_table_part(package.get(name) or b"")
            except (struct.error, IndexError, TypeError, ValueError):
                table = None
            if table is not None:
                existing.append((name, table))
    names = {table.displayName.casefold() for _, table in existing if table.displayName}
    next_id = max((table.id or 0 for _, table in existing), default=0) + 1
    # Tables of this sheet, loaded or added before: new ones may not overlap them.
    refs = [table.ref for table in sheet.tables if not any(table is new for new in added)]

    rids = []
    for item in added:
        table = _as_excel_table(item)
        if table.display_name.casefold() in names:
            raise ValueError(f"a table named {table.display_name!r} already exists in the workbook")
        for ref in refs:
            if _overlaps(table.ref, ref):
                raise ValueError(
                    f"table {table.display_name!r} ({table.ref}) overlaps the sheet's table ({ref})"
                )
        columns = table_column_names(table, _header_row(sheet, table))
        part = package.free_part("xl/tables", "table", "bin")
        package.set(part, cast(Any, XlsbWriter)._table_bin(XlsbWriter, table, next_id, columns))
        package.override(part, _TABLE_TYPE)
        rids.append(package.add_rel(rels, path, "table", part))
        names.add(table.display_name.casefold())
        refs.append(table.ref)
        next_id += 1

    data = package.get(path) or b""
    listed = []
    for record_id, payload in iter_record_payloads(data):
        if record_id == LIST_PART:
            listed.append(read_wide(payload, 0)[0])
    block = build_record(BEGIN_LIST_PARTS, struct.pack("<I", len(listed) + len(rids)))
    for rid in (*listed, *rids):
        block += build_record(LIST_PART, wide_string(rid))
    block += build_record(END_LIST_PARTS)
    package.set(path, replace_records(
        data, {BEGIN_LIST_PARTS, LIST_PART, END_LIST_PARTS}, block, before=_AFTER_LIST_PARTS,
    ))
    package.set(_rels_path(path), _bytes(rels))
    package.commit()
