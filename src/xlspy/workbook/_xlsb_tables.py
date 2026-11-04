"""Add, change and remove the tables of a sheet of a loaded XLSB package.

A table is a binary part (``xl/tables/tableN.bin``) and the sheet lists its parts
in ``BrtBeginListParts``. A new table is written with the XLSB writer's own encoder.
A table that was edited keeps its part and only the records the object model can
change are re-encoded (``patch_table_part``); a table that was not edited keeps its
bytes; a removed table takes its part, relationship and content type with it.
"""

from __future__ import annotations
from .._package_handle import PackageHandle

import re
import struct
from dataclasses import dataclass, replace
from typing import TYPE_CHECKING, Any, cast

from .._xlsb_metadata import read_wide, replace_records, wide_string
from .._xlsb_table_part import TableEdit, TablePart, TablePartStyle, parse_table_part, patch_table_part, table_catalog
from .._table_totals import totals_formula
from ..biff_utils import build_record, iter_record_payloads
from ..exceptions import UnsupportedFeatureError
from ..styles.core import table_column_names
from ..writer import XlsbWriter
from .._xlsx_edit import _bytes
from ._objects import (
    _Package, _as_excel_table, _overlaps, _resolve_part_target, _rels_path, _table_signature, _target,
)
from ._xlsb_sheet_meta import LIST_PART, read_table_part
from ..refs import column_label, parse_cell_range

if TYPE_CHECKING:
    from ..xlsb_updater import XlsbUpdater
    from ._worksheet import Worksheet

BEGIN_LIST_PARTS = 0x294
END_LIST_PARTS = 0x296
_TABLE_TYPE = "application/vnd.ms-excel.table"
_AFTER_LIST_PARTS = (0x25, 0x82)
# Parts that refer to a table by name: after a table is renamed or removed they would
# point at a table that is not there.
_NAMING_PARTS = ("xl/pivotCache/", "xl/slicerCaches/")


def _header_row(sheet: Worksheet, table: Any) -> list:
    """The values of the table's header row, indexed by column."""

    first_col, first_row, last_col, _ = parse_cell_range(table.ref)
    row: list = [None] * (last_col + 1)
    for column in range(first_col, last_col + 1):
        cell = sheet._existing_cell(first_row + 1, column + 1)
        row[column] = None if cell is None else cell.value
    return row


def _row_values(sheet: Worksheet, table: Any, row_index: int) -> list:
    """The values of one row of ``table``, left to right."""

    first_col, _, last_col, _ = parse_cell_range(table.ref)
    values: list = []
    for column in range(first_col, last_col + 1):
        cell = sheet._existing_cell(row_index + 1, column + 1)
        values.append(None if cell is None else cell.value)
    return values


def _data_range(ref: str, totals_rows: int) -> str:
    """The range a table's filter covers: the header and the data, not the totals."""
    first_col, first_row, last_col, last_row = parse_cell_range(ref)
    last_row -= 1 if totals_rows else 0
    return f"{column_label(first_col)}{first_row + 1}:{column_label(last_col)}{last_row + 1}"


class _Listed:
    """A table part the sheet lists: its relationship id, part and, if readable, its content."""

    def __init__(self, rid: str, part: str | None, data: bytes | None) -> None:
        self.rid = rid
        self.part = part
        self.data = data
        self.parsed: TablePart | None = None
        try:
            if data is not None and read_table_part(data) is not None:
                self.parsed = parse_table_part(data)
        except (struct.error, IndexError, TypeError, ValueError):
            self.parsed = None


def _listed(package: _Package, path: str, rels: Any, data: bytes) -> list[_Listed]:
    targets = {node.get("Id"): _resolve_part_target(path, _target(node)) for node in rels}
    listed = []
    for record_id, payload in iter_record_payloads(data):
        if record_id == LIST_PART:
            rid = read_wide(payload, 0)[0]
            part = targets.get(rid)
            listed.append(_Listed(rid, part, package.get(part) if part else None))
    return listed


def _model_column_names(table: Any) -> list[Any]:
    """The column names the model holds for ``table``.

    A table the workbook loaded keeps its columns in ``tableColumns``; a table added by an earlier
    save of this session is kept as the ``ExcelTable`` that was added, with ``columns``.
    """
    columns = getattr(table, "tableColumns", None)
    if columns is not None:
        return [column.name for column in columns]
    added = getattr(table, "columns", None)
    return list(added) if added is not None else []


def _edit_for(sheet: Worksheet, table: Any, part: TablePart) -> TableEdit:
    """What the part of a loaded table is to say after the edits of the model."""
    excel = _as_excel_table(table)
    totals_changed = bool(excel.totals_row) != bool(part.totals_rows)
    names = _model_column_names(table)
    if names:
        if len(names) != len(part.columns) or not all(isinstance(name, str) and name for name in names):
            raise ValueError(f"table {excel.display_name!r}: its columns do not match its range")
        excel = replace(excel, columns=tuple(names))
    names = table_column_names(excel, _header_row(sheet, excel) if excel.header_row else None)
    auto_filter = None
    if part.auto_filter_ref is not None:
        wanted = table.autoFilter.ref if table.autoFilter is not None else part.auto_filter_ref
        if wanted != part.auto_filter_ref:
            auto_filter = wanted
        elif (excel.ref != part.ref or totals_changed) and part.auto_filter_ref == _data_range(part.ref, part.totals_rows or 0):
            auto_filter = _data_range(excel.ref, int(excel.totals_row))
    style = None
    info = table.tableStyleInfo
    if info is not None:
        style = TablePartStyle(
            name=info.name, first_column=bool(info.showFirstColumn), last_column=bool(info.showLastColumn),
            row_stripes=bool(info.showRowStripes), column_stripes=bool(info.showColumnStripes),
        )
    totals_cells = None
    if excel.totals_row and (excel.ref != part.ref or totals_changed):
        totals_cells = _row_values(sheet, excel, parse_cell_range(excel.ref)[3])
        if any(totals_formula(cell) is not None for cell in totals_cells):
            raise UnsupportedFeatureError(
                f"table {excel.display_name!r}: changing the range of a table with a formula "
                "totals row cannot be saved into a loaded XLSB workbook yet"
            )
    return TableEdit(
        ref=excel.ref, name=excel.display_name, header_rows=int(excel.header_row),
        totals_rows=int(excel.totals_row), column_names=names, auto_filter_ref=auto_filter, style=style,
        totals_cells=totals_cells,
    )


def _is_referred_to(package: _Package, ident: int) -> bool:
    """True when a formula of the workbook holds a structured reference (PtgList) to table ``ident``.

    The tokens are looked for in the bytes of the sheets and of the workbook part, so a chance
    match of twelve bytes would refuse a removal that was safe, never allow an unsafe one.
    """
    # PtgExtend, PtgList, ixti (any extern-sheet index), two bytes of flags, the table id.
    pattern = re.compile(re.escape(b"\x18\x19") + b"...." + re.escape(struct.pack("<I", ident)), re.S)
    for name in package.names:
        if name == "xl/workbook.bin" or (name.startswith("xl/") and name.endswith(".bin") and "/worksheets/" in name):
            data = package.get(name)
            if data and pattern.search(data):
                return True
    return False


def _other_tables(package: _Package, own: set[str | None]) -> list[Any]:
    """The tables of the workbook's other sheets, whose names and ids are taken."""
    others = []
    for name in sorted(package.names):
        if name.startswith("xl/tables/") and name.endswith(".bin") and name not in package.removed and name not in own:
            try:
                table = read_table_part(package.get(name) or b"")
            except (struct.error, IndexError, TypeError, ValueError):
                table = None
            if table is not None:
                others.append(table)
    return others


@dataclass
class PendingTable:
    """A table the save adds: its id, name, column names and sheet, known before its part is written."""

    table: Any
    ident: int
    name: str
    columns: list[str] | None
    sheet: str


def plan_new_tables(workbook: Any, updater: XlsbUpdater) -> tuple[PendingTable, ...]:
    """Give each table the sheets gain an id, in workbook order, before any formula is written.

    A formula refers to a table by its id, and the formulas of a sheet are written before the
    sheet's tables are; so the ids are fixed first, and the table parts are written with them.
    """
    package = _Package(updater)
    catalog = table_catalog(package.names, package.get)
    next_id = max(catalog, default=0) + 1
    pending: list[PendingTable] = []
    for sheet in workbook._sheets:
        saved = {id(table) for table in getattr(sheet, "_saved_table_objects", ())}
        for table in sheet.tables:
            if id(table) in saved:
                continue
            excel = _as_excel_table(table)
            try:
                columns: list[str] | None = table_column_names(excel, _header_row(sheet, excel) if excel.header_row else None)
            except ValueError:
                columns = None  # the save reports it, in its own order of checks
            pending.append(PendingTable(table, next_id, excel.display_name, columns, sheet.title))
            next_id += 1
    return tuple(pending)


def sync_tables(updater: XlsbUpdater, sheet: Worksheet, pending: tuple[PendingTable, ...] = ()) -> None:
    """Save the tables the sheet gained, changed or lost since it was loaded or last saved."""
    saved = list(getattr(sheet, "_saved_table_objects", ()))
    signatures = (getattr(sheet, "_saved_object_signatures", None) or {}).get("tables", ())
    current = list(sheet.tables)
    current_ids = {id(table) for table in current}
    saved_ids = {id(table) for table in saved}
    added = [table for table in current if id(table) not in saved_ids]
    removed = {index for index, table in enumerate(saved) if id(table) not in current_ids}
    edited = {
        index for index, table in enumerate(saved)
        if id(table) in current_ids and _table_signature(table) != signatures[index]
    }
    if not (added or removed or edited):
        return

    handle = PackageHandle(updater)
    package = _Package(updater)
    path = handle.sheets[sheet.title]
    rels = package.rels(path)
    data = package.get(path) or b""
    listed = _listed(package, path, rels, data)
    # The model holds the tables it could read, in the order the sheet lists them.
    readable = [item for item in listed if item.parsed is not None]
    if len(readable) != len(saved):
        raise UnsupportedFeatureError("table changes cannot be saved: the sheet's tables do not match the file")
    parsed = [cast(TablePart, item.parsed) for item in readable]
    renamed = {index for index in edited if saved[index].displayName != parsed[index].name}
    if (removed or renamed) and any(name.startswith(_NAMING_PARTS) for name in package.names):
        raise UnsupportedFeatureError(
            "removing or renaming a table cannot be saved into a workbook with pivot tables or slicers yet: "
            "they would keep pointing at the old table"
        )

    for index in edited - removed:
        first_col, _, last_col, _ = parse_cell_range(saved[index].ref)
        if last_col - first_col + 1 != len(parsed[index].columns):
            raise UnsupportedFeatureError(
                f"table {parsed[index].name!r}: changing its number of columns cannot be saved "
                "into a loaded XLSB workbook yet"
            )

    for index in sorted(removed):
        if _is_referred_to(package, parsed[index].id):
            raise UnsupportedFeatureError(
                f"table {parsed[index].name!r} is used by a formula, so removing it cannot be saved "
                "into a loaded XLSB workbook yet: the formula would refer to a table that is gone"
            )

    # Every check comes before the first change.
    others = _other_tables(package, {item.part for item in listed})
    taken = {table.displayName.casefold() for table in others if table.displayName}
    kept = [index for index in range(len(saved)) if index not in removed]
    final = [(table, _as_excel_table(table)) for table in [*(saved[index] for index in kept), *added]]
    for _, excel in final:
        if excel.display_name.casefold() in taken:
            raise ValueError(f"a table named {excel.display_name!r} already exists in the workbook")
        taken.add(excel.display_name.casefold())
    changed = {id(saved[index]) for index in edited} | {id(table) for table in added}
    for position, (table, excel) in enumerate(final):
        for other_table, other in final[:position]:
            if (id(table) in changed or id(other_table) in changed) and _overlaps(excel.ref, other.ref):
                raise ValueError(
                    f"table {excel.display_name!r} ({excel.ref}) overlaps the sheet's table ({other.ref})"
                )
    patched = {}
    for index in sorted(edited - removed):
        part = readable[index].part
        assert part is not None and readable[index].data is not None
        patched[part] = patch_table_part(cast(bytes, readable[index].data), _edit_for(sheet, saved[index], parsed[index]))
    new_columns = [table_column_names(excel, _header_row(sheet, excel)) for table, excel in final[len(kept):]]
    new_totals = [
        _row_values(sheet, excel, parse_cell_range(excel.ref)[3]) if excel.totals_row else None
        for _, excel in final[len(kept):]
    ]

    for part, body in patched.items():
        package.set(part, body)
    rids = [item.rid for item in listed]
    for index in sorted(removed):
        item = readable[index]
        rids.remove(item.rid)
        for node in list(rels):
            if node.get("Id") == item.rid:
                rels.remove(node)
        if item.part:
            package.remove(item.part)
            package.remove(_rels_path(item.part))
            for node in list(package.types):
                if node.get("PartName") == "/" + item.part:
                    package.types.remove(node)
    planned = {id(item.table): item.ident for item in pending}
    next_id = max([table.id or 0 for table in others] + [part.id for part in parsed], default=0) + 1
    formula_context = handle.formula_context()
    for (table, excel), columns, totals in zip(final[len(kept):], new_columns, new_totals):
        part = package.free_part("xl/tables", "table", "bin")
        ident = planned.get(id(table), next_id)
        package.set(part, cast(Any, XlsbWriter)._table_bin(
            XlsbWriter, excel, ident, columns, totals, formula_context, sheet.title,
        ))
        package.override(part, _TABLE_TYPE)
        rids.append(package.add_rel(rels, path, "table", part))
        next_id = max(next_id, ident) + 1

    block = b""
    if rids:
        block = build_record(BEGIN_LIST_PARTS, struct.pack("<I", len(rids)))
        for rid in rids:
            block += build_record(LIST_PART, wide_string(rid))
        block += build_record(END_LIST_PARTS)
    package.set(path, replace_records(
        data, {BEGIN_LIST_PARTS, LIST_PART, END_LIST_PARTS}, block, before=_AFTER_LIST_PARTS,
    ))
    package.set(_rels_path(path), _bytes(rels))
    package.commit()
