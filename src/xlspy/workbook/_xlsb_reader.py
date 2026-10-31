"""Load an XLSB file into the editable workbook model (values first)."""

from __future__ import annotations

import os
import posixpath
import struct
from xml.etree import ElementTree as ET

from .._xlsb_formulas import FormulaContext
from .._xlsb_metadata import read_wide
from .._xlsb_read import iter_indexed_rows
from .._xlsb_styles import XlsbStyleTable
from .._xlsb_validation import dval_rule
from ..biff_utils import iter_record_payloads
from ..comments import rels_path
from ..hyperlinks import Hyperlink
from ..reader import ExcelReader
from ..refs import column_label, parse_cell_range
from ..xlsb_updater import XlsbUpdater
from ..chart_reader import read_drawing
from ..updater_utils import resolve_zip_target
from ._model import AutoFilter, DefinedNameList, Workbook, Worksheet
from ._reader import _read_custom_properties, _read_print_names, _read_properties, _read_sheet_relationships
from ._xlsb_layout import builtin_names
from ._xlsb_loaded import remember_saved_state, remember_workbook_state
from .._xlsb_cf import BRT_BEGIN_CF_RULE, BRT_BEGIN_CONDITIONAL_FORMATTING, read_conditional_formats
from ._xlsb_sheet_meta import (
    FILTER_RECORDS, LIST_PART, RECORDS as SHEET_RECORDS, ROW_FLAGS, ROW_HEADER,
    apply_sheet_records, read_filter_record, read_table_part, row_header,
)


def _external_links(package, sheet_path: str) -> dict[str, str]:
    """External relationship targets of a worksheet, by relationship id."""
    path = rels_path(sheet_path)
    if not package.has(path):
        return {}
    return {
        rel.get("Id"): rel.get("Target")
        for rel in ET.fromstring(package.get(path))
        if rel.get("TargetMode") == "External" and rel.get("Id") and rel.get("Target")
    }


def load_workbook_xlsb(path: str | os.PathLike[str], *, data_only: bool = False) -> Workbook:
    path = os.fspath(path)
    updater = XlsbUpdater(path)
    reader = ExcelReader(path)
    reader.open(path)
    workbook = Workbook()
    workbook._sheets.clear()
    workbook._loaded = True
    workbook._read_only = False
    workbook._updater = updater
    workbook._source_path = path
    workbook._format = "xlsb"
    workbook._loaded_worksheet_names = set()
    try:
        styles_path = reader._styles_location and reader._resolve_path(reader._styles_location)
        styles_bin = (
            updater._package.get(styles_path)
            if styles_path and updater._package.has(styles_path) else None
        )
        workbook._styles = XlsbStyleTable.from_bin(styles_bin)
        workbook._xlsb_styles_path = styles_path if styles_bin is not None else None
        formula_context = None
        for name in updater.get_sheet_names():
            sheet = Worksheet(workbook, name)
            sheet._loaded = True
            sheet._path = updater._sheets[name]
            sheet._original_title = name
            with updater._package.open_entry(sheet._path) as stream:
                sheet_data = stream.read()
            set_cell = sheet._set_cell_raw
            for row_number, cells in iter_indexed_rows(
                reader, name, data_only=data_only, data=sheet_data, with_styles=True,
            ):
                for column, value, style_id in cells:
                    if value is None:
                        if style_id:
                            set_cell(row_number, column + 1, None, style_id=style_id)
                    elif type(value) is str and value[:1] in ("=", "e"):
                        if value.startswith("="):
                            set_cell(row_number, column + 1, None, style_id=style_id,
                                     formula_text=value[1:])
                        else:
                            if value.startswith("error:"):
                                value = value[6:]
                            set_cell(row_number, column + 1, value, style_id=style_id)
                    else:
                        set_cell(row_number, column + 1, value, style_id=style_id)
            links = None
            validations = []
            metadata = []
            conditional = []
            list_parts = []
            for record_id, payload in iter_record_payloads(sheet_data):
                if record_id == ROW_HEADER:
                    if len(payload) > ROW_FLAGS and payload[ROW_FLAGS] & 0x7F:
                        row_header(sheet, payload)
                    continue
                if record_id in SHEET_RECORDS:
                    metadata.append((record_id, payload))
                    continue
                if record_id in (BRT_BEGIN_CONDITIONAL_FORMATTING, BRT_BEGIN_CF_RULE):
                    conditional.append((record_id, payload))
                    continue
                if record_id == LIST_PART:
                    list_parts.append(read_wide(payload, 0)[0])
                    continue
                if record_id == 0x40:  # BrtDVal
                    validations.append(payload)
                    continue
                if record_id == 0xB0 and len(payload) >= 16:  # BrtMergeCell
                    r1, r2, c1, c2 = struct.unpack_from("<4I", payload)
                    sheet._merged_cells.add(
                        f"{column_label(c1)}{r1 + 1}:{column_label(c2)}{r2 + 1}"
                    )
                    sheet._max_row = max(sheet._max_row, r2 + 1)
                    sheet._max_column = max(sheet._max_column, c2 + 1)
                    continue
                if record_id == 0xA1 and len(payload) >= 16:  # BrtBeginAFilter
                    r1, r2, c1, c2 = struct.unpack_from("<4I", payload)
                    if r1 <= r2 < 1048576 and c1 <= c2 < 16384:
                        sheet._auto_filter = AutoFilter(
                            f"{column_label(c1)}{r1 + 1}:{column_label(c2)}{r2 + 1}"
                        )
                    continue
                if record_id in FILTER_RECORDS:
                    read_filter_record(sheet._auto_filter, record_id, payload)
                    continue
                if record_id != 0x1EE or len(payload) < 16:
                    continue
                row, _row2, column, _column2 = struct.unpack_from("<4I", payload)
                offset = 16
                relationship, offset = read_wide(payload, offset)
                location, offset = read_wide(payload, offset)
                tooltip, offset = read_wide(payload, offset)
                display, offset = read_wide(payload, offset)
                # As for XLSX: an external relationship wins, else the location.
                target = None
                if relationship:
                    if links is None:
                        links = _external_links(updater._package, sheet._path)
                    target = links.get(relationship)
                if target is None and location:
                    target = "#" + location
                if not target:
                    continue
                try:
                    link = Hyperlink(target, text=display or None, tooltip=tooltip or None)
                except (TypeError, ValueError):
                    continue
                cell = sheet._existing_cell(row + 1, column + 1)
                if cell is None:
                    cell = sheet._set_cell_raw(row + 1, column + 1, None)
                cell._hyperlink = link
            apply_sheet_records(sheet, metadata)
            relationships = _read_sheet_relationships(updater._package, sheet._path)
            for target, _mode, kind in relationships.values():
                if kind.endswith("/drawing"):
                    read_drawing(
                        updater._package, sheet,
                        resolve_zip_target(target, posixpath.dirname(sheet._path)),
                    )
            for rid in list_parts:
                target = relationships.get(rid, (None,))[0]
                part = target and resolve_zip_target(target, posixpath.dirname(sheet._path))
                if part and updater._package.has(part):
                    try:
                        table = read_table_part(updater._package.get(part))
                    except (struct.error, IndexError, TypeError, ValueError):
                        table = None
                    if table is not None:
                        sheet._tables.append(table)
            if (validations or conditional) and formula_context is None:
                formula_context = FormulaContext(
                    updater.get_sheet_names(), updater._package.get("xl/workbook.bin"),
                )
            for sqref, rule in read_conditional_formats(conditional, formula_context, workbook._styles):
                sheet._conditional_formats.add(sqref, rule)
            if validations:
                for payload in validations:
                    rule = dval_rule(payload, formula_context)
                    if rule is not None:
                        sheet._data_validations.append(rule)
            for ref, comment in reader.get_comments(name).items():
                try:
                    col, row, _, _ = parse_cell_range(ref)
                except (TypeError, ValueError):
                    continue
                cell = sheet._existing_cell(row + 1, col + 1)
                if cell is None:
                    cell = sheet._set_cell_raw(row + 1, col + 1, None)
                cell._comment = comment
            workbook._sheets.append(sheet)
        names = reader.get_defined_names()
        workbook._defined_names = DefinedNameList(workbook, names)
        workbook._loaded_defined_names = {
            (item.name, item.formula, item.scope) for item in names
        }
        workbook._loaded_worksheet_names = set(workbook.sheetnames)
        states = updater.get_sheet_states()
        for sheet in workbook._sheets:
            sheet._sheet_state = states.get(sheet.title, "visible")
        workbook._active_index = min(updater.get_active_sheet(), len(workbook._sheets) - 1)
        workbook._loaded_active_index = workbook._active_index
        workbook._properties = _read_properties(updater._package)
        workbook.custom_doc_props = _read_custom_properties(updater._package)
        print_names = builtin_names(FormulaContext(updater._sheets, updater._package.get("xl/workbook.bin")))
        for sheet in workbook._sheets:
            _read_print_names(sheet, print_names.get(sheet.title, {}))
            sheet._clear_dirty()
            remember_saved_state(sheet)
        remember_workbook_state(workbook)
        return workbook
    except BaseException:
        reader.close()
        updater.close()
        raise
    finally:
        reader.close()
