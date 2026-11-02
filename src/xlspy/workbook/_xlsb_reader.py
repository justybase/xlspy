"""Load an XLSB file into the editable workbook model (values first)."""

from __future__ import annotations

from typing import Any

import os
import posixpath
import struct
from .._safe_xml import fromstring as safe_fromstring

from .._xlsb_formulas import FormulaContext
from .._xlsb_metadata import read_wide
from itertools import chain as _chain

from .._accel import xlsb_load_cells
from .._errors import CellError
from .._xlsb_read import cached_cell, formula_groups, formula_text, iter_xlsb_indexed_rows
from .._xlsb_styles import XlsbStyleTable
from .._xlsb_validation import dval_rule
from ..biff_utils import iter_record_payloads
from ..comments import rels_path
from ..external_links import read_external_links
from ..hyperlinks import Hyperlink
from ..reader import ExcelReader, _get_date_num_fmt_ids
from ..refs import column_label, parse_cell_range
from ..xlsb_updater import XlsbUpdater
from ..chart_reader import read_drawing
from ._objects import note_unread
from ..updater_utils import resolve_zip_target
from ._model import CELL_SLOTS, AutoFilter, Cell, DefinedNameList, NamedStyleList, Workbook, Worksheet
from ._reader import _read_custom_properties, _read_print_names, _read_properties, _read_sheet_relationships
from ._xlsb_layout import BEGIN_FILTER, END_FILTER, builtin_names
from ._xlsb_loaded import remember_saved_state, remember_workbook_state
from .._xlsb_cf import (
    BRT_BEGIN_CF_RULE, BRT_BEGIN_CONDITIONAL_FORMATTING, SCALE_RECORDS, _CF_RULE_EXT,
    read_conditional_formats,
)
from ._xlsb_sheet_meta import (
    FILTER_RECORDS, LIST_PART, RECORDS as SHEET_RECORDS, ROW_FLAGS, ROW_HEADER,
    apply_sheet_records, read_filter_record, read_table_part, row_header,
)


_CF_RECORDS = SCALE_RECORDS | {BRT_BEGIN_CONDITIONAL_FORMATTING, BRT_BEGIN_CF_RULE, _CF_RULE_EXT}


def _external_links(package, sheet_path: str) -> dict[str, str]:
    """External relationship targets of a worksheet, by relationship id."""
    path = rels_path(sheet_path)
    if not package.has(path):
        return {}
    return {
        rel.get("Id"): rel.get("Target")
        for rel in safe_fromstring(package.get(path))
        if rel.get("TargetMode") == "External" and rel.get("Id") and rel.get("Target")
    }


def _load_cells(sheet, reader, name, sheet_data, data_only):
    """Read the cells into ``sheet``. Return the ``(start, end)`` offsets of the
    sheet data when the native loader read it, else None (the records were read in
    Python, row headers excluded)."""
    if xlsb_load_cells is None:
        _load_cells_python(sheet, reader, name, sheet_data, data_only)
        return None
    max_row, max_column, formulas, headers, groups, start, end = xlsb_load_cells(
        sheet_data, reader._shared_string_array, reader._styles_cell_xfs_array,
        _get_date_num_fmt_ids(reader._number_formats_type_dict), reader._date_1904,
        Cell, CELL_SLOTS, sheet, sheet._cells,
    )
    sheet._max_row = max(sheet._max_row, max_row)
    sheet._max_column = max(sheet._max_column, max_column)
    for _row, payload in headers:
        row_header(sheet, payload)
    if formulas:
        context = None if data_only else FormulaContext(
            reader.get_sheet_names(), reader._archive.read("xl/workbook.bin"), reader.read_part,
        )
        shared = formula_groups(groups)
        for row, kind, payload in formulas:
            value, formula, style = cached_cell(kind, payload, reader)
            column = struct.unpack_from("<I", payload)[0]
            text = None
            if not data_only and formula is not None:
                text = formula_text(formula, shared, context, name, row - 1, column)
            _set_loaded_cell(sheet, row, column + 1, value, style, text)
    if start < 0:
        return 0, 0
    return start, end


def _set_loaded_cell(sheet, row, column, value, style_id, formula_text=None):
    if formula_text is not None:
        sheet._set_cell_raw(row, column, None, style_id=style_id, data_type="f",
                            formula_text=formula_text)
    elif value is None:
        if style_id:
            sheet._set_cell_raw(row, column, None, style_id=style_id)
    elif isinstance(value, CellError):
        sheet._set_cell_raw(row, column, value[6:], style_id=style_id, data_type="e")
    else:
        kind = "s" if isinstance(value, str) else "b" if isinstance(value, bool) else "n"
        sheet._set_cell_raw(row, column, value, style_id=style_id, data_type=kind)


def _load_cells_python(sheet, reader, name, sheet_data, data_only):
    for row_number, cells in iter_xlsb_indexed_rows(
        reader, name, data_only=data_only, data=sheet_data, with_styles=True,
    ):
        for column, value, style_id in cells:
            if type(value) is str and value.startswith("="):
                # iter_indexed_rows marks decoded formulas with "=".
                _set_loaded_cell(sheet, row_number, column + 1, None, style_id, value[1:])
            else:
                _set_loaded_cell(sheet, row_number, column + 1, value, style_id)


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
        named_styles: list[Any] = list(workbook._styles.named_style_defs.values())
        workbook.named_styles = NamedStyleList(named_styles or list(workbook._styles.named_styles) or ["Normal"])
        formula_context = None
        for name in updater.get_worksheet_names():
            sheet = Worksheet(workbook, name)
            sheet._loaded = True
            sheet._path = updater._sheets[name]
            sheet._cf_extensions = {}
            sheet._original_title = name
            with updater._package.open_entry(sheet._path) as stream:
                sheet_data = stream.read()
            span = _load_cells(sheet, reader, name, sheet_data, data_only)
            links = None
            validations = []
            metadata = []
            conditional = []
            list_parts = []
            filter_records = None
            # The native loader has read the sheet data (with its row headers):
            # only the records around it are left.
            outside = (
                iter_record_payloads(sheet_data) if span is None
                else _chain(iter_record_payloads(sheet_data[:span[0]]),
                            iter_record_payloads(sheet_data[span[1]:]))
            )
            for record_id, payload in outside:
                if filter_records is not None:
                    filter_records.append((record_id, payload))
                    if record_id == END_FILTER:
                        sheet._saved_filter_records = filter_records
                        filter_records = None
                if record_id == ROW_HEADER:
                    if len(payload) > ROW_FLAGS and payload[ROW_FLAGS] & 0x7F:
                        row_header(sheet, payload)
                    continue
                if record_id in SHEET_RECORDS:
                    metadata.append((record_id, payload))
                    continue
                if record_id in _CF_RECORDS:
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
                if record_id == BEGIN_FILTER and len(payload) >= 16:
                    filter_records = [(record_id, payload)]
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
                    updater._part_or_none,
                )
            read = 0
            extensions: dict = {}
            for sqref, rule in read_conditional_formats(
                conditional, formula_context, workbook._styles, extensions,
            ):
                extension = extensions.get(id(rule))
                rule = sheet._conditional_formats.add(sqref, rule)
                read += 1
                if extension is not None:
                    # Kept with the rule while it stays as it was read.
                    sheet._cf_extensions[id(rule)] = (rule, repr(rule), extension)
            note_unread(
                sheet, "conditional_formatting",
                sum(record_id == BRT_BEGIN_CF_RULE for record_id, _ in conditional) - read,
            )
            if validations:
                for payload in validations:
                    rule = dval_rule(payload, formula_context)
                    if rule is None:
                        note_unread(sheet, "data_validations")
                    else:
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
        # The file counts every sheet entry, chartsheets included.
        all_names = updater.get_sheet_names()
        active_name = all_names[min(updater.get_active_sheet(), len(all_names) - 1)]
        titles = [sheet.title for sheet in workbook._sheets]
        workbook._active_index = titles.index(active_name) if active_name in titles else 0
        from ..chart_reader import read_chartsheet
        from ._model import Chartsheet

        for index, name in enumerate(updater.get_chartsheet_names()):
            chartsheet = Chartsheet(workbook, name, index=index)
            chartsheet.sheet_state = states.get(name, "visible")
            chartsheet._part = updater._sheets[name]
            read_chartsheet(updater._package, chartsheet._part, name, chartsheet)
            workbook._chartsheets.append(chartsheet)
        workbook._loaded_active_index = workbook._active_index
        workbook._properties = _read_properties(updater._package)
        workbook.custom_doc_props = _read_custom_properties(updater._package)
        workbook.external_links = read_external_links(updater._package, binary=True)
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
