"""Add new chartsheets to a loaded XLSB package.

The chart and its drawing are the same DrawingML parts as in XLSX; the sheet
part is a short binary stream (as Excel writes it) and the sheet is listed in
``workbook.bin`` as one more BrtBundleSh record, after the other sheets.
"""

from __future__ import annotations

import struct

from .. import chart_parts
from .._xlsb_metadata import wide_string
from .._xlsb_sheets import _STATES, _Entry, _bundles
from ..biff_utils import build_record
from ..refs import column_label
from ._chartsheets import _CHART_TYPE, _DRAWING_TYPE, _DRAWING_XML, _set_rels, pending_chartsheets, source_title
from ._objects import _Package

_CHARTSHEET_TYPE = "application/vnd.ms-excel.chartsheet"
# A chartsheet as Excel writes it: properties, one view, page setup, margins,
# the drawing (relationship rId1) and an archive-id block.
_SHEET = b"".join((
    build_record(0x81),
    build_record(0x28B, bytes.fromhex("0100004000000000000000000000")),
    build_record(0x8B),
    build_record(0x8D, bytes.fromhex("03000a00000000000000")),
    build_record(0x8E),
    build_record(0x8C),
    build_record(0x29D, bytes(10)),
    build_record(0x1DC, struct.pack("<6d", 0.7, 0.7, 0.75, 0.75, 0.3, 0.3)),
    build_record(0x226, wide_string("rId1")),
    build_record(0x25, bytes.fromhex("010000100080")),
    build_record(0xC00, bytes.fromhex("67493498a1e88c4b87769e383fd428fd")),
    build_record(0x26),
    build_record(0x82),
))


def add_chartsheets_xlsb(updater, workbook) -> None:
    """Write the workbook's new chartsheets into the updater's package."""
    sheets = pending_chartsheets(workbook)
    if not sheets:
        return
    package = _Package(updater)
    book = package.get("xl/workbook.bin") or b""
    entries = _bundles(book)
    book_rels = package.rels("xl/workbook.bin")
    letters = [column_label(index) for index in range(16384)]
    tab_id = max((entry.tab_id for entry in entries), default=0)
    added = []
    for sheet in sheets:
        chart = package.free_part("xl/charts", "chart", "xml")
        package.set(chart, chart_parts.chart_xml(letters, sheet.charts[0], source_title(workbook, sheet)))
        package.override(chart, _CHART_TYPE)
        drawing = package.free_part("xl/drawings", "drawing", "xml")
        package.set(drawing, _DRAWING_XML)
        package.override(drawing, _DRAWING_TYPE)
        drawing_rels = package.rels(drawing)
        package.add_rel(drawing_rels, drawing, "chart", chart)
        part = package.free_part("xl/chartsheets", "sheet", "bin")
        package.set(part, _SHEET)
        package.override(part, _CHARTSHEET_TYPE)
        part_rels = package.rels(part)
        package.add_rel(part_rels, part, "drawing", drawing)
        for owner, rels in ((drawing, drawing_rels), (part, part_rels)):
            _set_rels(package, owner, rels)
        rid = package.add_rel(book_rels, "xl/workbook.bin", "chartsheet", part)
        tab_id += 1
        entries.append(_Entry(None, _STATES[sheet.sheet_state], tab_id, rid, sheet.title))
        added.append((sheet, part))
    package.set("xl/workbook.bin", updater._rewrite_sheet_list(book, entries, {}))
    _set_rels(package, "xl/workbook.bin", book_rels)
    package.commit()
    for sheet, part in added:
        sheet._part = part
        updater._sheets[sheet.title] = part
        updater._chartsheets.add(sheet.title)
