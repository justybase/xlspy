"""Worksheet metadata records of an XLSB part, read into the editable model.

The loader collects the records while it walks the sheet once and hands them
here in file order. The layouts mirror the encoders in ``writer.py`` and
``_xlsb_print.py``; what the model cannot express is left in the file.
"""

from __future__ import annotations

import struct

from .._xlsb_metadata import read_wide
from .._xlsb_print import (
    BRT_BEGIN_COL_BREAKS, BRT_BEGIN_HEADER_FOOTER, BRT_BEGIN_ROW_BREAKS, BRT_BREAK,
    BRT_END_COL_BREAKS, BRT_END_ROW_BREAKS, BRT_MARGINS, BRT_PAGE_SETUP,
    BRT_PRINT_OPTIONS, FIT_TO_PAGE_BIT, FIT_TO_PAGE_BYTE,
)
from .._xlsb_styles import _read_brt_color
from ..refs import column_label
from ._model import Break, ColBreak, Pane, RowBreak, Selection
from ._xlsb_layout import FORMAT_INFO as FORMAT_INFO

ROW_HEADER = 0x00
COL_INFO = 0x3C
WS_PROP = 0x93
BEGIN_WS_VIEW = 0x89
PANE = 0x97
SELECTION = 0x98
SHEET_PROTECTION = 0x217
# Records collected for apply_sheet_records (row headers are filtered first).
RECORDS = frozenset({
    COL_INFO, WS_PROP, BEGIN_WS_VIEW, PANE, SELECTION, FORMAT_INFO, SHEET_PROTECTION,
    BRT_PRINT_OPTIONS,
    BRT_MARGINS, BRT_PAGE_SETUP, BRT_BEGIN_HEADER_FOOTER, BRT_BEGIN_ROW_BREAKS,
    BRT_END_ROW_BREAKS, BRT_BEGIN_COL_BREAKS, BRT_END_COL_BREAKS, BRT_BREAK,
})
# BrtRowHdr flags byte: outline level (bits 0-2), collapsed, hidden, custom
# height (fUnsynced) and fGhostDirty (the row has a style).
ROW_FLAGS = 11
_PROTECTION_FIELDS = (
    "sheet", "objects", "scenarios", "formatCells", "formatColumns", "formatRows",
    "insertColumns", "insertRows", "insertHyperlinks", "deleteColumns", "deleteRows",
    "selectLockedCells", "sort", "autoFilter", "pivotTables", "selectUnlockedCells",
)
_VIEWS = {1: "pageBreakPreview", 2: "pageLayout"}
_PANES = {0: "bottomRight", 1: "topRight", 2: "bottomLeft"}
_HEADER_FOOTER_PARTS = (
    "oddHeader", "oddFooter", "evenHeader", "evenFooter", "firstHeader", "firstFooter",
)


def row_header(sheet, payload: bytes) -> None:
    """Row height, visibility, outline level and style from a BrtRowHdr."""
    row = struct.unpack_from("<I", payload, 0)[0] + 1
    style, height = struct.unpack_from("<IH", payload, 4)
    flags = payload[ROW_FLAGS]
    dimension = sheet.row_dimensions[row]
    if flags & 0x20:
        dimension.height = height / 20
    dimension.hidden = bool(flags & 0x10)
    dimension.outlineLevel = flags & 0x07
    dimension.collapsed = bool(flags & 0x08)
    if flags & 0x40:
        dimension.style = style


def _column_info(sheet, payload: bytes) -> None:
    first, last, width, style, flags = struct.unpack_from("<IIIIH", payload, 0)
    if first > last or last > 16383:
        return
    for index in range(first, last + 1):
        dimension = sheet.column_dimensions[column_label(index)]
        dimension.min = first + 1
        dimension.max = last + 1
        dimension.width = width / 256
        dimension.hidden = bool(flags & 0x01)
        dimension.bestFit = True if flags & 0x04 else None
        dimension.outlineLevel = (flags >> 8) & 0x07
        dimension.collapsed = bool(flags & 0x1000)
        if style:
            dimension.style = style


def _read_sheet_properties(sheet, payload: bytes) -> None:
    properties = sheet.sheet_properties
    color = _read_brt_color(payload, 3)
    if color is not None and not color.auto:
        properties.tabColor = color
    if payload[FIT_TO_PAGE_BYTE] & FIT_TO_PAGE_BIT:
        properties.pageSetUpPr.fitToPage = True


def _sheet_view(sheet, payload: bytes) -> None:
    view = sheet._sheet_view
    flags = payload[0]
    view.windowProtection = True if flags & 0x01 else None
    view.showFormulas = True if flags & 0x02 else None
    view.showGridLines = False if not flags & 0x04 else None
    view.showRowColHeaders = False if not flags & 0x08 else None
    view.showZeros = False if not flags & 0x10 else None
    view.rightToLeft = True if flags & 0x20 else None
    view.tabSelected = True if flags & 0x40 else None
    kind, top, left = struct.unpack_from("<III", payload, 2)
    view.view = _VIEWS.get(kind)
    view.topLeftCell = f"{column_label(left)}{top + 1}" if top or left else None
    if len(payload) >= 22:
        scale, normal = struct.unpack_from("<HH", payload, 18)
        view.zoomScale = scale if scale not in (0, 100) else None
        view.zoomScaleNormal = normal if normal not in (0, 100) else None


def _pane(sheet, payload: bytes) -> None:
    x_split, y_split = struct.unpack_from("<dd", payload, 0)
    top, left = struct.unpack_from("<II", payload, 16)
    flags = payload[28] if len(payload) > 28 else 0
    frozen = bool(flags & 0x03)
    if not (x_split or y_split):
        return
    cell = f"{column_label(left)}{top + 1}"
    sheet._sheet_view.pane = Pane(
        xSplit=x_split or None, ySplit=y_split or None, topLeftCell=cell,
        activePane="bottomRight" if x_split and y_split
        else "topRight" if x_split else "bottomLeft",
        state="frozen" if frozen else "split",
    )
    if frozen:
        sheet._freeze_panes = cell


def _selection(sheet, payload: bytes) -> None:
    pane, row, column, active_id, count = struct.unpack_from("<5I", payload, 0)
    ranges = []
    for index in range(count):
        r1, r2, c1, c2 = struct.unpack_from("<4I", payload, 20 + index * 16)
        first = f"{column_label(c1)}{r1 + 1}"
        ranges.append(first if (r1, c1) == (r2, c2) else f"{first}:{column_label(c2)}{r2 + 1}")
    sheet._sheet_view.selection.append(Selection(
        pane=_PANES.get(pane), activeCell=f"{column_label(column)}{row + 1}",
        activeCellId=active_id or None, sqref=" ".join(ranges) or "A1",
    ))


def _sheet_format(sheet, payload: bytes) -> None:
    width, base, height, flags, row_level, column_level = struct.unpack_from("<IHHHBB", payload, 0)
    fmt = sheet.sheet_format
    fmt.defaultColWidth = None if width == 0xFFFFFFFF else width / 256
    fmt.baseColWidth = base
    fmt.defaultRowHeight = height / 20
    fmt.customHeight = True if flags & 0x01 else None
    fmt.zeroHeight = True if flags & 0x02 else None
    fmt.thickTop = True if flags & 0x04 else None
    fmt.thickBottom = True if flags & 0x08 else None
    fmt.outlineLevelRow = row_level or None
    fmt.outlineLevelCol = column_level or None


def _protection(sheet, payload: bytes) -> None:
    values = struct.unpack_from("<H16I", payload, 0)
    password, flags = values[0], values[1:]
    protection = sheet.protection
    # Except fLocked (the sheet), the record stores what users may do: the
    # inverse of the OOXML attributes, which say what is protected.
    for index, (name, value) in enumerate(zip(_PROTECTION_FIELDS, flags)):
        setattr(protection, name, bool(value) if index == 0 else not value)
    protection.set_password(f"{password:04X}" if password else None, already_hashed=True)


def _read_xlsb_page_setup(sheet, payload: bytes) -> None:
    paper, scale, _hres, _vres, copies, first_page, width, height, flags = (
        struct.unpack_from("<8IH", payload, 0)
    )
    setup = sheet.page_setup
    if not flags & 0x40:
        setup.orientation = "landscape" if flags & 0x02 else "portrait"
    setup.paperSize = paper if 1 <= paper <= 255 else None
    setup.scale = scale if 10 <= scale <= 400 and scale != 100 else None
    setup.fitToWidth = width if width != 1 else None
    setup.fitToHeight = height if height != 1 else None
    if flags & 0x80:
        setup.useFirstPageNumber = True
        setup.firstPageNumber = first_page
    if copies > 1:
        setup.copies = copies


def _read_header_footer(sheet, payload: bytes) -> None:
    flags = struct.unpack_from("<H", payload, 0)[0]
    footer = sheet.HeaderFooter
    footer.differentOddEven = True if flags & 0x01 else None
    footer.differentFirst = True if flags & 0x02 else None
    offset = 2
    for name in _HEADER_FOOTER_PARTS:
        text, offset = read_wide(payload, offset, nullable=True)
        if text:
            getattr(footer, name).load_code(text)


# BrtBeginFilterColumn, BrtBeginFilters and BrtFilter of the sheet's auto filter.
FILTER_RECORDS = frozenset({0xA3, 0xA5, 0xA7})


def read_filter_record(auto_filter, record_id: int, payload: bytes) -> None:
    """Filter columns with value lists ([MS-XLSB] 2.4.380 ff.); other filters stay in the file."""
    from ._model import FilterColumn

    if auto_filter.ref is None:
        return
    try:
        if record_id == 0xA3:
            column, flags = struct.unpack_from("<IH", payload)
            auto_filter.filterColumn.append(FilterColumn(
                colId=column, hiddenButton=bool(flags & 0x01), showButton=not flags & 0x02,
            ))
        elif auto_filter.filterColumn and record_id == 0xA5:
            auto_filter.filterColumn[-1].blank = bool(struct.unpack_from("<I", payload)[0])
        elif auto_filter.filterColumn and record_id == 0xA7:
            auto_filter.filterColumn[-1].vals.append(read_wide(payload, 0)[0])
    except (struct.error, ValueError):
        pass


LIST_PART = 0x295
_BEGIN_LIST, _BEGIN_LIST_COL, _TABLE_STYLE_CLIENT, _BEGIN_AUTO_FILTER = 0x157, 0x15B, 0x201, 0xA1
_TOTALS_FUNCTIONS = {
    1: "average", 2: "count", 3: "countNums", 4: "max", 5: "min",
    6: "sum", 7: "stdDev", 8: "var", 9: "custom",
}


def _range(r1: int, r2: int, c1: int, c2: int) -> str:
    return f"{column_label(c1)}{r1 + 1}:{column_label(c2)}{r2 + 1}"


def read_table_part(data: bytes):
    """A ``Table`` from an ``xl/tables/tableN.bin`` part ([MS-XLSB] 2.4.100 BrtBeginList)."""
    from ..biff_utils import iter_record_payloads
    from ._model import AutoFilter, Table, TableColumn, TableStyleInfo

    table = None
    for record_id, payload in iter_record_payloads(data):
        if record_id == _BEGIN_LIST and table is None:
            r1, r2, c1, c2, _kind, ident, header, totals, _flags = struct.unpack_from("<9I", payload, 0)
            # Six dxf ids and the connection id precede the names.
            name, offset = read_wide(payload, 36 + 24 + 4, nullable=True)
            _display, offset = read_wide(payload, offset, nullable=True)
            comment, offset = read_wide(payload, offset, nullable=True)
            table = Table(
                id=ident, displayName=name, ref=_range(r1, r2, c1, c2), name=name,
                comment=comment or None, headerRowCount=header,
                totalsRowCount=totals or None,
            )
        elif table is None:
            continue
        elif record_id == _BEGIN_AUTO_FILTER and table.autoFilter is None:
            table.autoFilter = AutoFilter(_range(*struct.unpack_from("<4I", payload, 0)))
        elif record_id == _BEGIN_LIST_COL:
            ident, totals_function = struct.unpack_from("<II", payload, 0)
            name, offset = read_wide(payload, 24, nullable=True)
            caption, offset = read_wide(payload, offset, nullable=True)
            label, offset = read_wide(payload, offset, nullable=True)
            table.tableColumns.append(TableColumn(
                id=ident, name=caption or name,
                totalsRowFunction=_TOTALS_FUNCTIONS.get(totals_function),
                totalsRowLabel=label or None,
            ))
        elif record_id == _TABLE_STYLE_CLIENT:
            flags = struct.unpack_from("<H", payload, 0)[0]
            style, _ = read_wide(payload, 2, nullable=True)
            table.tableStyleInfo = TableStyleInfo(
                name=style,
                showFirstColumn=bool(flags & 0x01), showLastColumn=bool(flags & 0x02),
                showRowStripes=bool(flags & 0x04), showColumnStripes=bool(flags & 0x08),
            )
    return table


def apply_sheet_records(sheet, items) -> None:
    """Read the collected ``(record_id, payload)`` pairs into the sheet."""
    breaks: list[Break] | None = None
    for record_id, payload in items:
        try:
            if record_id == COL_INFO:
                _column_info(sheet, payload)
            elif record_id == WS_PROP:
                _read_sheet_properties(sheet, payload)
            elif record_id == BEGIN_WS_VIEW:
                _sheet_view(sheet, payload)
            elif record_id == PANE:
                _pane(sheet, payload)
            elif record_id == SELECTION:
                _selection(sheet, payload)
            elif record_id == FORMAT_INFO:
                _sheet_format(sheet, payload)
            elif record_id == SHEET_PROTECTION:
                _protection(sheet, payload)
            elif record_id == BRT_PRINT_OPTIONS:
                flags = struct.unpack_from("<H", payload, 0)[0]
                options = sheet.print_options
                options.horizontalCentered = True if flags & 0x01 else None
                options.verticalCentered = True if flags & 0x02 else None
                options.headings = True if flags & 0x04 else None
                options.gridLines = True if flags & 0x08 else None
            elif record_id == BRT_MARGINS:
                margins = sheet.page_margins
                (margins.left, margins.right, margins.top, margins.bottom,
                 margins.header, margins.footer) = struct.unpack_from("<6d", payload, 0)
            elif record_id == BRT_PAGE_SETUP:
                _read_xlsb_page_setup(sheet, payload)
            elif record_id == BRT_BEGIN_HEADER_FOOTER:
                _read_header_footer(sheet, payload)
            elif record_id in (BRT_BEGIN_ROW_BREAKS, BRT_BEGIN_COL_BREAKS):
                breaks = []
            elif record_id == BRT_BREAK and breaks is not None:
                value, low, high, manual, _pivot = struct.unpack_from("<5I", payload, 0)
                breaks.append(Break(id=value, min=low, max=high, man=bool(manual)))
            elif record_id == BRT_END_ROW_BREAKS and breaks is not None:
                sheet.row_breaks = RowBreak(breaks)
                breaks = None
            elif record_id == BRT_END_COL_BREAKS and breaks is not None:
                sheet.col_breaks = ColBreak(breaks)
                breaks = None
        except (struct.error, IndexError, TypeError, ValueError):
            # A record the model cannot express stays in the file untouched.
            continue
