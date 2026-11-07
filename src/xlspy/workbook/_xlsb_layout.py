"""Worksheet layout of the editable model written back into an XLSB part.

The counterpart of ``_xlsb_sheet_meta``: sheet properties, views, the default
format, column and row dimensions, protection, the auto filter and the print
setup. Each group is compared with what the file held when it was loaded and
only a changed group is re-encoded, so unrelated records keep their bytes.
Where a record holds fields the model does not expose, the original payload
is patched rather than rebuilt.
"""

from __future__ import annotations
from .._package_handle import PackageHandle

from ..exceptions import UnsupportedFeatureError
import re
import struct

from .._metadata import password_hash, DefinedName
from .._xlsb_metadata import (
    pane_record, read_wide, records, replace_records, selection_record, sheet_protection_record, wide_string,
)
from .._xlsb_print import (
    BRT_BEGIN_COL_BREAKS, BRT_BEGIN_HEADER_FOOTER, BRT_BEGIN_ROW_BREAKS, BRT_BREAK,
    BRT_END_COL_BREAKS, BRT_END_HEADER_FOOTER, BRT_END_ROW_BREAKS, BRT_MARGINS,
    BRT_PAGE_SETUP, BRT_PRINT_OPTIONS, breaks_block, header_footer_block, margins_record,
)
from .._xlsb_ignored_errors import (
    RECORD_BEGIN as BEGIN_IGNORED_ERRORS,
    RECORD_END as END_IGNORED_ERRORS,
    RECORD_ERROR,
    ignored_errors_bin,
)
from ..ignored_errors import snapshot as ignored_errors_snapshot
from .._xlsb_style_records import color_payload
from ..biff_utils import build_record
from .._xlsb_sort import (
    BEGIN_SORT_COND, BEGIN_SORT_COND14, BEGIN_SORT_STATE,
    END_SORT_COND, END_SORT_COND14, END_SORT_STATE, sort_state_records,
)
from ..printing import _paper_code, print_area_name
from ..refs import parse_cell_range, column_label, quote_sheet_name
from ..styles import _color_key
from ._loaded import _auto_filter_snapshot, _break_snapshot, _header_footer_snapshot, _snapshot
from ._model import Selection
from .._xlsb_formulas import decode_formula, _Name, parsed_formula, FormulaContext

WS_PROP, WS_DIM = 0x93, 0x94
BEGIN_VIEWS, END_VIEWS, BEGIN_VIEW, END_VIEW, PANE, SELECTION = 0x85, 0x86, 0x89, 0x8A, 0x97, 0x98
FORMAT_INFO = 0x1E5
BEGIN_COL_INFOS, COL_INFO, END_COL_INFOS = 0x186, 0x3C, 0x187
BEGIN_SHEET_DATA, END_SHEET_DATA, ROW_HEADER = 0x91, 0x92, 0x00
SHEET_PROTECTION, SHEET_PROTECTION_ISO = 0x217, 0x2A6
BEGIN_FILTER, END_FILTER = 0xA1, 0xA2
# The sort state lives inside BrtBeginAFilter..BrtEndAFilter; its record ids
# come from _xlsb_sort, which writes the block for the writer too.
# Records of a value-list or blank-only auto filter; any other record means custom criteria.
_VALUE_FILTER_RECORDS = frozenset({
    0xA1, 0xA2, 0xA3, 0xA4, 0xA5, 0xA6, 0xA7,
    BEGIN_SORT_STATE, END_SORT_STATE, BEGIN_SORT_COND, END_SORT_COND,
    BEGIN_SORT_COND14, END_SORT_COND14,
})
BEGIN_FILTER_COLUMN, END_FILTER_COLUMN, BEGIN_FILTERS, END_FILTERS, FILTER = 0xA3, 0xA4, 0xA5, 0xA6, 0xA7
END_SHEET = 0x82

_PRINT_RECORDS = frozenset({
    BRT_PRINT_OPTIONS, BRT_MARGINS, BRT_PAGE_SETUP, BRT_BEGIN_HEADER_FOOTER,
    BRT_END_HEADER_FOOTER, BRT_BEGIN_ROW_BREAKS, BRT_END_ROW_BREAKS,
    BRT_BEGIN_COL_BREAKS, BRT_END_COL_BREAKS, BRT_BREAK,
})
# Records that follow a group in worksheet order ([MS-XLSB] 2.1.7.62): a group
# the file does not have yet is inserted before the first of them.
_AFTER_PRINT = (0x1A7, 0x1A8, 0x1AD, 0x1B1, 0x1B3, 0x226, 0x227, 0x228, 0x229,
                0x27F, 0x280, 0x294, 0x25, END_SHEET)
_BRT_AFTER_HYPERLINKS_AND_PRINT = (BRT_PRINT_OPTIONS, BRT_MARGINS, BRT_PAGE_SETUP, BRT_BEGIN_HEADER_FOOTER) + _AFTER_PRINT
_AFTER_FILTER = (BEGIN_SORT_STATE, 0x1F0, 0xB1, 0x1CD, 0x23D, 0x1EE) + _BRT_AFTER_HYPERLINKS_AND_PRINT
_AFTER_PROTECTION = (0x1A5, 0x1A6, BEGIN_FILTER) + _AFTER_FILTER
_AFTER_VIEWS = (FORMAT_INFO, BEGIN_COL_INFOS, BEGIN_SHEET_DATA)

_PANE_IDS = {"bottomRight": 0, "topRight": 1, "bottomLeft": 2, "topLeft": 3}
_VIEW_KINDS = {None: 0, "normal": 0, "pageBreakPreview": 1, "pageLayout": 2}
# BrtBeginWsView flag bits and the value a field has when it is None.
_VIEW_FLAGS = (
    ("windowProtection", 0x01, False), ("showFormulas", 0x02, False),
    ("showGridLines", 0x04, True), ("showRowColHeaders", 0x08, True),
    ("showZeros", 0x10, True), ("rightToLeft", 0x20, False),
    ("tabSelected", 0x40, False), ("showRuler", 0x80, None),
    ("showOutlineSymbols", 0x100, None), ("defaultGridColor", 0x200, None),
    ("showWhiteSpace", 0x400, None),
)
# BrtWsProp bits (24-bit flags): (model path, bit); None keeps the file's bit.
_PROPERTY_BITS = (
    ("published", 3), ("outlinePr.applyStyles", 5), ("outlinePr.summaryBelow", 6),
    ("outlinePr.summaryRight", 7), ("outlinePr.showOutlineSymbols", 10),
    ("syncHorizontal", 12), ("syncVertical", 13), ("transitionEvaluation", 14),
    ("transitionEntry", 15), ("filterMode", 16), ("enableFormatConditionsCalculation", 17),
)
_FIT_TO_PAGE = 8
# BrtSheetProtection: fLocked, then the permissions, stored as "allowed" (the
# inverse of the OOXML attributes).
_PROTECTION_FLAGS_NO_SHEET = (
    "objects", "scenarios", "formatCells", "formatColumns", "formatRows",
    "insertColumns", "insertRows", "insertHyperlinks", "deleteColumns", "deleteRows",
    "selectLockedCells", "sort", "autoFilter", "pivotTables", "selectUnlockedCells",
)
_ERRORS = {None: None, "displayed": 0, "blank": 1, "dash": 2, "NA": 3}
_DEFAULT_PAGE_SETUP = struct.pack("<8IH", 1, 100, 600, 600, 1, 1, 1, 1, 0x40) + b"\xff\xff\xff\xff"
_DEFAULT_VIEW = bytes.fromhex("dc0300000000000000000000000040000000640000000000000000000000")


def _parse_sync_ref(ref: str) -> tuple[int, int]:
    column, row, _, _ = parse_cell_range(ref)
    return row, column


def _get(obj, path: str):
    for name in path.split("."):
        obj = getattr(obj, name)
    return obj


# --- what the model holds, per group -----------------------------------------

def _columns_state(sheet):
    return tuple(
        (letter, dim.min, dim.max, dim.width, dim.hidden, dim.outlineLevel, dim.collapsed,
         dim.bestFit, dim.style)
        for letter, dim in sorted(
            sheet.column_dimensions.items(), key=lambda item: _parse_sync_ref(f"{item[0]}1")[1],
        )
        if dim.width is not None or dim.hidden or dim.outlineLevel or dim.collapsed
        or dim.bestFit or dim.style
    )


def _rows_state(sheet):
    return tuple(
        (row, dim.height, dim.hidden, dim.outlineLevel, dim.collapsed, dim.style)
        for row, dim in sorted(sheet.row_dimensions.items())
        if dim.height is not None or dim.hidden or dim.outlineLevel or dim.collapsed or dim.style
    )


def _view_state(sheet):
    view = sheet.sheet_view
    return (
        _snapshot(view)[:-2],
        None if view.pane is None else _snapshot(view.pane),
        tuple(_snapshot(item) for item in view.selection),
        sheet.freeze_panes,
    )


def _print_state(sheet):
    return (
        _snapshot(sheet.page_setup), _snapshot(sheet.print_options), _snapshot(sheet.page_margins),
        _header_footer_snapshot(sheet.HeaderFooter),
        _break_snapshot(sheet.row_breaks), _break_snapshot(sheet.col_breaks),
        sheet.sheet_properties.pageSetUpPr.fitToPage,
    )


def _properties_state(sheet):
    properties = sheet.sheet_properties
    return (
        repr(properties.tabColor), properties.codeName, properties.syncRef,
        tuple(_get(properties, path) for path, _ in _PROPERTY_BITS),
        properties.pageSetUpPr.fitToPage,
    )


GROUPS = {
    "properties": _properties_state,
    "views": _view_state,
    "format": lambda sheet: _snapshot(sheet.sheet_format),
    "columns": _columns_state,
    "rows": _rows_state,
    "protection": lambda sheet: _snapshot(sheet.protection),
    "filter": lambda sheet: _auto_filter_snapshot(sheet.auto_filter),
    "print": _print_state,
    "print_names": lambda sheet: (sheet.print_area, sheet.print_title_rows, sheet.print_title_cols),
    "ignored_errors": lambda sheet: ignored_errors_snapshot(sheet.ignored_errors),
}


def layout_state(sheet) -> dict:
    return {name: state(sheet) for name, state in GROUPS.items()}


def changed_groups(sheet) -> set[str]:
    saved = getattr(sheet, "_saved_layout", None)
    if saved is None:
        return set(GROUPS)
    current = layout_state(sheet)
    return {name for name in GROUPS if current[name] != saved.get(name)}


# --- encoders ----------------------------------------------------------------

def _find(data: bytes, record_id: int) -> bytes | None:
    for record, payload in records(data):
        if record.record_id == record_id:
            return payload
    return None


def _sheet_properties_payload(sheet, original: bytes | None) -> bytes:
    payload = bytearray(original if original and len(original) >= 23
                        else bytes.fromhex("c904020040000000000000ffffffffffffffff00000000"))
    flags = int.from_bytes(payload[0:3], "little")
    properties = sheet.sheet_properties
    for path, bit in _PROPERTY_BITS:
        value = _get(properties, path)
        if value is not None:
            flags = flags | (1 << bit) if value else flags & ~(1 << bit)
    if properties.pageSetUpPr.fitToPage:
        flags |= 1 << _FIT_TO_PAGE
    else:
        flags &= ~(1 << _FIT_TO_PAGE)
    payload[0:3] = flags.to_bytes(3, "little")
    tab = properties.tabColor
    payload[3:11] = color_payload(_color_key(tab)) if tab is not None else bytes((0, 0x40, 0, 0, 0, 0, 0, 0))
    if properties.syncRef:
        row, column = _parse_sync_ref(properties.syncRef)
        struct.pack_into("<II", payload, 11, row, column)
    head = bytes(payload[:19])
    code_name = properties.codeName or ""
    return build_record(WS_PROP, head + wide_string(code_name))


def _view_record(sheet, original: bytes | None) -> bytes:
    view = sheet.sheet_view
    payload = bytearray(original if original and len(original) >= 30 else _DEFAULT_VIEW)
    flags = struct.unpack_from("<H", payload, 0)[0]
    for name, bit, default in _VIEW_FLAGS:
        value = getattr(view, name)
        if value is None:
            if default is None:
                continue
            value = default
        flags = flags | bit if value else flags & ~bit
    struct.pack_into("<H", payload, 0, flags)
    if view.view not in _VIEW_KINDS:
        raise ValueError(f"invalid sheet view {view.view!r}")
    top, left = _parse_sync_ref(view.topLeftCell) if view.topLeftCell else (0, 0)
    struct.pack_into("<III", payload, 2, _VIEW_KINDS[view.view], top, left)
    if view.colorId is not None:
        payload[14] = view.colorId
    struct.pack_into(
        "<HHHHI", payload, 18,
        view.zoomScale or 100, view.zoomScaleNormal or 0,
        view.zoomScaleSheetLayoutView or 0, view.zoomScalePageLayoutView or 0,
        view.workbookViewId or 0,
    )
    return build_record(BEGIN_VIEW, bytes(payload))


def _pane_record(pane) -> bytes:
    if pane is None or not (pane.xSplit or pane.ySplit):
        return b""
    top, left = _parse_sync_ref(pane.topLeftCell) if pane.topLeftCell else (0, 0)
    state = {"frozen": 0x03, "frozenSplit": 0x01}.get(pane.state, 0x00)
    return pane_record(float(pane.xSplit or 0), float(pane.ySplit or 0), top, left,
                       _PANE_IDS.get(pane.activePane, 3), state)


def _selection_records(view) -> bytes:
    selections = list(view.selection)
    if not selections:
        pane = view.pane.activePane if view.pane is not None and (view.pane.xSplit or view.pane.ySplit) else "topLeft"

        selections = [Selection(pane=pane)]
    output = []
    for item in selections:
        row, column = _parse_sync_ref(item.activeCell or "A1")
        ranges = []
        for part in (item.sqref or item.activeCell or "A1").split():
            c1, r1, c2, r2 = parse_cell_range(part)
            ranges.append((r1, r2, c1, c2))
        output.append(selection_record(_PANE_IDS.get(item.pane or "topLeft", 3), row, column,
                                       item.activeCellId or 0, ranges))
    return b"".join(output)


def _views_block(sheet, data: bytes) -> bytes:
    views = [payload for record, payload in records(data) if record.record_id == BEGIN_VIEW]
    if len(views) > 1:
        raise UnsupportedFeatureError("editing a sheet with several window views is not supported")
    return b"".join((
        build_record(BEGIN_VIEWS),
        _view_record(sheet, views[0] if views else None),
        _pane_record(sheet.sheet_view.pane),
        _selection_records(sheet.sheet_view),
        build_record(END_VIEW),
        build_record(END_VIEWS),
    ))


def _format_record(sheet) -> bytes:
    fmt = sheet.sheet_format
    width = 0xFFFFFFFF if fmt.defaultColWidth is None else round(fmt.defaultColWidth * 256)
    flags = (
        (0x01 if fmt.customHeight else 0) | (0x02 if fmt.zeroHeight else 0)
        | (0x04 if fmt.thickTop else 0) | (0x08 if fmt.thickBottom else 0)
    )
    return build_record(FORMAT_INFO, struct.pack(
        "<IHHHBB", width, fmt.baseColWidth if fmt.baseColWidth is not None else 8,
        round((fmt.defaultRowHeight or 15) * 20), flags,
        fmt.outlineLevelRow or 0, fmt.outlineLevelCol or 0,
    ))


def _column_records(sheet) -> bytes:
    fmt = sheet.sheet_format
    default = round((fmt.defaultColWidth or 8.43) * 256)
    entries: list[list] = []
    for letter, _, _, width, hidden, level, collapsed, best_fit, style in _columns_state(sheet):
        index = _parse_sync_ref(f"{letter}1")[1]
        flags = (
            (0x01 if hidden else 0) | (0x02 if width is not None else 0)
            | (0x04 if best_fit else 0) | ((level or 0) & 0x07) << 8
            | (0x1000 if collapsed else 0)
        )
        key = (default if width is None else round(width * 256), style or 0, flags)
        if entries and entries[-1][1] == index - 1 and entries[-1][2] == key:
            entries[-1][1] = index
        else:
            entries.append([index, index, key])
    if not entries:
        return b""
    return build_record(BEGIN_COL_INFOS) + b"".join(
        build_record(COL_INFO, struct.pack("<IIIIH", first, last, *key))
        for first, last, key in entries
    ) + build_record(END_COL_INFOS)


def _protection_record(sheet) -> bytes:
    protection = sheet.protection
    if not protection.sheet:
        return b""
    flags = [1] + [0 if getattr(protection, name) else 1 for name in _PROTECTION_FLAGS_NO_SHEET]
    iso = getattr(sheet, "_protection_iso", None)
    if iso is not None:
        if protection.password is not None:
            raise UnsupportedFeatureError(
                "changing the password of a hashed sheet protection cannot be saved into XLSB"
            )
        payload = bytearray(iso)
        if len(payload) < 4 + 16 * 4:
            raise ValueError("malformed XLSB ISO sheet protection")
        original_algorithm = getattr(sheet, "_protection_iso_algorithm", None)
        if protection.algorithmName != original_algorithm:
            raise UnsupportedFeatureError(
                "changing the hash algorithm of a hashed sheet protection cannot be saved into XLSB"
            )
        stored_spin = struct.unpack_from("<I", payload, 0)[0]
        if int(protection.spinCount or 0) != stored_spin:
            raise UnsupportedFeatureError(
                "changing the spin count of a hashed sheet protection cannot be saved into XLSB"
            )
        struct.pack_into("<16I", payload, 4, *flags)
        return build_record(SHEET_PROTECTION_ISO, bytes(payload)) + sheet_protection_record(0, flags)
    if protection.hashValue or protection.algorithmName:
        raise UnsupportedFeatureError("sheet protection with a hashed password cannot be edited in XLSB yet")
    password = protection.password
    if not password:
        value = 0
    elif re.fullmatch(r"[0-9A-Fa-f]{4}", password):
        value = int(password, 16)
    else:
        value = int(password_hash(password), 16)
    return sheet_protection_record(value, flags)


def _filter_block(sheet) -> bytes:
    auto_filter = sheet.auto_filter
    if not auto_filter.ref:
        return b""
    c1, r1, c2, r2 = parse_cell_range(auto_filter.ref)
    output = [build_record(BEGIN_FILTER, struct.pack("<IIII", r1, r2, c1, c2))]
    for column in auto_filter.filterColumn:
        flags = (0x01 if column.hiddenButton else 0) | (0 if column.showButton else 0x02)
        output.append(build_record(BEGIN_FILTER_COLUMN, struct.pack("<IH", column.colId, flags)))
        output.append(build_record(BEGIN_FILTERS, struct.pack("<II", 1 if column.blank else 0, 0)))
        output.extend(build_record(FILTER, wide_string(str(value))) for value in column.vals)
        output.append(build_record(END_FILTERS))
        output.append(build_record(END_FILTER_COLUMN))
    sort = auto_filter.sortState
    if sort is not None:
        output.append(sort_state_records(auto_filter.ref, sort))
    output.append(build_record(END_FILTER))
    return b"".join(output)


def _page_setup_record(sheet, original: bytes | None) -> bytes:
    setup = sheet.page_setup
    payload = bytearray(original if original and len(original) >= 34 else _DEFAULT_PAGE_SETUP)
    paper, scale, h_dpi, v_dpi, copies, first_page, fit_width, fit_height, flags = struct.unpack_from("<8IH", payload)
    if setup.paperSize is not None:
        paper = _paper_code(setup.paperSize)
    scale = 100 if setup.scale is None else setup.scale
    h_dpi = h_dpi if setup.horizontalDpi is None else setup.horizontalDpi
    v_dpi = v_dpi if setup.verticalDpi is None else setup.verticalDpi
    copies = 1 if setup.copies is None else setup.copies
    fit_width = 1 if setup.fitToWidth is None else setup.fitToWidth
    fit_height = 1 if setup.fitToHeight is None else setup.fitToHeight
    if setup.firstPageNumber is not None:
        first_page = setup.firstPageNumber

    def bit(mask, value):
        nonlocal flags
        if value is not None:
            flags = flags | mask if value else flags & ~mask

    if setup.orientation not in (None, "default", "portrait", "landscape"):
        raise ValueError(f"invalid orientation {setup.orientation!r}")
    bit(0x02, setup.orientation == "landscape")
    bit(0x40, setup.orientation in (None, "default"))
    bit(0x80, setup.useFirstPageNumber)
    bit(0x01, None if setup.pageOrder is None else setup.pageOrder == "overThenDown")
    bit(0x08, setup.blackAndWhite)
    bit(0x10, setup.draft)
    if setup.cellComments is not None:
        bit(0x20, setup.cellComments != "none")
        bit(0x200, setup.cellComments == "atEnd")
    errors = _ERRORS.get(setup.errors)
    if errors is not None:
        flags = (flags & ~0xC00) | errors << 10
    struct.pack_into("<8IH", payload, 0, paper, scale, h_dpi, v_dpi, copies, first_page,
                     fit_width, fit_height, flags)
    return build_record(BRT_PAGE_SETUP, bytes(payload))


def _header_footer_block(sheet) -> bytes:
    footer = sheet.HeaderFooter
    codes = [
        getattr(footer, name).code()
        for name in ("oddHeader", "oddFooter", "evenHeader", "evenFooter", "firstHeader", "firstFooter")
    ]
    flags = (
        (0x01 if footer.differentOddEven else 0) | (0x02 if footer.differentFirst else 0)
        | (0 if footer.scaleWithDoc is False else 0x04) | (0 if footer.alignWithMargins is False else 0x08)
    )
    if not any(codes) and flags == 0x0C:
        return b""
    return header_footer_block(flags, codes)


def _breaks_block(begin, end, container) -> bytes:
    return breaks_block(begin, end, [
        (item.id, item.min or 0, item.max or 0, bool(item.man), item.pt or 0) for item in container.brk
    ])


def _print_block(sheet, data: bytes) -> bytes:
    options = sheet.print_options
    original = _find(data, BRT_PRINT_OPTIONS)
    flags = struct.unpack_from("<H", original)[0] if original and len(original) >= 2 else 0x10
    for mask, value in (
        (0x01, options.horizontalCentered), (0x02, options.verticalCentered),
        (0x04, options.headings), (0x08, options.gridLines), (0x10, options.gridLinesSet),
    ):
        if value is not None or mask != 0x10:
            flags = flags | mask if value else flags & ~mask
    if options.gridLinesSet is None:
        flags |= 0x10
    margins = sheet.page_margins
    return b"".join((
        build_record(BRT_PRINT_OPTIONS, struct.pack("<H", flags)),
        margins_record((margins.left, margins.right, margins.top, margins.bottom, margins.header, margins.footer)),
        _page_setup_record(sheet, _find(data, BRT_PAGE_SETUP)),
        _header_footer_block(sheet),
        _breaks_block(BRT_BEGIN_ROW_BREAKS, BRT_END_ROW_BREAKS, sheet.row_breaks),
        _breaks_block(BRT_BEGIN_COL_BREAKS, BRT_END_COL_BREAKS, sheet.col_breaks),
    ))


def _replace_filter(data: bytes, block: bytes) -> bytes:
    """Replace the BrtBeginAFilter ... BrtEndAFilter block of the sheet itself."""
    if not block and not any(record.record_id == BEGIN_FILTER for record, _ in records(data)):
        return data
    output = []
    inserted = False
    depth = 0
    for record, payload in records(data):
        kind = record.record_id
        if depth:
            if kind == BEGIN_FILTER:
                depth += 1
            elif kind == END_FILTER:
                depth -= 1
            continue
        if kind == BEGIN_FILTER:
            depth = 1
            if not inserted:
                output.append(block)
                inserted = True
            continue
        if not inserted and kind in _AFTER_FILTER:
            output.append(block)
            inserted = True
        output.append(data[record.header_start:record.data_end])
    if not inserted:
        raise ValueError("BIFF12 insertion point was not found")
    return b"".join(output)


# --- rows ----------------------------------------------------------------------

def _row_payload(row: int, original: bytes | None, dim, default_height: int) -> bytes:
    if original is None:
        payload = bytearray(struct.pack("<IIH", row - 1, 0, default_height) + bytes(3)
                            + struct.pack("<III", 1, 0, 0))
    else:
        payload = bytearray(original)
    flags = payload[11] & 0x80  # keep fPhShow's neighbour bits outside the model
    if dim is not None:
        height, hidden, level, collapsed, style = dim
        flags |= (level or 0) & 0x07
        if collapsed:
            flags |= 0x08
        if hidden:
            flags |= 0x10
        if height is not None:
            flags |= 0x20
            struct.pack_into("<H", payload, 8, round(height * 20))
        elif original is not None and payload[11] & 0x20:
            struct.pack_into("<H", payload, 8, default_height)
        if style:
            flags |= 0x40
            struct.pack_into("<I", payload, 4, style)
        else:
            struct.pack_into("<I", payload, 4, 0)
    else:
        if payload[11] & 0x20:
            struct.pack_into("<H", payload, 8, default_height)
        struct.pack_into("<I", payload, 4, 0)
    payload[11] = flags
    return build_record(ROW_HEADER, bytes(payload))


def _patch_xlsb_rows(data: bytes, sheet, saved_rows) -> bytes:
    wanted = {row: values for row, *values in _rows_state(sheet)}
    touched = set(wanted) | {row for row, *_ in (saved_rows or ())}
    if not touched:
        return data
    default_height = round((sheet.sheet_format.defaultRowHeight or 15) * 20)
    output = []
    pending = sorted(row for row in wanted)
    next_pending = 0
    cursor = 0
    in_data = False
    for record, payload in records(data):
        kind = record.record_id
        if kind == BEGIN_SHEET_DATA:
            in_data = True
            continue
        if not in_data:
            continue
        if kind == END_SHEET_DATA:
            output.append(data[cursor:record.header_start])
            while next_pending < len(pending):
                row = pending[next_pending]
                next_pending += 1
                output.append(_row_payload(row, None, wanted[row], default_height))
            cursor = record.header_start
            break
        if kind != ROW_HEADER:
            continue
        row = struct.unpack_from("<I", payload)[0] + 1
        while next_pending < len(pending) and pending[next_pending] < row:
            new_row = pending[next_pending]
            next_pending += 1
            output.append(data[cursor:record.header_start])
            output.append(_row_payload(new_row, None, wanted[new_row], default_height))
            cursor = record.header_start
        if next_pending < len(pending) and pending[next_pending] == row:
            next_pending += 1
        if row in touched:
            output.append(data[cursor:record.header_start])
            output.append(_row_payload(row, payload, wanted.get(row), default_height))
            cursor = record.data_end
    output.append(data[cursor:])
    return b"".join(output)


# --- the sheet part --------------------------------------------------------------

def _drop_legacy_column_wrappers(data: bytes) -> bytes:
    """xlspy before 0.11 wrapped BrtColInfo in records 0x86/0x87 instead of
    0x186/0x187; 0x86 is also BrtEndWsViews. Drop that pair before the sheet
    data, keeping the views' own end record."""
    output = []
    seen_views_end = False
    changed = False
    for record, _payload in records(data):
        kind = record.record_id
        if kind == BEGIN_SHEET_DATA:
            output.append(data[record.header_start:])
            break
        if kind == END_VIEWS:
            if seen_views_end:
                changed = True
                continue
            seen_views_end = True
        elif kind == 0x87:
            changed = True
            continue
        output.append(data[record.header_start:record.data_end])
    return b"".join(output) if changed else data


def patch_layout(data: bytes, sheet, groups) -> bytes:
    """Re-encode the ``groups`` (names from ``GROUPS``) of ``sheet`` into ``data``."""
    saved = getattr(sheet, "_saved_layout", None) or {}
    if {"views", "columns", "format"} & set(groups):
        data = _drop_legacy_column_wrappers(data)
    if "properties" in groups or "print" in groups:
        data = replace_records(data, {WS_PROP}, _sheet_properties_payload(sheet, _find(data, WS_PROP)),
                               before=(WS_DIM, BEGIN_VIEWS) + _AFTER_VIEWS)
    if "views" in groups:
        data = replace_records(data, {BEGIN_VIEWS, END_VIEWS, BEGIN_VIEW, END_VIEW, PANE, SELECTION},
                               _views_block(sheet, data), before=_AFTER_VIEWS)
    if "format" in groups:
        data = replace_records(data, {FORMAT_INFO}, _format_record(sheet),
                               before=(BEGIN_COL_INFOS, BEGIN_SHEET_DATA))
    if "columns" in groups or "format" in groups:
        block = _column_records(sheet)
        if block or any(record.record_id == BEGIN_COL_INFOS for record, _ in records(data)):
            data = replace_records(data, {BEGIN_COL_INFOS, COL_INFO, END_COL_INFOS}, block,
                                   before=(BEGIN_SHEET_DATA,))
    if "rows" in groups or "format" in groups:
        data = _patch_xlsb_rows(data, sheet, saved.get("rows"))
    if "protection" in groups:
        block = _protection_record(sheet)
        if block or any(record.record_id in (SHEET_PROTECTION, SHEET_PROTECTION_ISO) for record, _ in records(data)):
            data = replace_records(data, {SHEET_PROTECTION, SHEET_PROTECTION_ISO}, block,
                                   before=_AFTER_PROTECTION)
    if "filter" in groups:
        loaded = getattr(sheet, "_saved_filter_records", None) or ()
        if sheet.auto_filter.ref and any(record_id not in _VALUE_FILTER_RECORDS for record_id, _ in loaded):
            raise UnsupportedFeatureError(
                "custom auto-filter criteria cannot be changed in a loaded XLSB workbook yet; "
                "keep the filter as it is or remove it"
            )
        data = _replace_filter(data, _filter_block(sheet))
    if "print" in groups:
        data = replace_records(data, _PRINT_RECORDS, _print_block(sheet, data), before=_AFTER_PRINT)
    if "ignored_errors" in groups:
        data = replace_records(
            data,
            {BEGIN_IGNORED_ERRORS, RECORD_ERROR, END_IGNORED_ERRORS},
            ignored_errors_bin(sheet.ignored_errors),
            before=(END_SHEET,),
        )
    return data


# --- built-in names: print area, print titles, filter database -------------------

_BUILTIN = 0x20
_HIDDEN = 0x01


def builtin_names(context) -> dict[str, dict[str, str]]:
    """``{sheet title: {"_xlnm.Print_Area": formula, ...}}`` from a FormulaContext."""

    result: dict[str, dict[str, str]] = {}
    for entry in context.names:
        if not entry.flags & _BUILTIN or entry.raw is None or not isinstance(entry.scope, int):
            continue
        try:
            scope = context.sheets[entry.scope]
            _, offset = read_wide(entry.raw, 9)
            formula = decode_formula(entry.raw[offset:], context, sheet=scope, name_formula=True)
        except (IndexError, NotImplementedError, ValueError, struct.error):
            continue
        result.setdefault(scope, {})["_xlnm." + entry.name] = formula
    return result


def set_builtin_name(context, name: str, formula: str | None, scope: str, *, hidden=False) -> None:
    """Set (or with ``formula=None`` remove) a built-in sheet name such as Print_Area."""

    for index, entry in enumerate(context.names):
        if (entry.flags & _BUILTIN and entry.name.casefold() == name.casefold()
                and context._scope(entry) == scope):
            if formula is None:
                _drop(context, index)
            else:
                entry.scope = scope
                entry.definition = DefinedName(name, formula, scope)
            return
    if formula is not None:
        flags = _BUILTIN | (_HIDDEN if hidden else 0)
        context.names.append(_Name(name, scope, flags, definition=DefinedName(name, formula, scope)))


def _drop(context, index: int) -> None:
    # Formulas refer to names by position: the slot stays, renamed and hidden.

    entry = context.names[index]
    placeholder = f"_xlspy_deleted_{index}_{len(context.names)}"
    entry.name, entry.flags, entry.definition, entry.scope = placeholder, _HIDDEN, None, None
    entry.raw = (struct.pack("<IBI", _HIDDEN, 0, 0xFFFFFFFF) + wide_string(placeholder)
                 + parsed_formula(b"\x1c\x1d") + b"\xff" * 4)


def save_builtin_names(updater, sheet, groups) -> None:
    """Write the print area, print titles and filter range of ``sheet``."""

    if not {"print_names", "filter"} & set(groups):
        return
    handle = PackageHandle(updater)
    original = handle.package.get("xl/workbook.bin")
    context = FormulaContext(handle.sheets, original)
    title = sheet.title
    quoted = quote_sheet_name(title)
    if "print_names" in groups:
        area = sheet.print_area
        set_builtin_name(context, "Print_Area", print_area_name(title, area) if area else None, title)
        titles = []
        for part in (sheet.print_title_rows, sheet.print_title_cols):
            if part:
                first, _, last = part.partition(":")
                titles.append(f"{quoted}!${first}:${last or first}")
        set_builtin_name(context, "Print_Titles", ",".join(titles) or None, title)
    if "filter" in groups:
        ref = sheet.auto_filter.ref
        formula = None
        if ref:
            c1, r1, c2, r2 = parse_cell_range(ref)
            formula = f"{quoted}!${column_label(c1)}${r1 + 1}:${column_label(c2)}${r2 + 1}"
        set_builtin_name(context, "_FilterDatabase", formula, title, hidden=True)
    updated = context.write_workbook(original)
    if updated != original:
        handle.apply_parts({"xl/workbook.bin": updated})
