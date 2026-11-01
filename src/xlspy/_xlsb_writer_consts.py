"""Module-level constants and record helpers of the XLSB writer (BIFF12 record ids)."""

import struct
import uuid
from .hyperlinks import Hyperlink


_REL_PIVOT_TABLE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable"
_REL_PIVOT_CACHE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition"
# BIFF12 record identifiers used in xl/styles.bin ([MS-XLSB] section 2.1.4).
_BRT_FONT = 0x002B
_BRT_FMT = 0x002C
_BRT_FILL = 0x002D
_BRT_BORDER = 0x002E
_BRT_XF = 0x002F
_BRT_BEGIN_FILLS = 0x025B
_BRT_END_FILLS = 0x025C
_BRT_BEGIN_FONTS = 0x0263
_BRT_END_FONTS = 0x0264
_BRT_BEGIN_BORDERS = 0x0265
_BRT_END_BORDERS = 0x0266
_BRT_BEGIN_FMTS = 0x0267
_BRT_END_FMTS = 0x0268
_BRT_BEGIN_CELL_XFS = 0x0269
_BRT_END_CELL_XFS = 0x026A
# BIFF12 record identifiers used in xl/worksheets/sheetN.bin.
_BRT_ROW_HDR = 0x0000
_BRT_COL_INFO = 0x003C
_BRT_BEGIN_COL_INFOS = 0x0186
_BRT_END_COL_INFOS = 0x0187
_BRT_PANE = 0x0097
_BRT_SEL = 0x0098
_BRT_MERGE_CELL = 0x00B0
_BRT_BEGIN_MERGE_CELLS = 0x00B1
_BRT_END_MERGE_CELLS = 0x00B2
# BrtDrawing ([MS-XLSB] 2.4.351) links a sheet to its Drawings part.
_BRT_DRAWING = 0x0226
# BrtHLink: one hyperlink, written directly under the worksheet (no container).
_BRT_HLINK = 0x01EE


def _hlink_payload(row: int, col: int, link: Hyperlink, rid: str | None) -> bytes:
    """A BrtHLink payload: the cell, then the relationship id, location, tooltip
    and display text as wide strings ('' when absent). The layout matches a
    file saved by Excel: an external link names its relationship and displays
    its URL; an internal link has no relationship and repeats its location."""
    from ._xlsb_metadata import wide_string
    if link.is_internal:
        relationship, location, display = "", link.location, link.location
    else:
        relationship, location, display = rid or "", "", link.target or ""
    return (
        struct.pack("<4I", row, row, col, col)
        + wide_string(relationship) + wide_string(location)
        + wide_string(link.tooltip or "") + wide_string(display)
    )


# Table (List) parts ([MS-XLSB] 2.1.7.51, 2.4.100).
_BRT_BEGIN_LIST = 343  # 0x157
_BRT_END_LIST = 344  # 0x158
_BRT_BEGIN_LIST_COLS = 345  # 0x159
_BRT_END_LIST_COLS = 346  # 0x15A
_BRT_BEGIN_LIST_COL = 347  # 0x15B
_BRT_END_LIST_COL = 348  # 0x15C
_BRT_BEGIN_LIST_PARTS = 660  # 0x294
_BRT_LIST_PART = 661  # 0x295
_BRT_END_LIST_PARTS = 662  # 0x296
_BRT_BEGIN_AUTOFILTER = 161  # 0xA1
_BRT_END_AUTOFILTER = 162  # 0xA2
_BRT_TABLE_STYLE_CLIENT = 513  # 0x201
# Conditional formatting ([MS-XLSB] 2.1.7.62).
_BRT_BEGIN_CF = 461
_BRT_END_CF = 462
_BRT_BEGIN_DXFS = 505
_BRT_END_DXFS = 506
_BRT_DXF = 507
# FRT wrappers used by table parts.
_BRT_AC_BEGIN = 0x25
_BRT_AC_END = 0x26
_BRT_FRT_BEGIN = 0xC00
_ACUID_NAMESPACE = uuid.UUID("6f1c2a52-7b1e-4f0b-9a7e-58a3c1d0b6e4")
# Active-pane identifiers for BrtPane/BrtSel ([MS-XLSB] 2.5.108 Pnn).
_PANE_ACTIVE = {"bottomRight": 0, "topRight": 1, "bottomLeft": 2, "topLeft": 3}
# The native row encoder hands rows with rich text to the Python encoder, which
# writes them as rich shared strings (see _write_rich_str). Hyperlinks are
# written as BrtHLink records (see _hlink_payload).
_MAX_STR_RUNS = 32767


def _capture_cells(rows, store: dict):
    """Yield rows unchanged, filling the chart-cache slots they cover."""
    columns = [(col_idx, first_row, slots) for col_idx, (first_row, slots) in store.items()]
    for row_idx, row in enumerate(rows):
        for col_idx, first_row, slots in columns:
            offset = row_idx - first_row
            if 0 <= offset < len(slots) and col_idx < len(row):
                slots[offset] = row[col_idx]
        yield row
