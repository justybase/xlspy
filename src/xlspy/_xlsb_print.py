"""Print settings as XLSB worksheet records ([MS-XLSB] 2.4: BrtPrintOptions ff.).

The field values mirror what Excel writes when it saves the same settings
from an XLSX file as XLSB, so both formats print alike.
"""

from __future__ import annotations
from collections.abc import Iterable, Sequence

import struct

from ._xlsb_metadata import wide_string
from .biff_utils import build_record
from .printing import PageMargins, PrintSettings, _paper_code

BRT_PRINT_OPTIONS = 0x1DD
BRT_MARGINS = 0x1DC
BRT_PAGE_SETUP = 0x1DE
BRT_BEGIN_HEADER_FOOTER = 0x1DF
BRT_END_HEADER_FOOTER = 0x1E0
BRT_BEGIN_ROW_BREAKS = 0x188
BRT_END_ROW_BREAKS = 0x189
BRT_BEGIN_COL_BREAKS = 0x18A
BRT_END_COL_BREAKS = 0x18B
BRT_BREAK = 0x18C

# BrtPrintOptions: 0x01 fHCenter, 0x02 fVCenter; Excel sets 0x10 in every file.
_PRINT_OPTIONS_DEFAULT = 0x10
# BrtPageSetup flags: fLandscape, and fNoOrient when no orientation is set.
_LANDSCAPE, _NO_ORIENTATION = 0x02, 0x40
# BrtBeginHeaderFooter: fHFScaleWithDoc and fHFAlignMargins, Excel's defaults.
_HEADER_FOOTER_FLAGS = 0x0C
_NULL_STRING = b"\xff\xff\xff\xff"
# BrtWsProp: fFitToPage is bit 8, the low bit of the second flags byte.
FIT_TO_PAGE_BYTE, FIT_TO_PAGE_BIT = 1, 0x01


def _nullable_text(text: str | None) -> bytes:
    return _NULL_STRING if text is None else wide_string(text)


def _xlsb_page_setup_record(settings: PrintSettings) -> bytes:
    """BrtPageSetup, or nothing when no page option is set (as in XLSX)."""
    if (settings.paper_size is None and not settings.fit_to_page
            and settings.scale is None and settings.orientation is None):
        return b""
    paper = 1 if settings.paper_size is None else _paper_code(settings.paper_size)
    fit_width = fit_height = 1
    if settings.fit_to_page:
        fit_width = 1 if settings.fit_to_width is None else settings.fit_to_width
        fit_height = 1 if settings.fit_to_height is None else settings.fit_to_height
    flags = {"landscape": _LANDSCAPE, "portrait": 0, None: _NO_ORIENTATION}[settings.orientation]
    payload = struct.pack(
        "<8IH", paper, settings.scale or 100, 600, 600, 1, 1, fit_width, fit_height, flags,
    ) + _NULL_STRING
    return build_record(BRT_PAGE_SETUP, payload)


def header_footer_block(flags: int, texts: Sequence[str | None]) -> bytes:
    """BrtBeginHeaderFooter/BrtEndHeaderFooter: the flags, then the odd, even and
    first-page header and footer texts (None for none), in that order."""
    payload = struct.pack("<H", flags) + b"".join(_nullable_text(text) for text in texts)
    return build_record(BRT_BEGIN_HEADER_FOOTER, payload) + build_record(BRT_END_HEADER_FOOTER)


def breaks_block(begin: int, end: int, breaks: Sequence[tuple[int, int, int, bool, int]]) -> bytes:
    """A row or column break list: ``(index, first, last, manual, pivot)`` per BrtBrk."""
    if not breaks:
        return b""
    manual = sum(1 for item in breaks if item[3])
    return b"".join((
        build_record(begin, struct.pack("<II", len(breaks), manual)),
        *(build_record(BRT_BREAK, struct.pack("<5I", index, first, last, 1 if man else 0, pivot))
          for index, first, last, man, pivot in breaks),
        build_record(end),
    ))


def margins_record(margins: Iterable[float]) -> bytes:
    """BrtMargins: left, right, top, bottom, header and footer, in inches."""
    return build_record(BRT_MARGINS, struct.pack("<6d", *margins))


def _header_footer_record(settings: PrintSettings) -> bytes:
    if settings.header is None and settings.footer is None:
        return b""
    # Even and first-page headers and footers are not written.
    return header_footer_block(_HEADER_FOOTER_FLAGS, (settings.header, settings.footer, None, None, None, None))


def _breaks(begin: int, end: int, breaks: Sequence[int] | None, last: int) -> bytes:
    # Each break spans the whole sheet and is manual.
    return breaks_block(begin, end, [(value, 0, last, True, 0) for value in breaks or ()])


def print_records(settings: PrintSettings) -> bytes:
    """The records from BrtPrintOptions to the column breaks, in Worksheet order."""
    options = _PRINT_OPTIONS_DEFAULT
    if settings.center_horizontally:
        options |= 0x01
    if settings.center_vertically:
        options |= 0x02
    margins = settings.margins or PageMargins()
    return b"".join((
        build_record(BRT_PRINT_OPTIONS, struct.pack("<H", options)),
        margins_record((margins.left, margins.right, margins.top, margins.bottom, margins.header, margins.footer)),
        _xlsb_page_setup_record(settings),
        _header_footer_record(settings),
        _breaks(BRT_BEGIN_ROW_BREAKS, BRT_END_ROW_BREAKS, settings.row_breaks, 16383),
        _breaks(BRT_BEGIN_COL_BREAKS, BRT_END_COL_BREAKS, settings.col_breaks, 1048575),
    ))
