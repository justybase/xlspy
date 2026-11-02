"""Table slicers: the filter buttons that drive an Excel table by one column.

This module holds the options and the XLSX form; ``_xlsb_slicers`` writes XLSB.
A slicer is written as four things, in the layout Excel itself produces:

* a slicer part (``xl/slicers/sliceN.xml``) with the caption and the cache name,
* a cache part (``xl/slicerCaches/slicerCacheN.xml``) that names the table and
  the column (its 1-based position in the table),
* a drawing object in the sheet's ``drawingN.xml``, with a text fallback for
  Excel versions without slicers,
* a ``Slicer_<column>`` defined name, which Excel uses to link the cache.

Only table slicers are written; pivot slicers and timelines are not supported.
"""

from __future__ import annotations
from typing import Any

from dataclasses import dataclass
import re
import xml.sax.saxutils as saxutils

from ._extlst import SLICER_DRAWING, SLICER_DRAWING_CHOICE, EXT_TABLE_SLICER_CACHE, X14, X15
from .refs import parse_cell_range
from .pivots import _MAIN_NS as _MAIN_NS
from ._text import _attr as _attr

_MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_EMU_PER_PIXEL = 9525
# Excel's default item row height for a slicer, in EMU.
_ITEM_ROW_HEIGHT = 241300
_FALLBACK_TEXT = (
    "This shape represents a table slicer. Table slicers are not supported "
    "in this version of Excel."
)


@dataclass(frozen=True, slots=True)
class Slicer:
    """A slicer filtering the table ``table`` by its column ``column``.

    ``table`` is the display name of an ``ExcelTable`` written to the same
    sheet in the same call, and ``column`` the header text of one of its
    columns. ``name`` (default: the column) and ``caption`` (default: the
    column) must be unique within the workbook for ``name``.
    """

    table: str
    column: str
    anchor: str = "H2"
    width: int = 144
    height: int = 180
    caption: str | None = None
    name: str | None = None
    columns: int = 1

    def __post_init__(self) -> None:
        for field_name in ("table", "column"):
            value = getattr(self, field_name)
            if not isinstance(value, str) or not value:
                raise ValueError(f"slicer {field_name} must be a non-empty string")
        if not isinstance(self.anchor, str):
            raise TypeError("slicer anchor must be an A1 cell such as 'H2'")
        first_col, first_row, last_col, last_row = parse_cell_range(self.anchor)
        if (first_col, first_row) != (last_col, last_row):
            raise ValueError(f"slicer anchor must be a single cell, got {self.anchor!r}")
        for field_name in ("width", "height", "columns"):
            value = getattr(self, field_name)
            if isinstance(value, bool) or not isinstance(value, int) or value <= 0:
                raise ValueError(f"slicer {field_name} must be a positive integer")
        for field_name in ("caption", "name"):
            value = getattr(self, field_name)
            if value is not None and (not isinstance(value, str) or not value):
                raise ValueError(f"slicer {field_name} must be a non-empty string or None")

    @property
    def display_name(self) -> str:
        """``name``, or the column name when none is set."""
        return self.name or self.column

    @property
    def display_caption(self) -> str:
        """``caption``, or the column name when none is set."""
        return self.caption or self.column

    @property
    def cache_name(self) -> str:
        """The defined name that links a slicer to its cache, e.g. ``Slicer_Region``."""
        return "Slicer_" + re.sub(r"\W", "_", self.column)


def normalize_slicers(slicers: Any, tables: Any, used_names: set[str], used_caches: set[str],
                      used_targets: set[tuple[str, str]]) -> list[Slicer]:
    """Validate slicers against this sheet's tables and the names already in the workbook.

    ``used_*`` hold casefolded names from earlier sheets; the caller adds the
    returned names to them once the sheet is accepted.
    """
    if not slicers:
        return []
    tables_by_name = {table.display_name.casefold(): table for table in tables}
    names = set(used_names)
    caches = set(used_caches)
    targets = set(used_targets)
    result = []
    for slicer in slicers:
        if not isinstance(slicer, Slicer):
            raise TypeError("slicers must be Slicer instances")
        table = tables_by_name.get(slicer.table.casefold())
        if table is None:
            raise ValueError(
                f"slicer table {slicer.table!r} is not a table written to this sheet"
            )
        if not table.header_row:
            raise ValueError(f"slicer table {slicer.table!r} needs a header row")
        if table.columns is not None and slicer.column not in table.columns:
            raise ValueError(f"slicer column {slicer.column!r} is not a column of {slicer.table!r}")
        key = (slicer.table.casefold(), slicer.column.casefold())
        if key in targets:
            raise ValueError(f"column {slicer.column!r} of {slicer.table!r} already has a slicer")
        if slicer.display_name.casefold() in names:
            raise ValueError(f"duplicate slicer name {slicer.display_name!r}; pass name=")
        if slicer.cache_name.casefold() in caches:
            raise ValueError(
                f"slicer cache name {slicer.cache_name!r} is already used by another slicer"
            )
        names.add(slicer.display_name.casefold())
        caches.add(slicer.cache_name.casefold())
        targets.add(key)
        result.append(slicer)
    return result


def slicer_part_xml(slicer: Slicer) -> str:
    """The ``slicers`` part: one ``slicer`` element with its caption and cache."""
    columns = f' columnCount="{slicer.columns}"' if slicer.columns != 1 else ""
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<slicers xmlns="{X14}">'
        f'<slicer name="{_attr(slicer.display_name)}" cache="{_attr(slicer.cache_name)}" '
        f'caption="{_attr(slicer.display_caption)}" rowHeight="{_ITEM_ROW_HEIGHT}"{columns}/>'
        "</slicers>"
    )


def slicer_cache_xml(slicer: Slicer, table_id: int, column_id: int) -> str:
    """The cache part: links the slicer to table ``table_id`` and column ``column_id``."""
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<slicerCacheDefinition xmlns="{X14}" xmlns:x="{_MAIN_NS}" '
        f'name="{_attr(slicer.cache_name)}" sourceName="{_attr(slicer.column)}">'
        f'<extLst><x:ext uri="{EXT_TABLE_SLICER_CACHE}" xmlns:x15="{X15}">'
        f'<x15:tableSlicerCache tableId="{table_id}" column="{column_id}"/>'
        "</x:ext></extLst></slicerCacheDefinition>"
    )


def slicer_frame_xml(slicer: Slicer, object_id: int) -> str:
    """The drawing object for one slicer: a ``sle`` frame, with a shape for old Excel.

    Excel reads the ``mc:Choice`` branch when it supports slicers and the
    ``mc:Fallback`` shape otherwise. The frame refers to the slicer by name,
    so it needs no relationship in the drawing part.
    """
    first_col, first_row, _, _ = parse_cell_range(slicer.anchor)
    cx = slicer.width * _EMU_PER_PIXEL
    cy = slicer.height * _EMU_PER_PIXEL
    name = _attr(slicer.display_name)
    return (
        "<oneCellAnchor>"
        f"<from><col>{first_col}</col><colOff>0</colOff>"
        f"<row>{first_row}</row><rowOff>0</rowOff></from>"
        f'<ext cx="{cx}" cy="{cy}"/>'
        f'<mc:AlternateContent xmlns:mc="{_MC_NS}">'
        f'<mc:Choice xmlns:sle15="{SLICER_DRAWING_CHOICE}" Requires="sle15">'
        '<graphicFrame macro=""><nvGraphicFramePr>'
        f'<cNvPr id="{object_id}" name="{name}"/>'
        "<cNvGraphicFramePr/></nvGraphicFramePr>"
        '<xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xfrm>'
        '<a:graphic><a:graphicData uri="' + SLICER_DRAWING + '">'
        f'<sle:slicer xmlns:sle="{SLICER_DRAWING}" name="{name}"/>'
        "</a:graphicData></a:graphic></graphicFrame>"
        "</mc:Choice>"
        "<mc:Fallback>"
        '<sp macro="" textlink=""><nvSpPr><cNvPr id="0" name=""/>'
        '<cNvSpPr><a:spLocks noTextEdit="1"/></cNvSpPr></nvSpPr>'
        '<spPr><a:xfrm><a:off x="0" y="0"/>'
        f'<a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '<a:solidFill><a:prstClr val="white"/></a:solidFill>'
        '<a:ln w="1"><a:solidFill><a:prstClr val="green"/></a:solidFill></a:ln></spPr>'
        '<txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip"/><a:lstStyle/>'
        f'<a:p><a:r><a:rPr lang="en-US" sz="1100"/><a:t>{saxutils.escape(_FALLBACK_TEXT)}</a:t>'
        "</a:r></a:p></txBody></sp>"
        "</mc:Fallback></mc:AlternateContent>"
        "<clientData/></oneCellAnchor>"
    )


__all__ = [
    "Slicer", "normalize_slicers", "slicer_part_xml", "slicer_cache_xml", "slicer_frame_xml",
]
