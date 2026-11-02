"""Module-level helpers of the XLSX writer: XML fragments, spool and flush constants."""
from __future__ import annotations

from typing import IO, Any, Protocol
from dataclasses import dataclass, field
from .printing import (
    PageMargins,
)
from ._xlsb_writer_consts import _REL_PIVOT_CACHE as _REL_PIVOT_CACHE, _REL_PIVOT_TABLE as _REL_PIVOT_TABLE


_ATTR_ENTITIES = {'"': "&quot;"}
# Rows sampled for automatic column widths (the historical limit).
_WIDTH_SAMPLE_ROWS = 1000  # rows measured for automatic column widths by default
# Encoded row bytes are handed to the spool once they grow past this size.
_FLUSH_BYTES = 1 << 20
# Shared-string entries written per chunk of sharedStrings.xml.
_SST_CHUNK_STRINGS = 4096
# Larger spools cost resident memory without making writes faster: 1 MiB
# and 8 MiB write 50,000 x 50 rows in the same time, 1 MiB with 14 MB less.
DEFAULT_SPOOL_BYTES = 1 << 20
_EMPTY_ROW_STYLES: dict = {}


class _ByteSink(Protocol):
    def write(self, data: bytes, /) -> object: ...


class _Tee:
    """Write the encoded rows to the spool and to the deflate segment."""

    __slots__ = ("_spool", "_segment")

    def __init__(self, spool: IO[bytes], segment: _ByteSink):
        self._spool = spool
        self._segment = segment

    def write(self, data: bytes) -> int:
        self._segment.write(data)
        return self._spool.write(data)


_REL_TABLE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"
_REL_DRAWING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
_REL_HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
_REL_WORKSHEET = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
_CT_PIVOT_TABLE = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"
_CT_PIVOT_CACHE = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"
_DEFAULT_MARGINS = PageMargins()


@dataclass
class _SheetBody:
    """What the row pass learned: sizes and the cells that need trailer entries."""

    num_rows: int
    max_cols: int
    head: list
    hyperlinks: list
    table_refs: set = field(default_factory=set)
    # For the autofilter, as in the XLSB writer: the columns up to the last one
    # holding a value (trailing None and formatted blanks do not count), and the
    # rows down to the last row of an array formula, which may reach past the data.
    value_cols: int = 0
    array_rows: int = 0
    tables: list = field(default_factory=list)
    header_rows: dict = field(default_factory=dict)  # row index -> cells, for table headers


def _format_number(value: object) -> str:
    """Shortest numeric literal: ``1.0`` -> ``"1"``, ``0.5`` -> ``"0.5"``."""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def _outline_col_attr(col_outline: Any, index: int) -> str:
    """``<col>`` attributes for a column's outline level and collapsed flag."""
    if col_outline is None:
        return ""
    level, _ = col_outline.state_at(index)
    attrs = f' outlineLevel="{level}"' if level else ""
    if col_outline.collapsed_at(index):
        attrs += ' collapsed="1"'
    return attrs


def _alignment_xml(align: Any) -> str:
    attrs = []
    if align.horizontal is not None:
        attrs.append(f'horizontal="{align.horizontal}"')
    if align.vertical is not None:
        attrs.append(f'vertical="{align.vertical}"')
    if align.wrap_text is not None:
        attrs.append(f'wrapText="{1 if align.wrap_text else 0}"')
    if align.shrink_to_fit is not None:
        attrs.append(f'shrinkToFit="{1 if align.shrink_to_fit else 0}"')
    if align.indent is not None:
        attrs.append(f'indent="{align.indent}"')
    if align.text_rotation is not None:
        attrs.append(f'textRotation="{align.text_rotation}"')
    if align.relative_indent is not None:
        attrs.append(f'relativeIndent="{align.relative_indent}"')
    if align.justify_last_line is not None:
        attrs.append(
            f'justifyLastLine="{1 if align.justify_last_line else 0}"'
        )
    if align.reading_order is not None:
        attrs.append(f'readingOrder="{align.reading_order}"')
    return f'<alignment {" ".join(attrs)}/>' if attrs else ""


def _protection_xml(prot: Any) -> str:
    attrs = []
    if prot.locked is not None:
        attrs.append(f'locked="{1 if prot.locked else 0}"')
    if prot.hidden is not None:
        attrs.append(f'hidden="{1 if prot.hidden else 0}"')
    return f'<protection {" ".join(attrs)}/>' if attrs else ""
