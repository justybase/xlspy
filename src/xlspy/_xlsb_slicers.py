"""Table slicers as BIFF12 parts and records: the XLSB form of ``slicers.py``.

The layout was mapped from files Excel saved (``tools/capture_xlsb.py --slicers``).
Like Excel, a sheet's slicers share one slicer part and each has its own cache:

* ``xl/slicers/slicerN.bin``: BrtBeginSlicers, (BrtBeginSlicer BrtEndSlicer)*,
  BrtEndSlicers. BrtBeginSlicer holds a flag byte (0x01 shows the caption),
  the first item, the column count, the level, the item row height in EMU and
  three strings: the slicer name, the cache name and the caption.
* ``xl/slicerCaches/slicerCacheN.bin``: the cache name and source column, then
  an Excel 2013 FRT block naming the table id and the 1-based column.
* the sheet lists its slicer part in an FRT block after LISTPARTS; the workbook
  lists every cache in an FRT block after BrtFileRecover.

The drawing frames are the same XML as in XLSX (``slicer_frame_xml``).
"""

from __future__ import annotations

import struct

from ._xlsb_metadata import wide_string
from .biff_utils import build_record
from .slicers import _ITEM_ROW_HEIGHT, Slicer

CT_SLICER_BIN = "application/vnd.ms-excel.slicer"
CT_SLICER_CACHE_BIN = "application/vnd.ms-excel.slicerCache"

_FRT_BEGIN, _FRT_END = 0x23, 0x24
_FRT_2013 = struct.pack("<I", 0x0F03)
_REL_ID_FLAG = 0x08  # FRT header with fRelID: a relationship id follows

_BEGIN_SLICERS, _END_SLICERS = 0x45B, 0x45C
_BEGIN_SLICER, _END_SLICER = 0x43B, 0x43C
_BEGIN_CACHE, _END_CACHE = 0x433, 0x434
_BEGIN_CACHE_DEF, _END_CACHE_DEF = 0x435, 0x436
_CACHE_PIVOT_TABLES = 0x43D
_BEGIN_TABLE_CACHE, _END_TABLE_CACHE = 0x81D, 0x81E
_BEGIN_SHEET_LIST, _END_SHEET_LIST = 0x437, 0x438
_SHEET_SLICER_PART, _END_SHEET_SLICER_PART = 0x439, 0x43A
_BEGIN_BOOK_CACHES, _BOOK_CACHE = 0x81B, 0x81C

_SHOW_CAPTION = 0x01
_SLICER_FLAGS = 0x02 | _SHOW_CAPTION  # as Excel writes a default slicer
_TABLE_CACHE_FLAGS = 0x15  # ascending sort, cross filtering: Excel's defaults


def _rel_id(rel_id: str) -> bytes:
    """An FRT header with a relationship id, its length a uint16 as Excel writes it."""
    encoded = rel_id.encode("utf-16le")
    return struct.pack("<IH", _REL_ID_FLAG, len(encoded) // 2) + encoded


def _slicer_payload(slicer: Slicer) -> bytes:
    return (bytes((_SLICER_FLAGS,)) + struct.pack("<IIII", 0, slicer.columns, 0, _ITEM_ROW_HEIGHT)
            + wide_string(slicer.display_name) + wide_string(slicer.cache_name)
            + wide_string(slicer.display_caption))


def slicers_part_bin(slicers: list[Slicer]) -> bytes:
    """The slicer part of one sheet, holding all of its slicers."""
    body = b"".join(build_record(_BEGIN_SLICER, _slicer_payload(slicer)) + build_record(_END_SLICER)
                    for slicer in slicers)
    return build_record(_BEGIN_SLICERS) + body + build_record(_END_SLICERS)


def slicer_cache_bin(slicer: Slicer, table_id: int, column_id: int) -> bytes:
    """The cache part: links the slicer to table ``table_id`` and its column ``column_id``."""
    table_cache = struct.pack("<IIIB", 0, column_id, table_id, _TABLE_CACHE_FLAGS)
    return (build_record(_BEGIN_CACHE)
            + build_record(_BEGIN_CACHE_DEF, wide_string(slicer.cache_name) + wide_string(slicer.column))
            + build_record(_CACHE_PIVOT_TABLES, struct.pack("<I", 0))
            + build_record(_FRT_BEGIN, _FRT_2013)
            + build_record(_BEGIN_TABLE_CACHE, table_cache) + build_record(_END_TABLE_CACHE)
            + build_record(_FRT_END)
            + build_record(_END_CACHE_DEF) + build_record(_END_CACHE))


def sheet_slicers_bin(rel_id: str) -> bytes:
    """The worksheet FRT block naming the sheet's slicer part by relationship id."""
    return (build_record(_FRT_BEGIN, _FRT_2013) + build_record(_BEGIN_SHEET_LIST)
            + build_record(_SHEET_SLICER_PART, _rel_id(rel_id))
            + build_record(_END_SHEET_SLICER_PART) + build_record(_END_SHEET_LIST)
            + build_record(_FRT_END))


def workbook_slicer_caches_bin(rel_ids: list[str]) -> bytes:
    """The workbook FRT block listing every slicer cache by relationship id."""
    if not rel_ids:
        return b""
    return (build_record(_FRT_BEGIN, _FRT_2013) + build_record(_BEGIN_BOOK_CACHES)
            + b"".join(build_record(_BOOK_CACHE, _rel_id(rel_id)) for rel_id in rel_ids)
            + build_record(_FRT_END))
