"""``ignoredErrors`` in an XLSB worksheet.

The published [MS-XLSB] specification has no record for ignored errors;
Excel 16 writes and reads the block captured from its own files:

- ``BrtBeginIgnoredErrors`` (0x0288), empty;
- one ``BrtIgnoredError`` (0x0289) per ``IgnoredError``, holding a 32-bit flag
  word, the number of ranges and each range as ``rwFirst``, ``rwLast``,
  ``colFirst``, ``colLast`` (0-based, inclusive);
- ``BrtEndIgnoredErrors`` (0x028A), empty.

Only the checks Excel shows as ignored are set: the binary form has no way to
store an explicit "not ignored" (``False``), which reads back as ``None``.
The bit of each flag was read back from Excel (see
``tests/test_excel_com_ignored_errors.py``).
"""

from __future__ import annotations

import struct
from collections.abc import Iterable
from typing import Any

from .biff_utils import build_record
from .refs import parse_cell_range

RECORD_BEGIN = 0x0288
RECORD_ERROR = 0x0289
RECORD_END = 0x028A

# Flag name -> bit, in the order Excel stores them.
BITS = (
    ("evalError", 0x001),
    ("emptyCellReference", 0x002),
    ("numberStoredAsText", 0x004),
    ("formulaRange", 0x008),
    ("formula", 0x010),
    ("twoDigitTextYear", 0x020),
    ("unlockedFormula", 0x040),
    ("listDataValidation", 0x080),
    ("calculatedColumn", 0x100),
)


def ignored_errors_bin(items: Iterable[Any]) -> bytes:
    """The records of ``items`` (``IgnoredError`` objects), or ``b""``."""
    if not items:
        return b""
    out = bytearray(build_record(RECORD_BEGIN))
    for item in items:
        flags = 0
        for name, bit in BITS:
            if getattr(item, name):
                flags |= bit
        ranges = [parse_cell_range(part) for part in item.sqref.split()]
        payload = bytearray(struct.pack("<II", flags, len(ranges)))
        for col_first, row_first, col_last, row_last in ranges:
            payload += struct.pack("<IIII", row_first, row_last, col_first, col_last)
        out += build_record(RECORD_ERROR, payload)
    out += build_record(RECORD_END)
    return bytes(out)


def ignored_errors_from_records(records: Iterable[tuple[int, bytes]]) -> list[Any]:
    """``IgnoredError`` objects from the parsed records of a sheet.

    ``records`` yields ``(record_id, payload)``; the block is read where it
    appears, so a sheet from Excel keeps the ranges and flags it shows.
    """
    from .ignored_errors import IgnoredError

    items = []
    inside = False
    for record_id, payload in records:
        if record_id == RECORD_BEGIN:
            inside = True
            continue
        if record_id == RECORD_END:
            inside = False
            continue
        if not inside or record_id != RECORD_ERROR or len(payload) < 8:
            continue
        flags, count = struct.unpack_from("<II", payload, 0)
        if count > (len(payload) - 8) // 16:
            continue
        ranges = []
        for index in range(count):
            row_first, row_last, col_first, col_last = struct.unpack_from(
                "<IIII", payload, 8 + index * 16
            )
            ranges.append(_reference(row_first, row_last, col_first, col_last))
        values = {
            name: True if flags & bit else None
            for name, bit in BITS
        }
        items.append(IgnoredError(sqref=" ".join(ranges), **values))
    return items


def _reference(row_first: int, row_last: int, col_first: int, col_last: int) -> str:
    from .utils.cell import get_column_letter

    first = f"{get_column_letter(col_first + 1)}{row_first + 1}"
    if (row_first, col_first) == (row_last, col_last):
        return first
    return f"{first}:{get_column_letter(col_last + 1)}{row_last + 1}"
