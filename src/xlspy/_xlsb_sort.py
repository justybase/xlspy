"""The BrtBeginSortState block of an XLSB auto filter.

Both the workbook writer and the loaded-model save encode the same records from
the object model's sort state.  Keeping the encoder here means the two paths
cannot drift apart, and neither has to import the other's object model.
"""

from __future__ import annotations

import struct
from typing import Any

from ._sort_state import EXTRA_ICON_SET_CODES, check_sort_state
from ._xlsb_cf import ICON_SET_CODES
from ._xlsb_metadata import wide_string
from .biff_utils import build_record
from .exceptions import UnsupportedFeatureError
from .refs import column_label, parse_cell_range, sort_state_ref

BEGIN_SORT_STATE, END_SORT_STATE = 0x212, 0x213
BEGIN_SORT_COND, END_SORT_COND = 0x214, 0x215
BEGIN_SORT_COND14, END_SORT_COND14 = 0x480, 0x481

_SORT_ON = {"value": 0, "cellColor": 1, "fontColor": 2, "icon": 3}


def range_ref(bounds: tuple[int, int, int, int]) -> str:
    """The A1 label of a ``(first_col, first_row, last_col, last_row)`` tuple."""
    first_col, first_row, last_col, last_row = bounds
    return (
        f"{column_label(first_col)}{first_row + 1}:"
        f"{column_label(last_col)}{last_row + 1}"
    )


def sort_state_records(filter_ref: str | None, sort_state: Any) -> bytes:
    """Encode BrtBeginSortState .. BrtEndSortState for ``sort_state``."""

    conditions = check_sort_state(sort_state, filter_ref)
    sort_ref = sort_state_ref(filter_ref, getattr(sort_state, "ref", None))
    if sort_ref is None:
        raise ValueError("a sort state needs a range")
    sc1, sr1, sc2, sr2 = parse_cell_range(sort_ref)
    flags = (
        (0x01 if getattr(sort_state, "columnSort", False) else 0)
        | (0x02 if getattr(sort_state, "caseSensitive", False) else 0)
        | (0x04 if getattr(sort_state, "sortMethod", None) in ("stroke", "pinYin") else 0)
    )
    output = [build_record(BEGIN_SORT_STATE, struct.pack("<HIIII", flags, sr1, sr2, sc1, sc2))]

    codes = {**ICON_SET_CODES, **EXTRA_ICON_SET_CODES}
    for condition in conditions:
        cc1, cr1, cc2, cr2 = parse_cell_range(condition.ref)
        sort_by = getattr(condition, "sortBy", None) or "value"
        sort_on = _SORT_ON[sort_by]
        condition_flags = (
            (0x01 if getattr(condition, "descending", False) else 0) | (sort_on << 1)
        )
        header = struct.pack("<HIIII", condition_flags, cr1, cr2, cc1, cc2)
        custom = getattr(condition, "customList", None)
        if sort_on == 0:
            union = bytes(8)
            tail = b"\xff" * 4 if not custom else wide_string(custom)
            output.append(build_record(BEGIN_SORT_COND, header + union + tail))
            output.append(build_record(END_SORT_COND))
            continue
        if sort_on in (1, 2):
            union = struct.pack("<II", getattr(condition, "dxfId", 0) or 0, 0)
            tail = b"\xff" * 4 if not custom else wide_string(custom)
            output.append(build_record(BEGIN_SORT_COND, header + union + tail))
            output.append(build_record(END_SORT_COND))
            continue

        icon_set = getattr(condition, "iconSet", None)
        assert isinstance(icon_set, str)
        code = codes.get(icon_set)
        if code is None:
            raise UnsupportedFeatureError(
                f"unsupported icon set {icon_set!r} for XLSB sort"
            )
        icon_id = getattr(condition, "iconId", None)
        assert isinstance(icon_id, int)
        union = struct.pack("<II", code, icon_id)
        # Custom lists are ignored for icon sorts, as in Excel.
        tail = b"\xff" * 4
        if code in EXTRA_ICON_SET_CODES.values():
            output.append(build_record(BEGIN_SORT_COND14, header + union + tail))
            output.append(build_record(END_SORT_COND14))
        else:
            output.append(build_record(BEGIN_SORT_COND, header + union + tail))
            output.append(build_record(END_SORT_COND))

    output.append(build_record(END_SORT_STATE))
    return b"".join(output)
