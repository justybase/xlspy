"""The rules a sort state has to follow in both workbook formats.

The XLSX and XLSB encoders take the object model's ``SortState`` by its
attributes.  Checking them here once means a bad range or sort key fails before
either writer has touched the output, with the same message in both formats.
"""

from __future__ import annotations

from typing import Any

from .exceptions import UnsupportedFeatureError
from .refs import parse_cell_range, sort_state_ref

SORT_BY_VALUES = ("value", "cellColor", "fontColor", "icon")
SORT_METHODS = ("stroke", "pinYin")

# Icon set name -> number of icons.  The last three are the sets Excel writes
# as BrtBeginSortCond14 (CFFlag14) in XLSB; XLSX names them the same way.
ICON_SET_SIZES = {
    "3Arrows": 3, "3ArrowsGray": 3, "3Flags": 3, "3TrafficLights1": 3,
    "3TrafficLights2": 3, "3Signs": 3, "3Symbols": 3, "3Symbols2": 3,
    "4Arrows": 4, "4ArrowsGray": 4, "4RedToBlack": 4, "4Rating": 4,
    "4TrafficLights": 4,
    "5Arrows": 5, "5ArrowsGray": 5, "5Rating": 5, "5Quarters": 5,
    "3Stars": 3, "3Triangles": 3, "5Boxes": 5,
}

# Icon sets stored in a CFFlag14 record rather than the classic CFFlag.
EXTRA_ICON_SET_CODES = {"3Stars": 0x11, "3Triangles": 0x12, "5Boxes": 0x13}


def _is_bool(value: Any) -> bool:
    return isinstance(value, bool)


def _check_int(condition: Any, field: str, *, minimum: int = 0) -> int | None:
    value = getattr(condition, field, None)
    if value is None:
        return None
    if not isinstance(value, int) or isinstance(value, bool):
        raise TypeError(f"{field} must be an int or None")
    if value < minimum:
        raise ValueError(f"{field} must be at least {minimum}")
    return value


def _check_condition(condition: Any) -> None:
    ref = getattr(condition, "ref", None)
    if ref is None:
        raise ValueError("a sort condition needs a range")
    if not isinstance(ref, str):
        raise TypeError("a sort condition range must be an A1 range string")
    parse_cell_range(ref)

    descending = getattr(condition, "descending", None)
    if descending is not None and not _is_bool(descending):
        raise TypeError("descending must be a bool or None")

    sort_by = getattr(condition, "sortBy", None)
    if sort_by is not None and sort_by not in SORT_BY_VALUES:
        raise ValueError(f"invalid sortBy {sort_by!r}")
    if sort_by is not None and not isinstance(sort_by, str):
        raise TypeError("sortBy must be a string or None")

    custom_list = getattr(condition, "customList", None)
    if custom_list is not None and not isinstance(custom_list, str):
        raise TypeError("customList must be a string or None")

    _check_int(condition, "dxfId")

    if sort_by == "icon":
        icon_set = getattr(condition, "iconSet", None)
        if icon_set is None:
            raise UnsupportedFeatureError("sorting by icon needs an iconSet and iconId")
        if icon_set not in ICON_SET_SIZES:
            raise UnsupportedFeatureError(
                f"unsupported icon set {icon_set!r} for a sort condition"
            )
        icon_id = getattr(condition, "iconId", None)
        if (
            not isinstance(icon_id, int) or isinstance(icon_id, bool)
            or not 0 <= icon_id < ICON_SET_SIZES[icon_set]
        ):
            raise ValueError(f"invalid iconId {icon_id!r}")


def check_sort_state(sort_state: Any, filter_ref: str | None = None) -> list[Any]:
    """Validate a duck-typed ``SortState`` and return its conditions.

    ``filter_ref`` is checked with the filter it belongs to when the caller
    knows it; otherwise only the sort state's own range is validated.
    """

    conditions = getattr(sort_state, "sortCondition", None)
    if conditions is None or not hasattr(sort_state, "ref"):
        raise TypeError("sort_state must be a SortState with ref and sortCondition")
    if not isinstance(conditions, (list, tuple)):
        raise TypeError("sort_state must hold a list of sort conditions")

    ref = getattr(sort_state, "ref", None)
    if ref is not None:
        if not isinstance(ref, str):
            raise TypeError("a sort state range must be an A1 range string")
        parse_cell_range(ref)
    if filter_ref is not None:
        # The range Excel stores is the filter's data body when the caller did
        # not give a narrower one; it still has to be a valid A1 range.
        sort_ref = sort_state_ref(filter_ref, ref)
        assert sort_ref is not None
        parse_cell_range(sort_ref)

    column_sort = getattr(sort_state, "columnSort", None)
    if column_sort is not None and not _is_bool(column_sort):
        raise TypeError("columnSort must be a bool or None")
    case_sensitive = getattr(sort_state, "caseSensitive", None)
    if case_sensitive is not None and not _is_bool(case_sensitive):
        raise TypeError("caseSensitive must be a bool or None")
    sort_method = getattr(sort_state, "sortMethod", None)
    if sort_method is not None:
        if not isinstance(sort_method, str):
            raise TypeError("sortMethod must be a string or None")
        if sort_method not in SORT_METHODS:
            raise ValueError(f"invalid sortMethod {sort_method!r}")

    for condition in conditions:
        _check_condition(condition)
    return list(conditions)
