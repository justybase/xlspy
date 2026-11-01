"""Sparklines: small line, column or win/loss charts drawn inside cells.

This module holds the options and the XLSX form; ``_xlsb_sparklines`` writes XLSB.

A ``SparklineGroup`` is one style applied to several sparklines. ``data`` is
the source range and ``location`` the cells that show the sparklines:

* one location cell draws one sparkline from all of ``data``;
* a column of location cells takes one data row each, so the row counts match;
* a row of location cells takes one data column each.

``data`` may name another sheet (``"'Sales data'!B2:M2"``); otherwise it is read
from the sheet the group is written to.
"""

from __future__ import annotations

from dataclasses import dataclass
import re
import xml.sax.saxutils as saxutils

from ._extlst import EXT_SPARKLINES, X14, XM, ext
from .refs import column_label, parse_cell_range, quote_sheet_name

# Excel's sparkline types: a win/loss sparkline is written as "stacked".
_KIND_TYPES = {"line": None, "column": "column", "win_loss": "stacked"}
_EMPTY_CELLS = {"gap": "gap", "zero": "zero", "connect": "span"}
_HEX_RE = re.compile(r"#?([0-9A-Fa-f]{6})")
_DEFAULT_LINE_WIDTH = 0.75
_DEFAULT_COLOR = "#376092"
_POINT_COLOR = "#D00000"


def _argb(value: str, name: str) -> str:
    match = _HEX_RE.fullmatch(value) if isinstance(value, str) else None
    if match is None:
        raise ValueError(f"{name} must be an RGB hex colour such as '#376092', got {value!r}")
    return "FF" + match[1].upper()


def _cell_ref(column: int, row: int) -> str:
    return f"{column_label(column)}{row + 1}"


@dataclass(frozen=True, slots=True)
class SparklineGroup:
    """One style of sparklines over ``data``, shown in ``location``."""

    data: str
    location: str
    kind: str = "line"
    color: str = _DEFAULT_COLOR
    negative_color: str = _POINT_COLOR
    axis_color: str = "#000000"
    marker_color: str = _POINT_COLOR
    high_color: str = _POINT_COLOR
    low_color: str = _POINT_COLOR
    first_color: str = _POINT_COLOR
    last_color: str = _POINT_COLOR
    line_width: float = _DEFAULT_LINE_WIDTH
    markers: bool = False
    high_point: bool = False
    low_point: bool = False
    first_point: bool = False
    last_point: bool = False
    negative_points: bool = False
    empty_cells: str = "gap"
    show_hidden: bool = False
    show_axis: bool = False
    right_to_left: bool = False
    # A number fixes the axis; "group" shares it across the group and
    # "individual" (the default) scales each sparkline on its own.
    axis_min: float | str | None = None
    axis_max: float | str | None = None

    def __post_init__(self):
        if self.kind not in _KIND_TYPES:
            raise ValueError(f"sparkline kind must be one of {sorted(_KIND_TYPES)}, got {self.kind!r}")
        if self.empty_cells not in _EMPTY_CELLS:
            raise ValueError(f"empty_cells must be one of {sorted(_EMPTY_CELLS)}, got {self.empty_cells!r}")
        if isinstance(self.line_width, bool) or not isinstance(self.line_width, (int, float)) \
                or self.line_width <= 0:
            raise ValueError("line_width must be a positive number of points")
        for name in ("color", "negative_color", "axis_color", "marker_color",
                     "high_color", "low_color", "first_color", "last_color"):
            _argb(getattr(self, name), name)
        for name in ("axis_min", "axis_max"):
            value = getattr(self, name)
            if value is None or value in ("group", "individual"):
                continue
            if isinstance(value, bool) or not isinstance(value, (int, float)):
                raise ValueError(f"{name} must be a number, 'group', 'individual' or None")
        if not isinstance(self.location, str) or not isinstance(self.data, str):
            raise TypeError("sparkline location and data must be A1 range strings")


def _expand(group: SparklineGroup, sheet_name: str) -> list[tuple[str, str]]:
    """``(location cell, qualified data range)`` for every sparkline in the group."""
    sheet, bang, ref = group.data.rpartition("!")
    prefix = sheet + "!" if bang else quote_sheet_name(sheet_name) + "!"
    dc1, dr1, dc2, dr2 = parse_cell_range(ref)
    lc1, lr1, lc2, lr2 = parse_cell_range(group.location)
    if (lc1, lr1) == (lc2, lr2):
        return [(_cell_ref(lc1, lr1), f"{prefix}{ref}")]
    if lc1 == lc2:
        count = lr2 - lr1 + 1
        if dr2 - dr1 + 1 != count:
            raise ValueError(
                f"location {group.location!r} has {count} cells but data {group.data!r} "
                f"has {dr2 - dr1 + 1} rows; a column of cells needs one data row each"
            )
        return [
            (_cell_ref(lc1, lr1 + i),
             f"{prefix}{_cell_ref(dc1, dr1 + i)}:{_cell_ref(dc2, dr1 + i)}")
            for i in range(count)
        ]
    if lr1 == lr2:
        count = lc2 - lc1 + 1
        if dc2 - dc1 + 1 != count:
            raise ValueError(
                f"location {group.location!r} has {count} cells but data {group.data!r} "
                f"has {dc2 - dc1 + 1} columns; a row of cells needs one data column each"
            )
        return [
            (_cell_ref(lc1 + i, lr1),
             f"{prefix}{_cell_ref(dc1 + i, dr1)}:{_cell_ref(dc1 + i, dr2)}")
            for i in range(count)
        ]
    raise ValueError(
        f"location {group.location!r} must be one cell, or a single row or column of cells"
    )


def normalize_sparklines(sparklines) -> list[SparklineGroup]:
    """Validate the groups, including their location and data shapes, before any row is written."""
    if not sparklines:
        return []
    groups = []
    for group in sparklines:
        if not isinstance(group, SparklineGroup):
            raise TypeError("sparklines must be SparklineGroup instances")
        _expand(group, "Sheet1")  # the shape check does not depend on the sheet name
        groups.append(group)
    return groups


def _sparkline_group_xml(group: SparklineGroup, sheet_name: str) -> str:
    attrs = []
    kind_type = _KIND_TYPES[group.kind]
    if kind_type is not None:
        attrs.append(f'type="{kind_type}"')
    attrs.append(f'displayEmptyCellsAs="{_EMPTY_CELLS[group.empty_cells]}"')
    if group.negative_points:
        attrs.append('negative="1"')
    if group.markers:
        attrs.append('markers="1"')
    for flag, name in ((group.high_point, "high"), (group.low_point, "low"),
                       (group.first_point, "first"), (group.last_point, "last"),
                       (group.show_hidden, "displayHidden"), (group.show_axis, "displayXAxis"),
                       (group.right_to_left, "rightToLeft")):
        if flag:
            attrs.append(f'{name}="1"')
    if group.line_width != _DEFAULT_LINE_WIDTH:
        attrs.append(f'lineWeight="{group.line_width:g}"')
    for name, value in (("min", group.axis_min), ("max", group.axis_max)):
        if value is None:
            continue
        if isinstance(value, str):
            attrs.append(f'{name}AxisType="{value}"')
        else:
            attrs.append(f'{name}AxisType="custom"')
            attrs.append(f'manual{name.capitalize()}="{value:g}"')
    colors = (
        ("colorSeries", group.color), ("colorNegative", group.negative_color),
        ("colorAxis", group.axis_color), ("colorMarkers", group.marker_color),
        ("colorFirst", group.first_color), ("colorLast", group.last_color),
        ("colorHigh", group.high_color), ("colorLow", group.low_color),
    )
    color_xml = "".join(f'<x14:{tag} rgb="{_argb(value, tag)}"/>' for tag, value in colors)
    sparkline_xml = "".join(
        f"<x14:sparkline><xm:f>{saxutils.escape(formula)}</xm:f>"
        f"<xm:sqref>{cell}</xm:sqref></x14:sparkline>"
        for cell, formula in _expand(group, sheet_name)
    )
    return (
        f"<x14:sparklineGroup {' '.join(attrs)}>{color_xml}"
        f"<x14:sparklines>{sparkline_xml}</x14:sparklines></x14:sparklineGroup>"
    )


def sparklines_ext(groups: list[SparklineGroup], sheet_name: str) -> str:
    """The sparkline ``<ext>`` for a sheet's ``extLst``, or ``""`` when there are no groups."""
    if not groups:
        return ""
    body = "".join(_sparkline_group_xml(group, sheet_name) for group in groups)
    return ext(
        EXT_SPARKLINES,
        f'<x14:sparklineGroups xmlns:xm="{XM}">{body}</x14:sparklineGroups>',
        ("x14", X14),
    )
