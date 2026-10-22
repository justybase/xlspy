"""Chart models: what a chart shows and how it is configured.

The models only validate and store settings. DrawingML serialization lives in
``chart_parts`` so that the XLSX and XLSB writers share one XML builder.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from .refs import parse_cell_range

CHART_TYPES = frozenset({"area", "bar", "barh", "doughnut", "line", "pie", "scatter"})
# Chart.grouping values and the chart types that accept them.
GROUPINGS = {"stacked": "stacked", "percent_stacked": "percentStacked"}
_GROUPED_TYPES = frozenset({"area", "bar", "barh", "line"})
# Types without markers, lines or axes of their own.
_FILLED_TYPES = frozenset({"area", "bar", "barh", "doughnut", "pie"})
_ROUND_TYPES = frozenset({"doughnut", "pie"})

_RGB_RE = re.compile(r"[0-9A-Fa-f]{6}")
MARKER_SYMBOLS = frozenset(
    {"circle", "square", "diamond", "triangle", "x", "star", "dot", "dash", "plus", "none"}
)
LEGEND_POSITIONS = {
    "right": "r",
    "left": "l",
    "top": "t",
    "bottom": "b",
    "top_right": "tr",
}
# User-facing label positions and their DrawingML spelling.
LABEL_POSITIONS = {
    "center": "ctr",
    "inside_end": "inEnd",
    "inside_base": "inBase",
    "outside_end": "outEnd",
    "above": "t",
    "below": "b",
    "left": "l",
    "right": "r",
    "best_fit": "bestFit",
}
_LABEL_POSITIONS_BY_TYPE = {
    "bar": {"center", "inside_end", "inside_base", "outside_end"},
    "barh": {"center", "inside_end", "inside_base", "outside_end"},
    "line": {"center", "above", "below", "left", "right"},
    "scatter": {"center", "above", "below", "left", "right"},
    "pie": {"center", "inside_end", "outside_end", "best_fit"},
    # Excel has no label positions for these two.
    "area": set(),
    "doughnut": set(),
}


def _validate_chart_range(reference: str, kind: str) -> tuple[int, int, int, int]:
    """Parse an A1 range on the chart's own sheet (no ``!`` allowed)."""
    if not isinstance(reference, str) or "!" in reference:
        raise ValueError(f"{kind} must be a range on the chart sheet, got {reference!r}")
    return parse_cell_range(reference)


def _normalize_rgb(value: str | None, kind: str) -> str | None:
    if value is None:
        return None
    if not isinstance(value, str):
        raise TypeError(f"{kind} must be a '#RRGGBB' string or None")
    digits = value.removeprefix("#")
    if not _RGB_RE.fullmatch(digits):
        raise ValueError(f"invalid {kind} {value!r}; expected '#RRGGBB'")
    return digits.upper()


def _check_number(value, kind: str, *, minimum=None, maximum=None):
    if value is None:
        return None
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise TypeError(f"{kind} must be a number or None")
    if minimum is not None and value < minimum:
        raise ValueError(f"{kind} must be at least {minimum}")
    if maximum is not None and value > maximum:
        raise ValueError(f"{kind} must be at most {maximum}")
    return float(value)


@dataclass(frozen=True, slots=True)
class ChartMarker:
    """Marker shape for line and scatter series.

    The marker is filled with the series ``color``.
    """

    symbol: str = "circle"
    size: int = 7

    def __post_init__(self):
        if self.symbol not in MARKER_SYMBOLS:
            raise ValueError(
                f"invalid marker symbol {self.symbol!r}; "
                f"expected one of {sorted(MARKER_SYMBOLS)}"
            )
        if isinstance(self.size, bool) or not isinstance(self.size, int):
            raise TypeError("marker size must be an int")
        if not 2 <= self.size <= 72:
            raise ValueError("marker size must be between 2 and 72")


@dataclass(frozen=True, slots=True)
class ChartAxis:
    """Settings for one chart axis.

    ``gridlines`` defaults to on for the value axis and off for the category
    or X axis when left as ``None``.
    """

    title: str | None = None
    minimum: float | None = None
    maximum: float | None = None
    major_unit: float | None = None
    number_format: str | None = None
    gridlines: bool | None = None
    reverse: bool = False
    visible: bool = True

    def __post_init__(self):
        if self.title is not None and not isinstance(self.title, str):
            raise TypeError("axis title must be a string or None")
        minimum = _check_number(self.minimum, "axis minimum")
        maximum = _check_number(self.maximum, "axis maximum")
        if minimum is not None and maximum is not None and minimum >= maximum:
            raise ValueError("axis minimum must be less than axis maximum")
        major_unit = _check_number(self.major_unit, "axis major_unit")
        if major_unit is not None and major_unit <= 0:
            raise ValueError("axis major_unit must be positive")
        if self.number_format is not None:
            if not isinstance(self.number_format, str) or not self.number_format:
                raise TypeError("axis number_format must be a non-empty string or None")
        if self.gridlines is not None and not isinstance(self.gridlines, bool):
            raise TypeError("axis gridlines must be a bool or None")
        if not isinstance(self.reverse, bool) or not isinstance(self.visible, bool):
            raise TypeError("axis reverse and visible must be bools")


@dataclass(frozen=True, slots=True)
class ChartLegend:
    """Legend placement; ``overlay`` lets the legend cover the plot area."""

    position: str = "right"
    overlay: bool = False

    def __post_init__(self):
        if self.position not in LEGEND_POSITIONS:
            raise ValueError(
                f"invalid legend position {self.position!r}; "
                f"expected one of {sorted(LEGEND_POSITIONS)}"
            )
        if not isinstance(self.overlay, bool):
            raise TypeError("legend overlay must be a bool")


@dataclass(frozen=True, slots=True)
class DataLabels:
    """Value labels drawn on the points or bars of a chart."""

    number_format: str | None = None
    position: str | None = None
    show_category: bool = False
    show_series_name: bool = False

    def __post_init__(self):
        if self.number_format is not None and (
            not isinstance(self.number_format, str) or not self.number_format
        ):
            raise TypeError("data label number_format must be a non-empty string or None")
        if self.position is not None and self.position not in LABEL_POSITIONS:
            raise ValueError(
                f"invalid data label position {self.position!r}; "
                f"expected one of {sorted(LABEL_POSITIONS)}"
            )
        if not isinstance(self.show_category, bool) or not isinstance(self.show_series_name, bool):
            raise TypeError("data label show_category and show_series_name must be bools")


@dataclass(frozen=True, slots=True)
class ChartSeries:
    """One data series of a ``Chart``.

    ``categories`` labels the points of bar, line and pie series. ``x_values``
    gives the numeric X positions of a scatter series. ``color`` is the fill of
    bars and pie slices, and the line or marker colour of line and scatter
    series. ``marker`` and ``line`` apply to line and scatter series only.
    """

    values: str
    title: str | None = None
    title_cell: str | None = None
    categories: str | None = None
    x_values: str | None = None
    color: str | None = None
    marker: ChartMarker | None = None
    line: bool | None = None

    def __post_init__(self):
        if not isinstance(self.values, str):
            raise TypeError("series values must be an A1 range such as 'B2:B5'")
        _validate_chart_range(self.values, "series values")
        if self.title is not None and not isinstance(self.title, str):
            raise TypeError("series title must be a string or None")
        if self.title_cell is not None:
            if not isinstance(self.title_cell, str):
                raise TypeError("series title_cell must be an A1 cell or None")
            first_col, first_row, last_col, last_row = _validate_chart_range(
                self.title_cell, "series title_cell"
            )
            if (first_col, first_row) != (last_col, last_row):
                raise ValueError(
                    f"series title_cell must be a single cell, got {self.title_cell!r}"
                )
        if self.title is not None and self.title_cell is not None:
            raise ValueError("series title and title_cell are mutually exclusive")
        if self.categories is not None:
            if not isinstance(self.categories, str):
                raise TypeError("series categories must be an A1 range or None")
            _validate_chart_range(self.categories, "series categories")
        if self.x_values is not None:
            if not isinstance(self.x_values, str):
                raise TypeError("series x_values must be an A1 range or None")
            _validate_chart_range(self.x_values, "series x_values")
        object.__setattr__(self, "color", _normalize_rgb(self.color, "series color"))
        if self.marker is not None and not isinstance(self.marker, ChartMarker):
            raise TypeError("series marker must be a ChartMarker or None")
        if self.line is not None and not isinstance(self.line, bool):
            raise TypeError("series line must be a bool or None")


@dataclass(frozen=True, slots=True)
class Chart:
    """A chart anchored to a cell of the sheet it is written to.

    ``x_axis`` is the category (or X value) axis and ``y_axis`` the value axis,
    whatever the orientation: in a horizontal bar chart the category axis is
    drawn vertically but is still ``x_axis``.

    ``legend`` accepts ``None`` (default, right-hand legend), ``False`` (no
    legend), ``True`` (default legend) or a ``ChartLegend``.

    ``grouping`` stacks the series of a bar, horizontal bar, line or area
    chart: ``"stacked"`` adds them up, ``"percent_stacked"`` shows each
    category as 100%. ``None`` (default) draws them side by side.
    """

    chart_type: str
    series: list[ChartSeries] | tuple[ChartSeries, ...]
    title: str | None = None
    anchor: str = "E2"
    width: float = 15.0
    height: float = 7.5
    style: int | None = None
    x_axis: ChartAxis | None = None
    y_axis: ChartAxis | None = None
    legend: ChartLegend | bool | None = None
    data_labels: DataLabels | None = None
    show_data_labels: bool = False
    title_size: float | None = None
    grouping: str | None = None

    def __post_init__(self):
        if self.chart_type not in CHART_TYPES:
            raise ValueError(
                f"invalid chart type {self.chart_type!r}; "
                f"expected one of {sorted(CHART_TYPES)}"
            )
        if not isinstance(self.series, (list, tuple)) or not self.series:
            raise ValueError("a chart needs at least one series")
        series = tuple(self.series)
        for item in series:
            if not isinstance(item, ChartSeries):
                raise TypeError("chart series must be ChartSeries instances")
        object.__setattr__(self, "series", series)
        self._check_series_for_type(series)
        if self.title is not None and not isinstance(self.title, str):
            raise TypeError("chart title must be a string or None")
        if not isinstance(self.anchor, str):
            raise TypeError("chart anchor must be an A1 cell such as 'E2'")
        first_col, first_row, last_col, last_row = _validate_chart_range(
            self.anchor, "chart anchor"
        )
        if (first_col, first_row) != (last_col, last_row):
            raise ValueError(f"chart anchor must be a single cell, got {self.anchor!r}")
        for name in ("width", "height"):
            value = getattr(self, name)
            if isinstance(value, bool) or not isinstance(value, (int, float)):
                raise TypeError(f"chart {name} must be a number (cm)")
            if value <= 0:
                raise ValueError(f"chart {name} must be positive")
        if self.style is not None:
            if isinstance(self.style, bool) or not isinstance(self.style, int):
                raise TypeError("chart style must be an int or None")
            if not 1 <= self.style <= 48:
                raise ValueError("chart style must be between 1 and 48")
        self._check_axes()
        self._check_legend()
        self._check_labels()
        object.__setattr__(self, "title_size", _check_number(
            self.title_size, "chart title_size", minimum=1, maximum=400,
        ))
        if not isinstance(self.show_data_labels, bool):
            raise TypeError("show_data_labels must be a bool")
        if self.grouping is not None:
            if self.grouping not in GROUPINGS:
                raise ValueError(
                    f"invalid chart grouping {self.grouping!r}; "
                    f"expected None, {', '.join(repr(g) for g in GROUPINGS)}"
                )
            if self.chart_type not in _GROUPED_TYPES:
                raise ValueError(
                    f"grouping is only valid for {sorted(_GROUPED_TYPES)} charts, "
                    f"not {self.chart_type!r}"
                )

    def _check_series_for_type(self, series) -> None:
        kind = self.chart_type
        for item in series:
            if kind == "scatter":
                if item.categories is not None:
                    raise ValueError("scatter series take x_values, not categories")
            else:
                if item.x_values is not None:
                    raise ValueError(f"x_values are only valid for scatter charts, not {kind!r}")
            if kind in _FILLED_TYPES and item.marker is not None:
                raise ValueError(f"marker is only valid for line and scatter charts, not {kind!r}")
            if kind in _FILLED_TYPES and item.line is not None:
                raise ValueError(f"line is only valid for line and scatter charts, not {kind!r}")
            if kind in _ROUND_TYPES and item.color is not None:
                raise ValueError(f"{kind} charts colour slices from the theme; series color is not supported")
            if kind == "line" and item.line is False and item.marker is None:
                raise ValueError("a line series with line=False needs a marker")

    def _check_axes(self) -> None:
        for name in ("x_axis", "y_axis"):
            value = getattr(self, name)
            if value is not None and not isinstance(value, ChartAxis):
                raise TypeError(f"{name} must be a ChartAxis or None")
        if self.chart_type in _ROUND_TYPES and (self.x_axis or self.y_axis):
            raise ValueError(f"{self.chart_type} charts have no axes; x_axis and y_axis are not valid")

    def _check_legend(self) -> None:
        legend = self.legend
        if legend is None or isinstance(legend, (bool, ChartLegend)):
            return
        raise TypeError("legend must be a ChartLegend, True, False or None")

    def _check_labels(self) -> None:
        if self.data_labels is not None and not isinstance(self.data_labels, DataLabels):
            raise TypeError("data_labels must be a DataLabels or None")
        if self.data_labels is not None and self.data_labels.position is not None:
            allowed = _LABEL_POSITIONS_BY_TYPE[self.chart_type]
            if self.data_labels.position not in allowed:
                raise ValueError(
                    f"data label position {self.data_labels.position!r} is not valid "
                    f"for {self.chart_type!r} charts"
                )

    @property
    def labels_enabled(self) -> bool:
        return self.show_data_labels or self.data_labels is not None

    @property
    def uses_extended_features(self) -> bool:
        """True when the chart uses anything beyond the original model.

        The XLSB writer uses this to reject, rather than silently drop, the
        options it does not serialize.
        """
        if self.chart_type in ("area", "barh", "doughnut", "scatter"):
            return True
        if self.grouping is not None:
            return True
        if self.x_axis is not None or self.y_axis is not None:
            return True
        if self.legend is not None:
            return True
        if self.labels_enabled or self.title_size is not None:
            return True
        return any(
            item.x_values is not None or item.color is not None
            or item.marker is not None or item.line is not None
            for item in self.series
        )
