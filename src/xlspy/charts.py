"""Chart models: what a chart shows and how it is configured.

The models only validate and store settings. DrawingML serialization lives in
``chart_parts`` so that the XLSX and XLSB writers share one XML builder.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from ._text import check_xml_text
from .refs import parse_cell_range
from .chart_edit import LEGEND_POSITIONS as LEGEND_POSITIONS, get_axis_title, get_legend_position, get_series_titles, set_axis_title, set_chart_title, set_legend_position, set_series_color, set_series_title

CHART_TYPES = frozenset({
    "area", "area3d", "bar", "bar3d", "barh", "bubble", "doughnut",
    "line", "line3d", "pie", "pie3d", "projected_pie", "radar", "scatter",
    "stock", "surface", "surface3d",
})
# The types the DrawingML writer can serialise. The rest of CHART_TYPES is
# refused when a chart is built, rather than being saved as some other chart.
SUPPORTED_CHART_TYPES = frozenset({
    "area", "bar", "barh", "doughnut", "line", "pie", "radar", "scatter",
})
# Chart.grouping values and the chart types that accept them.
GROUPINGS = {"stacked": "stacked", "percent_stacked": "percentStacked"}
_GROUPED_TYPES = frozenset({"area", "bar", "barh", "line"})
# Types without markers, lines or axes of their own.
_FILLED_TYPES = frozenset({"area", "bar", "barh", "doughnut", "pie"})
_ROUND_TYPES = frozenset({"doughnut", "pie"})
# CT_PieSer and CT_RadarSer have no trendline or error-bar element, so the
# schema rejects them for these chart types (Excel would repair the chart).
_NO_SERIES_EXTRAS_TYPES = _ROUND_TYPES | frozenset({"radar"})
# Types a combined chart can mix, and that can use a secondary value axis.
COMBINABLE_TYPES = frozenset({"area", "bar", "line"})
TRENDLINE_TYPES = {
    "linear": "linear", "exponential": "exp", "logarithmic": "log",
    "polynomial": "poly", "power": "power", "moving_average": "movingAvg",
}
ERROR_BAR_KINDS = {
    "fixed": "fixedVal", "percentage": "percentage", "std_dev": "stdDev", "std_err": "stdErr",
}
ERROR_BAR_DIRECTIONS = frozenset({"both", "plus", "minus"})

_RGB_RE = re.compile(r"[0-9A-Fa-f]{6}")
MARKER_SYMBOLS = frozenset(
    {"circle", "square", "diamond", "triangle", "x", "star", "dot", "dash", "plus", "none"}
)
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
    # Excel has no label positions for these three.
    "area": set(),
    "doughnut": set(),
    "radar": set(),
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


@dataclass(slots=True, eq=False)
class LoadedChart:
    """A chart read back from a workbook, preserved as its original parts.

    The title, the axis titles, the legend position and the titles and colours
    of the series (``set_series_title()``, ``set_series_color()``) can be
    changed; the rest of the chart is saved as it was read. ``x_axis_title`` is
    the title of the category (or X) axis and ``y_axis_title`` that of the
    value axis; ``legend_position`` is ``None`` when the chart has no legend.
    Setting ``None`` removes the title or the legend.
    """

    chart_type: str | None = None
    title: str | None = None
    series_count: int = 0
    anchor: str | None = None
    part: str | None = None
    xml: bytes = b""
    _saved_xml: bytes = field(default=b"", init=False, repr=False)

    def __post_init__(self):
        self._saved_xml = self.xml

    @property
    def edited(self) -> bool:
        """True when the chart XML changed since it was loaded."""
        return self.xml != self._saved_xml

    def _mark_saved(self) -> None:
        self._saved_xml = self.xml

    def __setattr__(self, name, value):
        if name == "title" and getattr(self, "xml", b"") and hasattr(self, "_saved_xml"):
            if value is not None and not isinstance(value, str):
                raise TypeError("chart title must be a string or None")
            object.__setattr__(self, "xml", set_chart_title(self.xml, value or None))
        object.__setattr__(self, name, value)

    @property
    def x_axis_title(self) -> str | None:
        """The title of the first axis (usually the category axis), or None."""
        return get_axis_title(self.xml, 0) if self.xml else None

    @x_axis_title.setter
    def x_axis_title(self, value: str | None) -> None:
        self._set_axis_title(0, value)

    @property
    def y_axis_title(self) -> str | None:
        """The title of the second axis (usually the value axis), or None."""
        return get_axis_title(self.xml, 1) if self.xml else None

    @y_axis_title.setter
    def y_axis_title(self, value: str | None) -> None:
        self._set_axis_title(1, value)

    def _set_axis_title(self, index: int, value: str | None) -> None:
        if value is not None and not isinstance(value, str):
            raise TypeError("axis title must be a string or None")
        self.xml = set_axis_title(self.xml, index, value or None)

    @property
    def series_titles(self) -> list[str | None]:
        """The literal title of each series (``None`` when it reads a cell)."""
        return get_series_titles(self.xml) if self.xml else []

    def set_series_title(self, index: int, text: str) -> None:
        """Give series ``index`` a literal title."""
        if not isinstance(text, str):
            raise TypeError("series title must be a string")
        self.xml = set_series_title(self.xml, index, text)

    def set_series_color(self, index: int, color: str) -> None:
        """Fill series ``index`` (or draw its line) in ``color``, a ``'#RRGGBB'`` string."""
        rgb = _normalize_rgb(color, "series color")
        if rgb is None:
            # Writing "" produced <a:srgbClr val=""/>, which is not a colour.
            raise TypeError("series color must be a '#RRGGBB' string, not None")
        self.xml = set_series_color(self.xml, index, rgb)

    @property
    def legend_position(self) -> str | None:
        """The legend position (``"r"``, ``"b"`` and so on), or None without a legend."""
        return get_legend_position(self.xml) if self.xml else None

    @legend_position.setter
    def legend_position(self, value: str | None) -> None:
        self.xml = set_legend_position(self.xml, value)


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
        if self.title is not None:
            check_xml_text(self.title, "axis title")
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
class Trendline:
    """A trendline fitted to a series.

    ``order`` (2 to 6) is for ``"polynomial"`` and ``period`` (at least 2) for
    ``"moving_average"``. ``forward`` and ``backward`` extend the line by that
    many categories (X units on a scatter chart).
    """

    type: str = "linear"
    order: int | None = None
    period: int | None = None
    forward: float | None = None
    backward: float | None = None
    display_equation: bool = False
    display_r_squared: bool = False
    name: str | None = None
    color: str | None = None

    def __post_init__(self):
        if self.type not in TRENDLINE_TYPES:
            raise ValueError(
                f"invalid trendline type {self.type!r}; expected one of {sorted(TRENDLINE_TYPES)}"
            )
        if self.order is not None:
            if self.type != "polynomial":
                raise ValueError("trendline order is only valid for a polynomial trendline")
            if isinstance(self.order, bool) or not isinstance(self.order, int) or not 2 <= self.order <= 6:
                raise ValueError("trendline order must be an int between 2 and 6")
        elif self.type == "polynomial":
            object.__setattr__(self, "order", 2)
        if self.period is not None:
            if self.type != "moving_average":
                raise ValueError("trendline period is only valid for a moving average")
            if isinstance(self.period, bool) or not isinstance(self.period, int) or self.period < 2:
                raise ValueError("trendline period must be an int of at least 2")
        elif self.type == "moving_average":
            object.__setattr__(self, "period", 2)
        object.__setattr__(self, "forward", _check_number(self.forward, "trendline forward", minimum=0))
        object.__setattr__(self, "backward", _check_number(self.backward, "trendline backward", minimum=0))
        if not isinstance(self.display_equation, bool) or not isinstance(self.display_r_squared, bool):
            raise TypeError("trendline display_equation and display_r_squared must be bools")
        if self.name is not None and not isinstance(self.name, str):
            raise TypeError("trendline name must be a string or None")
        if self.name is not None:
            check_xml_text(self.name, "trendline name")
        object.__setattr__(self, "color", _normalize_rgb(self.color, "trendline color"))


@dataclass(frozen=True, slots=True)
class ErrorBars:
    """Error bars on a series.

    ``kind`` is ``"fixed"`` (``value`` units), ``"percentage"`` (``value`` percent),
    ``"std_dev"`` (``value`` standard deviations) or ``"std_err"`` (no value).
    ``direction`` picks ``"both"``, ``"plus"`` or ``"minus"`` bars; ``axis`` is
    ``"y"``, or ``"x"`` on a scatter series.
    """

    kind: str = "fixed"
    value: float | None = None
    direction: str = "both"
    axis: str = "y"
    end_cap: bool = True
    color: str | None = None

    def __post_init__(self):
        if self.kind not in ERROR_BAR_KINDS:
            raise ValueError(
                f"invalid error bar kind {self.kind!r}; expected one of {sorted(ERROR_BAR_KINDS)}"
            )
        if self.kind == "std_err":
            if self.value is not None:
                raise ValueError("standard error bars take no value")
        else:
            if self.value is None and self.kind == "std_dev":
                object.__setattr__(self, "value", 1.0)
            elif self.value is None:
                raise ValueError(f"{self.kind} error bars need a value")
            object.__setattr__(self, "value", _check_number(self.value, "error bar value", minimum=0))
        if self.direction not in ERROR_BAR_DIRECTIONS:
            raise ValueError(f"invalid error bar direction {self.direction!r}; expected both, plus or minus")
        if self.axis not in ("x", "y"):
            raise ValueError("error bar axis must be 'x' or 'y'")
        if not isinstance(self.end_cap, bool):
            raise TypeError("error bar end_cap must be a bool")
        object.__setattr__(self, "color", _normalize_rgb(self.color, "error bar color"))


@dataclass(frozen=True, slots=True)
class ChartSeries:
    """One data series of a ``Chart``.

    ``categories`` labels the points of bar, line and pie series. ``x_values``
    gives the numeric X positions of a scatter series. ``color`` is the fill of
    bars and pie slices, and the line or marker colour of line and scatter
    series. ``marker`` and ``line`` apply to line and scatter series only.

    In a combined chart ``chart_type`` (``"bar"``, ``"line"`` or ``"area"``)
    draws this series as another type than the chart's, and ``secondary_axis``
    reads it against a second value axis on the right.
    """

    values: str
    title: str | None = None
    title_cell: str | None = None
    categories: str | None = None
    x_values: str | None = None
    color: str | None = None
    marker: ChartMarker | None = None
    line: bool | None = None
    chart_type: str | None = None
    secondary_axis: bool = False
    trendline: Trendline | None = None
    error_bars: ErrorBars | None = None

    def __post_init__(self):
        if not isinstance(self.values, str):
            raise TypeError("series values must be an A1 range such as 'B2:B5'")
        _validate_chart_range(self.values, "series values")
        if self.title is not None and not isinstance(self.title, str):
            raise TypeError("series title must be a string or None")
        if self.title is not None:
            check_xml_text(self.title, "series title")
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
        if self.chart_type is not None and self.chart_type not in COMBINABLE_TYPES:
            raise ValueError(
                f"a series can be drawn as {sorted(COMBINABLE_TYPES)}, not {self.chart_type!r}"
            )
        if not isinstance(self.secondary_axis, bool):
            raise TypeError("series secondary_axis must be a bool")
        if self.trendline is not None and not isinstance(self.trendline, Trendline):
            raise TypeError("series trendline must be a Trendline or None")
        if self.error_bars is not None and not isinstance(self.error_bars, ErrorBars):
            raise TypeError("series error_bars must be an ErrorBars or None")


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

    A chart mixes types when some series set ``ChartSeries.chart_type`` (bar,
    line and area only); series with ``secondary_axis`` use ``y2_axis``, the
    value axis on the right.
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
    y2_axis: ChartAxis | None = None
    legend: ChartLegend | bool | None = None
    data_labels: DataLabels | None = None
    show_data_labels: bool = False
    title_size: float | None = None
    grouping: str | None = None
    # Formatting. None keeps the chart's default; colours are '#RRGGBB'.
    gap_width: int | None = None
    hole_size: int | None = None
    fill: str | None = None
    border: str | None = None
    plot_fill: str | None = None
    font_size: float | None = None
    font_color: str | None = None

    def __post_init__(self):
        if self.chart_type not in CHART_TYPES:
            raise ValueError(
                f"invalid chart type {self.chart_type!r}; "
                f"expected one of {sorted(CHART_TYPES)}"
            )
        if self.chart_type not in SUPPORTED_CHART_TYPES:
            raise ValueError(
                f"chart type {self.chart_type!r} is not supported yet; "
                f"use one of {sorted(SUPPORTED_CHART_TYPES)}"
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
        if self.title is not None:
            check_xml_text(self.title, "chart title")
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
        self._check_formatting()

    def _check_formatting(self) -> None:
        if self.gap_width is not None:
            if isinstance(self.gap_width, bool) or not isinstance(self.gap_width, int) \
                    or not 0 <= self.gap_width <= 500:
                raise ValueError("chart gap_width must be an int from 0 to 500")
            if not {"bar", "barh"} & set(self.series_types):
                raise ValueError("gap_width is only valid for bar and horizontal bar charts")
        if self.hole_size is not None:
            if isinstance(self.hole_size, bool) or not isinstance(self.hole_size, int) \
                    or not 10 <= self.hole_size <= 90:
                raise ValueError("chart hole_size must be an int from 10 to 90")
            if self.chart_type != "doughnut":
                raise ValueError("hole_size is only valid for doughnut charts")
        for name in ("fill", "border", "plot_fill", "font_color"):
            object.__setattr__(self, name, _normalize_rgb(getattr(self, name), f"chart {name}"))
        object.__setattr__(self, "font_size", _check_number(
            self.font_size, "chart font_size", minimum=1, maximum=400,
        ))

    def _check_series_for_type(self, series) -> None:
        for item in series:
            kind = item.chart_type or self.chart_type
            if item.chart_type is not None and self.chart_type not in COMBINABLE_TYPES:
                raise ValueError(
                    f"series chart_type needs a bar, line or area chart, not {self.chart_type!r}"
                )
            if item.secondary_axis and kind not in COMBINABLE_TYPES:
                raise ValueError(f"a secondary axis is only valid for bar, line and area series, not {kind!r}")
            if item.trendline is not None and kind in _NO_SERIES_EXTRAS_TYPES:
                raise ValueError(f"trendlines are not valid for {kind!r} charts")
            if item.error_bars is not None:
                if kind in _NO_SERIES_EXTRAS_TYPES:
                    raise ValueError(f"error bars are not valid for {kind!r} charts")
                if item.error_bars.axis == "x" and kind != "scatter":
                    raise ValueError("x error bars are only valid for scatter series")
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
        for name in ("x_axis", "y_axis", "y2_axis"):
            value = getattr(self, name)
            if value is not None and not isinstance(value, ChartAxis):
                raise TypeError(f"{name} must be a ChartAxis or None")
        if self.y2_axis is not None and not any(item.secondary_axis for item in self.series):
            raise ValueError("y2_axis needs a series with secondary_axis=True")
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
            for kind in self.series_types:
                if self.data_labels.position not in _LABEL_POSITIONS_BY_TYPE[kind]:
                    raise ValueError(
                        f"data label position {self.data_labels.position!r} is not valid "
                        f"for {kind!r} charts"
                    )

    @property
    def series_types(self) -> list[str]:
        """The distinct types the series are drawn as, in series order."""
        return list(dict.fromkeys(item.chart_type or self.chart_type for item in self.series))

    @property
    def labels_enabled(self) -> bool:
        """True when the chart shows data labels."""
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
        if any(value is not None for value in (
            self.gap_width, self.hole_size, self.fill, self.border,
            self.plot_fill, self.font_size, self.font_color,
        )):
            return True
        if self.x_axis is not None or self.y_axis is not None:
            return True
        if self.legend is not None:
            return True
        if self.labels_enabled or self.title_size is not None or self.y2_axis is not None:
            return True
        return any(
            item.x_values is not None or item.color is not None
            or item.marker is not None or item.line is not None
            or item.chart_type is not None or item.secondary_axis
            or item.trendline is not None or item.error_bars is not None
            for item in self.series
        )
