"""Shared DrawingML chart/drawing builders for the XLSX and XLSB writers.

Both package formats store drawings and charts as the same XML parts
(``xl/drawings/drawingN.xml``, ``xl/charts/chartN.xml``); only the sheet
records that link to them differ (``<drawing>`` element vs. ``BrtDrawing``
record). Keeping the XML generation in one place keeps the writers in sync.

The XLSB path (``xlsb=True``) additionally mirrors what real Excel files
contain: cached values (``strCache``/``numCache``) inside every reference,
text categories via ``strRef``, an explicit ``varyColors`` flag, and numeric
chart styles spelled out through the ``c14:style`` extension. Excel's XLSB
loader evidently depends on the caches for series structure; without them
single-series charts over vertical ranges render transposed (one series per
row, category names as legend), and it does not apply a bare legacy
``style`` element the way the XLSX loader does. The XLSB path also pins the
classic Office 2007 accent palette and 18 pt titles: the loader's modern
style engine would otherwise paint series with the theme accents and use
14 pt titles, visibly unlike the legacy-style rendering of the XLSX loader.
The XLSX path stays byte-identical to the openpyxl oracle and is
intentionally left cache-free.
"""

from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal
from xml.sax import saxutils

from .charts import (
    GROUPINGS,
    LABEL_POSITIONS,
    LEGEND_POSITIONS,
    ChartAxis,
    ChartMarker,
    DataLabels,
)
from .refs import parse_cell_range, quote_sheet_name  # noqa: F401

_CHART_AXES_XML = (
    "<catAx><axId val=\"10\"/><scaling><orientation val=\"minMax\"/></scaling>"
    "<axPos val=\"l\"/><majorTickMark val=\"none\"/><minorTickMark val=\"none\"/>"
    "<crossAx val=\"100\"/><lblOffset val=\"100\"/></catAx>"
    "<valAx><axId val=\"100\"/><scaling><orientation val=\"minMax\"/></scaling>"
    "<axPos val=\"l\"/><majorGridlines/><majorTickMark val=\"none\"/>"
    "<minorTickMark val=\"none\"/><crossAx val=\"10\"/></valAx>"
)

# The XLSB loader in Excel does not apply the same attribute defaults as the
# XLSX one: without explicit delete/axPos/tickLblPos the axis (and therefore
# its tick labels) is not drawn. Real Excel-generated XLSB charts spell these
# out, so the XLSB path mirrors them. The XLSX path stays byte-identical to
# the openpyxl oracle, which renders correctly without them.
_XLSB_CHART_AXES_XML = (
    "<catAx><axId val=\"10\"/><scaling><orientation val=\"minMax\"/></scaling>"
    "<delete val=\"0\"/><axPos val=\"b\"/>"
    "<majorTickMark val=\"none\"/><minorTickMark val=\"none\"/>"
    "<tickLblPos val=\"nextTo\"/><crossAx val=\"100\"/><lblOffset val=\"100\"/></catAx>"
    "<valAx><axId val=\"100\"/><scaling><orientation val=\"minMax\"/></scaling>"
    "<delete val=\"0\"/><axPos val=\"l\"/><majorGridlines/>"
    "<majorTickMark val=\"none\"/><minorTickMark val=\"none\"/>"
    "<tickLblPos val=\"nextTo\"/><crossAx val=\"10\"/></valAx>"
)

# [MS-ODRAWXML] 2.7.3.5: the modern c14 style value is the legacy value plus
# 100 (Excel writes style 10 as 110 and the default style 2 as 102). A bare
# legacy ``style`` element is enough for the XLSX loader, but the XLSB loader
# renders the chart with the default style unless the extension is spelled
# out, so Excel-written XLSB files use the mc:AlternateContent form with the
# legacy element as fallback. The XLSX path keeps the plain element, matching
# the openpyxl oracle.
_MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_C14_NS = "http://schemas.microsoft.com/office/drawing/2007/8/2/chart"

# Excel's XLSX loader draws charts that carry a legacy ``style`` element with
# the classic Office 2007 accent palette and 18 pt titles. The XLSB loader
# only knows the modern style engine: its colours come from the theme (the
# Office 2013 accents of our theme) and titles default to 14 pt. Pinning the
# classic palette and title size in the XLSB path keeps a chart visually
# identical in both formats, whatever style engine Excel applies on top.
_LEGACY_ACCENTS = ("4F81BD", "C0504D", "9BBB59", "8064A2", "4BACC6", "F79646")


def column_letters() -> list[str]:
    """Column labels ``A`` … ``XFD`` (shared with the writers)."""
    letters = []
    for i in range(26):
        letters.append(chr(65 + i))
    for i in range(26):
        for j in range(26):
            letters.append(chr(65 + i) + chr(65 + j))
    for i in range(24):
        for j in range(26):
            for k in range(26):
                if i == 23 and j > 24:
                    break
                letters.append(chr(65 + i) + chr(65 + j) + chr(65 + k))
    return letters


def chart_ref(letters: list[str], sheet_name: str, ref: str, *,
              absolute: bool = True) -> str:
    first_col, first_row, last_col, last_row = parse_cell_range(ref)
    if absolute:
        start = f"${letters[first_col]}${first_row + 1}"
        end = f"${letters[last_col]}${last_row + 1}"
        range_ref = f"{start}:{end}"
    elif (first_col, first_row) == (last_col, last_row):
        range_ref = f"{letters[first_col]}{first_row + 1}"
    else:
        range_ref = (
            f"{letters[first_col]}{first_row + 1}"
            f":{letters[last_col]}{last_row + 1}"
        )
    return f"{quote_sheet_name(sheet_name)}!{range_ref}"


def chart_cell_spans(charts) -> dict[int, tuple[int, int]]:
    """Zero-based column -> (first row, last row) that the charts' ranges read.

    The XLSB writer keeps only those cells while the rows stream past, in one
    list per column (see ``new_cell_store``), so the sheet is never held.
    """
    spans: dict[int, tuple[int, int]] = {}
    for chart in charts:
        for series in chart.series:
            for ref in (series.values, series.categories, series.x_values,
                        series.title_cell):
                if ref is None:
                    continue
                first_col, first_row, last_col, last_row = parse_cell_range(ref)
                for col_idx in range(first_col, last_col + 1):
                    low, high = spans.get(col_idx, (first_row, last_row))
                    spans[col_idx] = (min(low, first_row), max(high, last_row))
    return spans


def new_cell_store(spans: dict[int, tuple[int, int]]) -> dict[int, tuple[int, list]]:
    """Empty capture store: column -> (first row, one slot per row in the span)."""
    return {
        col_idx: (first_row, [None] * (last_row - first_row + 1))
        for col_idx, (first_row, last_row) in spans.items()
    }


def _range_values(cells: dict | None, ref: str) -> list:
    """Cell values covered by ``ref`` in row-major order.

    ``cells`` comes from ``new_cell_store``; cells outside the captured span
    read as ``None``.
    """
    if cells is None:
        raise ValueError("chart caches need the sheet cells; pass cells=...")
    first_col, first_row, last_col, last_row = parse_cell_range(ref)
    values = []
    for row_idx in range(first_row, last_row + 1):
        for col_idx in range(first_col, last_col + 1):
            value = None
            column = cells.get(col_idx)
            if column is not None:
                start, slots = column
                if start <= row_idx < start + len(slots):
                    value = slots[row_idx - start]
            values.append(value)
    return values


def _chart_number_text(value) -> str:
    """Shortest plain-text literal for a cached chart number."""
    from .excel_dates import datetime_to_serial

    if isinstance(value, bool):
        return "1" if value else "0"
    if isinstance(value, Decimal):
        value = float(value)
    if isinstance(value, (datetime, date)):
        value = datetime_to_serial(value)
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def _cached_value(raw):
    """What a chart cache shows for a sheet cell.

    A formula shows its cached result (``None`` when it has none), a hyperlink
    its label; ``(value, format)`` and ``(value, Style)`` show the value.
    """
    from .formats import _unwrap_cell
    from .formulas import Formula
    from .hyperlinks import Hyperlink
    from .rich_text import CellRichText

    value = _unwrap_cell(raw)
    if isinstance(value, Formula):
        return value.value
    if isinstance(value, Hyperlink):
        return value.display_text
    if isinstance(value, CellRichText):
        # Rich text is text: a strRef point, never a numCache value.
        return str(value)
    return value


def _chart_text(value) -> str:
    """Plain-text form of a cached chart string."""
    value = _cached_value(value)
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    return str(value)


def _str_cache_xml(values: list) -> str:
    points = "".join(
        f"<pt idx=\"{i}\"><v>{saxutils.escape(_chart_text(v))}</v></pt>"
        for i, v in enumerate(values)
    )
    return f"<strCache><ptCount val=\"{len(values)}\"/>{points}</strCache>"


def _num_cache_xml(values: list) -> str:
    points = []
    for i, raw in enumerate(values):
        value = _cached_value(raw)
        if value is None:
            # Excel stores an empty cell as a missing point, drawn as a gap.
            continue
        if isinstance(value, str):
            try:
                text = _chart_number_text(float(value))
            except ValueError:
                text = "0"
        else:
            text = _chart_number_text(value)
        points.append(f"<pt idx=\"{i}\"><v>{text}</v></pt>")
    joined = "".join(points)
    return (
        "<numCache><formatCode>General</formatCode>"
        f"<ptCount val=\"{len(values)}\"/>{joined}</numCache>"
    )


def _is_numeric_range(values: list) -> bool:
    seen = False
    for raw in values:
        value = _cached_value(raw)
        if value is None:
            continue
        if isinstance(value, (str, bool)):
            return False
        seen = True
    return seen


def _attr(value: str) -> str:
    return saxutils.escape(value, {'"': "&quot;"})


def _solid(color: str) -> str:
    return f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'


def _ref_xml(tag: str, letters, sheet_name: str, ref: str, cache: str = "") -> str:
    formula = chart_ref(letters, sheet_name, ref)
    return f"<{tag}><numRef><f>{formula}</f>{cache}</numRef></{tag}>"


def _series_title_xml(letters, series, sheet_name, cells, xlsb: bool) -> str:
    if series.title_cell is not None:
        formula = chart_ref(
            letters, sheet_name, series.title_cell, absolute=False,
        )
        if xlsb:
            title_value = _range_values(cells, series.title_cell)[0]
            return (
                "<tx><strRef><f>"
                f"{formula}"
                "</f>"
                f"{_str_cache_xml([title_value])}</strRef></tx>"
            )
        return (
            "<tx><strRef><f>"
            f"{formula}"
            "</f></strRef></tx>"
        )
    if series.title is not None:
        return (
            "<tx><v>"
            f"{saxutils.escape(series.title)}"
            "</v></tx>"
        )
    return ""


def _categories_xml(letters, series, sheet_name, cells, xlsb: bool) -> str:
    if series.categories is None:
        return ""
    if not xlsb:
        return _ref_xml("cat", letters, sheet_name, series.categories)
    return _cached_ref_xml("cat", letters, sheet_name, series.categories, cells)


def _cached_ref_xml(tag: str, letters, sheet_name: str, ref: str, cells) -> str:
    """A reference with its cache, as ``numRef`` for numbers, else ``strRef``."""
    formula = chart_ref(letters, sheet_name, ref)
    values = _range_values(cells, ref)
    if _is_numeric_range(values):
        return f"<{tag}><numRef><f>{formula}</f>{_num_cache_xml(values)}</numRef></{tag}>"
    return f"<{tag}><strRef><f>{formula}</f>{_str_cache_xml(values)}</strRef></{tag}>"


def _line_sp_xml(color, line_on: bool) -> str:
    if not line_on:
        return '<spPr><a:ln w="19050"><a:noFill/></a:ln></spPr>'
    if color is None:
        return '<spPr><a:ln><a:prstDash val="solid"/></a:ln></spPr>'
    return (
        f'<spPr><a:ln w="28575">{_solid(color)}'
        '<a:prstDash val="solid"/></a:ln></spPr>'
    )


def _marker_xml(marker, color) -> str:
    if marker.symbol == "none":
        return '<marker><symbol val="none"/></marker>'
    sp_pr = ""
    if color is not None:
        sp_pr = f'<spPr>{_solid(color)}<a:ln w="9525">{_solid(color)}</a:ln></spPr>'
    return (
        f'<marker><symbol val="{marker.symbol}"/><size val="{marker.size}"/>'
        f"{sp_pr}</marker>"
    )


def _series_style_xml(series, chart_type: str, index: int, cells, xlsb: bool):
    """Return ``(spPr, dPt, marker)`` XML for one series."""
    if xlsb:
        # Every colour is explicit: the XLSB loader would otherwise paint
        # series with the theme accents (see _LEGACY_ACCENTS).
        color = series.color or _LEGACY_ACCENTS[index % len(_LEGACY_ACCENTS)]
        line_on = series.line is True if chart_type == "scatter" else series.line is not False
        if chart_type in ("line", "scatter"):
            if line_on:
                sp_pr = (
                    "<spPr><a:ln>"
                    f"<a:solidFill><a:srgbClr val=\"{color}\"/></a:solidFill>"
                    "<a:prstDash val=\"solid\"/></a:ln></spPr>"
                )
            else:
                sp_pr = '<spPr><a:ln w="19050"><a:noFill/></a:ln></spPr>'
        else:
            sp_pr = (
                f"<spPr><a:solidFill><a:srgbClr val=\"{color}\"/></a:solidFill>"
                "<a:ln><a:prstDash val=\"solid\"/></a:ln></spPr>"
            )
        points = ""
        if chart_type in ("pie", "doughnut"):
            value_count = len(_range_values(cells, series.values))
            points = "".join(
                f"<dPt><idx val=\"{point}\"/><spPr><a:solidFill>"
                f"<a:srgbClr val=\"{_LEGACY_ACCENTS[point % len(_LEGACY_ACCENTS)]}\"/>"
                "</a:solidFill><a:ln><a:prstDash val=\"solid\"/></a:ln></spPr></dPt>"
                for point in range(value_count)
            )
        marker = ""
        if chart_type == "scatter":
            marker = _marker_xml(series.marker or ChartMarker(), color)
        elif chart_type == "line":
            if series.marker is None:
                marker = (
                    "<marker><symbol val=\"none\"/>"
                    "<spPr><a:ln><a:prstDash val=\"solid\"/></a:ln></spPr></marker>"
                )
            else:
                marker = _marker_xml(series.marker, color)
        return sp_pr, points, marker

    color = series.color
    points = ""
    marker = ""
    if chart_type in ("area", "bar", "barh"):
        if color is None:
            sp_pr = '<spPr><a:ln><a:prstDash val="solid"/></a:ln></spPr>'
        else:
            sp_pr = (
                f'<spPr>{_solid(color)}'
                '<a:ln><a:prstDash val="solid"/></a:ln></spPr>'
            )
    elif chart_type == "line":
        sp_pr = _line_sp_xml(color, series.line is not False)
        if series.marker is None:
            marker = (
                "<marker><symbol val=\"none\"/>"
                "<spPr><a:ln><a:prstDash val=\"solid\"/></a:ln></spPr></marker>"
            )
        else:
            marker = _marker_xml(series.marker, color)
    elif chart_type == "scatter":
        sp_pr = _line_sp_xml(color, series.line is True)
        marker = _marker_xml(series.marker or ChartMarker(), color)
    else:
        sp_pr = "<spPr><a:ln><a:prstDash val=\"solid\"/></a:ln></spPr>"
    return sp_pr, points, marker


def chart_series_xml(letters: list[str], series, sheet_name: str,
                     index: int, chart_type: str, cells=None,
                     xlsb: bool = False) -> str:
    title = _series_title_xml(letters, series, sheet_name, cells, xlsb)
    sp_pr, points, marker = _series_style_xml(series, chart_type, index, cells, xlsb)
    if chart_type == "scatter":
        x_xml = ""
        if series.x_values is not None:
            if xlsb:
                x_xml = _cached_ref_xml("xVal", letters, sheet_name, series.x_values, cells)
            else:
                x_xml = _ref_xml("xVal", letters, sheet_name, series.x_values)
        if xlsb:
            y_xml = _ref_xml(
                "yVal", letters, sheet_name, series.values,
                cache=_num_cache_xml(_range_values(cells, series.values)),
            )
        else:
            y_xml = _ref_xml("yVal", letters, sheet_name, series.values)
        return (
            f"<ser><idx val=\"{index}\"/><order val=\"{index}\"/>{title}"
            f"{sp_pr}{marker}{x_xml}{y_xml}<smooth val=\"0\"/></ser>"
        )
    categories = _categories_xml(letters, series, sheet_name, cells, xlsb)
    smooth = ""
    if xlsb:
        values_xml = _ref_xml(
            "val", letters, sheet_name, series.values,
            cache=_num_cache_xml(_range_values(cells, series.values)),
        )
        if chart_type == "line":
            # The XLSB loader smooths a line series that does not say otherwise.
            smooth = "<smooth val=\"0\"/>"
    else:
        values_xml = _ref_xml("val", letters, sheet_name, series.values)
    return (
        f"<ser><idx val=\"{index}\"/><order val=\"{index}\"/>{title}"
        f"{sp_pr}{points}{marker}{categories}"
        f"{values_xml}{smooth}</ser>"
    )


def _dlbls_xml(labels) -> str:
    parts = ["<dLbls>"]
    if labels.number_format is not None:
        parts.append(
            f'<numFmt formatCode="{_attr(labels.number_format)}" sourceLinked="0"/>'
        )
    if labels.position is not None:
        parts.append(f'<dLblPos val="{LABEL_POSITIONS[labels.position]}"/>')
    parts.append(
        '<showLegendKey val="0"/>'
        '<showVal val="1"/>'
        f'<showCatName val="{int(labels.show_category)}"/>'
        f'<showSerName val="{int(labels.show_series_name)}"/>'
        '<showPercent val="0"/><showBubbleSize val="0"/>'
        "</dLbls>"
    )
    return "".join(parts)


def _axis_title_xml(text: str) -> str:
    return (
        "<title><tx><rich><a:bodyPr/><a:p><a:pPr><a:defRPr/></a:pPr>"
        f"<a:r><a:t>{saxutils.escape(text)}</a:t></a:r></a:p></rich></tx>"
        '<overlay val="0"/></title>'
    )


def _axis_xml(kind: str, ax_id: int, cross_id: int, position: str, axis,
              default_gridlines: bool, cross_between: str = "between",
              xlsb: bool = False) -> str:
    """One ``catAx`` or ``valAx`` element for the extended chart model.

    The XLSB loader draws an axis only with an explicit ``delete val="0"``
    (see ``_XLSB_CHART_AXES_XML``).
    """
    axis = axis if axis is not None else ChartAxis()
    gridlines = default_gridlines if axis.gridlines is None else axis.gridlines
    scaling = f'<orientation val="{"maxMin" if axis.reverse else "minMax"}"/>'
    if kind == "valAx":
        if axis.maximum is not None:
            scaling += f'<max val="{_chart_number_text(axis.maximum)}"/>'
        if axis.minimum is not None:
            scaling += f'<min val="{_chart_number_text(axis.minimum)}"/>'
    parts = [f'<{kind}><axId val="{ax_id}"/><scaling>{scaling}</scaling>']
    if not axis.visible:
        parts.append('<delete val="1"/>')
    elif xlsb:
        parts.append('<delete val="0"/>')
    parts.append(f'<axPos val="{position}"/>')
    if gridlines:
        parts.append("<majorGridlines/>")
    if axis.title is not None:
        parts.append(_axis_title_xml(axis.title))
    if axis.number_format is not None:
        parts.append(f'<numFmt formatCode="{_attr(axis.number_format)}" sourceLinked="0"/>')
    parts.append(
        '<majorTickMark val="none"/><minorTickMark val="none"/>'
        '<tickLblPos val="nextTo"/>'
        f'<crossAx val="{cross_id}"/>'
    )
    if kind == "catAx":
        parts.append('<lblOffset val="100"/>')
    else:
        parts.append(f'<crossBetween val="{cross_between}"/>')
        if axis.major_unit is not None:
            parts.append(f'<majorUnit val="{_chart_number_text(axis.major_unit)}"/>')
    parts.append(f"</{kind}>")
    return "".join(parts)


def _axes_xml(chart, xlsb: bool) -> str:
    """Axis elements for a bar, horizontal bar, line or scatter chart."""
    kind = chart.chart_type
    customised = chart.x_axis is not None or chart.y_axis is not None
    if kind in ("bar", "line") and not customised:
        return _XLSB_CHART_AXES_XML if xlsb else _CHART_AXES_XML
    if kind == "area":
        # Excel's own area charts start the plot at the value axis.
        return (
            _axis_xml("catAx", 10, 100, "b", chart.x_axis, False, xlsb=xlsb)
            + _axis_xml("valAx", 100, 10, "l", chart.y_axis, True, "midCat", xlsb)
        )
    if kind == "scatter":
        return (
            _axis_xml("valAx", 10, 100, "b", chart.x_axis, False, "midCat", xlsb)
            + _axis_xml("valAx", 100, 10, "l", chart.y_axis, True, "midCat", xlsb)
        )
    if kind == "barh":
        return (
            _axis_xml("catAx", 10, 100, "l", chart.x_axis, False, xlsb=xlsb)
            + _axis_xml("valAx", 100, 10, "b", chart.y_axis, True, xlsb=xlsb)
        )
    return (
        _axis_xml("catAx", 10, 100, "b", chart.x_axis, False, xlsb=xlsb)
        + _axis_xml("valAx", 100, 10, "l", chart.y_axis, True, xlsb=xlsb)
    )


def _legend_xml(legend) -> str:
    if legend is False:
        return ""
    if legend is None or legend is True:
        return '<legend><legendPos val="r"/><overlay val="0"/></legend>'
    # An absent overlay element means the legend covers the plot in Excel.
    overlay = '<overlay val="1"/>' if legend.overlay else '<overlay val="0"/>'
    return (
        f'<legend><legendPos val="{LEGEND_POSITIONS[legend.position]}"/>'
        f"{overlay}</legend>"
    )


def _title_xml(chart, xlsb: bool) -> str:
    title_text = saxutils.escape(chart.title)
    if xlsb:
        size = "1800" if chart.title_size is None else _chart_number_text(chart.title_size * 100)
        return (
            "<title><tx><rich><a:bodyPr/><a:p>"
            f"<a:pPr><a:defRPr sz=\"{size}\" b=\"1\"/></a:pPr>"
            f"<a:r><a:rPr lang=\"en-US\" sz=\"{size}\" b=\"1\"/>"
            f"<a:t>{title_text}</a:t></a:r>"
            "</a:p></rich></tx><overlay val=\"0\"/></title>"
        )
    if chart.title_size is None:
        return (
            "<title><tx><rich><a:bodyPr/><a:p><a:pPr><a:defRPr/></a:pPr>"
            f"<a:r><a:t>{title_text}</a:t></a:r></a:p></rich></tx>"
            '<overlay val="0"/></title>'
        )
    size = _chart_number_text(chart.title_size * 100)
    return (
        "<title><tx><rich><a:bodyPr/><a:p>"
        f'<a:pPr><a:defRPr sz="{size}"/></a:pPr>'
        f'<a:r><a:rPr lang="en-US" sz="{size}"/><a:t>{title_text}</a:t></a:r>'
        '</a:p></rich></tx><overlay val="0"/></title>'
    )


def chart_xml(letters: list[str], chart, sheet_name: str, cells=None,
              xlsb: bool = False) -> str:
    parts = []
    if chart.style is not None:
        if xlsb:
            parts.append(
                f"<mc:AlternateContent xmlns:mc=\"{_MC_NS}\">"
                f"<mc:Choice Requires=\"c14\" xmlns:c14=\"{_C14_NS}\">"
                f"<c14:style val=\"{chart.style + 100}\"/></mc:Choice>"
                f"<mc:Fallback><style val=\"{chart.style}\"/></mc:Fallback>"
                "</mc:AlternateContent>"
            )
        else:
            parts.append(f"<style val=\"{chart.style}\"/>")
    parts.append("<chart>")
    if chart.title is not None:
        parts.append(_title_xml(chart, xlsb))
    else:
        # The automatic title: the series name of a single-series chart. Excel
        # saves it as an empty title; with autoTitleDeleted alone an XLSB
        # opened while another workbook is open usually loses it.
        if len(chart.series) == 1:
            parts.append("<title><overlay val=\"0\"/></title>")
        parts.append("<autoTitleDeleted val=\"0\"/>")
    series_xml = "".join(
        chart_series_xml(
            letters, item, sheet_name, index, chart.chart_type,
            cells=cells, xlsb=xlsb,
        )
        for index, item in enumerate(chart.series)
    )
    vary = "<varyColors val=\"0\"/>" if xlsb else ""
    labels = ""
    if chart.labels_enabled:
        labels = _dlbls_xml(chart.data_labels or DataLabels())
    parts.append("<plotArea>")
    grouping = GROUPINGS.get(chart.grouping)
    if chart.chart_type in ("bar", "barh"):
        bar_dir = "col" if chart.chart_type == "bar" else "bar"
        # Stacked bars sit on top of each other: overlap 100%.
        overlap = "<overlap val=\"100\"/>" if grouping else ""
        parts.append(
            f"<barChart><barDir val=\"{bar_dir}\"/>"
            f"<grouping val=\"{grouping or 'clustered'}\"/>"
            f"{vary}{series_xml}{labels}"
            f"<gapWidth val=\"150\"/>{overlap}<axId val=\"10\"/><axId val=\"100\"/></barChart>"
            f"{_axes_xml(chart, xlsb)}"
        )
    elif chart.chart_type in ("line", "area"):
        tag = f"{chart.chart_type}Chart"
        parts.append(
            f"<{tag}><grouping val=\"{grouping or 'standard'}\"/>"
            f"{vary}{series_xml}{labels}"
            f"<axId val=\"10\"/><axId val=\"100\"/></{tag}>"
            f"{_axes_xml(chart, xlsb)}"
        )
    elif chart.chart_type == "scatter":
        parts.append(
            "<scatterChart><scatterStyle val=\"lineMarker\"/>"
            f"<varyColors val=\"0\"/>{series_xml}{labels}"
            "<axId val=\"10\"/><axId val=\"100\"/></scatterChart>"
            f"{_axes_xml(chart, xlsb)}"
        )
    elif chart.chart_type == "doughnut":
        parts.append(
            f"<doughnutChart><varyColors val=\"1\"/>{series_xml}{labels}"
            "<firstSliceAng val=\"0\"/><holeSize val=\"50\"/></doughnutChart>"
        )
    else:
        parts.append(
            f"<pieChart><varyColors val=\"1\"/>{series_xml}{labels}"
            "<firstSliceAng val=\"0\"/></pieChart>"
        )
    parts.append("</plotArea>")
    parts.append(
        f"{_legend_xml(chart.legend)}"
        "<plotVisOnly val=\"1\"/><dispBlanksAs val=\"gap\"/>"
    )
    parts.append("</chart>")
    return (
        "<chartSpace "
        "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" "
        "xmlns=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">"
        f"{''.join(parts)}</chartSpace>"
    )


EMU_PER_PIXEL = 9525


def _chart_frame_xml(chart, position: int) -> str:
    first_col, first_row, _, _ = parse_cell_range(chart.anchor)
    width_emu = round(chart.width * 360000)
    height_emu = round(chart.height * 360000)
    return (
        "<oneCellAnchor>"
        f"<from><col>{first_col}</col><colOff>0</colOff>"
        f"<row>{first_row}</row><rowOff>0</rowOff></from>"
        f"<ext cx=\"{width_emu}\" cy=\"{height_emu}\"/>"
        "<graphicFrame><nvGraphicFramePr>"
        f"<cNvPr id=\"{position}\" name=\"Chart {position}\"/>"
        "<cNvGraphicFramePr/></nvGraphicFramePr><xfrm/>"
        "<a:graphic><a:graphicData "
        "uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">"
        f"<c:chart r:id=\"rId{position}\"/>"
        "</a:graphicData></a:graphic></graphicFrame>"
        "<clientData/></oneCellAnchor>"
    )


def _picture_frame_xml(image, position: int) -> str:
    """A ``pic`` anchor; ``image`` exposes anchor, pixel size and description."""
    first_col, first_row, _, _ = parse_cell_range(image.anchor)
    cx = image.pixel_width * EMU_PER_PIXEL
    cy = image.pixel_height * EMU_PER_PIXEL
    descr = (
        f' descr="{_attr(image.description)}"' if image.description is not None else ""
    )
    return (
        "<oneCellAnchor>"
        f"<from><col>{first_col}</col><colOff>0</colOff>"
        f"<row>{first_row}</row><rowOff>0</rowOff></from>"
        f"<ext cx=\"{cx}\" cy=\"{cy}\"/>"
        "<pic><nvPicPr>"
        f"<cNvPr id=\"{position}\" name=\"Picture {position}\"{descr}/>"
        "<cNvPicPr><a:picLocks noChangeAspect=\"1\"/></cNvPicPr></nvPicPr>"
        f"<blipFill><a:blip r:embed=\"rId{position}\"/>"
        "<a:stretch><a:fillRect/></a:stretch></blipFill>"
        "<spPr><a:xfrm><a:off x=\"0\" y=\"0\"/>"
        f"<a:ext cx=\"{cx}\" cy=\"{cy}\"/></a:xfrm>"
        "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></spPr>"
        "</pic><clientData/></oneCellAnchor>"
    )


def drawing_xml(charts: list, chart_ids: list[int], images=()) -> str:
    """``wsDr`` part; relationship ids run charts first, then pictures."""
    frames = [
        _chart_frame_xml(chart, position)
        for position, (chart, _chart_id) in enumerate(zip(charts, chart_ids), start=1)
    ]
    offset = len(charts)
    frames.extend(
        _picture_frame_xml(image, offset + position)
        for position, image in enumerate(images, start=1)
    )
    return (
        "<wsDr "
        "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" "
        "xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" "
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" "
        "xmlns=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\">"
        f"{''.join(frames)}</wsDr>"
    )


def drawing_rels_xml(chart_ids: list[int], image_names=()) -> str:
    parts = "".join(
        f'<Relationship Id="rId{position}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" '
        f'Target="../charts/chart{chart_id}.xml"/>'
        for position, chart_id in enumerate(chart_ids, start=1)
    )
    offset = len(chart_ids)
    parts += "".join(
        f'<Relationship Id="rId{offset + position}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
        f'Target="../media/{name}"/>'
        for position, name in enumerate(image_names, start=1)
    )
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
{parts}
</Relationships>'''


# Standard Office theme (Excel 365) bundled with chart-enabled workbooks.
# Chart default styles and accent colors resolve through the theme; without
# it the XLSX and XLSB parsers in Excel fall back to different defaults
# (observed: per-point colors + category legend on style-less line charts
# in XLSB). Content-verbatim copy of ``xl/theme/theme1.xml`` from an Excel file.
THEME_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="5B9BD5"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="4472C4"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light" panose="020F0302020204030204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="Yu Gothic Light"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="等线 Light"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:majorFont><a:minorFont><a:latin typeface="Calibri" panose="020F0502020204030204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="Yu Gothic"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="等线"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/><a:extLst><a:ext uri="{05A4C25C-085E-4340-85A3-A5531E510DB2}"><thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="Office Theme" id="{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}" vid="{4A3C46E8-61CC-4603-A589-7422A47A8E4A}"/></a:ext></a:extLst></a:theme>
"""
