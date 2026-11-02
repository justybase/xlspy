"""Edit the title, axis titles and legend of a chart read from a workbook.

The chart part is edited as text, not through an XML tree, so the prefixes and
``mc:Ignorable`` declarations of the file stay as Excel wrote them.
"""

from __future__ import annotations

import re
from xml.sax.saxutils import escape, unescape


# Legend positions and their DrawingML spelling (also exported by xlspy.charts).
LEGEND_POSITIONS = {
    "right": "r",
    "left": "l",
    "top": "t",
    "bottom": "b",
    "top_right": "tr",
}

_CHART_SPACE_ROOT = re.compile(r"<(?:(\w+):)?chartSpace\b")
_AXES = ("catAx", "valAx", "dateAx", "serAx")


def _chart_prefix(xml: str) -> str:
    match = _CHART_SPACE_ROOT.search(xml)
    return f"{match[1]}:" if match and match[1] else ""


def _drawing_prefix(xml: str) -> str:
    """Prefix of the DrawingML main namespace (``a:`` in Excel's files)."""
    match = re.search(r'xmlns:(\w+)="http://schemas\.openxmlformats\.org/drawingml/2006/main"', xml)
    return f"{match[1]}:" if match else ""


def _edit_title_xml(c: str, a: str, text: str) -> str:
    return (
        f"<{c}title><{c}tx><{c}rich><{a}bodyPr/><{a}p><{a}r><{a}t>{escape(text)}</{a}t>"
        f"</{a}r></{a}p></{c}rich></{c}tx><{c}overlay val=\"0\"/></{c}title>"
    )


def _element(xml: str, c: str, name: str, start: int = 0) -> tuple[int, int] | None:
    """``(start, end)`` of the first ``<c:name>`` element at or after ``start``."""
    opening = re.compile(rf"<{c}{name}(?=[\s>/])[^>]*?(/?)>")
    match = opening.search(xml, start)
    if match is None:
        return None
    if match[1]:
        return match.start(), match.end()
    depth = 0
    tag = re.compile(rf"<(/?){c}{name}(?=[\s>/])[^>]*?(/?)>")
    for item in tag.finditer(xml, match.start()):
        if item[2]:
            continue
        depth += -1 if item[1] else 1
        if depth == 0:
            return match.start(), item.end()
    return None


def _set_title_in(xml: str, c: str, a: str, span: tuple[int, int], text: str | None) -> str:
    """Set the title element that starts at ``span`` (None removes it)."""
    start, end = span
    old = xml[start:end]
    if text is None:
        return xml[:start] + xml[end:]
    run = re.search(rf"<{a}t>.*?</{a}t>|<{a}t/>", old, re.S)
    if run is None:
        return xml[:start] + _edit_title_xml(c, a, text) + xml[end:]
    # Keep the formatting of the first run: its text changes, other runs go.
    paragraph = re.search(rf"<{a}p>.*?</{a}p>", old, re.S)
    first_run = re.search(rf"<{a}r>.*?</{a}r>", old, re.S)
    new = f"<{a}t>{escape(text)}</{a}t>"
    edited = old[:run.start()] + new + old[run.end():]
    if paragraph and first_run:
        shift = len(new) - (run.end() - run.start())
        runs = list(re.finditer(rf"<{a}r>.*?</{a}r>", edited, re.S))
        for extra in reversed(runs[1:]):
            if paragraph.start() <= extra.start() < paragraph.end() + shift:
                edited = edited[:extra.start()] + edited[extra.end():]
    return xml[:start] + edited + xml[end:]


def _auto_title_deleted(xml: str, c: str, deleted: bool) -> str:
    value = "1" if deleted else "0"
    updated, count = re.subn(rf'(<{c}autoTitleDeleted\s+val=")[^"]*(")', rf"\g<1>{value}\2", xml, count=1)
    if count or not deleted:
        return updated
    # Schema order: title, autoTitleDeleted, pivotFmts, view3D, ..., plotArea.
    anchor = re.search(rf"<{c}(?:pivotFmts|view3D|floor|sideWall|backWall|plotArea)\b", xml)
    if anchor is None:
        return xml
    return xml[:anchor.start()] + f'<{c}autoTitleDeleted val="1"/>' + xml[anchor.start():]


def set_chart_title(xml: bytes, text: str | None) -> bytes:
    source = xml.decode("utf-8")
    c = _chart_prefix(source)
    a = _drawing_prefix(source) or "a:"
    chart = _element(source, c, "chart")
    if chart is None:
        raise ValueError("the chart part has no chart element")
    body = re.compile(rf"<{c}chart(?=[\s>])[^>]*>").match(source, chart[0])
    if body is None:
        raise ValueError("the chart part has no chart element")
    inner = body.end()
    title = None
    if re.compile(rf"\s*<{c}title(?=[\s>/])").match(source, inner):
        found = _element(source, c, "title", inner)
        title = found
    if title is not None:
        source = _set_title_in(source, c, a, title, text)
    elif text is not None:
        source = source[:inner] + _edit_title_xml(c, a, text) + source[inner:]
    return _auto_title_deleted(source, c, text is None).encode("utf-8")


def _axis_span(source: str, c: str, index: int) -> tuple[int, int] | None:
    spans = []
    for name in _AXES:
        position = 0
        while (found := _element(source, c, name, position)) is not None:
            spans.append(found)
            position = found[1]
    spans.sort()
    return spans[index] if index < len(spans) else None


def set_axis_title(xml: bytes, index: int, text: str | None) -> bytes:
    """Set the title of the ``index``-th axis in document order (0: X, 1: Y)."""
    source = xml.decode("utf-8")
    c = _chart_prefix(source)
    a = _drawing_prefix(source) or "a:"
    axis = _axis_span(source, c, index)
    if axis is None:
        raise ValueError("the chart has no such axis")
    start, end = axis
    body = source[start:end]
    found = _element(body, c, "title")
    if found is not None:
        body = _set_title_in(body, c, a, found, text)
    elif text is not None:
        # Schema order: ... axPos, gridlines, title, numFmt, ...
        grid_end = max(
            (m.end() for m in re.finditer(rf"</{c}(?:major|minor)Gridlines>|<{c}(?:major|minor)Gridlines\s*/>", body)),
            default=0,
        )
        anchor = re.compile(
            rf"<{c}(?:numFmt|majorTickMark|minorTickMark|tickLblPos|spPr|txPr|crossAx)\b",
        ).search(body, grid_end)
        at = anchor.start() if anchor else body.rindex(f"</{c}")
        body = body[:at] + _edit_title_xml(c, a, text) + body[at:]
    return (source[:start] + body + source[end:]).encode("utf-8")


def get_axis_title(xml: bytes, index: int) -> str | None:
    source = xml.decode("utf-8")
    c = _chart_prefix(source)
    a = _drawing_prefix(source) or "a:"
    axis = _axis_span(source, c, index)
    if axis is None:
        return None
    body = source[axis[0]:axis[1]]
    found = _element(body, c, "title")
    if found is None:
        return None
    runs = re.findall(rf"<{a}t>(.*?)</{a}t>", body[found[0]:found[1]], re.S)
    return unescape("".join(runs)) or None


def get_legend_position(xml: bytes) -> str | None:
    source = xml.decode("utf-8")
    c = _chart_prefix(source)
    found = _element(source, c, "legend")
    if found is None:
        return None
    match = re.search(rf'<{c}legendPos\s+val="(\w+)"', source[found[0]:found[1]])
    code = match[1] if match else "r"
    for name, value in LEGEND_POSITIONS.items():
        if value == code:
            return name
    return "right"


def set_legend_position(xml: bytes, position: str | None) -> bytes:
    """Move, add (``position`` given) or remove (None) the chart legend."""
    if position is not None and position not in LEGEND_POSITIONS:
        raise ValueError(
            f"invalid legend position {position!r}; expected one of {sorted(LEGEND_POSITIONS)}"
        )
    source = xml.decode("utf-8")
    c = _chart_prefix(source)
    found = _element(source, c, "legend")
    if position is None:
        if found is not None:
            source = source[:found[0]] + source[found[1]:]
        return source.encode("utf-8")
    code = LEGEND_POSITIONS[position]
    if found is not None:
        body = source[found[0]:found[1]]
        if re.search(rf"<{c}legendPos\b", body):
            body = re.sub(rf'(<{c}legendPos\s+val=")[^"]*(")', rf"\g<1>{code}\2", body, count=1)
        else:
            body = re.sub(rf"(<{c}legend>)", rf'\1<{c}legendPos val="{code}"/>', body, count=1)
        return (source[:found[0]] + body + source[found[1]:]).encode("utf-8")
    plot = _element(source, c, "plotArea")
    if plot is None:
        raise ValueError("the chart has no plot area")
    legend = f'<{c}legend><{c}legendPos val="{code}"/><{c}overlay val="0"/></{c}legend>'
    after = _element(source, c, "dTable", plot[1])
    at = plot[1] if after is None or after[0] != plot[1] else after[1]
    return (source[:at] + legend + source[at:]).encode("utf-8")


# --- series -----------------------------------------------------------------

_SERIES_BODY = (
    "marker", "dPt", "dLbls", "trendline", "errBars", "cat", "val", "xVal", "yVal",
    "invertIfNegative", "pictureOptions", "explosion", "smooth", "bubbleSize",
)
_LINE_GROUPS = frozenset({"lineChart", "line3DChart", "scatterChart", "radarChart", "stockChart"})
_FILLS = r"(?:solidFill|noFill|gradFill|pattFill|blipFill|grpFill)"


def _series_span(source: str, c: str, index: int) -> tuple[int, int]:
    position = 0
    found = None
    for _ in range(index + 1):
        found = _element(source, c, "ser", position)
        if found is None:
            break
        position = found[1]
    if found is None:
        raise IndexError("the chart has no such series")
    return found


def _series_end_of_header(body: str, c: str) -> int:
    """Offset in a ``<c:ser>`` element where idx, order, tx and spPr end."""
    names = "|".join(_SERIES_BODY)
    match = re.search(rf"<{c}(?:{names})\b", body)
    return match.start() if match else body.rindex(f"</{c}ser")


def set_series_title(xml: bytes, index: int, text: str) -> bytes:
    source = xml.decode("utf-8")
    c = _chart_prefix(source)
    start, end = _series_span(source, c, index)
    body = source[start:end]
    tx = f"<{c}tx><{c}v>{escape(text)}</{c}v></{c}tx>"
    found = _element(body, c, "tx")
    header = _series_end_of_header(body, c)
    if found is not None and found[0] < header:
        body = body[:found[0]] + tx + body[found[1]:]
    else:
        order = re.search(rf"<{c}order\b[^>]*/>", body)
        if order is None:
            raise ValueError("the series has no order element")
        body = body[:order.end()] + tx + body[order.end():]
    return (source[:start] + body + source[end:]).encode("utf-8")


def get_series_titles(xml: bytes) -> list[str | None]:
    source = xml.decode("utf-8")
    c = _chart_prefix(source)
    titles: list[str | None] = []
    position = 0
    while (found := _element(source, c, "ser", position)) is not None:
        body = source[found[0]:found[1]]
        position = found[1]
        tx = _element(body, c, "tx")
        if tx is None or tx[0] > _series_end_of_header(body, c):
            titles.append(None)
            continue
        text = re.search(rf"<{c}v>(.*?)</{c}v>", body[tx[0]:tx[1]], re.S)
        titles.append(unescape(text[1]) if text else None)
    return titles


def set_series_color(xml: bytes, index: int, color: str) -> bytes:
    """Fill a bar, area or pie series, or draw a line or scatter series, in ``color`` (RRGGBB)."""
    source = xml.decode("utf-8")
    c = _chart_prefix(source)
    a = _drawing_prefix(source) or "a:"
    start, end = _series_span(source, c, index)
    groups = re.findall(rf"<{c}(\w+Chart)\b", source[:start])
    line = bool(groups) and groups[-1] in _LINE_GROUPS
    fill = f'<{a}solidFill><{a}srgbClr val="{color}"/></{a}solidFill>'
    body = source[start:end]
    header = _series_end_of_header(body, c)
    sp = _element(body, c, "spPr")
    if sp is None or sp[0] >= header:
        shape = f"<{c}spPr>{f'<{a}ln>{fill}</{a}ln>' if line else fill}</{c}spPr>"
        anchor = max(
            (m.end() for m in re.finditer(rf"<{c}(?:order|tx)\b[^>]*/>|</{c}tx>", body[:header])),
            default=None,
        )
        body = body[:anchor] + shape + body[anchor:]
        return (source[:start] + body + source[end:]).encode("utf-8")
    old = body[sp[0]:sp[1]]
    inner_start = old.index(">") + 1
    inner = old[inner_start:old.rindex("<")] if not old.endswith("/>") else ""
    ln = re.search(rf"<{a}ln\b[^>]*?(?:/>|>.*?</{a}ln>)", inner, re.S)
    strip = re.compile(rf"<{a}{_FILLS}\b[^>]*?(?:/>|>.*?</{a}{_FILLS}>)", re.S)
    if line:
        if ln is None:
            new_ln = f"<{a}ln>{fill}</{a}ln>"
            inner = _insert_before_effects(inner, a, new_ln)
        else:
            text = ln[0]
            if text.endswith("/>"):
                text = text[:-2] + f"></{a}ln>"
            opening = text[:text.index(">") + 1]
            rest = strip.sub("", text[len(opening):])
            inner = inner[:ln.start()] + opening + fill + rest + inner[ln.end():]
    else:
        without_ln = inner if ln is None else inner[:ln.start()] + "\0" + inner[ln.end():]
        without_ln = strip.sub("", without_ln)
        if "\0" in without_ln:
            inner = without_ln.replace("\0", fill + ln[0])  # type: ignore[index]
        else:
            inner = _insert_before_effects(without_ln, a, fill)
    opening = old[:inner_start] if not old.endswith("/>") else old[:-2] + ">"
    shape = f"{opening}{inner}</{c}spPr>"
    body = body[:sp[0]] + shape + body[sp[1]:]
    return (source[:start] + body + source[end:]).encode("utf-8")


def _insert_before_effects(inner: str, a: str, fragment: str) -> str:
    effects = re.search(rf"<{a}(?:effectLst|effectDag|scene3d|sp3d|extLst)\b", inner)
    at = effects.start() if effects else len(inner)
    return inner[:at] + fragment + inner[at:]
