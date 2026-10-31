"""Read charts and pictures from worksheet/chartsheet drawings."""

from __future__ import annotations

import posixpath
from xml.etree import ElementTree as ET

from .charts import LoadedChart
from .images import Image
from .refs import column_label

_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

_CHART_TAGS = {
    "areaChart": "area",
    "area3DChart": "area3d",
    "barChart": "bar",
    "bar3DChart": "bar3d",
    "bubbleChart": "bubble",
    "doughnutChart": "doughnut",
    "lineChart": "line",
    "line3DChart": "line3d",
    "ofPieChart": "projected_pie",
    "pieChart": "pie",
    "pie3DChart": "pie3d",
    "radarChart": "radar",
    "scatterChart": "scatter",
    "stockChart": "stock",
    "surfaceChart": "surface",
    "surface3DChart": "surface3d",
}


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _relationships(package, part: str) -> dict[str, tuple[str, str]]:
    path = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
    if not package.has(path):
        return {}
    root = ET.fromstring(package.get(path))
    result = {}
    for node in root:
        rid = node.get("Id")
        target = node.get("Target")
        if rid and target:
            resolved = (
                target[1:] if target.startswith("/")
                else posixpath.normpath(posixpath.join(posixpath.dirname(part), target))
            )
            result[rid] = (resolved, node.get("Type", ""))
    return result


def _anchor_cell(anchor) -> str:
    marker = next((child for child in anchor if _local(child.tag) == "from"), None)
    if marker is None:
        return "A1"
    column = 0
    row = 0
    for child in marker:
        name = _local(child.tag)
        if name == "col" and child.text:
            column = int(child.text)
        elif name == "row" and child.text:
            row = int(child.text)
    return f"{column_label(column)}{row + 1}"


def _chart_facts(xml: bytes) -> tuple[str | None, str | None, int]:
    root = ET.fromstring(xml)
    chart = next((node for node in root.iter() if _local(node.tag) == "chart"), None)
    if chart is None:
        return None, None, 0
    chart_type = None
    for node in chart.iter():
        if _local(node.tag) in _CHART_TAGS:
            chart_type = _CHART_TAGS[_local(node.tag)]
            break
    title = None
    # The chart's own title, not that of an axis.
    title_node = next((node for node in chart if _local(node.tag) == "title"), None)
    if title_node is not None:
        title = "".join(
            node.text or "" for node in title_node.iter() if _local(node.tag) == "t"
        ) or None
    series = sum(1 for node in chart.iter() if _local(node.tag) == "ser")
    return chart_type, title, series


def read_drawing(package, sheet, drawing_part: str) -> None:
    """Append charts and images from one drawing part to ``sheet``."""
    if not package.has(drawing_part):
        return
    relationships = _relationships(package, drawing_part)
    root = ET.fromstring(package.get(drawing_part))
    # Where each object is in the file, so that a save can remove it.
    sources = sheet.__dict__.setdefault("_drawing_sources", [])
    position = -1
    for anchor in root:
        if _local(anchor.tag) not in {"oneCellAnchor", "twoCellAnchor", "absoluteAnchor"}:
            continue
        position += 1
        cell = _anchor_cell(anchor)
        for node in anchor.iter():
            name = _local(node.tag)
            if name == "chart":
                rid = node.get(f"{{{_REL}}}id")
                part = relationships.get(rid, (None, ""))[0]
                if not part or not package.has(part):
                    continue
                chart_xml = package.get(part)
                chart_type, title, series = _chart_facts(chart_xml)
                sheet._charts.append(
                    LoadedChart(chart_type, title, series, cell, part, chart_xml)
                )
                sources.append((drawing_part, position, sheet._charts[-1]))
            elif name == "pic":
                blip = next(
                    (item for item in node.iter() if _local(item.tag) == "blip"), None
                )
                rid = blip.get(f"{{{_REL}}}embed") if blip is not None else None
                part = relationships.get(rid, (None, ""))[0]
                if not part or not package.has(part):
                    continue
                description = None
                non_visual = next(
                    (item for item in node.iter() if _local(item.tag) == "cNvPr"), None
                )
                if non_visual is not None:
                    description = non_visual.get("descr")
                try:
                    image = Image(
                        package.get(part), anchor=cell, description=description
                    )
                except (TypeError, ValueError):
                    continue
                sheet._images.append(image)
                sources.append((drawing_part, position, image))


def read_chartsheet(package, part: str, title: str, chartsheet) -> None:
    relationships = _relationships(package, part)
    drawing = next(
        (
            target for target, kind in relationships.values()
            if "drawing" in kind and package.has(target)
        ),
        None,
    )
    if drawing is not None:
        read_drawing(package, chartsheet, drawing)
