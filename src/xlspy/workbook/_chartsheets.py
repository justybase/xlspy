"""Add new chartsheets to an XLSX package.

A chartsheet is a sheet entry in workbook.xml whose part holds one drawing
with one chart. The chart reads its data from a worksheet (the chartsheet's
``source``, by default the first worksheet).
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING

from .. import chart_parts
from .._metadata import REL
from ..refs import column_label
from ._objects import _Package
from ..external_links import _MAIN as _MAIN

if TYPE_CHECKING:
    from ._workbook import Workbook
    from ._worksheet import Chartsheet

_CHARTSHEET_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"
_DRAWING_TYPE = "application/vnd.openxmlformats-officedocument.drawing+xml"
_CHART_TYPE = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"

_CHARTSHEET_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    f'<chartsheet xmlns="{_MAIN}" xmlns:r="{REL}"><sheetPr/>'
    '<sheetViews><sheetView zoomScale="100" zoomToFit="1" workbookViewId="0"/></sheetViews>'
    '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
    '<drawing r:id="rId1"/></chartsheet>'
)
# The chart fills the sheet: Excel sizes an absolute anchor of a chartsheet itself.
_DRAWING_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
    '<xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="9293679" cy="6068786"/>'
    '<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/>'
    '<xdr:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></xdr:cNvGraphicFramePr>'
    '</xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>'
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
    '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
    f'xmlns:r="{REL}" r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame>'
    '<xdr:clientData/></xdr:absoluteAnchor></xdr:wsDr>'
)


def pending_chartsheets(workbook) -> list:
    return [sheet for sheet in workbook._chartsheets if sheet._part is None]


def check_chartsheets(workbook) -> None:
    """Fail before anything is written on a chartsheet that cannot be saved."""
    for sheet in pending_chartsheets(workbook):
        if len(sheet.charts) != 1:
            raise ValueError(f"chartsheet {sheet.title!r} needs exactly one chart")
        source_title(workbook, sheet)


def source_title(workbook: Workbook, sheet: Chartsheet) -> str:
    source = sheet._source
    if source is None:
        if not workbook._sheets:
            raise ValueError(f"chartsheet {sheet.title!r} has no worksheet to read its data from")
        return workbook._sheets[0].title
    title = source if isinstance(source, str) else source.title
    if title not in workbook.sheetnames:
        raise ValueError(f"chartsheet {sheet.title!r} reads from a missing worksheet {title!r}")
    return title


def _namespace_prefix(xml: str, namespace: str) -> str | None:
    match = re.search(rf'xmlns:([\w.-]+)="{re.escape(namespace)}"', xml)
    return match[1] if match else None


def add_chartsheets(updater, workbook) -> None:
    """Write the workbook's new chartsheets into the updater's package."""
    sheets = pending_chartsheets(workbook)
    if not sheets:
        return
    package = _Package(updater)
    workbook_xml = (package.get("xl/workbook.xml") or b"").decode("utf-8")
    workbook_rels = package.rels("xl/workbook.xml")
    letters = [column_label(index) for index in range(16384)]
    close = re.search(r"</([\w.-]+:)?sheets\s*>", workbook_xml)
    if close is None:
        raise ValueError("workbook.xml has no sheets element")
    tag = f"{close[1] or ''}sheet"
    rel_prefix = _namespace_prefix(workbook_xml, REL)
    sheet_id = max((int(value) for value in re.findall(r'\bsheetId="(\d+)"', workbook_xml)), default=0)
    entries = []
    for sheet in sheets:
        chart = package.free_part("xl/charts", "chart", "xml")
        package.set(chart, chart_parts.chart_xml(letters, sheet.charts[0], source_title(workbook, sheet)))
        package.override(chart, _CHART_TYPE)
        drawing = package.free_part("xl/drawings", "drawing", "xml")
        package.set(drawing, _DRAWING_XML)
        package.override(drawing, _DRAWING_TYPE)
        drawing_rels = package.rels(drawing)
        package.add_rel(drawing_rels, drawing, "chart", chart)
        part = package.free_part("xl/chartsheets", "sheet", "xml")
        package.set(part, _CHARTSHEET_XML)
        package.override(part, _CHARTSHEET_TYPE)
        part_rels = package.rels(part)
        package.add_rel(part_rels, part, "drawing", drawing)
        for owner, rels in ((drawing, drawing_rels), (part, part_rels)):
            _set_rels(package, owner, rels)
        rid = package.add_rel(workbook_rels, "xl/workbook.xml", "chartsheet", part)
        sheet_id += 1
        state = "" if sheet.sheet_state == "visible" else f' state="{sheet.sheet_state}"'
        name = sheet.title.replace("&", "&amp;").replace('"', "&quot;").replace("<", "&lt;")
        if rel_prefix:
            entries.append(f'<{tag} name="{name}" sheetId="{sheet_id}"{state} {rel_prefix}:id="{rid}"/>')
        else:
            entries.append(f'<{tag} xmlns:r="{REL}" name="{name}" sheetId="{sheet_id}"{state} r:id="{rid}"/>')
        sheet._part = part
    workbook_xml = workbook_xml[:close.start()] + "".join(entries) + workbook_xml[close.start():]
    package.set("xl/workbook.xml", workbook_xml)
    _set_rels(package, "xl/workbook.xml", workbook_rels)
    package.commit()


def _set_rels(package, owner, rels) -> None:
    from xml.etree import ElementTree as ET

    from .._metadata import PKG_REL
    from .._xlsx_edit import _bytes
    from ._objects import _rels_path

    ET.register_namespace("", PKG_REL)
    package.set(_rels_path(owner), _bytes(rels))
