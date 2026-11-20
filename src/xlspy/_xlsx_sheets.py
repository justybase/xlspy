"""Sheet visibility, active sheet and order for ``XlsxUpdater``.

The XLSB updater keeps these in ``workbook.bin``; here they are attributes on
``<sheet>`` and ``<workbookView>`` in ``xl/workbook.xml``. Reordering has to
follow two indexes stored elsewhere in that part: ``localSheetId`` on defined
names and ``activeTab`` in the workbook view.
"""

from __future__ import annotations

import posixpath
import re
from collections.abc import Callable, Iterable, Mapping, MutableMapping
from typing import TYPE_CHECKING, Any
from xml.etree import ElementTree as ET
from xml.sax.saxutils import quoteattr

from ._metadata import CT, MAIN, PKG_REL, REL, element_pattern, replace_element, tag_attributes
from ._safe_xml import fromstring as safe_fromstring
from ._xlsx_edit import _bytes as _xml
from ._xml_text import replace_or_add_xml_attribute
from .refs import check_sheet_name

if TYPE_CHECKING:
    from .updater_utils import ZipPackage


_SHEET_STATES = frozenset({"visible", "hidden", "veryHidden"})
_WORKSHEET_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
)
_EMPTY_WORKSHEET = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f'<worksheet xmlns="{MAIN}"><sheetData/></worksheet>'
).encode()
_WORKBOOK_VIEW_SUCCESSORS = (
    "sheets", "functionGroups", "externalReferences", "definedNames",
    "calcPr", "oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr",
    "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects",
    "extLst",
)


class XlsxSheetsMixin:
    """Sheet management for ``XlsxUpdater``."""

    if TYPE_CHECKING:
        # Provided by XlsxUpdater.
        _package: ZipPackage
        _sheets: dict[str, str]

        def _apply_parts(self, parts: Mapping[str, Any], removed: Iterable[str] = ()) -> None: ...
        def _recalculate(self, updates: MutableMapping[str, Any], removed: list[str]) -> None: ...

    def _workbook_xml(self) -> str:
        return self._package.get("xl/workbook.xml").decode("utf-8")

    def get_sheet_states(self) -> dict[str, str]:
        """Sheet visibility (``visible``, ``hidden`` or ``veryHidden``) by name."""
        root = safe_fromstring(self._workbook_xml())
        states = {}
        for node in root.findall(f"{{{MAIN}}}sheets/{{{MAIN}}}sheet"):
            name = node.get("name")
            if name is None:
                continue
            state = node.get("state", "visible")
            states[name] = state if state in _SHEET_STATES else "visible"
        return states

    def get_active_sheet(self) -> int:
        """Index of the sheet Excel shows when the workbook opens."""
        root = safe_fromstring(self._workbook_xml())
        view = root.find(f"{{{MAIN}}}bookViews/{{{MAIN}}}workbookView")
        if view is None:
            return 0
        return _int_attribute(view.get("activeTab")) or 0

    def set_active_sheet(self, index: int) -> None:
        """Make the sheet at 0-based ``index`` the active one."""
        if not 0 <= index < len(self._sheets):
            raise IndexError("active sheet index out of range")
        xml = self._workbook_xml()
        match = element_pattern("workbookView").search(xml)
        first = None
        if match is not None:
            first = _int_attribute(tag_attributes(_opening_tag(match.group(0))).get("firstSheet"))
        updated = _write_view(xml, index, first)
        if updated != xml:
            self._apply_parts({"xl/workbook.xml": updated})

    def set_sheet_state(self, name: str, state: str) -> None:
        """Set a sheet to ``"visible"``, ``"hidden"`` or ``"veryHidden"``."""
        if state not in _SHEET_STATES:
            raise ValueError(f"invalid sheet state {state!r}")
        xml = self._workbook_xml()
        found = False

        def change(match: re.Match[str]) -> str:
            nonlocal found
            opening = _opening_tag(match.group(0))
            if tag_attributes(opening).get("name") != name:
                return match.group(0)
            found = True
            return (
                replace_or_add_xml_attribute(opening, "state", state)
                + match.group(0)[len(opening):]
            )

        updated = element_pattern("sheet").sub(change, xml)
        if not found:
            raise KeyError(f"Worksheet not found: {name}")
        if updated != xml:
            self._apply_parts({"xl/workbook.xml": updated})

    def add_sheet(
        self, name: str, index: int | None = None, *, state: str = "visible"
    ) -> str:
        """Add an empty worksheet at ``index`` (default: last); return its part path."""
        check_sheet_name(name, self._sheets)
        if state not in _SHEET_STATES:
            raise ValueError(f"invalid sheet state {state!r}")

        current = list(self._sheets)
        position = len(current) if index is None else max(0, min(index, len(current)))
        path = _new_sheet_path(self._package)
        xml = self._workbook_xml()
        root = safe_fromstring(xml)
        sheet_id = max(
            (
                int(node.get("sheetId", "0") or 0)
                for node in root.findall(f"{{{MAIN}}}sheets/{{{MAIN}}}sheet")
            ),
            default=0,
        ) + 1

        rels_path = "xl/_rels/workbook.xml.rels"
        rels = safe_fromstring(self._package.get(rels_path))
        rid = _next_relationship_id(rels)
        ET.SubElement(
            rels,
            f"{{{PKG_REL}}}Relationship",
            {
                "Id": rid,
                "Type": REL + "/worksheet",
                "Target": posixpath.relpath(path, "xl"),
            },
        )
        types = safe_fromstring(self._package.get("[Content_Types].xml"))
        ET.SubElement(
            types,
            f"{{{CT}}}Override",
            {"PartName": "/" + path, "ContentType": _WORKSHEET_CONTENT_TYPE},
        )

        state_attribute = "" if state == "visible" else f' state="{state}"'
        fragment = (
            f'<sheet xmlns:r="{REL}" name={quoteattr(name)} sheetId="{sheet_id}" '
            f'r:id="{rid}"{state_attribute}/>'
        )
        xml = _splice_sheet_entry(xml, fragment, position)

        order = current[:position] + [name] + current[position:]
        mapping = {old: order.index(existing) for old, existing in enumerate(current)}
        previous_active = self.get_active_sheet()
        active = mapping.get(previous_active, previous_active)
        states = self.get_sheet_states()
        states[name] = state
        if not 0 <= active < len(order) or states.get(order[active], "visible") != "visible":
            active = next(
                (place for place, sheet in enumerate(order)
                 if states.get(sheet, "visible") == "visible"),
                0,
            )
        view = element_pattern("workbookView").search(xml)
        first = None
        if view is not None:
            opening = _opening_tag(view.group(0))
            first = _int_attribute(tag_attributes(opening).get("firstSheet"))
            if first is not None:
                first = min(mapping.get(first, 0), active)
        xml = _write_view(xml, active, first)

        updates: dict[str, Any] = {
            "xl/workbook.xml": xml,
            path: _EMPTY_WORKSHEET,
            rels_path: _xml(rels),
            "[Content_Types].xml": _xml(types),
        }
        removed: list[str] = []
        self._recalculate(updates, removed)
        self._apply_parts(updates, removed)
        self._sheets = {
            sheet: (path if sheet == name else self._sheets[sheet]) for sheet in order
        }
        return path

    def reorder_sheets(self, order: Iterable[str]) -> None:
        """Put the sheets in ``order`` (every sheet name, once)."""
        order = list(order)
        current = list(self._sheets)
        if sorted(order) != sorted(current):
            raise ValueError("order must list every sheet exactly once")
        if order == current:
            return

        xml = self._workbook_xml()
        container = element_pattern("sheets").search(xml)
        if container is None:
            raise ValueError("workbook.xml sheets element was not found")
        body = container.group(0)
        by_name = {
            tag_attributes(match.group(0)).get("name"): match.group(0)
            for match in element_pattern("sheet").finditer(body)
        }
        if any(name not in by_name for name in order):
            raise ValueError("order must list every sheet exactly once")
        opening_end = body.index(">") + 1
        closing_start = body.rfind("</")
        reordered = "".join(by_name[name] for name in order)
        body = body[:opening_end] + reordered + body[closing_start:]
        xml = xml[:container.start()] + body + xml[container.end():]

        mapping = {}
        for new, name in enumerate(order):
            mapping[current.index(name)] = new
        xml = element_pattern("definedName").sub(
            lambda match: _remap_local_sheet_id(match, mapping), xml
        )
        active = self.get_active_sheet()
        active = mapping.get(active, active)
        states = self.get_sheet_states()
        if not 0 <= active < len(order) or states.get(order[active], "visible") != "visible":
            active = next(
                (position for position, name in enumerate(order)
                 if states.get(name, "visible") == "visible"),
                0,
            )
        view = element_pattern("workbookView").search(xml)
        first = None
        if view is not None:
            opening = _opening_tag(view.group(0))
            first = _int_attribute(tag_attributes(opening).get("firstSheet"))
            if first is not None:
                first = min(mapping.get(first, 0), active)
        xml = _write_view(xml, active, first)

        updates: dict[str, Any] = {"xl/workbook.xml": xml}
        removed: list[str] = []
        self._recalculate(updates, removed)
        self._apply_parts(updates, removed)
        self._sheets = {name: self._sheets[name] for name in order}


def _new_sheet_path(package: ZipPackage) -> str:
    highest = 0
    for name in package.names():
        match = re.fullmatch(r"xl/worksheets/sheet(\d+)\.xml", name)
        if match:
            highest = max(highest, int(match.group(1)))
    return f"xl/worksheets/sheet{highest + 1}.xml"


def _next_relationship_id(rels: ET.Element) -> str:
    highest = 0
    for node in rels:
        match = re.fullmatch(r"rId(\d+)", node.get("Id", ""))
        if match:
            highest = max(highest, int(match.group(1)))
    return f"rId{highest + 1}"


def _splice_sheet_entry(workbook_xml: str, fragment: str, index: int) -> str:
    match = element_pattern("sheets").search(workbook_xml)
    if match is None:
        raise ValueError("workbook.xml sheets element was not found")
    body = match.group(0)
    rows = list(element_pattern("sheet").finditer(body))
    position = rows[index].start() if index < len(rows) else body.rfind("</")
    if position < 0:
        position = len(body)
    body = body[:position] + fragment + body[position:]
    return workbook_xml[:match.start()] + body + workbook_xml[match.end():]


def _patch_opening(xml: str, tag: str, change: Callable[[str], str]) -> str:
    match = element_pattern(tag).search(xml)
    if match is None:
        return xml
    opening = _opening_tag(match.group(0))
    return (
        xml[:match.start()] + change(opening)
        + xml[match.start() + len(opening):]
    )


def _set_view_attributes(opening: str, active: int, first: int | None) -> str:
    updated = replace_or_add_xml_attribute(opening, "activeTab", active)
    if first is not None:
        updated = replace_or_add_xml_attribute(updated, "firstSheet", min(first, active))
    return updated


def _write_view(xml: str, active: int, first: int | None) -> str:
    if element_pattern("workbookView").search(xml) is None:
        first_attribute = "" if first is None else f' firstSheet="{min(first, active)}"'
        fragment = (
            f'<bookViews xmlns="{MAIN}">'
            f'<workbookView activeTab="{active}"{first_attribute}/>'
            "</bookViews>"
        )
        return replace_element(xml, "bookViews", fragment, before=_WORKBOOK_VIEW_SUCCESSORS)
    return _patch_opening(
        xml, "workbookView",
        lambda opening: _set_view_attributes(opening, active, first),
    )


def _opening_tag(fragment: str) -> str:
    match = re.match(r"<[^>]+>", fragment)
    if match is None:
        raise ValueError("XML opening tag was not found")
    return match.group(0)


def _int_attribute(value: str | None) -> int | None:
    if value is None:
        return None
    try:
        return int(value)
    except ValueError:
        return None


def _remap_local_sheet_id(match: re.Match[str], mapping: dict[int, int]) -> str:
    opening = _opening_tag(match.group(0))
    old = _int_attribute(tag_attributes(opening).get("localSheetId"))
    if old is None:
        return match.group(0)
    new = mapping.get(old)
    if new is None:
        return match.group(0)
    return replace_or_add_xml_attribute(opening, "localSheetId", new) + match.group(0)[len(opening):]
