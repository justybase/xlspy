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
from xml.sax.saxutils import escape as _xml_escape, quoteattr

from ._metadata import CT, MAIN, PKG_REL, REL, element_pattern, replace_element, tag_attributes
from ._safe_xml import fromstring as safe_fromstring
from ._sheet_common import SHEET_STATES, check_sheet_state, next_sheet_part
from ._xlsx_edit import _bytes as _xml
from ._xml_text import replace_or_add_xml_attribute, unescape_xml_text
from .comments import rels_path as _part_rels_path
from .exceptions import UnsupportedFeatureError
from .refs import check_sheet_name
from .updater_utils import (
    rename_formula_sheet_references,
    renamed_chart_parts,
    sheet_has_pivots,
    sheet_has_slicers,
)

if TYPE_CHECKING:
    from .updater_utils import ZipPackage


_FORMULA_TAGS = r"(?:f|formula|formula1|formula2|calculatedColumnFormula)"
_FORMULA_ELEMENT = re.compile(
    rf"(<(?:[A-Za-z_][\w.-]*:)?{_FORMULA_TAGS}\b[^>]*?(?<!/)>)"
    rf"(.*?)(</(?:[A-Za-z_][\w.-]*:)?{_FORMULA_TAGS}\s*>)",
    re.DOTALL,
)
_HYPERLINK_LOCATION = re.compile(
    r"(<(?:[A-Za-z_][\w.-]*:)?hyperlink\b[^>]*?\blocation\s*=\s*)"
    r"(\"[^\"]*\"|'[^']*')",
    re.DOTALL,
)
_DEFINED_NAME_ELEMENT = re.compile(
    r"(<(?:[A-Za-z_][\w.-]*:)?definedName\b[^>]*>)"
    r"(.*?)(</(?:[A-Za-z_][\w.-]*:)?definedName\s*>)",
    re.DOTALL,
)
_REFERENCE_PART = re.compile(r"xl/(?:worksheets|chartsheets|tables)/[^/]+\.xml")
_WORKSHEET_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
_EMPTY_WORKSHEET = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="{MAIN}"><sheetData/></worksheet>'
).encode()
_WORKBOOK_VIEW_SUCCESSORS = (
    "sheets",
    "functionGroups",
    "externalReferences",
    "definedNames",
    "calcPr",
    "oleSize",
    "customWorkbookViews",
    "pivotCaches",
    "smartTagPr",
    "smartTagTypes",
    "webPublishing",
    "fileRecoveryPr",
    "webPublishObjects",
    "extLst",
)


class XlsxSheetsMixin:
    """Sheet management for ``XlsxUpdater``."""

    if TYPE_CHECKING:
        # Provided by XlsxUpdater.
        _package: ZipPackage
        _sheets: dict[str, str]
        _chartsheets: set[str]

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
            states[name] = state if state in SHEET_STATES else "visible"
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
        check_sheet_state(state)
        xml = self._workbook_xml()
        found = False

        def change(match: re.Match[str]) -> str:
            nonlocal found
            opening = _opening_tag(match.group(0))
            if tag_attributes(opening).get("name") != name:
                return match.group(0)
            found = True
            return replace_or_add_xml_attribute(opening, "state", state) + match.group(0)[len(opening) :]

        updated = element_pattern("sheet").sub(change, xml)
        if not found:
            raise KeyError(f"Worksheet not found: {name}")
        if updated != xml:
            self._apply_parts({"xl/workbook.xml": updated})

    def add_sheet(self, name: str, index: int | None = None, *, state: str = "visible") -> str:
        """Add an empty worksheet at ``index`` (default: last); return its part path."""
        check_sheet_name(name, self._sheets)
        check_sheet_state(state)

        current = list(self._sheets)
        position = len(current) if index is None else max(0, min(index, len(current)))
        path = next_sheet_part(self._package.names(), "xml")
        xml = self._workbook_xml()
        root = safe_fromstring(xml)
        sheet_id = (
            max(
                (int(node.get("sheetId", "0") or 0) for node in root.findall(f"{{{MAIN}}}sheets/{{{MAIN}}}sheet")),
                default=0,
            )
            + 1
        )

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
        fragment = f'<sheet xmlns:r="{REL}" name={quoteattr(name)} sheetId="{sheet_id}" r:id="{rid}"{state_attribute}/>'
        xml = _splice_sheet_entry(xml, fragment, position)

        order = current[:position] + [name] + current[position:]
        mapping = {old: order.index(existing) for old, existing in enumerate(current)}
        previous_active = self.get_active_sheet()
        active = mapping.get(previous_active, previous_active)
        states = self.get_sheet_states()
        states[name] = state
        xml = _remap_active_and_first(xml, mapping, order, states, active)

        updates: dict[str, Any] = {
            "xl/workbook.xml": xml,
            path: _EMPTY_WORKSHEET,
            rels_path: _xml(rels),
            "[Content_Types].xml": _xml(types),
        }
        removed: list[str] = []
        self._recalculate(updates, removed)
        self._apply_parts(updates, removed)
        self._sheets = {sheet: (path if sheet == name else self._sheets[sheet]) for sheet in order}
        return path

    def remove_sheet(self, name: str) -> None:
        """Remove a sheet, its part and the parts only it refers to.

        References to it become ``#REF!`` and names local to it are dropped.
        """
        path = self._sheets.get(name)
        if path is None:
            raise KeyError(f"Worksheet not found: {name}")
        if len(self._sheets) == 1:
            raise ValueError("a workbook must keep at least one sheet")
        package = self._package
        # As for XLSB: slicers and pivots own workbook-level caches that would
        # be left pointing at a removed sheet.
        if sheet_has_slicers(package, path):
            raise UnsupportedFeatureError(f"sheet {name!r} holds table slicers; removing it is not supported yet")
        if sheet_has_pivots(package, path):
            raise UnsupportedFeatureError(f"sheet {name!r} holds pivot tables; removing it is not supported yet")

        current = list(self._sheets)
        removed_index = current.index(name)
        order = [sheet for sheet in current if sheet != name]
        xml = self._workbook_xml()
        xml, rid = _remove_sheet_entry(xml, name)
        xml = _strip_local_sheet_names(xml, removed_index)
        mapping = {old: order.index(sheet) for old, sheet in enumerate(current) if sheet != name}
        previous_active = self.get_active_sheet()
        active = mapping.get(previous_active, min(previous_active, len(order) - 1))
        states = self.get_sheet_states()
        states.pop(name, None)
        xml = _remap_active_and_first(xml, mapping, order, states, active)

        rels_path = "xl/_rels/workbook.xml.rels"
        rels = safe_fromstring(package.get(rels_path))
        for node in list(rels):
            if rid is not None and node.get("Id") == rid:
                rels.remove(node)
        types = safe_fromstring(package.get("[Content_Types].xml"))

        removed: list[str] = []
        skip = {_part_rels_path(path)}
        _drop_sheet_part_tree(package, path, removed, skip)
        _remove_part_overrides(types, removed)

        updates: dict[str, Any] = {
            "xl/workbook.xml": xml,
            rels_path: _xml(rels),
            "[Content_Types].xml": _xml(types),
        }
        self._recalculate(updates, removed)
        self._apply_parts(updates, removed)
        self._sheets = {sheet: self._sheets[sheet] for sheet in order}
        self._chartsheets.discard(name)

    def rename_sheet(self, old: str, new: str) -> None:
        """Rename sheet ``old`` to ``new``; raises KeyError when ``old`` is absent."""
        if old not in self._sheets:
            raise KeyError(f"Worksheet not found: {old}")
        if old == new:
            return
        check_sheet_name(new, [name for name in self._sheets if name != old])
        workbook_xml = self._workbook_xml()
        updated, found = _rename_sheet_entry(workbook_xml, old, new)
        if not found:
            raise KeyError(f"Worksheet not found: {old}")
        updated = _rename_defined_names(updated, old, new)
        updates: dict[str, Any] = {"xl/workbook.xml": updated}
        updates.update(_renamed_reference_parts(self._package, old, new))
        updates.update(renamed_chart_parts(self._package, old, new))
        self._apply_parts(updates)
        self._sheets = {(new if name == old else name): path for name, path in self._sheets.items()}
        if old in self._chartsheets:
            self._chartsheets.discard(old)
            self._chartsheets.add(new)

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
        xml = xml[: container.start()] + body + xml[container.end() :]

        mapping = {}
        for new, name in enumerate(order):
            mapping[current.index(name)] = new
        xml = element_pattern("definedName").sub(lambda match: _remap_local_sheet_id(match, mapping), xml)
        active = self.get_active_sheet()
        active = mapping.get(active, active)
        states = self.get_sheet_states()
        xml = _remap_active_and_first(xml, mapping, order, states, active)

        updates: dict[str, Any] = {"xl/workbook.xml": xml}
        removed: list[str] = []
        self._recalculate(updates, removed)
        self._apply_parts(updates, removed)
        self._sheets = {name: self._sheets[name] for name in order}


def _remap_active_and_first(
    xml: str,
    mapping: dict[int, int],
    order: list[str],
    states: dict[str, str],
    active: int,
) -> str:
    """Carry the active tab and the tab-bar origin through a sheet remap.

    ``mapping`` maps an old sheet index to its new one; ``order`` and ``states``
    describe the workbook after the change. A hidden active sheet falls back to
    the first visible one.
    """
    if not 0 <= active < len(order) or states.get(order[active], "visible") != "visible":
        active = next(
            (place for place, sheet in enumerate(order) if states.get(sheet, "visible") == "visible"),
            0,
        )
    view = element_pattern("workbookView").search(xml)
    first = None
    if view is not None:
        opening = _opening_tag(view.group(0))
        first = _int_attribute(tag_attributes(opening).get("firstSheet"))
        if first is not None:
            first = min(mapping.get(first, 0), active)
    return _write_view(xml, active, first)


def _rename_sheet_entry(workbook_xml: str, old: str, new: str) -> tuple[str, bool]:
    for match in list(element_pattern("sheet").finditer(workbook_xml)):
        if tag_attributes(match.group(0)).get("name") != old:
            continue
        opening = _opening_tag(match.group(0))
        renamed = replace_or_add_xml_attribute(opening, "name", new)
        return (
            workbook_xml[: match.start()] + renamed + workbook_xml[match.start() + len(opening) :],
            True,
        )
    return workbook_xml, False


def _rename_formula_text(text: str, old: str, new: str) -> str:
    decoded = unescape_xml_text(text)
    updated = rename_formula_sheet_references(decoded, old, new)
    return text if updated == decoded else _xml_escape(updated)


def _rename_hyperlink_location(match: re.Match[str], old: str, new: str) -> str:
    quoted = match.group(2)
    decoded = unescape_xml_text(quoted[1:-1])
    updated = rename_formula_sheet_references(decoded, old, new)
    if updated == decoded:
        return match.group(0)
    return match.group(1) + quoteattr(updated)


def _renamed_reference_parts(package: ZipPackage, old: str, new: str) -> dict[str, bytes]:
    parts: dict[str, bytes] = {}
    for name in package.names():
        if _REFERENCE_PART.fullmatch(name) is None:
            continue
        xml = package.get(name).decode("utf-8")
        updated = _FORMULA_ELEMENT.sub(
            lambda match: match.group(1) + _rename_formula_text(match.group(2), old, new) + match.group(3),
            xml,
        )
        updated = _HYPERLINK_LOCATION.sub(
            lambda match: _rename_hyperlink_location(match, old, new),
            updated,
        )
        if updated != xml:
            parts[name] = updated.encode("utf-8")
    return parts


def _rename_defined_names(workbook_xml: str, old: str, new: str) -> str:
    return _DEFINED_NAME_ELEMENT.sub(
        lambda match: match.group(1) + _rename_formula_text(match.group(2), old, new) + match.group(3),
        workbook_xml,
    )


def _remove_sheet_entry(workbook_xml: str, name: str) -> tuple[str, str | None]:
    for match in list(element_pattern("sheet").finditer(workbook_xml)):
        attrs = tag_attributes(match.group(0))
        if attrs.get("name") != name:
            continue
        rid = next(
            (value for key, value in attrs.items() if key.rsplit(":", 1)[-1] == "id"),
            None,
        )
        return (
            workbook_xml[: match.start()] + workbook_xml[match.end() :],
            rid,
        )
    return workbook_xml, None


def _strip_local_sheet_names(workbook_xml: str, removed_index: int) -> str:
    def change(match: re.Match[str]) -> str:
        opening = _opening_tag(match.group(0))
        value = _int_attribute(tag_attributes(opening).get("localSheetId"))
        if value is None:
            return match.group(0)
        if value == removed_index:
            return ""
        if value > removed_index:
            return replace_or_add_xml_attribute(opening, "localSheetId", value - 1) + match.group(0)[len(opening) :]
        return match.group(0)

    return element_pattern("definedName").sub(change, workbook_xml)


def _sheet_part_target(source: str, target: str) -> str:
    if target.startswith("/"):
        return target[1:]
    return posixpath.normpath(posixpath.join(posixpath.dirname(source), target))


def _sheet_part_referenced(package: ZipPackage, part: str, skip: set[str]) -> bool:
    for name in package.names():
        if not name.endswith(".rels") or name in skip:
            continue
        try:
            root = safe_fromstring(package.get(name))
        except Exception:
            return True
        if name.startswith("_rels/"):
            source = ""
        else:
            source = name[:-5].replace("/_rels/", "/")
        for node in root:
            if node.get("TargetMode") == "External":
                continue
            target = node.get("Target")
            if target and _sheet_part_target(source, target) == part:
                return True
    return False


def _drop_sheet_part_tree(
    package: ZipPackage,
    part: str,
    removed: list[str],
    skip: set[str],
) -> None:
    if part in removed or not package.has(part):
        return
    removed.append(part)
    rels_name = _part_rels_path(part)
    skip.add(rels_name)
    targets: list[str] = []
    if package.has(rels_name):
        root = safe_fromstring(package.get(rels_name))
        for node in root:
            if node.get("TargetMode") == "External":
                continue
            target = node.get("Target")
            if target:
                targets.append(_sheet_part_target(part, target))
        removed.append(rels_name)
    for target in targets:
        if target not in removed and not _sheet_part_referenced(package, target, skip):
            _drop_sheet_part_tree(package, target, removed, skip)


def _remove_part_overrides(types: ET.Element, parts: Iterable[str]) -> None:
    names = {"/" + part for part in parts}
    for node in list(types):
        if node.tag == f"{{{CT}}}Override" and node.get("PartName") in names:
            types.remove(node)


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
    return workbook_xml[: match.start()] + body + workbook_xml[match.end() :]


def _patch_opening(xml: str, tag: str, change: Callable[[str], str]) -> str:
    match = element_pattern(tag).search(xml)
    if match is None:
        return xml
    opening = _opening_tag(match.group(0))
    return xml[: match.start()] + change(opening) + xml[match.start() + len(opening) :]


def _set_view_attributes(opening: str, active: int, first: int | None) -> str:
    updated = replace_or_add_xml_attribute(opening, "activeTab", active)
    if first is not None:
        updated = replace_or_add_xml_attribute(updated, "firstSheet", min(first, active))
    return updated


def _write_view(xml: str, active: int, first: int | None) -> str:
    if element_pattern("workbookView").search(xml) is None:
        first_attribute = "" if first is None else f' firstSheet="{min(first, active)}"'
        fragment = f'<bookViews xmlns="{MAIN}"><workbookView activeTab="{active}"{first_attribute}/></bookViews>'
        return replace_element(xml, "bookViews", fragment, before=_WORKBOOK_VIEW_SUCCESSORS)
    return _patch_opening(
        xml,
        "workbookView",
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
    return replace_or_add_xml_attribute(opening, "localSheetId", new) + match.group(0)[len(opening) :]
