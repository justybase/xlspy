"""Sheet visibility, active sheet and order for ``XlsxUpdater``.

The XLSB updater keeps these in ``workbook.bin``; here they are attributes on
``<sheet>`` and ``<workbookView>`` in ``xl/workbook.xml``. Reordering has to
follow two indexes stored elsewhere in that part: ``localSheetId`` on defined
names and ``activeTab`` in the workbook view.
"""

from __future__ import annotations

import re
from collections.abc import Iterable, Mapping
from typing import TYPE_CHECKING, Any

from ._metadata import MAIN, element_pattern, replace_element, tag_attributes
from ._safe_xml import fromstring as safe_fromstring
from ._xml_text import replace_or_add_xml_attribute

if TYPE_CHECKING:
    from .updater_utils import ZipPackage


_SHEET_STATES = frozenset({"visible", "hidden", "veryHidden"})
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
        if match is not None:
            opening = _opening_tag(match.group(0))
            updated_opening = replace_or_add_xml_attribute(opening, "activeTab", index)
            first = _int_attribute(tag_attributes(opening).get("firstSheet"))
            if first is not None and first > index:
                updated_opening = replace_or_add_xml_attribute(updated_opening, "firstSheet", index)
            updated = (
                xml[:match.start()] + updated_opening
                + xml[match.start() + len(opening):]
            )
        else:
            fragment = (
                f'<bookViews xmlns="{MAIN}">'
                f'<workbookView activeTab="{index}"/>'
                "</bookViews>"
            )
            updated = replace_element(xml, "bookViews", fragment, before=_WORKBOOK_VIEW_SUCCESSORS)
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
        if view is not None:
            opening = _opening_tag(view.group(0))
            updated_opening = replace_or_add_xml_attribute(opening, "activeTab", active)
            first = _int_attribute(tag_attributes(opening).get("firstSheet"))
            if first is not None:
                first = min(mapping.get(first, 0), active)
                updated_opening = replace_or_add_xml_attribute(updated_opening, "firstSheet", first)
            xml = (
                xml[:view.start()] + updated_opening
                + xml[view.start() + len(opening):]
            )

        self._apply_parts({"xl/workbook.xml": xml})
        self._sheets = {name: self._sheets[name] for name in order}


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
