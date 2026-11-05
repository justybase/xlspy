"""Targeted XLSX/XLSM edits, staged atomically through ZipPackage."""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING, Any

from collections.abc import Collection, Iterable, Mapping, MutableMapping, Sequence
import datetime
from decimal import Decimal
import math
import numbers
import re
from xml.etree import ElementTree as ET
from . import _accel
from ._safe_xml import check_well_formed, fromstring as safe_fromstring
from xml.sax.saxutils import escape, quoteattr

from ._sheet_data import insert_rows, patch_rows, set_dimension_ref, sheet_data_match
from ._metadata import (
    tag_attributes,
    MAIN, REL, PKG_REL, CT, DefinedName, SheetProtection, WorkbookProtection,
    cell_ref, element_pattern, replace_element, append_child, read_defined_names, xml_text,
)
from ._text import check_cell_text, xstring_escape
from .comments import (
    Comment, COMMENTS_TYPE, VML_TYPE, V, comments_xml, vml_xml, normalize_comments,
    NotePartsEdit, comment_part_targets, find_comment_parts, read_comments, resolve_part, rels_path,
)
from .excel_dates import datetime_to_serial
from .exceptions import UnsupportedFeatureError
from .formulas import Formula
from .hyperlinks import Hyperlink
from .refs import RangeIndex, _column_index, parse_cell_range, column_label
from .updater_utils import unwrap_cell, replace_or_add_xml_attribute
from .rich_text import CellRichText

if TYPE_CHECKING:
    from .updater_utils import DiskStringIndex, ZipPackage


def _bytes(root: ET.Element) -> bytes:
    data: bytes = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    return data


def _required_match(match: re.Match[str] | None, what: str) -> re.Match[str]:
    """A regex match the file structure guarantees; a corrupt part fails clearly."""
    if match is None:
        raise ValueError(f"{what} was not found")
    return match


_HYPERLINK_TYPE = REL + "/hyperlink"
# The <worksheet> children that follow <hyperlinks> (CT_Worksheet).
_AFTER_HYPERLINKS = (
    "printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks",
    "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing",
    "legacyDrawing", "legacyDrawingHF", "drawingHF", "picture", "oleObjects", "controls",
    "webPublishItems", "tableParts", "extLst",
)


class XlsxEditMixin:
    if TYPE_CHECKING:
        # Provided by the host updater (XlsxUpdater or XlsbUpdater). Annotations
        # only, plus stubs that the host's own methods replace at runtime.
        _package: ZipPackage
        _sheets: dict[str, str]
        _date_1904: bool
        _shared_strings_loaded: bool
        _shared_strings_total: int
        _stream_shared_strings: DiskStringIndex | None

        @staticmethod
        def _tag_prefix(xml_data: str, tag_name: str) -> str: ...

    def _apply_parts(self, parts: Mapping[str, Any], removed: Iterable[str] = ()) -> None:
        # The package owns the write-back transaction (ZipPackage.apply_parts).
        self._package.apply_parts(parts, removed)

    def _sheet_text(self, sheet_name: str) -> tuple[str, str]:
        """The part path and XML of a sheet, unchecked and unparsed."""
        try:
            path = self._sheets[sheet_name]
        except KeyError as exc:
            raise KeyError(f"Worksheet not found: {sheet_name}") from exc
        return path, self._package.get(path).decode("utf-8")

    def _sheet_root(self, sheet_name: str) -> tuple[str, str, Any]:
        path, xml = self._sheet_text(sheet_name)
        root = safe_fromstring(xml)
        if root.tag != f"{{{MAIN}}}worksheet":
            raise ValueError(f"{sheet_name!r} is not a worksheet")
        return path, xml, root

    def _sheet_scan_source(self, sheet_name: str) -> tuple[str, str, Any, _PlainSheet | None]:
        """What update_cells scans: the XML checked, and either its parsed tree or,
        for the plain markup the native scan reads, the places in the text. A tree
        is built only when the text is not plain; a part that is not well formed, or
        not a worksheet, is refused here either way."""
        path, xml = self._sheet_text(sheet_name)
        fast = _fast_sheet(xml) if _accel.xlsx_edit_scan is not None else None
        if fast is not None:
            check_well_formed(xml)
            return path, xml, None, fast
        root = safe_fromstring(xml)
        if root.tag != f"{{{MAIN}}}worksheet":
            raise ValueError(f"{sheet_name!r} is not a worksheet")
        return path, xml, root, None

    def _sheet_xml(self, sheet_name: str) -> tuple[str, str]:
        path, xml, _ = self._sheet_root(sheet_name)
        return path, xml

    def get_comments(self, sheet_name: str) -> dict[str, Comment]:
        """Read classic note text and authors from the current staged package."""
        path, _ = self._sheet_xml(sheet_name)
        _, _, parts = find_comment_parts(self._package.get, self._package.names(), path)
        return read_comments(self._package.get(parts["comments"][0])) if "comments" in parts else {}

    def set_comments(self, sheet_name: str, comments: Mapping[str, Comment | None]) -> None:
        """Set selected classic notes; None removes a note, other notes are kept."""
        path, sheet, sheet_root = self._sheet_root(sheet_name)
        merges = [n.attrib["ref"] for n in sheet_root.findall(f"{{{MAIN}}}mergeCells/{{{MAIN}}}mergeCell")]
        changes = normalize_comments(comments, allow_none=True, merges=merges)
        if not changes:
            return
        names = self._package.names()
        legacy = sheet_root.find(f"{{{MAIN}}}legacyDrawing")
        legacy_rid = None if legacy is None else legacy.get(f"{{{REL}}}id")
        rp, relationships, parts = find_comment_parts(self._package.get, names, path, legacy_rid)
        for relationship in relationships:
            if relationship.get("Type", "").rsplit("/", 1)[-1] in {"threadedComment", "threadedComments"}:
                threaded = safe_fromstring(self._package.get(resolve_part(path, relationship.attrib["Target"])))
                occupied = {cell_ref(node.attrib["ref"]) for node in threaded if "ref" in node.attrib}
                if any(ref in occupied and value is not None for ref, value in changes.items()):
                    raise ValueError("a classic note cannot replace a threaded comment")
        current = read_comments(self._package.get(parts["comments"][0])) if "comments" in parts else {}
        changes = {ref: value for ref, value in changes.items() if value is not None or ref in current}
        if not changes:
            return
        for other_path in self._sheets.values():
            if other_path == path:
                continue
            other = comment_part_targets(self._package.get, names, other_path)
            if any(value[0] in other for value in parts.values()):
                raise ValueError("editing comment/VML parts shared by worksheets is unsupported")
        if legacy is not None and "vmlDrawing" not in parts:
            raise ValueError("unsupported legacy drawing relationship")

        edit = NotePartsEdit(path, relationships, parts, names)
        cp = edit.ensure("comments", "xl/comments", "comment", "xml")
        vp = edit.ensure("vmlDrawing", "xl/drawings", "commentsDrawing", "vml")
        comment_data = comments_xml(changes, self._package.get(cp) if cp in names else None)
        vml_data = vml_xml(changes, self._package.get(vp) if vp in names else None)
        remaining = read_comments(comment_data)
        removed: list[str] = []
        updates: dict[str, Any] = {cp: comment_data, vp: vml_data}
        types = safe_fromstring(self._package.get("[Content_Types].xml"))

        if not remaining:
            edit.drop(cp, "comments", types, updates, removed)
        elif not any(n.get("PartName") == "/" + cp for n in types):
            ET.SubElement(types, f"{{{CT}}}Override", {"PartName": "/" + cp, "ContentType": COMMENTS_TYPE})
        has_shapes = any(n.tag == f"{{{V}}}shape" for n in safe_fromstring(vml_data).iter())
        if not has_shapes:
            edit.drop(vp, "vmlDrawing", types, updates, removed)
            sheet = replace_element(sheet, "legacyDrawing", "")
        else:
            # Use a part override: it cannot alter another VML part's type.
            if not any(n.get("PartName") == "/" + vp for n in types) and not any(n.get("Extension") == "vml" for n in types):
                ET.SubElement(types, f"{{{CT}}}Override", {"PartName": "/" + vp, "ContentType": VML_TYPE})
            if legacy is None:
                fragment = f'<legacyDrawing xmlns="{MAIN}" xmlns:r="{REL}" r:id="{parts["vmlDrawing"][1]}"/>'
                sheet = replace_element(sheet, "legacyDrawing", fragment, before=("legacyDrawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst"))
        updates[path] = sheet
        if len(relationships):
            updates[rp] = _bytes(relationships)
        elif rp in names:
            removed.append(rp)
        updates["[Content_Types].xml"] = _bytes(types)
        self._apply_parts(updates, removed)

    def set_hyperlinks(self, sheet_name: str, links: Mapping[str, Hyperlink | None]) -> None:
        """Set selected cells' hyperlinks; None removes one, other links are kept.

        Only the link changes: the cell's value and style stay as they are. A
        link whose range starts at a given cell is replaced as a whole.
        """
        path, sheet = self._sheet_xml(sheet_name)
        changes = {}
        for ref, link in links.items():
            if link is not None and not isinstance(link, Hyperlink):
                raise TypeError("hyperlinks must be xlspy Hyperlink objects or None")
            changes[cell_ref(ref)] = link
        if not changes:
            return
        rp = rels_path(path)
        names = self._package.names()
        relationships = (
            safe_fromstring(self._package.get(rp)) if rp in names
            else ET.Element(f"{{{PKG_REL}}}Relationships")
        )
        prefix = self._tag_prefix(sheet, "worksheet")
        container = element_pattern("hyperlinks").search(sheet)
        kept: list[Any] = []
        dropped: set[Any] = set()
        if container:
            for match in element_pattern("hyperlink").finditer(container[0]):
                attrs = tag_attributes(match[0])
                try:
                    first_col, first_row, _, _ = parse_cell_range(attrs.get("ref", "").split()[0])
                    start = f"{column_label(first_col)}{first_row + 1}"
                except (IndexError, ValueError):
                    start = None
                if start in changes:
                    dropped.update(value for key, value in attrs.items() if key.endswith(":id"))
                else:
                    kept.append(match[0])

        used_ids = {node.get("Id") for node in relationships}
        added = []
        index = 1  # Ids only get taken here, so the first free one never moves back.
        for ref, link in changes.items():
            if link is None:
                continue
            text = f' ref="{ref}"'
            if link.is_internal:
                text += f" location={quoteattr(xml_text(link.location))}"
            else:
                while f"rId{index}" in used_ids:
                    index += 1
                rid = f"rId{index}"
                used_ids.add(rid)
                ET.SubElement(relationships, f"{{{PKG_REL}}}Relationship", {
                    "Id": rid, "Type": _HYPERLINK_TYPE,
                    "Target": xml_text(link.target), "TargetMode": "External",
                })
                text += f' xmlns:r="{REL}" r:id="{rid}"'
            if link.text is not None:
                text += f" display={quoteattr(xml_text(link.text))}"
            if link.tooltip is not None:
                text += f" tooltip={quoteattr(xml_text(link.tooltip))}"
            added.append(f"<{prefix}hyperlink{text}/>")

        items = kept + added
        fragment = f"<{prefix}hyperlinks>{''.join(items)}</{prefix}hyperlinks>" if items else ""
        sheet = replace_element(sheet, "hyperlinks", fragment, before=_AFTER_HYPERLINKS)
        # A dropped link's relationship goes unless something else uses it.
        if dropped:
            still_used = set(re.findall(r"""[\w.-]+:id=["']([^"']*)["']""", sheet))
            for node in list(relationships):
                rid = node.get("Id")
                if rid in dropped and rid not in still_used and node.get("Type") == _HYPERLINK_TYPE:
                    relationships.remove(node)
        updates: dict[str, str | bytes] = {path: sheet}
        removed: list[str] = []
        ET.register_namespace("", PKG_REL)
        if len(relationships):
            updates[rp] = _bytes(relationships)
        elif rp in names:
            removed.append(rp)
        self._apply_parts(updates, removed)

    def get_defined_names(self) -> list[DefinedName]:
        """Return user-defined global and worksheet-local names."""
        return read_defined_names(self._package.get("xl/workbook.xml"), self._sheets)

    def set_defined_name(self, definition: DefinedName) -> None:
        """Add a defined name, replacing one with the same name and scope."""
        if not isinstance(definition, DefinedName):
            raise TypeError("definition must be a DefinedName")
        definition = definition.resolved(self._sheets)
        fragment = definition.to_xml(self._sheets)
        xml = self._package.get("xl/workbook.xml").decode("utf-8")
        xml = self._remove_name(xml, definition.name, definition.scope)
        if element_pattern("definedNames").search(xml):
            xml = append_child(xml, "definedNames", fragment)
        else:
            xml = replace_element(xml, "definedNames", f'<definedNames xmlns="{MAIN}">{fragment}</definedNames>', before=("calcPr", "oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr", "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst"))
        self._apply_name_change(xml)

    def _apply_name_change(self, xml: str) -> None:
        """Save a changed workbook part. Formulas may use the name, so their
        cached values are stale: recalculate on load, as for XLSB."""
        updates: dict[str, Any] = {"xl/workbook.xml": xml}
        removed: list[str] = []
        self._recalculate(updates, removed)
        self._apply_parts(updates, removed)

    def _remove_name(self, xml: str, name: str, scope: str | None) -> str:
        scope_id = None if scope is None else str(list(self._sheets).index(scope))
        def remove(match: re.Match[str]) -> str:
            attrs = tag_attributes(match[0])
            return "" if attrs.get("name", "").casefold() == name.casefold() and attrs.get("localSheetId") == scope_id else match[0]
        updated = element_pattern("definedName").sub(remove, xml)
        if not element_pattern("definedName").search(updated):
            updated = replace_element(updated, "definedNames", "")
        return updated

    def remove_defined_name(self, name: str, scope: str | None = None) -> None:
        """Remove a defined name; with ``scope``, the one local to that sheet."""
        # Validate the name, including the reserved built-in namespace.
        DefinedName(name, "0", scope).to_xml(self._sheets)
        xml = self._package.get("xl/workbook.xml").decode("utf-8")
        updated = self._remove_name(xml, name, scope)
        if updated != xml:
            self._apply_name_change(updated)

    def set_sheet_protection(self, sheet_name: str, protection: SheetProtection) -> None:
        """Set or replace the protection of a sheet."""
        if not isinstance(protection, SheetProtection):
            raise TypeError("protection must be SheetProtection")
        path, xml = self._sheet_xml(sheet_name)
        xml = replace_element(xml, "sheetProtection", protection.to_xml(), before=("protectedRanges", "scenarios", "autoFilter", "sortState", "dataConsolidate", "customSheetViews", "mergeCells", "phoneticPr", "conditionalFormatting", "dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks", "drawing", "legacyDrawing", "tableParts", "extLst"))
        self._apply_parts({path: xml})

    def set_workbook_protection(self, protection: WorkbookProtection) -> None:
        """Set or replace the workbook protection."""
        if not isinstance(protection, WorkbookProtection):
            raise TypeError("protection must be WorkbookProtection")
        xml = self._package.get("xl/workbook.xml").decode("utf-8")
        xml = replace_element(xml, "workbookProtection", protection.to_xml(), before=("bookViews", "sheets"))
        self._apply_parts({"xl/workbook.xml": xml})

    def update_cells(self, sheet_name: str, cells: Mapping[str, object]) -> None:
        """Update sparse cells, preserving styles, notes and untouched cell XML."""
        if not isinstance(cells, Mapping):
            raise TypeError("cells must map cell references to values")
        path, xml, root, fast = self._sheet_scan_source(sheet_name)
        values = {}
        for ref, value in cells.items():
            ref = cell_ref(ref)
            if ref in values:
                raise ValueError(f"duplicate cell reference: {ref}")
            values[ref] = unwrap_cell(value)
        if not values:
            return
        scan = None if fast is None else _fast_scan(xml, fast, list(values))
        if scan is None:
            scan = _tree_scan(root if root is not None else safe_fromstring(xml), values)
        if scan.merges:
            merged = RangeIndex()
            for merge in scan.merges:
                merged.add(parse_cell_range(merge))
            for ref in values:
                col, row, _, _ = parse_cell_range(ref)
                for c1, r1, _, _, _ in merged.overlapping((col, row, col, row)):
                    if (col, row) != (c1, r1):
                        raise ValueError("merged cells must be updated through their top-left cell")
        # Refuse edits into array/shared formula groups; writing one member alone
        # would leave the other cells dependent on a missing formula master.
        _refuse_formula_groups(scan, values)
        for ref, value in values.items():
            if isinstance(value, Formula) and value.array_ref is not None:
                raise ValueError("update_cells supports scalar formulas only")
            if isinstance(value, Formula) and value.spill_ref is not None:
                raise UnsupportedFeatureError(
                    "update_cells cannot write dynamic-array formulas"
                )
        sheet_data = fast.sheet_data if fast is not None else sheet_data_match(xml)
        if sheet_data is None:
            raise ValueError("Worksheet sheetData element was not found")
        prefix = sheet_data["prefix"] or ""
        # Validate and encode all input before staging any mutation.
        encoded = {ref: _cell_xml_value(value, self._date_1904, prefix) for ref, value in values.items()}
        pending = dict(encoded)
        by_row_targets: dict[int, dict[str, Any]] = {}
        for ref, value in encoded.items():
            by_row_targets.setdefault(parse_cell_range(ref)[1] + 1, {})[ref] = value

        def patch_row(row_num: int, row_xml: str) -> str:
            # Popped, so a second <row> with the same number stays as it is.
            return _patch_row(row_xml, by_row_targets.pop(row_num), pending, prefix, row_num)

        body = patch_rows(sheet_data[0], patch_row, by_row_targets.keys())
        # The cells left in ``pending`` are in rows the sheet does not have yet.
        by_row: dict[int, list[str]] = {}
        for ref in pending:
            by_row.setdefault(parse_cell_range(ref)[1] + 1, []).append(ref)
        new_rows = {}
        for row_num, refs in by_row.items():
            content = "".join(_cell_xml(ref, encoded[ref], prefix) for ref in sorted(refs, key=lambda r: parse_cell_range(r)[0]))
            new_rows[row_num] = f'<{prefix}row r="{row_num}">{content}</{prefix}row>'
        body = insert_rows(body, new_rows)
        xml = xml[:sheet_data.start()] + body + xml[sheet_data.end():]
        # Expand dimensions, never shrink existing coverage or filters/tables.
        refs = list(values)
        if scan.dimension:
            refs.extend(scan.dimension.split(":"))
        spans: list[Sequence[int]] = [parse_cell_range(ref) for ref in refs]
        if scan.bounds is not None:
            spans.append(scan.bounds)
        first = f"{column_label(min(b[0] for b in spans))}{min(b[1] for b in spans) + 1}"
        last = f"{column_label(max(b[2] for b in spans))}{max(b[3] for b in spans) + 1}"
        xml = set_dimension_ref(xml, f"{first}:{last}")
        updates = {path: xml}
        removed: list[str] = []
        removed_shared_refs = sum(kind == "s" for _, kind, _ in scan.edited)
        if removed_shared_refs:
            rels = safe_fromstring(self._package.get("xl/_rels/workbook.xml.rels"))
            for relationship in rels:
                if relationship.get("Type") == REL + "/sharedStrings":
                    part = resolve_part("xl/workbook.xml", relationship.attrib["Target"])
                    strings = self._package.get(part).decode("utf-8")
                    opening = re.search(r"<(?:[\w.-]+:)?sst\b[^>]*>", strings)
                    if opening is None:
                        raise ValueError("sharedStrings element was not found")
                    attrs = tag_attributes(opening[0])
                    if "count" in attrs:
                        updated_tag = replace_or_add_xml_attribute(opening[0], "count", max(0, int(attrs["count"]) - removed_shared_refs))
                        updates[part] = strings[:opening.start()] + updated_tag + strings[opening.end():]
        old_formula = any(has_formula for _, _, has_formula in scan.edited)
        if old_formula or any(isinstance(value, Formula) for value in values.values()):
            self._recalculate(updates, removed)
        self._apply_parts(updates, removed)
        if self._shared_strings_loaded:
            self._shared_strings_total = max(0, self._shared_strings_total - removed_shared_refs)
        if self._stream_shared_strings is not None:
            self._stream_shared_strings.total_count = max(0, self._stream_shared_strings.total_count - removed_shared_refs)

    def _drop_stale_calc_chain(self, *, had_formulas: bool, has_formulas: bool) -> None:
        """Drop calcChain when a sheet's formula cells are being replaced.

        calcChain lists every formula cell. An entry that names a cell without a
        formula makes Excel repair the workbook. Dropping the part is safe because
        Excel rebuilds it, and fullCalcOnLoad makes the formulas recalculate.
        """
        if not (had_formulas or has_formulas):
            return
        updates: dict = {}
        removed: list = []
        self._recalculate(updates, removed)
        # Written straight to the package: the streaming path already runs inside
        # a package transaction, and _apply_parts cannot nest one.
        for path, data in updates.items():
            self._package.set(path, data.encode("utf-8") if isinstance(data, str) else data)
        for path in removed:
            self._package.remove(path)

    def _recalculate(self, updates: MutableMapping[str, Any], removed: list[str]) -> None:
        xml = updates.get("xl/workbook.xml") or self._package.get("xl/workbook.xml").decode("utf-8")
        existing = element_pattern("calcPr").search(xml)
        attrs = tag_attributes(existing[0]) if existing else {}
        attrs.update(fullCalcOnLoad="1", forceFullCalc="1", calcMode="auto")
        fragment = f'<calcPr xmlns="{MAIN}" ' + " ".join(f"{key}={quoteattr(value)}" for key, value in attrs.items() if not key.startswith("xmlns")) + "/>"
        updates["xl/workbook.xml"] = replace_element(xml, "calcPr", fragment, before=("oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr", "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst"))
        rp = "xl/_rels/workbook.xml.rels"
        rels = safe_fromstring(self._package.get(rp))
        types = safe_fromstring(self._package.get("[Content_Types].xml"))
        for node in list(rels):
            if node.get("Type") == REL + "/calcChain":
                part = resolve_part("xl/workbook.xml", node.attrib["Target"])
                removed.append(part)
                rels.remove(node)
                for ct in list(types):
                    if ct.get("PartName") == "/" + part:
                        types.remove(ct)
        updates[rp] = _bytes(rels)
        updates["[Content_Types].xml"] = _bytes(types)


@dataclass
class _EditScan:
    """What update_cells needs to know of a sheet before it edits it."""

    bounds: tuple[int, int, int, int] | None  # of all cells, zero-based
    edited: list[tuple[str, str | None, bool]]  # reference, ``t`` and whether it has a formula
    shared_hit: int | None  # index, among all <f>, of the first shared formula of an edited group
    arrays: list[tuple[str, int]]  # reference and index of each array / data-table formula
    merges: list[str]
    dimension: str | None


@dataclass
class _PlainSheet:
    """A worksheet part in the plain markup the native scan reads."""

    sheet_data: Any
    merges: list[str]
    dimension: str | None


_ROOT_OPEN = re.compile(r"\s*(?:<\?[^>]*\?>\s*)?<worksheet(?=[\s>])([^>]*)>")
_DIMENSION_REF = re.compile(r"<dimension\s[^>]*?\bref=\"([^\"]*)\"")
_MERGE_CELL = re.compile(r"<mergeCell\s[^>]*?\bref=\"([^\"]*)\"")
_FORMULA_OPEN = re.compile(r"<f[\s>/]")


def _fast_sheet(xml: str) -> _PlainSheet | None:
    """The places of a worksheet in its text, or None when its markup is not plain:
    unprefixed names, the main namespace declared once, in the root, and no formula
    outside the cells. The caller still has to check that it is well formed."""
    sheet_data = sheet_data_match(xml)
    if sheet_data is None or sheet_data["prefix"]:
        return None
    head, tail = xml[:sheet_data.start()], xml[sheet_data.end():]
    root = _ROOT_OPEN.match(head)
    if root is None or f'xmlns="{MAIN}"' not in root[1] or xml.count("xmlns") != root[1].count("xmlns"):
        return None
    if _FORMULA_OPEN.search(head) or _FORMULA_OPEN.search(tail):
        return None
    dimension = _DIMENSION_REF.search(head)
    merges = _MERGE_CELL.findall(tail)
    if (dimension is not None and "&" in dimension[1]) or any("&" in ref for ref in merges):
        return None
    return _PlainSheet(sheet_data, merges, None if dimension is None else dimension[1])


def _fast_scan(xml: str, fast: _PlainSheet, targets: list[str]) -> _EditScan | None:
    """The native scan of the cells, or None for markup it leaves to the tree."""
    scanned = _accel.xlsx_edit_scan(fast.sheet_data[0], targets)
    if scanned is None:
        return None
    bounds, edited, shared_hit, arrays = scanned
    return _EditScan(bounds, edited, shared_hit, arrays, fast.merges, fast.dimension)


def _tree_scan(root: Any, values: Mapping[str, object]) -> _EditScan:
    """The twin of the native scan, over the element tree: it answers the same for
    every sheet the native scan reads (tests/test_xlsx_edit_scan.py compares them)."""
    # The cells being edited, and the bounds of all of them: a sheet has far
    # more cells than an edit has targets.
    edited_nodes = []
    edited_rows = {parse_cell_range(ref)[1] + 1 for ref in values}
    cell_bounds = None
    next_row = 1
    for row_node in root.findall(f"{{{MAIN}}}sheetData/{{{MAIN}}}row"):
        row_number = int(row_node.get("r", next_row))
        next_row = row_number + 1
        next_column = 0
        for cell in row_node.findall(f"{{{MAIN}}}c"):
            position, column, row = _located_cell(cell.get("r"), next_column, row_number, edited_rows)
            next_column = column + 1
            if position in values:
                edited_nodes.append((position, cell))
            if cell_bounds is None:
                cell_bounds = [column, row, column, row]
            else:
                if column < cell_bounds[0]:
                    cell_bounds[0] = column
                elif column > cell_bounds[2]:
                    cell_bounds[2] = column
                if row < cell_bounds[1]:
                    cell_bounds[1] = row
                elif row > cell_bounds[3]:
                    cell_bounds[3] = row
    groups = {f.get("si") for _, c in edited_nodes for f in c.findall(f"{{{MAIN}}}f") if f.get("t") == "shared"}
    first_in_group: dict[str | None, int] = {}
    arrays = []
    for index, formula in enumerate(root.iter(f"{{{MAIN}}}f")):
        kind = formula.get("t")
        if kind == "shared":
            first_in_group.setdefault(formula.get("si"), index)
        elif kind in {"array", "dataTable"} and formula.get("ref"):
            arrays.append((formula.attrib["ref"], index))
    hits = [first_in_group[group] for group in groups if group in first_in_group]
    dimension = root.find(f"{{{MAIN}}}dimension")
    return _EditScan(
        None if cell_bounds is None else tuple(cell_bounds),  # type: ignore[arg-type]
        [(str(position), cell.get("t"), cell.find(f"{{{MAIN}}}f") is not None) for position, cell in edited_nodes],
        min(hits) if hits else None,
        arrays,
        [node.attrib["ref"] for node in root.findall(f"{{{MAIN}}}mergeCells/{{{MAIN}}}mergeCell")],
        dimension.get("ref") if dimension is not None else None,
    )


def _refuse_formula_groups(scan: _EditScan, values: Mapping[str, object]) -> None:
    """Refuse an edit inside a shared formula group or an array formula, in the order
    the formulas come in the sheet (the first one to forbid it gives the message)."""
    targets = [parse_cell_range(ref) for ref in values]
    for ref, index in scan.arrays:
        if scan.shared_hit is not None and scan.shared_hit < index:
            break
        c1, r1, c2, r2 = parse_cell_range(ref)
        if any(c1 <= column <= c2 and r1 <= row <= r2 for column, row, _, _ in targets):
            raise ValueError("editing members of array/data-table formulas is unsupported")
    if scan.shared_hit is not None:
        raise ValueError("editing members of shared formula groups is unsupported")


def _cell_xml_value(value: Any, date_1904: bool, prefix: str) -> tuple[str | None, str]:
    if value is None:
        return None, ""
    if isinstance(value, Formula):
        xml_text(value.text)
        check_cell_text(value.text)
        if len(value.text) > 8192:
            raise ValueError("formula exceeds Excel's 8192-character limit")
        cached = value.value
        kind = "str" if isinstance(cached, str) else None
        if isinstance(cached, str):
            check_cell_text(cached)
        content = f'<{prefix}f>{escape(value.text)}</{prefix}f>'
        if cached is not None:
            text = xstring_escape(cached) if isinstance(cached, str) else str(cached)
            content += f'<{prefix}v>{escape(text)}</{prefix}v>'
        return kind, content
    if isinstance(value, bool):
        return "b", f'<{prefix}v>{int(value)}</{prefix}v>'
    if isinstance(value, (datetime.datetime, datetime.date)):
        return None, f'<{prefix}v>{datetime_to_serial(value, date_1904):.17g}</{prefix}v>'
    if isinstance(value, (numbers.Real, Decimal)):
        if not math.isfinite(value):
            raise ValueError("numeric cell values must be finite")
        text = str(value) if isinstance(value, (int, Decimal)) else repr(float(value))
        return None, f'<{prefix}v>{text}</{prefix}v>'

    if isinstance(value, CellRichText):
        runs = value.runs_xml(prefix)
        runs.encode("utf-8")
        return "inlineStr", f'<{prefix}is>{runs}</{prefix}is>'
    if not isinstance(value, str):
        raise TypeError(f"unsupported cell value: {type(value).__name__}")
    check_cell_text(value)
    # encode also rejects lone surrogates before any package mutation.
    text = escape(xstring_escape(value))
    text.encode("utf-8")
    return "inlineStr", f'<{prefix}is><{prefix}t xml:space="preserve">{text}</{prefix}t></{prefix}is>'


def _cell_xml(ref: str, encoded: tuple[str | None, str], prefix: str, old: str | None = None) -> str:
    kind, content = encoded
    attrs = tag_attributes(old) if old else {"r": ref}
    attrs["r"] = ref
    attrs.pop("t", None)
    if kind:
        attrs["t"] = kind
    opening = f'<{prefix}c ' + " ".join(f"{key}={quoteattr(value)}" for key, value in attrs.items()) + ">"
    # Keep cell extensions and other metadata when replacing the value.
    if old:
        inner = re.sub(r"^<[^>]+>|</[^>]+>$", "", old) if not old.endswith("/>") else ""
        for tag in ("f", "v", "is"):
            inner = element_pattern(tag).sub("", inner)
        content += inner
    return opening + content + f'</{prefix}c>'


# An already normal reference, the usual case: no ``$``, upper case, no range.
_PLAIN_CELL_REF = re.compile(r"([A-Z]{1,3})([1-9][0-9]{0,6})")


def _located_cell(reference: str | None, next_column: int, row_number: int,
                   wanted_rows: Collection[int]) -> tuple[str | None, int, int]:
    """A cell's normal reference and zero-based column and row. Without an ``r``
    attribute it is at the column after the previous cell in the row, and the
    reference is only made for the ``wanted_rows`` (None otherwise).
    """
    if reference is not None:
        match = _PLAIN_CELL_REF.fullmatch(reference)
        if match is not None:
            column = _column_index(match[1])
            row = int(match[2]) - 1
            if column < 16384 and row < 1048576:
                return reference, column, row
    else:
        if next_column < 16384 and 1 <= row_number <= 1048576 and row_number not in wanted_rows:
            return None, next_column, row_number - 1
        reference = f"{column_label(next_column)}{row_number}"
    reference = cell_ref(reference)
    column, row, _, _ = parse_cell_range(reference)
    return reference, column, row


def _patch_row(row_xml: str, targets: dict[str, Any], pending: dict[str, Any], prefix: str, row_num: int) -> str:
    additions = dict(targets)
    next_col = 0
    def patch_cell(match: re.Match[str]) -> str:
        nonlocal next_col
        ref = cell_ref(tag_attributes(match[0]).get("r", f"{column_label(next_col)}{row_num}"))
        next_col = parse_cell_range(ref)[0] + 1
        result = []
        col = parse_cell_range(ref)[0]
        for new_ref in sorted(additions, key=lambda r: parse_cell_range(r)[0]):
            if parse_cell_range(new_ref)[0] < col:
                result.append(_cell_xml(new_ref, additions.pop(new_ref), prefix))
                pending.pop(new_ref)
        if ref in additions:
            result.append(_cell_xml(ref, additions.pop(ref), prefix, match[0]))
            pending.pop(ref)
        else:
            result.append(match[0])
        return "".join(result)
    row_xml = element_pattern("c").sub(patch_cell, row_xml)
    if additions:
        content = "".join(_cell_xml(ref, additions[ref], prefix) for ref in sorted(additions, key=lambda r: parse_cell_range(r)[0]))
        row_xml = append_child(row_xml, "row", content)
        for ref in additions:
            pending.pop(ref)
    # Existing spans describe the original cells, so discard this optional hint.
    opening_end = row_xml.index(">")
    opening = re.sub(r'''\s+spans\s*=\s*(?:"[^"]*"|'[^']*')''', "", row_xml[:opening_end])
    return opening + row_xml[opening_end:]
