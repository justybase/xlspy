"""Targeted XLSX/XLSM edits, staged atomically through ZipPackage."""

from __future__ import annotations

from collections.abc import Mapping
import datetime
from decimal import Decimal
import math
import numbers
import posixpath
import re
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape, quoteattr

from ._metadata import (
    MAIN, REL, PKG_REL, CT, DefinedName, SheetProtection, WorkbookProtection,
    cell_ref, element_pattern, replace_element, append_child, read_defined_names, xml_text,
)
from ._text import check_cell_text, xstring_escape
from .comments import (
    Comment, COMMENTS_TYPE, VML_TYPE, V, comments_xml, vml_xml, normalize_comments,
    comment_part_targets, find_comment_parts, read_comments, resolve_part, rels_path,
)
from .excel_dates import datetime_to_serial
from .formulas import Formula
from .refs import parse_cell_range, column_label
from .updater_utils import unwrap_cell, replace_or_add_xml_attribute


def _bytes(root):
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


class XlsxEditMixin:
    def _apply_parts(self, parts, removed=()):
        package = self._package
        package.begin_transaction()
        try:
            for path, data in parts.items():
                package.set(path, data.encode("utf-8") if isinstance(data, str) else data)
            for path in removed:
                package.remove(path)
        except BaseException:
            package.rollback_transaction()
            raise
        else:
            package.commit_transaction()

    def _sheet_xml(self, sheet_name):
        try:
            path = self._sheets[sheet_name]
        except KeyError as exc:
            raise KeyError(f"Worksheet not found: {sheet_name}") from exc
        xml = self._package.get(path).decode("utf-8")
        if ET.fromstring(xml).tag != f"{{{MAIN}}}worksheet":
            raise ValueError(f"{sheet_name!r} is not a worksheet")
        return path, xml

    def get_comments(self, sheet_name: str) -> dict[str, Comment]:
        """Read classic note text and authors from the current staged package."""
        path, _ = self._sheet_xml(sheet_name)
        _, _, parts = find_comment_parts(self._package.get, self._package.names(), path)
        return read_comments(self._package.get(parts["comments"][0])) if "comments" in parts else {}

    def set_comments(self, sheet_name: str, comments: Mapping[str, Comment | None]) -> None:
        """Set selected classic notes; None removes a note, other notes are kept."""
        path, sheet = self._sheet_xml(sheet_name)
        merges = [n.attrib["ref"] for n in ET.fromstring(sheet).findall(f"{{{MAIN}}}mergeCells/{{{MAIN}}}mergeCell")]
        changes = normalize_comments(comments, allow_none=True, merges=merges)
        if not changes:
            return
        names = self._package.names()
        legacy = ET.fromstring(sheet).find(f"{{{MAIN}}}legacyDrawing")
        legacy_rid = None if legacy is None else legacy.get(f"{{{REL}}}id")
        rp, relationships, parts = find_comment_parts(self._package.get, names, path, legacy_rid)
        for relationship in relationships:
            if relationship.get("Type", "").rsplit("/", 1)[-1] in {"threadedComment", "threadedComments"}:
                threaded = ET.fromstring(self._package.get(resolve_part(path, relationship.attrib["Target"])))
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

        used_ids = {rel.get("Id") for rel in relationships}
        occupied = set(names)

        def ensure(kind, directory, stem, suffix):
            if kind in parts:
                return parts[kind][0]
            index = 1
            while f"{directory}/{stem}{index}.{suffix}" in occupied:
                index += 1
            part = f"{directory}/{stem}{index}.{suffix}"
            occupied.add(part)
            index = 1
            while f"rId{index}" in used_ids:
                index += 1
            rid = f"rId{index}"
            used_ids.add(rid)
            ET.SubElement(relationships, f"{{{PKG_REL}}}Relationship", {
                "Id": rid, "Type": REL + "/" + kind,
                "Target": posixpath.relpath(part, posixpath.dirname(path)),
            })
            parts[kind] = (part, rid)
            return part

        cp = ensure("comments", "xl/comments", "comment", "xml")
        vp = ensure("vmlDrawing", "xl/drawings", "commentsDrawing", "vml")
        comment_data = comments_xml(changes, self._package.get(cp) if cp in names else None)
        vml_data = vml_xml(changes, self._package.get(vp) if vp in names else None)
        remaining = read_comments(comment_data)
        removed = []
        updates = {cp: comment_data, vp: vml_data}
        types = ET.fromstring(self._package.get("[Content_Types].xml"))

        def drop(part, kind):
            updates.pop(part, None)
            removed.append(part)
            # Only this part's relationship: other vmlDrawing parts may exist.
            for node in list(relationships):
                if node.get("Id") == parts[kind][1]:
                    relationships.remove(node)
            for node in list(types):
                if node.get("PartName") == "/" + part:
                    types.remove(node)

        if not remaining:
            drop(cp, "comments")
        elif not any(n.get("PartName") == "/" + cp for n in types):
            ET.SubElement(types, f"{{{CT}}}Override", {"PartName": "/" + cp, "ContentType": COMMENTS_TYPE})
        has_shapes = any(n.tag == f"{{{V}}}shape" for n in ET.fromstring(vml_data).iter())
        if not has_shapes:
            drop(vp, "vmlDrawing")
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

    def get_defined_names(self) -> list[DefinedName]:
        """Return user-defined global and worksheet-local names."""
        return read_defined_names(self._package.get("xl/workbook.xml"), self._sheets)

    def set_defined_name(self, definition: DefinedName) -> None:
        if not isinstance(definition, DefinedName):
            raise TypeError("definition must be a DefinedName")
        fragment = definition.to_xml(self._sheets)
        xml = self._package.get("xl/workbook.xml").decode("utf-8")
        xml = self._remove_name(xml, definition.name, definition.scope)
        if element_pattern("definedNames").search(xml):
            xml = append_child(xml, "definedNames", fragment)
        else:
            xml = replace_element(xml, "definedNames", f'<definedNames xmlns="{MAIN}">{fragment}</definedNames>', before=("calcPr", "oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr", "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst"))
        self._apply_parts({"xl/workbook.xml": xml})

    def _remove_name(self, xml, name, scope):
        scope_id = None if scope is None else str(list(self._sheets).index(scope))
        def remove(match):
            attrs = _attributes(match[0])
            return "" if attrs.get("name", "").casefold() == name.casefold() and attrs.get("localSheetId") == scope_id else match[0]
        updated = element_pattern("definedName").sub(remove, xml)
        if not element_pattern("definedName").search(updated):
            updated = replace_element(updated, "definedNames", "")
        return updated

    def remove_defined_name(self, name: str, scope: str | None = None) -> None:
        # Validate the name, including the reserved built-in namespace.
        DefinedName(name, "0", scope).to_xml(self._sheets)
        xml = self._package.get("xl/workbook.xml").decode("utf-8")
        updated = self._remove_name(xml, name, scope)
        if updated != xml:
            self._apply_parts({"xl/workbook.xml": updated})

    def set_sheet_protection(self, sheet_name: str, protection: SheetProtection) -> None:
        if not isinstance(protection, SheetProtection):
            raise TypeError("protection must be SheetProtection")
        path, xml = self._sheet_xml(sheet_name)
        xml = replace_element(xml, "sheetProtection", protection.to_xml(), before=("protectedRanges", "scenarios", "autoFilter", "sortState", "dataConsolidate", "customSheetViews", "mergeCells", "phoneticPr", "conditionalFormatting", "dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks", "drawing", "legacyDrawing", "tableParts", "extLst"))
        self._apply_parts({path: xml})

    def set_workbook_protection(self, protection: WorkbookProtection) -> None:
        if not isinstance(protection, WorkbookProtection):
            raise TypeError("protection must be WorkbookProtection")
        xml = self._package.get("xl/workbook.xml").decode("utf-8")
        xml = replace_element(xml, "workbookProtection", protection.to_xml(), before=("bookViews", "sheets"))
        self._apply_parts({"xl/workbook.xml": xml})

    def update_cells(self, sheet_name: str, cells: Mapping[str, object]) -> None:
        """Update sparse cells, preserving styles, notes and untouched cell XML."""
        if not isinstance(cells, Mapping):
            raise TypeError("cells must map cell references to values")
        path, xml = self._sheet_xml(sheet_name)
        values = {}
        for ref, value in cells.items():
            ref = cell_ref(ref)
            if ref in values:
                raise ValueError(f"duplicate cell reference: {ref}")
            values[ref] = unwrap_cell(value)
        if not values:
            return
        root = ET.fromstring(xml)
        cell_nodes = []
        next_row = 1
        for row_node in root.findall(f"{{{MAIN}}}sheetData/{{{MAIN}}}row"):
            row_number = int(row_node.get("r", next_row))
            next_row = row_number + 1
            next_column = 0
            for cell in row_node.findall(f"{{{MAIN}}}c"):
                ref = cell_ref(cell.get("r", f"{column_label(next_column)}{row_number}"))
                next_column = parse_cell_range(ref)[0] + 1
                cell_nodes.append((ref, cell))
        for merge in root.findall(f"{{{MAIN}}}mergeCells/{{{MAIN}}}mergeCell"):
            bounds = parse_cell_range(merge.attrib["ref"])
            for ref in values:
                col, row, _, _ = parse_cell_range(ref)
                if bounds[0] <= col <= bounds[2] and bounds[1] <= row <= bounds[3] and (col, row) != bounds[:2]:
                    raise ValueError("merged cells must be updated through their top-left cell")
        # Refuse edits into array/shared formula groups; writing one member alone
        # would leave the other cells dependent on a missing formula master.
        shared_ids = {f.get("si") for ref, c in cell_nodes
                      if ref in values for f in c.findall(f"{{{MAIN}}}f") if f.get("t") == "shared"}
        for formula in root.iter(f"{{{MAIN}}}f"):
            if formula.get("t") == "shared" and formula.get("si") in shared_ids:
                raise ValueError("editing members of shared formula groups is unsupported")
            if formula.get("t") in {"array", "dataTable"} and formula.get("ref"):
                c1, r1, c2, r2 = parse_cell_range(formula.attrib["ref"])
                if any(c1 <= parse_cell_range(ref)[0] <= c2 and r1 <= parse_cell_range(ref)[1] <= r2 for ref in values):
                    raise ValueError("editing members of array/data-table formulas is unsupported")
        for ref, value in values.items():
            if isinstance(value, Formula) and value.array_ref is not None:
                raise ValueError("update_cells supports scalar formulas only")
        prefix_match = re.search(r"<([\w.-]+:)?sheetData\b", xml)
        if not prefix_match:
            raise ValueError("Worksheet sheetData element was not found")
        prefix = prefix_match[1] or ""
        # Validate and encode all input before staging any mutation.
        encoded = {ref: _cell_value(value, self._date_1904, prefix) for ref, value in values.items()}
        sheet_data = element_pattern("sheetData").search(xml)
        body = sheet_data[0]
        pending = dict(encoded)
        next_row = 1
        row_numbers = []

        def patch_row(match):
            nonlocal next_row
            row_xml = match[0]
            row_num = int(_attributes(row_xml).get("r", next_row))
            next_row = row_num + 1
            row_numbers.append(row_num)
            targets = {ref: value for ref, value in pending.items() if parse_cell_range(ref)[1] + 1 == row_num}
            if not targets:
                return row_xml
            return _patch_row(row_xml, targets, pending, prefix, row_num)

        body = element_pattern("row").sub(patch_row, body)
        # Insert new rows at their numeric position without changing old rows.
        by_row = {}
        for ref in pending:
            by_row.setdefault(parse_cell_range(ref)[1] + 1, []).append(ref)
        new_rows = {}
        for row_num, refs in by_row.items():
            content = "".join(_cell_xml(ref, encoded[ref], prefix) for ref in sorted(refs, key=lambda r: parse_cell_range(r)[0]))
            new_rows[row_num] = f'<{prefix}row r="{row_num}">{content}</{prefix}row>'
        number_iterator = iter(row_numbers)
        def insert_rows(match):
            row_num = next(number_iterator)
            before = "".join(new_rows.pop(n) for n in sorted(list(new_rows)) if n < row_num)
            return before + match[0]
        body = element_pattern("row").sub(insert_rows, body)
        if new_rows:
            tail = "".join(new_rows[n] for n in sorted(new_rows))
            # append_child also expands an empty <sheetData/>.
            body = append_child(body, "sheetData", tail)
        xml = xml[:sheet_data.start()] + body + xml[sheet_data.end():]
        # Expand dimensions, never shrink existing coverage or filters/tables.
        refs = list(values) + [ref for ref, _ in cell_nodes]
        dimension = root.find(f"{{{MAIN}}}dimension")
        if dimension is not None and dimension.get("ref"):
            refs.extend(dimension.attrib["ref"].split(":"))
        bounds = [parse_cell_range(ref) for ref in refs]
        first = f"{column_label(min(b[0] for b in bounds))}{min(b[1] for b in bounds) + 1}"
        last = f"{column_label(max(b[2] for b in bounds))}{max(b[3] for b in bounds) + 1}"
        xml = replace_element(xml, "dimension", f'<dimension xmlns="{MAIN}" ref="{first}:{last}"/>', before=("sheetViews", "sheetFormatPr", "cols", "sheetData"))
        updates = {path: xml}
        removed = []
        removed_shared_refs = sum(ref in values and c.get("t") == "s" for ref, c in cell_nodes)
        if removed_shared_refs:
            rels = ET.fromstring(self._package.get("xl/_rels/workbook.xml.rels"))
            for relationship in rels:
                if relationship.get("Type") == REL + "/sharedStrings":
                    part = resolve_part("xl/workbook.xml", relationship.attrib["Target"])
                    strings = self._package.get(part).decode("utf-8")
                    opening = re.search(r"<(?:[\w.-]+:)?sst\b[^>]*>", strings)
                    attrs = _attributes(opening[0])
                    if "count" in attrs:
                        updated_tag = replace_or_add_xml_attribute(opening[0], "count", max(0, int(attrs["count"]) - removed_shared_refs))
                        updates[part] = strings[:opening.start()] + updated_tag + strings[opening.end():]
        old_formula = any(ref in values and c.find(f"{{{MAIN}}}f") is not None for ref, c in cell_nodes)
        if old_formula or any(isinstance(value, Formula) for value in values.values()):
            self._recalculate(updates, removed)
        self._apply_parts(updates, removed)
        if self._shared_strings_loaded:
            self._shared_strings_total = max(0, self._shared_strings_total - removed_shared_refs)
        if self._stream_shared_strings is not None:
            self._stream_shared_strings.total_count = max(0, self._stream_shared_strings.total_count - removed_shared_refs)

    def _recalculate(self, updates, removed):
        xml = self._package.get("xl/workbook.xml").decode("utf-8")
        existing = element_pattern("calcPr").search(xml)
        attrs = _attributes(existing[0]) if existing else {}
        attrs.update(fullCalcOnLoad="1", forceFullCalc="1", calcMode="auto")
        fragment = f'<calcPr xmlns="{MAIN}" ' + " ".join(f"{key}={quoteattr(value)}" for key, value in attrs.items() if not key.startswith("xmlns")) + "/>"
        updates["xl/workbook.xml"] = replace_element(xml, "calcPr", fragment, before=("oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr", "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst"))
        rp = "xl/_rels/workbook.xml.rels"
        rels = ET.fromstring(self._package.get(rp))
        types = ET.fromstring(self._package.get("[Content_Types].xml"))
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


def _attributes(fragment):
    opening = re.match(r"<[\w:.-]+\b([^>]*?)(?:/?>)", fragment)
    if not opening:
        raise ValueError("XML opening tag was not found")
    # Attribute values can contain escaped XML; ET resolves those consistently.
    attrs = {}
    for match in re.finditer(r'''([\w:.-]+)\s*=\s*("[^"]*"|'[^']*')''', opening[1]):
        from xml.sax.saxutils import unescape
        attrs[match[1]] = unescape(match[2][1:-1], {"&quot;": '"', "&apos;": "'"})
    return attrs


def _cell_value(value, date_1904, prefix):
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
    if not isinstance(value, str):
        raise TypeError(f"unsupported cell value: {type(value).__name__}")
    check_cell_text(value)
    # encode also rejects lone surrogates before any package mutation.
    text = escape(xstring_escape(value))
    text.encode("utf-8")
    return "inlineStr", f'<{prefix}is><{prefix}t xml:space="preserve">{text}</{prefix}t></{prefix}is>'


def _cell_xml(ref, encoded, prefix, old=None):
    kind, content = encoded
    attrs = _attributes(old) if old else {"r": ref}
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


def _patch_row(row_xml, targets, pending, prefix, row_num):
    additions = dict(targets)
    next_col = 0
    def patch_cell(match):
        nonlocal next_col
        ref = cell_ref(_attributes(match[0]).get("r", f"{column_label(next_col)}{row_num}"))
        next_col = parse_cell_range(ref)[0] + 1
        result = []
        col = parse_cell_range(ref)[0]
        for new_ref in sorted(list(additions), key=lambda r: parse_cell_range(r)[0]):
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
