"""Opt-in XLSB template features. Bulk data loops never call these codecs."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from .exceptions import UnsupportedFeatureError
from collections.abc import Callable, Iterable, Mapping, Sequence
import datetime
from decimal import Decimal
import math
import numbers
import posixpath
import struct
from xml.etree import ElementTree as ET
from ._safe_xml import fromstring as safe_fromstring

from ._metadata import CT, REL, PKG_REL, DefinedName, SheetProtection, WorkbookProtection, cell_ref
from ._text import check_cell_text
from ._xlsb_metadata import (
    _XLSB_COMMENTS_TYPE, LEGACY_DRAWING, SHEET_PROTECTION, SHEET_PROTECTION_ISO,
    BOOK_PROTECTION, BOOK_PROTECTION_ISO, records, replace_records, parse_comments,
    comments_bin, wide_string, read_wide, sheet_protection_bin,
    workbook_protection_bin, check_merged_refs,
)
from ._xlsb_formulas import FormulaContext, formula_cell
from ._xlsx_edit import XlsxEditMixin
from .biff_utils import build_record
from .comments import Comment, V, VML_TYPE, comment_part_targets, find_comment_parts, normalize_comments, vml_xml, rels_path, resolve_part
from .excel_dates import datetime_to_serial
from .formulas import Formula
from .refs import parse_cell_range
from .updater_utils import unwrap_cell
from ._xlsx_edit import _HYPERLINK_TYPE as _HYPERLINK_TYPE


_MERGE_RECORDS = frozenset({0xB0, 0xB1, 0xB2})
# Worksheet records that follow the merged cells / the hyperlinks in schema
# order ([MS-XLSB] 2.1.7.62): a new block is inserted before the first one found.
_BRT_AFTER_HYPERLINKS = (0x1DD, 0x1DC, 0x1DE, 0x1DF, 0x188, 0x18A, 0x226, 0x227, 0x294, 0x82)
_AFTER_MERGES = (0x1CD, 0x23D, 0x1EE) + _BRT_AFTER_HYPERLINKS


def request_recalculation(data):
    chunks = []
    for record, payload in records(data):
        if record.record_id == 0x9D:
            if len(payload) != 26:
                raise ValueError("invalid XLSB calculation properties")
            value = bytearray(payload)
            struct.pack_into('<II', value, 0, 0, 1)
            flags = struct.unpack_from('<H', value, 24)[0]
            struct.pack_into('<H', value, 24, flags | 1 | 0x10)
            chunks.append(build_record(0x9D, value))
        else:
            chunks.append(data[record.header_start:record.data_end])
    return b''.join(chunks)


class XlsbEditMixin(XlsxEditMixin):
    # Set by XlsbUpdater; None until the shared strings part is read.
    _shared_strings_data: bytes | None

    if TYPE_CHECKING:
        from .data_validation import DataValidation
        from .hyperlinks import Hyperlink

        # Provided by XlsbUpdater; the attributes it shares with XlsxEditMixin are
        # declared there.
        _shared_strings_loaded: bool

        def _part_or_none(self, path: str) -> bytes | None: ...

    def _sheet_bin(self, sheet_name: str) -> tuple[str, bytes]:
        try:
            path = self._sheets[sheet_name]
        except KeyError as exc:
            raise KeyError(f"Worksheet not found: {sheet_name}") from exc
        data = self._package.get(path)
        if not any(record.record_id == 0x91 for record, _ in records(data)):
            raise ValueError(f"{sheet_name!r} is not a worksheet")
        return path, data

    def get_comments(self, sheet_name: str) -> dict[str, Comment]:
        path, _ = self._sheet_bin(sheet_name)
        _, _, parts = find_comment_parts(self._package.get, self._package.names(), path)
        return parse_comments(self._package.get(parts['comments'][0]))[0] if 'comments' in parts else {}

    def set_comments(self, sheet_name: str, comments: Mapping[str, Comment | None]) -> None:
        path, data = self._sheet_bin(sheet_name)
        changes = normalize_comments(comments, allow_none=True)
        check_merged_refs(data, [ref for ref, value in changes.items() if value is not None])
        names = self._package.names()
        legacy = [read_wide(p)[0] for r, p in records(data) if r.record_id == LEGACY_DRAWING]
        if len(legacy) > 1:
            raise ValueError("unsupported legacy drawing relationship")
        rp, relationships, parts = find_comment_parts(self._package.get, names, path, legacy[0] if legacy else None)
        for node in relationships:
            if node.get('Type', '').rsplit('/', 1)[-1] in {'threadedComment', 'threadedComments'}:
                root = safe_fromstring(self._package.get(resolve_part(path, node.attrib['Target'])))
                occupied = {cell_ref(n.attrib['ref']) for n in root if 'ref' in n.attrib}
                if occupied & changes.keys():
                    raise ValueError("threaded comment editing is unsupported")
        current = parse_comments(self._package.get(parts['comments'][0]))[0] if 'comments' in parts else {}
        changes = {ref: value for ref, value in changes.items() if value is not None or ref in current}
        if not changes:
            return
        for other in self._sheets.values():
            if other == path:
                continue
            other_parts = comment_part_targets(self._package.get, names, other)
            if any(p[0] in other_parts for p in parts.values()):
                raise ValueError("comment/VML parts shared by sheets cannot be edited")
        if legacy and 'vmlDrawing' not in parts:
            raise ValueError("unsupported legacy drawing relationship")
        used = {node.get('Id') for node in relationships}
        occupied = set(names)
        def ensure(kind, directory, stem, extension):
            if kind in parts:
                return parts[kind][0]
            index = 1
            while f'{directory}/{stem}{index}.{extension}' in occupied:
                index += 1
            part = f'{directory}/{stem}{index}.{extension}'
            occupied.add(part)
            index = 1
            while f'rId{index}' in used:
                index += 1
            rid = f'rId{index}'
            used.add(rid)
            ET.SubElement(relationships, f'{{{PKG_REL}}}Relationship', {
                'Id': rid, 'Type': REL + '/' + kind,
                'Target': posixpath.relpath(part, posixpath.dirname(path)),
            })
            parts[kind] = (part, rid)
            return part
        cp = ensure('comments', 'xl/comments', 'comment', 'bin')
        vp = ensure('vmlDrawing', 'xl/drawings', 'commentsDrawing', 'vml')
        note_data = comments_bin(changes, self._package.get(cp) if cp in names else None)
        drawing = vml_xml(changes, self._package.get(vp) if vp in names else None)
        remaining = parse_comments(note_data)[0]
        types = safe_fromstring(self._package.get('[Content_Types].xml'))
        updates, removed = {cp: note_data, vp: drawing}, []
        def drop(part, kind):
            updates.pop(part, None)
            removed.append(part)
            # Only this part's relationship: other vmlDrawing parts may exist.
            for node in list(relationships):
                if node.get('Id') == parts[kind][1]:
                    relationships.remove(node)
            for node in list(types):
                if node.get('PartName') == '/' + part:
                    types.remove(node)
        if remaining:
            if not any(node.get('PartName') == '/' + cp for node in types):
                ET.SubElement(types, f'{{{CT}}}Override', {'PartName': '/' + cp, 'ContentType': _XLSB_COMMENTS_TYPE})
        else:
            drop(cp, 'comments')
        if any(node.tag == f'{{{V}}}shape' for node in safe_fromstring(drawing).iter()):
            if not any(node.get('PartName') == '/' + vp or node.get('Extension') == 'vml' for node in types):
                ET.SubElement(types, f'{{{CT}}}Override', {'PartName': '/' + vp, 'ContentType': VML_TYPE})
            if not legacy:
                data = replace_records(data, {LEGACY_DRAWING}, build_record(LEGACY_DRAWING, wide_string(parts['vmlDrawing'][1])), before=(0x228, 0x229, 0x27F, 0x294, 0x82))
        else:
            drop(vp, 'vmlDrawing')
            data = replace_records(data, {LEGACY_DRAWING}, before=(0x82,))
        updates[path] = data
        if len(relationships):
            updates[rp] = ET.tostring(relationships, encoding='utf-8')
        elif rp in names:
            removed.append(rp)
        updates['[Content_Types].xml'] = ET.tostring(types, encoding='utf-8')
        self._apply_parts(updates, removed)

    def get_defined_names(self) -> list[DefinedName]:
        return FormulaContext(self._sheets, self._package.get('xl/workbook.bin'), self._part_or_none).defined_names()

    def set_defined_name(self, definition: DefinedName) -> None:
        if not isinstance(definition, DefinedName):
            raise TypeError("definition must be DefinedName")
        if definition.scope is not None and definition.scope not in self._sheets:
            raise KeyError(f"Worksheet not found: {definition.scope}")
        original = self._package.get('xl/workbook.bin')
        context = FormulaContext(self._sheets, original)
        context.set_name(definition)
        updates = {'xl/workbook.bin': context.write_workbook(original)}
        removed: list[str] = []
        self._invalidate_calculation(updates, removed)
        self._apply_parts(updates, removed)

    def remove_defined_name(self, name: str, scope: str | None = None) -> None:
        if scope is not None and scope not in self._sheets:
            raise KeyError(f"Worksheet not found: {scope}")
        original = self._package.get('xl/workbook.bin')
        context = FormulaContext(self._sheets, original)
        context.remove_name(name, scope)
        updates = {'xl/workbook.bin': context.write_workbook(original)}
        removed: list[str] = []
        self._invalidate_calculation(updates, removed)
        self._apply_parts(updates, removed)

    def set_sheet_protection(self, sheet_name: str, protection: SheetProtection) -> None:
        if not isinstance(protection, SheetProtection):
            raise TypeError("protection must be SheetProtection")
        path, data = self._sheet_bin(sheet_name)
        data = replace_records(data, {SHEET_PROTECTION, SHEET_PROTECTION_ISO}, sheet_protection_bin(protection), before=(0xA1, 0xB1, 0x1CD, 0x1EE, 0x1DC, 0x226, 0x227, 0x294, 0x82))
        self._apply_parts({path: data})

    def set_merged_cells(self, sheet_name: str, ranges: Iterable[str]) -> None:
        """Replace the sheet's merged ranges with ``ranges`` (A1 range strings)."""
        from .writer import XlsbWriter

        path, data = self._sheet_bin(sheet_name)
        refs = XlsbWriter._normalize_merges(list(ranges)) or []
        block = b""
        if refs:
            block = build_record(0xB1, struct.pack('<I', len(refs)))
            for ref in refs:
                c1, r1, c2, r2 = parse_cell_range(ref)
                block += build_record(0xB0, struct.pack('<IIII', r1, r2, c1, c2))
            block += build_record(0xB2)
        if not block and not any(r.record_id in _MERGE_RECORDS for r, _ in records(data)):
            return
        self._apply_parts({path: replace_records(data, _MERGE_RECORDS, block, before=_AFTER_MERGES)})

    def set_hyperlinks(self, sheet_name: str,  # type: ignore[override]  # XLSB replaces every link
                       links: Mapping[str, Hyperlink]) -> None:
        """Replace all of the sheet's hyperlinks with ``links`` ({cell: Hyperlink}).

        External targets reuse the sheet's hyperlink relationships, new ones get
        fresh ids and relationships no longer used are dropped.
        """
        from .hyperlinks import Hyperlink
        from .writer import _hlink_payload

        path, data = self._sheet_bin(sheet_name)
        entries = []
        for ref, link in links.items():
            if not isinstance(link, Hyperlink):
                raise TypeError("hyperlinks must be xlspy Hyperlink objects")
            col, row, _, _ = parse_cell_range(cell_ref(ref))
            entries.append((row, col, link))
        entries.sort(key=lambda entry: entry[:2])

        rels_name = rels_path(path)
        has_rels = self._package.has(rels_name)
        root = safe_fromstring(self._package.get(rels_name)) if has_rels else ET.Element(f"{{{PKG_REL}}}Relationships")
        by_target = {
            rel.get('Target'): rel.get('Id') for rel in root
            if rel.get('Type') == _HYPERLINK_TYPE and rel.get('TargetMode') == 'External'
        }
        taken = {rel.get('Id') for rel in root}
        used = set()
        rels_changed = False
        block = []
        for row, col, link in entries:
            rid = None
            if not link.is_internal:
                rid = by_target.get(link.target)
                if rid is None:
                    number = 1
                    while f"rId{number}" in taken:
                        number += 1
                    rid = f"rId{number}"
                    taken.add(rid)
                    ET.SubElement(root, f"{{{PKG_REL}}}Relationship", {
                        'Id': rid, 'Type': _HYPERLINK_TYPE, 'Target': link.target,
                        'TargetMode': 'External',
                    })
                    by_target[link.target] = rid
                    rels_changed = True
                used.add(rid)
            block.append(build_record(0x1EE, _hlink_payload(row, col, link, rid)))
        for rel in list(root):
            if rel.get('Type') == _HYPERLINK_TYPE and rel.get('Id') not in used:
                root.remove(rel)
                rels_changed = True
        if not block and not any(r.record_id == 0x1EE for r, _ in records(data)):
            return
        parts = {path: replace_records(data, {0x1EE}, b''.join(block), before=_BRT_AFTER_HYPERLINKS)}
        if rels_changed:
            ET.register_namespace("", PKG_REL)
            parts[rels_name] = (
                b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                + ET.tostring(root, encoding='utf-8', xml_declaration=False)
            )
        self._apply_parts(parts)

    def set_data_validations(self, sheet_name: str, rules: Sequence[DataValidation]) -> None:
        """Replace the sheet's data validation rules with ``rules``."""
        from ._xlsb_validation import BRT_BEGIN_DVALS, BRT_DVAL, BRT_END_DVALS, validations_bin
        from .data_validation import DataValidation

        rules = list(rules)
        for rule in rules:
            if not isinstance(rule, DataValidation):
                raise TypeError("data validations must be xlspy DataValidation objects")
        path, data = self._sheet_bin(sheet_name)
        targets = {BRT_BEGIN_DVALS, BRT_DVAL, BRT_END_DVALS}
        if not rules and not any(r.record_id in targets for r, _ in records(data)):
            return
        original_workbook = self._package.get('xl/workbook.bin')
        context = FormulaContext(self._sheets, original_workbook)
        block = validations_bin(rules, context)
        self._apply_parts({
            path: replace_records(data, targets, block, before=(0x1EE,) + _BRT_AFTER_HYPERLINKS),
            'xl/workbook.bin': context.write_workbook(original_workbook),
        })

    def set_conditional_formats(self, sheet_name: str, formats: Iterable[tuple[str, bool | None, list[Any]]],
                                dxf_id: Callable[[Any], int],
                                links: Mapping[int, bytes] | None = None) -> None:
        """Replace the sheet's conditional formats with ``formats``, an
        iterable of ``(sqref, pivot, rules)``; ``dxf_id(dxf)`` returns the
        index of a rule's differential style in the style sheet. ``links`` maps
        ``id(rule)`` to the ``BrtCFRuleExt`` payload of a rule whose Excel 2010
        extension is kept as it is in the file."""
        from ._xlsb_cf import conditional_formats_bin, insert_extension, without_conditional_formats

        path, data = self._sheet_bin(sheet_name)
        original_workbook = self._package.get('xl/workbook.bin')
        context = FormulaContext(self._sheets, original_workbook)
        block = conditional_formats_bin(formats, context, dxf_id, links)
        data, groups = without_conditional_formats(data, frozenset((links or {}).values()))
        if block:
            data = replace_records(data, (), block, before=(0x23D, 0x1EE) + _BRT_AFTER_HYPERLINKS)
        if groups:
            data = insert_extension(data, groups)
        self._apply_parts({path: data, 'xl/workbook.bin': context.write_workbook(original_workbook)})

    def set_workbook_protection(self, protection: WorkbookProtection) -> None:
        if not isinstance(protection, WorkbookProtection):
            raise TypeError("protection must be WorkbookProtection")
        data = self._package.get('xl/workbook.bin')
        data = replace_records(data, {BOOK_PROTECTION, BOOK_PROTECTION_ISO}, workbook_protection_bin(protection), before=(0x87, 0x8F))
        self._apply_parts({'xl/workbook.bin': data})

    def _invalidate_calculation(self, updates, removed):
        updates['xl/workbook.bin'] = request_recalculation(updates.get('xl/workbook.bin', self._package.get('xl/workbook.bin')))
        rp = 'xl/_rels/workbook.bin.rels'
        rels = safe_fromstring(updates.get(rp, self._package.get(rp)))
        types = safe_fromstring(updates.get('[Content_Types].xml', self._package.get('[Content_Types].xml')))
        for node in list(rels):
            if node.get('Type') == REL + '/calcChain':
                part = resolve_part('xl/workbook.bin', node.attrib['Target'])
                removed.append(part)
                rels.remove(node)
                for ct in list(types):
                    if ct.get('PartName') == '/' + part:
                        types.remove(ct)
        updates[rp] = ET.tostring(rels, encoding='utf-8')
        updates['[Content_Types].xml'] = ET.tostring(types, encoding='utf-8')

    def update_cells(self, sheet_name: str, cells: Mapping[str, object], *,
                     styles: Mapping[str, int] | None = None) -> None:
        """Set cell values; ``styles`` maps addresses to cell XF indexes.

        An address only in ``styles`` keeps its value (or becomes a styled blank).
        """
        if not isinstance(cells, Mapping):
            raise TypeError("cells must map single-cell addresses to values")
        if styles is not None and not isinstance(styles, Mapping):
            raise TypeError("styles must map single-cell addresses to XF indexes")
        path, data = self._sheet_bin(sheet_name)
        values = {}
        for ref, value in cells.items():
            ref = cell_ref(ref)
            if ref in values:
                raise ValueError(f"duplicate cell address {ref}")
            values[ref] = unwrap_cell(value)
        targets = dict(values)
        style_ids = {}
        for ref, xf in (styles or {}).items():
            ref = cell_ref(ref)
            if not isinstance(xf, int) or isinstance(xf, bool) or not 0 <= xf < 0xFFFFFF:
                raise ValueError(f"invalid cell style index {xf!r} for {ref}")
            style_ids[ref] = xf
            targets.setdefault(ref, KEEP)
        if not targets:
            return
        all_records = list(records(data))
        check_merged_refs(data, values, [
            struct.unpack_from('<IIII', payload) for record, payload in all_records if record.record_id == 0xB0
        ])
        for record, payload in all_records:
            if record.record_id in (0x1AA, 0x1AB, 0x1AC):
                r1, r2, c1, c2 = struct.unpack_from('<IIII', payload)
                if any(c1 <= parse_cell_range(ref)[0] <= c2 and r1 <= parse_cell_range(ref)[1] <= r2 for ref in values):
                    raise ValueError("editing members of array/shared/data-table formulas is unsupported")
        has_formulas = any(isinstance(value, Formula) for value in values.values())
        original_workbook = self._package.get('xl/workbook.bin') if has_formulas else None
        context = FormulaContext(self._sheets, original_workbook) if has_formulas else None
        # Encode everything before staging any changes, including all formulas.
        for ref, value in values.items():
            col, row, _, _ = parse_cell_range(ref)
            encode_cell(value, col, 0, context, sheet_name, row, self._date_1904)
        begin = next(r for r, _ in all_records if r.record_id == 0x91)
        end = next(r for r, _ in all_records if r.record_id == 0x92)
        row_starts = [(r, payload) for r, payload in all_records if r.record_id == 0 and begin.data_end <= r.header_start < end.header_start]
        by_row: dict[Any, Any] = {}
        for ref, value in targets.items():
            col, row, _, _ = parse_cell_range(ref)
            by_row.setdefault(row, {})[col] = (value, style_ids.get(ref))
        output = [data[:begin.data_end]]
        cursor = begin.data_end
        removed_shared_refs = 0
        pending = sorted(by_row)
        next_pending = 0
        for index, (row_record, payload) in enumerate(row_starts):
            row = struct.unpack_from('<I', payload)[0]
            row_end = row_starts[index + 1][0].header_start if index + 1 < len(row_starts) else end.header_start
            output.append(data[cursor:row_record.header_start])
            while next_pending < len(pending) and pending[next_pending] < row:
                new_row = pending[next_pending]
                next_pending += 1
                if new_row in by_row:
                    output.append(new_row_bytes(new_row, by_row.pop(new_row), context, sheet_name, self._date_1904))
            if row in by_row:
                edited, count = patch_row(data[row_record.header_start:row_end], by_row.pop(row), context, sheet_name, row, self._date_1904)
                output.append(edited)
                removed_shared_refs += count
            else:
                output.append(data[row_record.header_start:row_end])
            cursor = row_end
        output.append(data[cursor:end.header_start])
        for row in sorted(by_row):
            output.append(new_row_bytes(row, by_row[row], context, sheet_name, self._date_1904))
        output.append(data[end.header_start:])
        patched = b''.join(output)
        coordinates = [parse_cell_range(ref) for ref in targets]
        # Existing cells count too: XlsbWriter stores an A1:A1 placeholder.
        current_row, previous_column = None, -1
        for record, payload in all_records:
            if not begin.data_end <= record.header_start < end.header_start:
                continue
            if record.record_id == 0:
                current_row, previous_column = struct.unpack_from('<I', payload)[0], -1
            elif current_row is not None and 1 <= record.record_id <= 0x12:
                column = previous_column + 1 if record.record_id >= 0xC else struct.unpack_from('<I', payload)[0]
                previous_column = column
                if column < 16384:
                    coordinates.append((column, current_row, column, current_row))
        dimension = next(((record, payload) for record, payload in all_records if record.record_id == 0x94), None)
        if dimension is not None:
            r1, r2, c1, c2 = struct.unpack('<IIII', dimension[1])
            coordinates.append((c1, r1, c2, r2))
        bounds = struct.pack('<IIII', min(v[1] for v in coordinates), max(v[3] for v in coordinates), min(v[0] for v in coordinates), max(v[2] for v in coordinates))
        if dimension is not None and dimension[0].data_end <= begin.header_start:
            # Before the sheet data, so its offsets are unchanged in ``patched``.
            record = dimension[0]
            patched = patched[:record.header_start] + build_record(0x94, bounds) + patched[record.data_end:]
        else:
            patched = replace_records(patched, {0x94}, build_record(0x94, bounds), before=(0x85, 0x91))
        updates: dict[str, Any] = {path: patched}
        removed: list[str] = []
        if context is not None:
            updates['xl/workbook.bin'] = context.write_workbook(original_workbook)
        if removed_shared_refs and self._package.has('xl/sharedStrings.bin'):
            sst = self._package.get('xl/sharedStrings.bin')
            chunks = []
            for record, payload in records(sst):
                if record.record_id == 0x9F:
                    count, unique = struct.unpack('<II', payload)
                    chunks.append(build_record(0x9F, struct.pack('<II', max(0, count - removed_shared_refs), unique)))
                else:
                    chunks.append(sst[record.header_start:record.data_end])
            updates['xl/sharedStrings.bin'] = b''.join(chunks)
        self._invalidate_calculation(updates, removed)
        # Binary indexes contain offsets into row data. Discard the affected
        # cache rather than leave stale offsets; Excel can regenerate it.
        rp, rels, _ = find_comment_parts(self._package.get, self._package.names(), path)
        types = safe_fromstring(updates['[Content_Types].xml'])
        index_removed = False
        for node in list(rels):
            if node.get('Type', '').endswith('/xlBinaryIndex'):
                part = resolve_part(path, node.attrib['Target'])
                removed.append(part)
                rels.remove(node)
                index_removed = True
                for ct in list(types):
                    if ct.get('PartName') == '/' + part:
                        types.remove(ct)
        if index_removed:
            updates[rp] = ET.tostring(rels, encoding='utf-8')
            updates['[Content_Types].xml'] = ET.tostring(types, encoding='utf-8')
        self._apply_parts(updates, removed)
        if self._shared_strings_loaded:
            self._shared_strings_total = max(0, self._shared_strings_total - removed_shared_refs)
            if 'xl/sharedStrings.bin' in updates:
                self._shared_strings_data = updates['xl/sharedStrings.bin']
        if self._stream_shared_strings is not None:
            self._stream_shared_strings.total_count = max(0, self._stream_shared_strings.total_count - removed_shared_refs)


def encode_cell(value, column, style, context, sheet, row, date_1904=False):
    header = struct.pack('<II', column, style)
    if value is None:
        return build_record(1, header)
    if isinstance(value, Formula):
        if value.array_ref is not None:
            raise ValueError("update_cells supports scalar formulas only")
        return formula_cell(value, column, style, context, sheet=sheet, row=row)
    if isinstance(value, bool):
        return build_record(4, header + bytes([value]))
    if isinstance(value, (datetime.date, datetime.datetime)):
        return build_record(5, header + struct.pack('<d', datetime_to_serial(value, date_1904)))
    if isinstance(value, (numbers.Real, Decimal)):
        if not math.isfinite(value):
            raise ValueError("numeric cell values must be finite")
        return build_record(5, header + struct.pack('<d', float(value)))
    if not isinstance(value, str):
        raise TypeError(f"unsupported cell value {type(value).__name__}")
    check_cell_text(value)
    return build_record(6, header + wide_string(value))


# A style-only edit: the cell keeps its record (and value), only its XF changes.
KEEP = object()


def _encode_entry(entry, column, current_style, context, sheet, row, date_1904):
    """Encode a ``(value, style)`` edit of a cell that has no record to keep."""
    value, style = entry
    style = current_style if style is None else style
    return encode_cell(None if value is KEEP else value, column, style, context, sheet, row, date_1904)


def new_row_bytes(row, values, context, sheet, date_1904=False):
    payload = bytearray(25)
    struct.pack_into('<I', payload, 0, row)
    struct.pack_into('<I', payload, 8, 300)
    payload[13] = 1
    struct.pack_into('<II', payload, 17, min(values), max(values))
    return build_record(0, payload) + b''.join(
        _encode_entry(values[col], col, 0, context, sheet, row, date_1904) for col in sorted(values)
    )


def patch_row(data, values, context, sheet, row, date_1904):
    pending = dict(values)
    output = []
    removed_shared = 0
    previous_column = -1
    last_column = max(values)
    row_header = None
    for record, payload in records(data):
        kind = record.record_id
        if kind == 0:
            row_header = bytearray(payload)
            continue
        if kind in (0x31, 0x32):
            raise UnsupportedFeatureError("point edits in rows with cell metadata are unsupported")
        if 1 <= kind <= 0x12:
            if kind >= 0xC:
                column = previous_column + 1
                full_kind = kind - 0xB
                full_payload = struct.pack('<I', column) + payload
            else:
                column = struct.unpack_from('<I', payload)[0]
                full_kind, full_payload = kind, payload
            previous_column = column
            last_column = max(last_column, column)
            for new_column in sorted(col for col in pending if col < column):
                output.append(_encode_entry(pending.pop(new_column), new_column, 0, context, sheet, row, date_1904))
            if column in pending:
                value, style = pending.pop(column)
                current = struct.unpack_from('<I', full_payload, 4)[0]
                if value is KEEP:
                    # Keep the upper flag bits (fPhShow) of the style field.
                    new_payload = bytearray(full_payload)
                    struct.pack_into('<I', new_payload, 4, (current & 0xFF000000) | style)
                    output.append(build_record(full_kind, bytes(new_payload)))
                else:
                    output.append(encode_cell(
                        value, column, current if style is None else (current & 0xFF000000) | style,
                        context, sheet, row, date_1904,
                    ))
                    removed_shared += full_kind == 7
            else:
                output.append(build_record(full_kind, full_payload) if kind >= 0xC else data[record.header_start:record.data_end])
        else:
            output.append(data[record.header_start:record.data_end])
    for column in sorted(pending):
        output.append(_encode_entry(pending[column], column, 0, context, sheet, row, date_1904))
    if row_header is None:
        raise ValueError("row header was not found")
    if len(row_header) >= 25:
        first_column = struct.unpack_from('<I', row_header, 17)[0]
        struct.pack_into('<II', row_header, 17, min(first_column, min(values)), last_column)
    return build_record(0, row_header) + b''.join(output), removed_shared
