"""Add, remove, rename and reorder the sheets of an XLSB package.

Sheets are listed in ``xl/workbook.bin`` as BrtBundleSh records ([MS-XLSB]
2.4.304). Formulas never hold sheet names: they point into the ExternSheet
table (XTI entries with sheet indexes) and local names store the index of
their sheet. Renaming therefore only rewrites the BrtBundleSh record, while
removing or reordering sheets remaps those indexes, as Excel does: an XTI
entry of a removed sheet becomes ``-1`` (``#REF!``) and names local to it
are dropped.
"""

from __future__ import annotations

from .exceptions import UnsupportedFeatureError
import posixpath
import struct
from xml.etree import ElementTree as ET
from ._safe_xml import fromstring as safe_fromstring

from ._metadata import CT, PKG_REL, REL
from ._xlsx_edit import _bytes as _xml
from ._xlsb_formulas import SUP_SELF, FormulaContext, parsed_formula
from ._xlsb_metadata import read_wide, records, wide_string
from .biff_utils import build_record
from .refs import check_sheet_name
from .updater_utils import renamed_chart_parts

BUNDLE_SH = 0x9C
BOOK_VIEW = 0x9E
WORKSHEET_TYPE = "application/vnd.ms-excel.worksheet"
_STATES = {"visible": 0, "hidden": 1, "veryHidden": 2}
_STATE_NAMES = {value: key for key, value in _STATES.items()}
_DELETED = 1  # flag of a name slot left behind by FormulaContext.remove_name

# A worksheet as Excel writes an empty sheet: properties, dimension, one
# view, the default format, empty sheet data, print options and margins.
_EMPTY_SHEET = b"".join((
    build_record(0x81),
    build_record(0x93, bytes.fromhex("c904020040000000000000ffffffffffffffff00000000")),
    build_record(0x94, bytes(16)),
    build_record(0x85),
    build_record(0x89, bytes.fromhex("dc0300000000000000000000000040000000640000000000000000000000")),
    build_record(0x98, struct.pack("<9I", 3, 0, 0, 0, 1, 0, 0, 0, 0)),
    build_record(0x8A),
    build_record(0x86),
    build_record(0x1E5, struct.pack("<iHHHBB", -1, 8, 300, 0, 0, 0)),
    build_record(0x91),
    build_record(0x92),
    build_record(0x1DD, struct.pack("<H", 0x10)),
    build_record(0x1DC, struct.pack("<6d", 0.7, 0.7, 0.75, 0.75, 0.3, 0.3)),
    build_record(0x82),
))


class _Entry:
    __slots__ = ("old_index", "state", "tab_id", "rid", "name")

    def __init__(self, old_index, state, tab_id, rid, name):
        self.old_index = old_index
        self.state = state
        self.tab_id = tab_id
        self.rid = rid
        self.name = name

    def payload(self) -> bytes:
        return struct.pack("<II", self.state, self.tab_id) + wide_string(self.rid) + wide_string(self.name)


def _bundles(workbook: bytes) -> list[_Entry]:
    entries = []
    for record, payload in records(workbook):
        if record.record_id == BUNDLE_SH:
            state, tab_id = struct.unpack_from("<II", payload)
            rid, offset = read_wide(payload, 8)
            name, _ = read_wide(payload, offset)
            entries.append(_Entry(len(entries), state, tab_id, rid, name))
    return entries


def _next_id(taken, prefix: str) -> str:
    number = 1
    while f"{prefix}{number}" in taken:
        number += 1
    return f"{prefix}{number}"


class XlsbSheetsMixin:
    """Sheet management for ``XlsbUpdater``."""

    # Sheet name -> part path; set by the updater's loader.
    _sheets: dict[str, str]

    def get_sheet_states(self) -> dict[str, str]:
        """Sheet visibility (``visible``, ``hidden`` or ``veryHidden``) by name."""
        return {
            entry.name: _STATE_NAMES.get(entry.state, "visible")
            for entry in _bundles(self._package.get("xl/workbook.bin"))
        }

    def get_active_sheet(self) -> int:
        """Index of the sheet Excel shows when the workbook opens."""
        for record, payload in records(self._package.get("xl/workbook.bin")):
            if record.record_id == BOOK_VIEW and len(payload) >= 28:
                return struct.unpack_from("<I", payload, 24)[0]
        return 0

    def add_sheet(self, name: str, index: int | None = None, *, state: str = "visible") -> str:
        """Add an empty worksheet at ``index`` (default: last); return its part path."""
        check_sheet_name(name, self._sheets)
        if state not in _STATES:
            raise ValueError(f"invalid sheet state {state!r}")
        workbook = self._package.get("xl/workbook.bin")
        entries = _bundles(workbook)
        names = set(self._package.names())
        number = 1
        while f"xl/worksheets/sheet{number}.bin" in names:
            number += 1
        path = f"xl/worksheets/sheet{number}.bin"

        rels_path = "xl/_rels/workbook.bin.rels"
        rels = safe_fromstring(self._package.get(rels_path))
        rid = _next_id({node.get("Id") for node in rels}, "rId")
        ET.SubElement(rels, f"{{{PKG_REL}}}Relationship", {
            "Id": rid, "Type": REL + "/worksheet",
            "Target": posixpath.relpath(path, "xl"),
        })
        types = safe_fromstring(self._package.get("[Content_Types].xml"))
        ET.SubElement(types, f"{{{CT}}}Override", {"PartName": "/" + path, "ContentType": WORKSHEET_TYPE})

        tab_id = max((entry.tab_id for entry in entries), default=0) + 1
        new = _Entry(None, _STATES[state], tab_id, rid, name)
        index = len(entries) if index is None else max(0, min(index, len(entries)))
        entries.insert(index, new)
        parts = {
            path: _EMPTY_SHEET,
            rels_path: _xml(rels),
            "[Content_Types].xml": _xml(types),
        }
        parts["xl/workbook.bin"] = self._rewrite_sheet_list(workbook, entries, parts)
        self._apply_parts(parts)
        self._sheets = {entry.name: (path if entry is new else self._sheets[entry.name]) for entry in entries}
        return path

    def remove_sheet(self, name: str) -> None:
        """Remove a sheet, its part and its relationships.

        References to it become ``#REF!`` and names local to it are dropped.
        """
        path = self._sheets.get(name)
        if path is None:
            raise KeyError(f"Worksheet not found: {name}")
        if len(self._sheets) == 1:
            raise ValueError("a workbook must keep at least one sheet")
        from .workbook._loaded import sheet_has_pivots, sheet_has_slicers

        # As for XLSX: the slicer caches would be left pointing at a removed table.
        if sheet_has_slicers(self._package, path):
            raise UnsupportedFeatureError(f"sheet {name!r} holds table slicers; removing it is not supported yet")
        # And a pivot table, its cache and the workbook's cache list would be left dangling.
        if sheet_has_pivots(self._package, path):
            raise UnsupportedFeatureError(f"sheet {name!r} holds pivot tables; removing it is not supported yet")
        workbook = self._package.get("xl/workbook.bin")
        entries = _bundles(workbook)
        removed = next(entry for entry in entries if entry.name == name)
        entries.remove(removed)

        rels_path = "xl/_rels/workbook.bin.rels"
        rels = safe_fromstring(self._package.get(rels_path))
        for node in list(rels):
            if node.get("Id") == removed.rid:
                rels.remove(node)
        types = safe_fromstring(self._package.get("[Content_Types].xml"))
        for node in list(types):
            if node.get("PartName") == "/" + path:
                types.remove(node)
        parts = {rels_path: _xml(rels), "[Content_Types].xml": _xml(types)}
        parts["xl/workbook.bin"] = self._rewrite_sheet_list(workbook, entries, parts)
        dropped = [path]
        sheet_rels = posixpath.join(posixpath.dirname(path), "_rels", posixpath.basename(path) + ".rels")
        if self._package.has(sheet_rels):
            dropped.append(sheet_rels)
        removed_parts = []
        self._invalidate_calculation(parts, removed_parts)
        self._apply_parts(parts, dropped + removed_parts)
        del self._sheets[name]
        self._chartsheets.discard(name)

    def rename_sheet(self, old: str, new: str) -> None:
        if old not in self._sheets:
            raise KeyError(f"Worksheet not found: {old}")
        if old == new:
            return
        check_sheet_name(new, [name for name in self._sheets if name != old])
        workbook = self._package.get("xl/workbook.bin")
        entries = _bundles(workbook)
        for entry in entries:
            if entry.name == old:
                entry.name = new
        parts = renamed_chart_parts(self._package, old, new)
        parts["xl/workbook.bin"] = self._rewrite_sheet_list(workbook, entries, parts)
        self._apply_parts(parts)
        self._sheets = {(new if name == old else name): path for name, path in self._sheets.items()}
        if old in self._chartsheets:
            self._chartsheets.discard(old)
            self._chartsheets.add(new)

    def reorder_sheets(self, order) -> None:
        """Put the sheets in ``order`` (every sheet name, once)."""
        order = list(order)
        if sorted(order) != sorted(self._sheets):
            raise ValueError("order must list every sheet exactly once")
        workbook = self._package.get("xl/workbook.bin")
        by_name = {entry.name: entry for entry in _bundles(workbook)}
        entries = [by_name[name] for name in order]
        if [entry.old_index for entry in entries] == sorted(entry.old_index for entry in entries):
            return
        parts = {}
        parts["xl/workbook.bin"] = self._rewrite_sheet_list(workbook, entries, parts)
        self._apply_parts(parts)
        self._sheets = {name: self._sheets[name] for name in order}

    def set_sheet_state(self, name: str, state: str) -> None:
        if state not in _STATES:
            raise ValueError(f"invalid sheet state {state!r}")
        workbook = self._package.get("xl/workbook.bin")
        entries = _bundles(workbook)
        entry = next((entry for entry in entries if entry.name == name), None)
        if entry is None:
            raise KeyError(f"Worksheet not found: {name}")
        if entry.state == _STATES[state]:
            return
        entry.state = _STATES[state]
        self._apply_parts({"xl/workbook.bin": self._rewrite_sheet_list(workbook, entries, {})})

    def set_active_sheet(self, index: int) -> None:
        workbook = self._package.get("xl/workbook.bin")
        entries = _bundles(workbook)
        if not 0 <= index < len(entries):
            raise IndexError("active sheet index out of range")
        updated = _patch_book_views(workbook, lambda first, current: (min(first, index), index))
        if updated != workbook:
            self._apply_parts({"xl/workbook.bin": updated})

    def _rewrite_sheet_list(self, workbook: bytes, entries: list[_Entry], parts) -> bytes:
        """``workbook.bin`` with the BrtBundleSh records of ``entries``.

        Sheet indexes in the ExternSheet table, local names and book views are
        remapped from the old to the new positions.
        """
        old_names = [entry.name for entry in _bundles(workbook)]
        mapping = {entry.old_index: new for new, entry in enumerate(entries) if entry.old_index is not None}
        identity = all(old == new for old, new in mapping.items()) and len(mapping) == len(old_names)
        if not identity:
            context = FormulaContext(old_names, workbook)
            _remap_context(context, mapping)
            context._sheets = [entry.name for entry in entries]
            workbook = context.write_workbook(workbook)
        output = []
        wrote = False
        for record, payload in records(workbook):
            if record.record_id == BUNDLE_SH:
                if not wrote:
                    output.extend(build_record(BUNDLE_SH, entry.payload()) for entry in entries)
                    wrote = True
                continue
            output.append(workbook[record.header_start:record.data_end])
        workbook = b"".join(output)
        if not identity:
            count = len(entries)

            def remap_view(first, current):
                current = mapping.get(current, min(current, count - 1))
                first = mapping.get(first, 0)
                # The active sheet must be visible.
                if entries[current].state != 0:
                    current = next((i for i, entry in enumerate(entries) if entry.state == 0), 0)
                return min(first, current), current

            workbook = _patch_book_views(workbook, remap_view)
        return workbook


def _remap_context(context: FormulaContext, mapping: dict[int, int]) -> None:
    links = context.links
    remapped = []
    for entry in context.xti:
        if isinstance(entry, tuple):
            link, first, last = entry
            if link < len(links) and links[link] == SUP_SELF and first >= 0 and last >= 0:
                first, last = mapping.get(first, -1), mapping.get(last, -1)
                if first < 0 or last < 0:
                    first = last = -1
                elif first > last:
                    first, last = last, first
            entry = (link, first, last)
        remapped.append(entry)
    context.xti = remapped
    for index, name in enumerate(context.names):
        if not isinstance(name.scope, int):
            continue
        new_scope = mapping.get(name.scope)
        if new_scope is None:
            # As Excel does: a name local to a removed sheet is dropped. The
            # slot stays, because formulas refer to names by position.
            placeholder = f"_xlspy_deleted_{index}_{len(context.names)}"
            name.name, name.flags, name.scope, name.definition = placeholder, _DELETED, None, None
            name.raw = (
                struct.pack("<IBI", _DELETED, 0, 0xFFFFFFFF) + wide_string(placeholder)
                + parsed_formula(b"\x1c\x1d") + b"\xff" * 4
            )
            continue
        name.scope = new_scope
        if name.raw is not None:
            raw = bytearray(name.raw)
            struct.pack_into("<I", raw, 5, new_scope)
            name.raw = bytes(raw)


def _patch_book_views(workbook: bytes, change) -> bytes:
    output = []
    for record, payload in records(workbook):
        if record.record_id == BOOK_VIEW and len(payload) >= 28:
            first, current = struct.unpack_from("<II", payload, 20)
            first, current = change(first, current)
            value = bytearray(payload)
            struct.pack_into("<II", value, 20, first, current)
            output.append(build_record(BOOK_VIEW, bytes(value)))
        else:
            output.append(workbook[record.header_start:record.data_end])
    return b"".join(output)

