"""Small, format-neutral helpers used by the existing-file updaters.

The readers and writers in ``xlspy`` deliberately have very different
implementations for XML and BIFF12.  The updater layer only needs a few
shared operations: materialising a row iterable, keeping a table header
valid, and generating safe XML text. ZIP handling lives in
``_zip_package.py`` and the XML text primitives in ``_xml_text.py``; the
names are re-exported here for the existing import paths.
"""

from __future__ import annotations

import mmap
import os
import posixpath
import re
import struct

from dataclasses import dataclass
from typing import IO, Any
from collections.abc import Callable, Iterable, Sequence
from xml.etree import ElementTree as ET
from ._safe_xml import fromstring as safe_fromstring, harden_expat
from xml.parsers import expat
from xml.sax import saxutils

from ._text import xml_space_attribute
from .exceptions import UnsupportedFeatureError
from .comments import rels_path
from .formulas import ArrayFormula, DataTableFormula, Formula
from .refs import RangeIndex, parse_cell_range

from ._zip_package import (
    STREAM_BUFFER_SIZE as STREAM_BUFFER_SIZE,
    DiskStringIndex as DiskStringIndex,
    ZipPackage as ZipPackage,
    copy_file_range as copy_file_range,
    copy_stream as copy_stream,
)
from ._xml_text import (
    column_index_to_letter as column_index_to_letter,
    column_letter_to_index as column_letter_to_index,
    escape_xml_text as escape_xml_text,
    local_xml_name as local_xml_name,
    parse_relationships_xml as parse_relationships_xml,
    replace_or_add_xml_attribute as replace_or_add_xml_attribute,
    resolve_zip_target as resolve_zip_target,
    unescape_xml_text as unescape_xml_text,
    xml_attribute as xml_attribute,
)


UPDATER_MAX_COLUMNS = 16384


def dominant_styles(counts: dict[int, dict[int, int]]) -> dict[int, int]:
    """Per column, the style used most often (the lowest style index on a tie)."""

    return {column: max(styles.items(), key=lambda item: (item[1], -item[0]))[0] for column, styles in counts.items()}


def date_like_format(format_code: str) -> bool:
    # Remove quoted literals and escaped characters before looking for date
    # tokens.  This is intentionally conservative: a false positive only
    # chooses an existing date style for a date value.
    simplified = re.sub(r'"(?:[^"]|"")*"', "", format_code)
    simplified = re.sub(r"\\.", "", simplified).lower()
    simplified = re.sub(r"\[[^\]]*\]", "", simplified)
    return bool(re.search(r"[dy]", simplified) or re.search(r"m", simplified))


def rewrite_shared_strings_xml(
    source_path: str | os.PathLike[str],
    output_path: str | os.PathLike[str],
    pending_values: Iterable[str],
    *,
    total_count: int,
    unique_count: int,
) -> None:
    """Rewrite an XLSX SST without materialising its XML payload."""

    with open(source_path, "rb") as source_stream:
        size = os.fstat(source_stream.fileno()).st_size
        if size == 0:
            raise ValueError("Shared strings XML is empty")
        with mmap.mmap(source_stream.fileno(), 0, access=mmap.ACCESS_READ) as source:
            opening_match = re.search(rb"<(?P<prefix>[A-Za-z_][\w.-]*:)?sst\b[^>]*>", source)
            closing_match = re.search(rb"</(?:[A-Za-z_][\w.-]*:)?sst\s*>\s*$", source)
            if not opening_match or not closing_match:
                raise ValueError("Malformed XLSX sharedStrings.xml")

            opening = opening_match.group(0).decode("utf-8")
            opening = replace_or_add_xml_attribute(opening, "count", total_count)
            opening = replace_or_add_xml_attribute(opening, "uniqueCount", unique_count)
            prefix_match = re.match(r"<(?P<prefix>[A-Za-z_][\w.-]*:)?sst\b", opening)
            if prefix_match is None:
                raise ValueError("Malformed XLSX sharedStrings.xml")
            prefix = prefix_match.group("prefix") or ""

            with open(output_path, "wb") as target:
                copy_file_range(source_path, target, 0, opening_match.start())
                target.write(opening.encode("utf-8"))
                copy_file_range(source_path, target, opening_match.end(), closing_match.start())
                for value in pending_values:
                    target.write(
                        (
                            f"<{prefix}si><{prefix}t{xml_space_attribute(value)}>"
                            f"{escape_xml_text(value)}</{prefix}t></{prefix}si>"
                        ).encode()
                    )
                copy_file_range(source_path, target, closing_match.start(), size)


def row_cells(row: Any) -> list[Any]:
    """The cells of one source row, as a list; text is refused, not split."""

    if isinstance(row, (str, bytes, bytearray)):
        raise TypeError("Each row must be a sequence of cells, not text")
    try:
        return list(row)
    except TypeError as exc:
        raise TypeError("Each row must be an iterable of cells") from exc


def materialize_rows(rows: Iterable[Sequence[Any]]) -> list[list[Any]]:
    """Materialise an iterable of rows while keeping its cell boundaries."""

    return [row_cells(row) for row in rows]


def trim_trailing_empty_rows(rows: list[list[Any]]) -> list[list[Any]]:
    """Drop only rows whose cells are all literally ``None``.

    Empty rows in the middle of a range remain in place.  This mirrors the
    behaviour of the reference implementation and keeps an explicit blank
    row meaningful for downstream spreadsheet consumers.
    """

    last = len(rows)
    while last > 0 and all(cell is None for cell in rows[last - 1]):
        last -= 1
    return rows[:last]


class TrailingRowTrimmer:
    """Trailing empty rows of a streaming replacement, cut off at the end.

    The rows are written to ``output`` as they arrive; ``finish`` truncates the
    file after the last row with a value, which is what
    ``trim_trailing_empty_rows`` does to materialized rows. ``width`` grows to
    the widest row kept (an empty row counts once a later row is kept).
    """

    def __init__(self, output: IO[bytes], width: int) -> None:
        self._output = output
        self.width = width
        self.kept_rows = 0  # data rows up to the last one with a value
        self._rows_seen = 0
        self._kept_end = output.tell()
        self._pending_width = 0

    def keep(self) -> None:
        """Keep everything written so far (the header row)."""
        self._kept_end = self._output.tell()

    def wrote(self, row: Sequence[Any]) -> None:
        """Record a data row just written to the output."""
        self._rows_seen += 1
        if all(cell is None for cell in row):
            self._pending_width = max(self._pending_width, len(row))
        else:
            self.width = max(self.width, self._pending_width, len(row))
            self._pending_width = 0
            self.kept_rows = self._rows_seen
            self._kept_end = self._output.tell()

    def finish(self) -> None:
        self._output.truncate(self._kept_end)


def is_formatted_cell(cell: Any) -> bool:
    """Return whether a writer-style ``(value, format_string)`` wrapper exists."""

    return isinstance(cell, tuple) and len(cell) == 2 and isinstance(cell[1], str)


def unwrap_cell(cell: Any) -> Any:
    """Return the value from the optional writer-style format wrapper."""

    return cell[0] if is_formatted_cell(cell) else cell


def check_row_columns(row: Sequence[Any]) -> None:
    """Refuse a row that carries a value past Excel's last column."""

    if len(row) <= UPDATER_MAX_COLUMNS:
        return
    for column in range(UPDATER_MAX_COLUMNS, len(row)):
        if unwrap_cell(row[column]) is not None:
            raise ValueError(
                f"row has a value in column {column + 1}; Excel holds at most {UPDATER_MAX_COLUMNS} columns"
            )


def check_last_column(last_col: int) -> None:
    """Refuse an XLSB row header whose last column is past Excel's last column."""

    if last_col >= UPDATER_MAX_COLUMNS:
        raise ValueError(f"last column is {last_col + 1}; Excel holds at most {UPDATER_MAX_COLUMNS} columns")


def formula_range_bounds(cell: Any) -> tuple[int, int, int, int] | None:
    """The zero-based bounds of a multi-cell formula, else None.

    ``ArrayFormula``, ``DataTableFormula`` and ``Formula(..., array_ref=...)``
    / ``spill_ref=...`` cover a range; a scalar formula or a plain value does
    not.
    """

    value = unwrap_cell(cell)
    if isinstance(value, (ArrayFormula, DataTableFormula)):
        return parse_cell_range(value.ref)
    if isinstance(value, Formula):
        reference = value.array_ref or value.spill_ref
        return None if reference is None else parse_cell_range(reference)
    return None


def check_formula_anchor(bounds: tuple[int, int, int, int], column: int, row: int) -> None:
    """Refuse an array or spill formula outside the top-left cell of its range."""

    c1, r1, _c2, _r2 = bounds
    if (column, row) != (c1, r1):
        raise ValueError("array formula must be written in its top-left cell")


def normalize_merges(merges: Iterable[str]) -> list[str]:
    """The merged ranges in order, refusing two ranges that overlap.

    Excel repairs a workbook whose merged ranges overlap, so both the writers
    and the updaters reject them before writing anything.
    """

    refs: list[tuple[str, int, int, int, int]] = []
    if not merges:
        return []
    index = RangeIndex()
    clash = False
    for ref in merges:
        first_col, first_row, last_col, last_row = parse_cell_range(ref)
        refs.append((ref, first_col, first_row, last_col, last_row))
        bounds = (first_col, first_row, last_col, last_row)
        clash = clash or bool(index.overlapping(bounds))
        index.add(bounds)
    if not clash:
        return [ref for ref, *_ in refs]
    for i, (ref_a, c1a, r1a, c2a, r2a) in enumerate(refs):
        for ref_b, c1b, r1b, c2b, r2b in refs[i + 1 :]:
            if c1a <= c2b and c1b <= c2a and r1a <= r2b and r1b <= r2a:
                raise ValueError(f"overlapping merged ranges {ref_a!r} and {ref_b!r}")
    return [ref for ref, *_ in refs]


_CHART_PART = re.compile(r"xl/charts/chart[^/]*\.xml")
_CHART_FORMULA = re.compile(r"(<(?:\w+:)?f(?:\s[^>]*)?>)([^<]*)(</(?:\w+:)?f>)")


def rename_formula_sheet_references(text: str, old: str, new: str) -> str:
    """Point the sheet references in formula text at a renamed sheet.

    The formula parser is imported lazily: importing ``xlspy`` must not load the
    object model or the formula package.
    """
    from .formula.sheets import rename_sheet_references

    return rename_sheet_references(text, old, new)


def renamed_chart_parts(package: ZipPackage, old: str, new: str) -> dict[str, bytes]:
    """The chart parts whose formulas (``<c:f>``) name sheet ``old``, pointed at ``new``."""

    def rename(match: re.Match[str]) -> str:
        text = unescape_xml_text(match[2])
        updated = rename_formula_sheet_references(text, old, new)
        if updated == text:
            return match[0]
        return match[1] + saxutils.escape(updated) + match[3]

    parts = {}
    for name in package.names():
        if _CHART_PART.fullmatch(name):
            xml = package.get(name).decode("utf-8")
            updated = _CHART_FORMULA.sub(rename, xml)
            if updated != xml:
                parts[name] = updated.encode("utf-8")
    return parts


_WORKSHEET_PART = re.compile(r"xl/worksheets/sheet[^/]*\.xml")
# Sparkline formulas live in the sheet's extLst under the x14 ``xm`` prefix.
_SPARKLINE_FORMULA = re.compile(r"(<xm:f>)([^<]*)(</xm:f>)")


def renamed_sparkline_parts(package: ZipPackage, old: str, new: str) -> dict[str, bytes]:
    """The worksheet parts whose sparkline formulas (``<xm:f>``) name sheet ``old``."""

    def rename(match: re.Match[str]) -> str:
        text = unescape_xml_text(match[2])
        updated = rename_formula_sheet_references(text, old, new)
        if updated == text:
            return match[0]
        return match[1] + saxutils.escape(updated) + match[3]

    parts = {}
    for name in package.names():
        if _WORKSHEET_PART.fullmatch(name):
            xml = package.get(name).decode("utf-8")
            updated = _SPARKLINE_FORMULA.sub(rename, xml)
            if updated != xml:
                parts[name] = updated.encode("utf-8")
    return parts


def parse_shared_strings_xml(xml_data: bytes | str) -> list[str]:
    """Parse plain and rich-text ``<si>`` entries from an XLSX SST.

    Expat performs entity decoding for us and this parser intentionally joins
    all ``<t>`` runs inside one shared-string item, which is what Excel
    displays for rich text.
    """

    if isinstance(xml_data, bytes):
        source = xml_data
    else:
        source = xml_data.encode("utf-8")

    values: list[str] = []
    in_shared_item = False
    in_text = False
    text_parts: list[str] = []

    parser = harden_expat(expat.ParserCreate())

    def start_element(name: str, attrs: dict[str, str]) -> None:
        nonlocal in_shared_item, in_text, text_parts
        local = local_xml_name(name)
        if local == "si":
            in_shared_item = True
            text_parts = []
        elif local == "t" and in_shared_item:
            in_text = True

    def end_element(name: str) -> None:
        nonlocal in_shared_item, in_text
        local = local_xml_name(name)
        if local == "t":
            in_text = False
        elif local == "si" and in_shared_item:
            values.append("".join(text_parts))
            in_shared_item = False

    def character_data(data: str) -> None:
        if in_shared_item and in_text:
            text_parts.append(data)

    parser.StartElementHandler = start_element
    parser.EndElementHandler = end_element
    parser.CharacterDataHandler = character_data
    parser.Parse(source, True)
    return values


@dataclass(frozen=True)
class TableHeader:
    """The header row of one table on a sheet (zero-based row and column)."""

    display_name: str
    row: int
    first_col: int
    columns: tuple[str, ...]


def table_header(display_name: str, ref: str, has_header_row: bool, column_names: Sequence[str]) -> TableHeader | None:
    """Build a ``TableHeader``, or ``None`` for a table with no usable header."""

    if not has_header_row or not column_names:
        return None
    try:
        first_col, first_row, _, _ = parse_cell_range(ref)
    except ValueError:
        return None
    return TableHeader(
        display_name=display_name,
        row=first_row,
        first_col=first_col,
        columns=tuple(column_names),
    )


def sheet_table_headers(
    package: ZipPackage,
    sheet_path: str,
    read_table: Callable[[bytes], TableHeader | None],
) -> list[TableHeader]:
    """Header rows of the tables that belong to ``sheet_path``.

    ``read_table`` parses one table part for the format (XML for XLSX, BIFF12
    for XLSB).  A table that cannot be read contributes nothing.
    """

    sheet_dir = posixpath.dirname(sheet_path)
    rels_path = posixpath.join(sheet_dir, "_rels", posixpath.basename(sheet_path) + ".rels")
    if not package.has(rels_path):
        return []
    headers: list[TableHeader] = []
    try:
        relationships = list(safe_fromstring(package.get(rels_path)))
    except ET.ParseError:
        return []
    for node in relationships:
        target = node.get("Target")
        if not (node.get("Type") or "").endswith("/table") or not target:
            continue
        table_path = resolve_zip_target(target, sheet_dir or "xl")
        if not package.has(table_path):
            continue
        try:
            header = read_table(package.get(table_path))
        except (ET.ParseError, ValueError, IndexError, TypeError, struct.error):
            continue
        if header is not None:
            headers.append(header)
    return headers


def xml_table_header(table_xml: bytes) -> TableHeader | None:
    """The header row of an XLSX table part (``xl/tables/tableN.xml``)."""

    root = safe_fromstring(table_xml)
    has_header = (root.get("headerRowCount") or "1").strip() != "0"
    columns = [node.get("name") or "" for node in root.iter() if node.tag.rsplit("}", 1)[-1] == "tableColumn"]
    return table_header(
        root.get("displayName") or root.get("name") or "",
        root.get("ref") or "",
        has_header,
        columns,
    )


def header_row_texts(headers: Sequence[Any]) -> list[str]:
    """The text the updaters write for a header row (``None`` becomes empty)."""

    return ["" if unwrap_cell(value) is None else str(unwrap_cell(value)) for value in headers]


class TableHeaderGuard:
    """Refuse a rewrite that would change a table's header cells.

    Excel repairs a table whose header cells do not hold its column names, so a
    replacement must keep those cells equal (ignoring case, as the writers do).
    Call ``check_row`` for each new row, zero-based, before it is written,
    and ``finish`` once all rows are seen: a header row the new data never
    reaches is empty, which is a mismatch too.  Nothing is written before the
    check, so a failure leaves the package unchanged.
    """

    def __init__(self, sheet_name: str, tables: Sequence[TableHeader]):
        self._sheet_name = sheet_name
        self._pending: dict[int, list[TableHeader]] = {}
        for table in tables:
            self._pending.setdefault(table.row, []).append(table)

    def check_row(self, row: int, values: Sequence[Any]) -> None:
        tables = self._pending.pop(row, None)
        if tables is None:
            return
        for table in tables:
            self._check(table, values)

    def finish(self) -> None:
        if self._pending:
            table = min(
                (table for tables in self._pending.values() for table in tables),
                key=lambda table: table.row,
            )
            self._check(table, ())

    def _check(self, table: TableHeader, values: Sequence[Any]) -> None:
        for offset, name in enumerate(table.columns):
            col = table.first_col + offset
            cell = unwrap_cell(values[col]) if col < len(values) else None
            if isinstance(cell, str) and cell.casefold() == name.casefold():
                continue
            ref = f"{column_index_to_letter(col)}{table.row + 1}"
            raise UnsupportedFeatureError(
                f"sheet {self._sheet_name!r}: table {table.display_name!r} has column "
                f"{name!r} but the new data puts {cell!r} in header cell {ref}; Excel "
                "requires each header cell to equal its column name, and "
                "replace_sheet_data cannot rename table columns"
            )


def _sheet_relates_to(package: ZipPackage, sheet_path: str | None, part_type: bytes) -> bool:
    """True when the sheet part has a relationship of ``part_type``."""
    if not sheet_path:
        return False
    rels = rels_path(sheet_path)
    if rels not in package.names():
        return False
    return re.search(rb'Type="[^"]*/' + part_type + rb'"', package.get(rels)) is not None


def sheet_has_slicers(package: ZipPackage, sheet_path: str | None) -> bool:
    """True when the sheet part relates to a table slicer."""
    return _sheet_relates_to(package, sheet_path, b"slicer")


def sheet_has_pivots(package: ZipPackage, sheet_path: str | None) -> bool:
    """True when the sheet part relates to a pivot table."""
    return _sheet_relates_to(package, sheet_path, b"pivotTable")
