"""Finding and replacing the ``<sheetData>`` element of a worksheet part.

Three places used to locate it with their own regular expressions: the
updater's in-memory patch (``XlsxUpdater._patch_sheet_xml``), the stage-2 cell
edit (``_xlsx_edit``) and the object-model loader (``workbook/_loaded``). They
all come through here now, so the element is recognised the same way every
time: with or without a namespace prefix, empty or not.

``patch_rows`` and ``insert_rows`` are the row walk shared by the edits of an
existing sheet (``update_cells`` and the object model's row dimensions): what a
row becomes stays with the caller.
"""

from __future__ import annotations

import re
from collections.abc import Callable, Collection, Iterator, Mapping
from typing import Literal, overload

from ._metadata import append_child, element_pattern, tag_attributes
from .refs import column_label, parse_cell_range
from .updater_utils import escape_xml_text, replace_or_add_xml_attribute

# Opening tags, to walk a worksheet a tag at a time: ``.*?`` through 25 MB of cells
# is what these spare. ``number`` is the usual form of a row's ``r`` attribute.
_ROW_OPENING = re.compile(r'<(?P<prefix>[A-Za-z_][\w.-]*:)?row\b(?:\s+r="(?P<number>[0-9]+)")?[^>]*?(?:(?P<empty>/>)|>)')
_SHEET_DATA_OPENING = re.compile(r"<(?P<prefix>[A-Za-z_][\w.-]*:)?sheetData\b[^>]*?(?:(?P<empty>/>)|>)")
_DIMENSION = re.compile(r"<(?:[A-Za-z_][\w.-]*:)?dimension\b[^>]*>")
_WORKSHEET = re.compile(r"<(?:[A-Za-z_][\w.-]*:)?worksheet\b[^>]*>")
# CT_Worksheet: sheetPr, dimension, then these.
_AFTER_DIMENSION = re.compile(r"<(?:[A-Za-z_][\w.-]*:)?(?:sheetViews|sheetFormatPr|cols|sheetData)\b")
_CELL_ELEMENT = element_pattern("c")


def tag_prefix(xml: str, tag: str) -> str:
    """The namespace prefix (with its colon) used for ``tag``, or ``""``."""
    match = re.search(rf"<(?P<prefix>[A-Za-z_][\w.-]*:)?{re.escape(tag)}\b", xml)
    return match.group("prefix") or "" if match else ""


class SheetDataElement:
    """Where ``sheet_data_match`` found the element: ``start()``, ``end()``, ``[0]``
    (its XML) and ``["prefix"]`` (with its colon, or None), as on a ``re.Match``.
    """

    __slots__ = ("_xml", "_start", "_end", "_prefix")

    def __init__(self, xml: str, start: int, end: int, prefix: str | None) -> None:
        self._xml = xml
        self._start = start
        self._end = end
        self._prefix = prefix

    def start(self) -> int:
        return self._start

    def end(self) -> int:
        return self._end

    @overload
    def __getitem__(self, key: Literal[0]) -> str: ...
    @overload
    def __getitem__(self, key: Literal["prefix"]) -> str | None: ...
    def __getitem__(self, key: int | str) -> str | None:
        return self._xml[self._start:self._end] if key == 0 else self._prefix


def _element_end(xml: str, tag: str, opening: re.Match[str]) -> int:
    """Where the element opened by ``opening`` ends, or -1 without a closing tag. The
    same rule as ``element_pattern``: its own prefix first, then none.
    """
    if opening["empty"]:
        return opening.end()
    prefix = opening["prefix"]
    for closing in ((f"</{prefix}{tag}", f"</{tag}") if prefix else (f"</{tag}",)):
        position = xml.find(closing, opening.end())
        while position >= 0:
            end = position + len(closing)
            while xml[end:end + 1].isspace():
                end += 1
            if xml[end:end + 1] == ">":
                return end + 1
            position = xml.find(closing, end)
    return -1


def sheet_data_match(xml: str) -> SheetDataElement | None:
    """The whole ``<sheetData>`` element, empty or not, or None."""
    position = 0
    while (opening := _SHEET_DATA_OPENING.search(xml, position)) is not None:
        end = _element_end(xml, "sheetData", opening)
        if end >= 0:
            return SheetDataElement(xml, opening.start(), end, opening["prefix"])
        position = opening.start() + 1
    return None


def replace_sheet_data_body(sheet_xml: str, rows_xml: str) -> str:
    """Put ``rows_xml`` inside ``<sheetData>``, keeping its attributes and prefix.

    An empty ``<sheetData/>`` is expanded so the rows have an element to go into.
    Everything outside the element is left as it was.
    """
    match = sheet_data_match(sheet_xml)
    if match is None:
        raise ValueError("Worksheet sheetData element was not found")
    element = match[0]
    if element.endswith("/>"):
        opening = element[:-2] + ">"
        closing_tag = f"</{match['prefix'] or ''}sheetData>"
        replacement = opening + rows_xml + closing_tag
    else:
        body_start = element.index(">") + 1
        body_end = element.rindex("</")
        replacement = element[:body_start] + rows_xml + element[body_end:]
    return sheet_xml[:match.start()] + replacement + sheet_xml[match.end():]


def _iter_rows(element: str) -> Iterator[tuple[int, int, int]]:
    """``(start, end, number)`` of each ``<row>`` of ``element``. A sheet has a row
    per few cells, so this runs for every row of every edit: a tag at a time here,
    not ``.*?`` through each row's cells.

    A row's number is its ``r`` attribute, or the previous row's number plus one.
    """
    next_row = 1
    position = 0
    search = _ROW_OPENING.search
    while True:
        opening = search(element, position)
        if opening is None:
            return
        end = _element_end(element, "row", opening)
        if end < 0:
            position = opening.start() + 1
            continue
        if opening["number"]:
            number = int(opening["number"])
        elif opening.end() - opening.start() == len("<row>"):
            number = next_row
        else:
            number = int(tag_attributes(opening[0]).get("r", next_row))
        next_row = number + 1
        yield opening.start(), end, number
        position = end


def patch_rows(element: str, patch: Callable[[int, str], str], rows: Collection[int] | None = None) -> str:
    """``element`` (a ``<sheetData>`` element) with each ``<row>`` replaced by
    ``patch(row_number, row_xml)``; given ``rows``, only the rows with those numbers.
    """
    pieces = []
    done = 0
    for start, end, number in _iter_rows(element):
        if rows is None or number in rows:
            pieces.append(element[done:start])
            pieces.append(patch(number, element[start:end]))
            done = end
    pieces.append(element[done:])
    return "".join(pieces)


def insert_rows(element: str, new_rows: Mapping[int, str]) -> str:
    """``element`` (a ``<sheetData>`` element) with ``new_rows`` (row number -> row
    XML, for rows it does not have) added in row order.

    Each new row goes before the first existing row with a higher number; the rest
    go at the end, expanding an empty ``<sheetData/>``.
    """
    pending = sorted(new_rows.items())
    if not pending:
        return element
    pieces = []
    done = 0
    next_new = 0
    for start, _, number in _iter_rows(element):
        if next_new < len(pending) and pending[next_new][0] < number:
            pieces.append(element[done:start])
            done = start
            while next_new < len(pending) and pending[next_new][0] < number:
                pieces.append(pending[next_new][1])
                next_new += 1
    pieces.append(element[done:])
    element = "".join(pieces)
    if next_new < len(pending):
        element = append_child(element, "sheetData", "".join(xml for _, xml in pending[next_new:]))
    return element


def patch_cells(xml: str, patch: Callable[[str, str], str], rows: Collection[int]) -> str:
    """``xml`` (a worksheet part) with each ``<c>`` of the 1-based ``rows`` of its
    ``<sheetData>`` replaced by ``patch(ref, cell_xml)``.

    A cell without an ``r`` attribute gets the reference of its position, as
    Excel reads it: the column after the previous cell in the row.
    """
    match = sheet_data_match(xml)
    if match is None:
        return xml

    def patch_row(row_number: int, row_xml: str) -> str:
        opening = row_xml[: row_xml.index(">") + 1]
        next_column = 0

        def patch_cell(cell_match: re.Match[str]) -> str:
            nonlocal next_column
            cell_xml = cell_match[0]
            ref = tag_attributes(cell_xml).get("r") or f"{column_label(next_column)}{row_number}"
            next_column = parse_cell_range(ref)[0] + 1
            return patch(ref, cell_xml)

        return opening + _CELL_ELEMENT.sub(patch_cell, row_xml[len(opening):])

    body = patch_rows(match[0], patch_row, rows)
    return xml[:match.start()] + body + xml[match.end():]


def set_dimension_ref(sheet_xml: str, ref: str) -> str:
    """``sheet_xml`` with its ``<dimension>`` set to ``ref``.

    An existing element keeps its prefix and other attributes; a missing one is
    added where CT_Worksheet puts it, with the worksheet's prefix.
    """
    match = _DIMENSION.search(sheet_xml)
    if match:
        opening = replace_or_add_xml_attribute(match[0], "ref", ref)
        return sheet_xml[:match.start()] + opening + sheet_xml[match.end():]
    worksheet = _WORKSHEET.search(sheet_xml)
    if worksheet is None:
        raise ValueError("Worksheet root element was not found")
    element = f'<{tag_prefix(sheet_xml, "worksheet")}dimension ref="{escape_xml_text(ref)}"/>'
    following = _AFTER_DIMENSION.search(sheet_xml)
    position = following.start() if following else worksheet.end()
    return sheet_xml[:position] + element + sheet_xml[position:]
