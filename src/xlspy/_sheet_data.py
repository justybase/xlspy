"""Finding and replacing the ``<sheetData>`` element of a worksheet part.

Three places used to locate it with their own regular expressions: the
updater's in-memory patch (``XlsxUpdater._patch_sheet_xml``), the stage-2 cell
edit (``_xlsx_edit``) and the object-model loader (``workbook/_loaded``). They
all come through here now, so the element is recognised the same way every
time: with or without a namespace prefix, empty or not.

The row-level code that builds or edits the rows inside the element stays
where it is; it differs per caller.
"""

from __future__ import annotations

import re

from ._metadata import element_pattern


def tag_prefix(xml: str, tag: str) -> str:
    """The namespace prefix (with its colon) used for ``tag``, or ``""``."""
    match = re.search(rf"<(?P<prefix>[A-Za-z_][\w.-]*:)?{re.escape(tag)}\b", xml)
    return match.group("prefix") or "" if match else ""


def sheet_data_match(xml: str) -> re.Match[str] | None:
    """The whole ``<sheetData>`` element, empty or not, or None.

    The match has a ``prefix`` group (with its colon, or None).
    """
    return element_pattern("sheetData").search(xml)


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
