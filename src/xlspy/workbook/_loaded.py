"""Save edits back into an existing XLSX/XLSM package.

Only changed worksheet parts and the workbook/style metadata are rewritten.
Untouched parts (drawings, charts, pivots, macros, embedded objects) are copied
byte-for-byte by ``ZipPackage``.
"""

from __future__ import annotations

import datetime as _datetime
import math
import re
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape, quoteattr

from .._metadata import MAIN, REL, PKG_REL, CT, element_pattern, replace_element, serialize_like
from .._xlsx_edit import _attributes, _bytes, _cell_value
from ..formulas import Formula
from ..refs import column_label, parse_cell_range
from ..styles import _coerce_color, _color_key
from ..updater_utils import replace_or_add_xml_attribute, resolve_zip_target
from ._model import Workbook, Worksheet, _coordinate, _normalize_range
from ._styles import XlsxStyleTable

_CONTENT_WORKSHEET = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
_CONTENT_COMMENTS = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"
_CONTENT_VML = "application/vnd.openxmlformats-officedocument.vmlDrawing"


def _q(tag: str) -> str:
    return f"{{{MAIN}}}{tag}"


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _remove_attr(opening: str, name: str) -> str:
    return re.sub(
        rf'\s+{re.escape(name)}\s*=\s*(?:"[^"]*"|\'[^\']*\')',
        "",
        opening,
        count=1,
    )


def _set_attr(opening: str, name: str, value) -> str:
    return replace_or_add_xml_attribute(opening, name, value)


def _row_dimension_attrs(dimension) -> dict[str, str]:
    attrs = {}
    if dimension.height is not None:
        attrs["ht"] = f"{dimension.height:g}"
        attrs["customHeight"] = "1"
    if dimension.hidden:
        attrs["hidden"] = "1"
    if dimension.outlineLevel:
        attrs["outlineLevel"] = str(dimension.outlineLevel)
    if dimension.collapsed:
        attrs["collapsed"] = "1"
    return attrs


def _meta_signature(sheet: Worksheet):
    """A hashable snapshot of the sheet metadata edited through this model."""
    cols = tuple(
        (letter, dim.width, dim.hidden, dim.outlineLevel, dim.collapsed)
        for letter, dim in sorted(
            sheet.column_dimensions.items(),
            key=lambda item: parse_cell_range(f"{item[0]}1")[0],
        )
    )
    rows = tuple(
        (row, dim.height, dim.hidden, dim.outlineLevel, dim.collapsed)
        for row, dim in sorted(sheet.row_dimensions.items())
    )
    view = sheet.sheet_view
    return (
        cols, rows, sheet.freeze_panes, sheet.auto_filter.ref,
        tuple(sorted(str(value) for value in sheet.merged_cells.ranges)),
        repr(sheet.sheet_properties.tabColor),
        (
            view.showGridLines, view.showRowColHeaders, view.zoomScale,
            view.zoomScaleNormal, view.tabSelected, view.rightToLeft,
        ),
    )


def _dimension_ref(sheet: Worksheet) -> str:
    max_row = sheet._max_row
    max_column = sheet._max_column
    for merged in sheet._merged_cells.ranges:
        _, _, c2, r2 = parse_cell_range(merged)
        max_row = max(max_row, r2 + 1)
        max_column = max(max_column, c2 + 1)
    if max_row <= 0 or max_column <= 0:
        return "A1"
    return f"A1:{_coordinate(max_row, max_column)}"


def _sheet_views_xml(sheet: Worksheet, prefix: str) -> str:
    attrs = []
    view = sheet.sheet_view
    if view.showGridLines is False:
        attrs.append('showGridLines="0"')
    if view.showRowColHeaders is False:
        attrs.append('showRowColHeaders="0"')
    if view.rightToLeft:
        attrs.append('rightToLeft="1"')
    if view.zoomScale is not None and view.zoomScale != 100:
        attrs.append(f'zoomScale="{view.zoomScale}"')
    if view.zoomScaleNormal is not None and view.zoomScaleNormal != 100:
        attrs.append(f'zoomScaleNormal="{view.zoomScaleNormal}"')
    if view.tabSelected:
        attrs.append('tabSelected="1"')
    pane = ""
    selection = ""
    if sheet.freeze_panes:
        col, row, _, _ = parse_cell_range(sheet.freeze_panes)
        split = []
        if col:
            split.append(f'xSplit="{col}"')
        if row:
            split.append(f'ySplit="{row}"')
        split.append(f'topLeftCell="{sheet.freeze_panes}"')
        if col and row:
            active = "bottomRight"
        elif col:
            active = "topRight"
        else:
            active = "bottomLeft"
        split.append(f'activePane="{active}"')
        split.append('state="frozen"')
        pane = f'<{prefix}pane {" ".join(split)}/>'
        selection = (
            f'<{prefix}selection pane="{active}" '
            f'activeCell="{sheet.freeze_panes}" sqref="{sheet.freeze_panes}"/>'
        )
    view_attrs = (" " + " ".join(attrs)) if attrs else ""
    return (
        f'<{prefix}sheetViews><{prefix}sheetView{view_attrs} workbookViewId="0">'
        f"{pane}{selection}</{prefix}sheetView></{prefix}sheetViews>"
    )


def _cols_xml(sheet: Worksheet, prefix: str) -> str:
    parts = []
    for letter in sorted(sheet.column_dimensions, key=lambda value: parse_cell_range(f"{value}1")[0]):
        dimension = sheet.column_dimensions[letter]
        if (
            dimension.width is None
            and not dimension.hidden
            and not dimension.outlineLevel
            and not dimension.collapsed
        ):
            continue
        index = parse_cell_range(f"{letter}1")[0] + 1
        attrs = [f'min="{index}"', f'max="{index}"']
        if dimension.width is not None:
            attrs.append(f'width="{dimension.width:g}"')
            attrs.append('customWidth="1"')
        else:
            attrs.append('width="8.43"')
            attrs.append('customWidth="0"')
        if dimension.hidden:
            attrs.append('hidden="1"')
        if dimension.outlineLevel:
            attrs.append(f'outlineLevel="{dimension.outlineLevel}"')
        if dimension.collapsed:
            attrs.append('collapsed="1"')
        parts.append(f'<{prefix}col {" ".join(attrs)}/>')
    if not parts:
        return ""
    return f'<{prefix}cols>{"".join(parts)}</{prefix}cols>'


def _patch_rows(xml: str, sheet: Worksheet, prefix: str) -> str:
    dimensions = {
        row: dimension
        for row, dimension in sheet.row_dimensions.items()
        if _row_dimension_attrs(dimension)
    }
    if not dimensions:
        return xml
    match = element_pattern("sheetData").search(xml)
    if match is None:
        return xml
    body = match[0]
    seen = set()
    next_row = 1

    def patch_row(row_match):
        nonlocal next_row
        row_xml = row_match[0]
        opening_match = re.match(r"<[^>]+>", row_xml)
        if opening_match is None:
            return row_xml
        opening = opening_match[0]
        attrs = _attributes(opening)
        row_number = int(attrs.get("r", next_row))
        next_row = row_number + 1
        seen.add(row_number)
        dimension = dimensions.get(row_number)
        if dimension is None:
            return row_xml
        new_opening = opening
        for name in ("ht", "customHeight", "hidden", "outlineLevel", "collapsed"):
            new_opening = _remove_attr(new_opening, name)
        wanted = _row_dimension_attrs(dimension)
        for name, value in wanted.items():
            new_opening = _set_attr(new_opening, name, value)
        return new_opening + row_xml[len(opening):]

    body = element_pattern("row").sub(patch_row, body)
    missing = sorted(row for row in dimensions if row not in seen)
    for row in reversed(missing):
        dimension = dimensions[row]
        attrs = {"r": str(row), **_row_dimension_attrs(dimension)}
        fragment = f'<{prefix}row ' + " ".join(
            f"{key}={quoteattr(value)}" for key, value in attrs.items()
        ) + "/>"
        rows = list(element_pattern("row").finditer(body))
        position = None
        for candidate in rows:
            opening = re.match(r"<[^>]+>", candidate[0])[0]
            number = int(_attributes(opening).get("r", "0") or 0)
            if number > row:
                position = candidate.start()
                break
        if position is None:
            closing = body.rfind("</")
            position = closing if closing >= 0 else len(body)
        body = body[:position] + fragment + body[position:]
    return xml[:match.start()] + body + xml[match.end():]


def _tab_color_xml(sheet: Worksheet, prefix: str) -> str:
    color = sheet.sheet_properties.tabColor
    if color is None:
        return ""
    try:
        key = _color_key(_coerce_color(color, "tabColor"))
    except (TypeError, ValueError):
        return ""
    if key is None:
        return ""
    return f'<{prefix}tabColor {XlsxStyleTable._color_attrs(key)}/>'


def _patch_tab_color(xml: str, sheet: Worksheet, prefix: str) -> str:
    tab = _tab_color_xml(sheet, prefix)
    match = element_pattern("sheetPr").search(xml)
    if match is None:
        if not tab:
            return xml
        fragment = f'<{prefix}sheetPr>{tab}</{prefix}sheetPr>'
        return replace_element(
            xml, "sheetPr", fragment,
            before=("dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData"),
        )
    body = match[0]
    body = element_pattern("tabColor").sub("", body)
    if tab:
        if body.endswith("/>"):
            body = body[:-2] + ">" + tab + f"</{prefix}sheetPr>"
        else:
            closing = body.rfind("</")
            body = body[:closing] + tab + body[closing:]
    return xml[:match.start()] + body + xml[match.end():]


def _merge_xml(sheet: Worksheet, prefix: str) -> str:
    ranges = [str(value) for value in sheet.merged_cells.ranges]
    if not ranges:
        return ""
    body = "".join(f'<{prefix}mergeCell ref="{ref}"/>' for ref in ranges)
    return f'<{prefix}mergeCells count="{len(ranges)}">{body}</{prefix}mergeCells>'


def _auto_filter_xml(sheet: Worksheet, prefix: str) -> str:
    ref = sheet.auto_filter.ref
    if not ref:
        return ""
    return f'<{prefix}autoFilter ref="{_normalize_range(ref)}"/>'


def patch_sheet_metadata(updater, sheet: Worksheet, name: str) -> None:
    path, xml = updater._sheet_xml(name)
    prefix = updater._tag_prefix(xml, "sheetData")
    xml = _patch_tab_color(xml, sheet, prefix)
    xml = replace_element(
        xml, "sheetViews", _sheet_views_xml(sheet, prefix),
        before=("sheetFormatPr", "cols", "sheetData"),
    )
    cols = _cols_xml(sheet, prefix)
    xml = replace_element(xml, "cols", cols, before=("sheetData",))
    xml = _patch_rows(xml, sheet, prefix)
    xml = replace_element(
        xml, "dimension",
        f'<{prefix}dimension ref="{_dimension_ref(sheet)}"/>',
        before=("sheetViews", "sheetFormatPr", "cols", "sheetData"),
    )
    xml = replace_element(
        xml, "autoFilter", _auto_filter_xml(sheet, prefix),
        before=("mergeCells", "conditionalFormatting", "dataValidations", "hyperlinks",
                "printOptions", "pageMargins", "pageSetup", "headerFooter",
                "drawing", "legacyDrawing", "tableParts", "extLst"),
    )
    xml = replace_element(
        xml, "mergeCells", _merge_xml(sheet, prefix),
        before=("conditionalFormatting", "dataValidations", "hyperlinks",
                "printOptions", "pageMargins", "pageSetup", "headerFooter",
                "drawing", "legacyDrawing", "tableParts", "extLst"),
    )
    updater._apply_parts({path: xml})


def _patch_cells(xml: str, patch) -> str:
    """Call ``patch(ref, cell_xml)`` for each cell, resolving omitted ``r``
    attributes from the cell's position as Excel does."""
    match = element_pattern("sheetData").search(xml)
    if match is None:
        return xml
    next_row = 1

    def patch_row(row_match):
        nonlocal next_row
        row_xml = row_match[0]
        opening = re.match(r"<[^>]+>", row_xml)[0]
        row_number = int(_attributes(opening).get("r", next_row))
        next_row = row_number + 1
        next_column = 0

        def patch_cell(cell_match):
            nonlocal next_column
            cell_xml = cell_match[0]
            ref = _attributes(cell_xml).get("r") or f"{column_label(next_column)}{row_number}"
            next_column = parse_cell_range(ref)[0] + 1
            return patch(ref, cell_xml)

        return opening + element_pattern("c").sub(patch_cell, row_xml[len(opening):])

    body = element_pattern("row").sub(patch_row, match[0])
    return xml[:match.start()] + body + xml[match.end():]


def patch_cell_styles(updater, name: str, styles: dict[str, int]) -> None:
    if not styles:
        return
    path, xml = updater._sheet_xml(name)
    existing = set()
    _patch_cells(xml, lambda ref, cell_xml: existing.add(ref) or cell_xml)
    missing = [ref for ref in styles if ref not in existing]
    if missing:
        updater.update_cells(name, {ref: None for ref in missing})
        path, xml = updater._sheet_xml(name)

    def patch_cell(ref, cell_xml):
        if ref not in styles:
            return cell_xml
        opening = re.match(r"<[^>]+>", cell_xml)[0]
        return _set_attr(opening, "s", styles[ref]) + cell_xml[len(opening):]

    xml = _patch_cells(xml, patch_cell)
    updater._apply_parts({path: xml})


def _time_serial(value: _datetime.time) -> float:
    return (
        value.hour * 3600 + value.minute * 60 + value.second
        + value.microsecond / 1_000_000
    ) / 86400.0


def _value_for_save(cell, date_1904: bool):
    if isinstance(cell._formula, Formula):
        return cell._formula
    if isinstance(cell._formula, str):
        try:
            return Formula(cell._formula)
        except (TypeError, ValueError):
            raise ValueError(
                "this formula cannot be rewritten by the compatibility editor; "
                "use XlsxUpdater.update_cells with a Formula object instead"
            ) from None
    value = cell._value
    if isinstance(value, _datetime.time):
        if cell.number_format == "General":
            cell.number_format = "h:mm:ss"
        return _time_serial(value)
    if isinstance(value, _datetime.timedelta):
        if cell.number_format == "General":
            cell.number_format = "[h]:mm:ss"
        return value.total_seconds() / 86400.0
    return value


def _sheet_xml(sheet: Worksheet, date_1904: bool, *, prefix: str = "") -> str:
    rows = []
    max_row = max(sheet._max_row, 1)
    max_column = max(sheet._max_column, 1)
    for row_number in range(1, max_row + 1):
        cells = []
        for column in range(1, max_column + 1):
            cell = sheet._existing_cell(row_number, column)
            if cell is None:
                continue
            value = _value_for_save(cell, date_1904)
            kind, content = _cell_value(value, date_1904, prefix)
            attrs = [f'r="{_coordinate(row_number, column)}"']
            if cell._style_id:
                attrs.append(f's="{cell._style_id}"')
            if kind:
                attrs.append(f't="{kind}"')
            if content:
                cells.append(f'<{prefix}c {" ".join(attrs)}>{content}</{prefix}c>')
            elif cell._style_id:
                cells.append(f'<{prefix}c {" ".join(attrs)}/>')
        row_attrs = [f'r="{row_number}"']
        dimension = sheet.row_dimensions.get(row_number)
        if dimension is not None:
            for key, value in _row_dimension_attrs(dimension).items():
                row_attrs.append(f"{key}={quoteattr(value)}")
        if cells:
            rows.append(
                f'<{prefix}row {" ".join(row_attrs)}>{"".join(cells)}</{prefix}row>'
            )
        elif dimension is not None and _row_dimension_attrs(dimension):
            rows.append(f'<{prefix}row {" ".join(row_attrs)}/>')
    sheet_data = (
        f'<{prefix}sheetData>{"".join(rows)}</{prefix}sheetData>'
        if rows else f'<{prefix}sheetData/>'
    )
    tab = _tab_color_xml(sheet, prefix)
    sheet_pr = f'<{prefix}sheetPr>{tab}</{prefix}sheetPr>' if tab else ""
    dimension = f'<{prefix}dimension ref="{_dimension_ref(sheet)}"/>'
    views = _sheet_views_xml(sheet, prefix)
    cols = _cols_xml(sheet, prefix)
    merges = _merge_xml(sheet, prefix)
    auto = _auto_filter_xml(sheet, prefix)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<{prefix}worksheet xmlns="{MAIN}" xmlns:r="{REL}">'
        f"{sheet_pr}{dimension}{views}<{prefix}sheetFormatPr defaultRowHeight=\"15\"/>"
        f"{cols}{sheet_data}{auto}{merges}<{prefix}pageMargins left=\"0.7\" right=\"0.7\" "
        f'top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
        f"</{prefix}worksheet>"
    )


def _next_sheet_path(package) -> str:
    highest = 0
    for name in package.names():
        match = re.fullmatch(r"xl/worksheets/sheet(\d+)\.xml", name)
        if match:
            highest = max(highest, int(match.group(1)))
    return f"xl/worksheets/sheet{highest + 1}.xml"


def _next_rel_id(rels_root) -> str:
    highest = 0
    for node in rels_root:
        match = re.fullmatch(r"rId(\d+)", node.get("Id", ""))
        if match:
            highest = max(highest, int(match.group(1)))
    return f"rId{highest + 1}"


def _insert_sheet_fragment(workbook_xml: str, fragment: str, index: int) -> str:
    match = element_pattern("sheets").search(workbook_xml)
    if match is None:
        raise ValueError("workbook.xml sheets element was not found")
    body = match[0]
    rows = list(element_pattern("sheet").finditer(body))
    if index < len(rows):
        position = rows[index].start()
    else:
        position = body.rfind("</")
        if position < 0:
            position = len(body)
    body = body[:position] + fragment + body[position:]
    return workbook_xml[:match.start()] + body + workbook_xml[match.end():]


def add_sheet(updater, workbook: Workbook, sheet: Worksheet, index: int) -> None:
    package = updater._package
    path = _next_sheet_path(package)
    xml = _sheet_xml(sheet, updater._date_1904)
    package.set(path, xml.encode("utf-8"))

    content_types = ET.fromstring(package.get("[Content_Types].xml"))
    ET.SubElement(
        content_types, f"{{{CT}}}Override",
        {"PartName": "/" + path, "ContentType": _CONTENT_WORKSHEET},
    )
    package.set("[Content_Types].xml", _bytes(content_types))

    rels_path = "xl/_rels/workbook.xml.rels"
    rels = ET.fromstring(package.get(rels_path))
    rid = _next_rel_id(rels)
    ET.SubElement(
        rels, f"{{{PKG_REL}}}Relationship",
        {
            "Id": rid,
            "Type": REL + "/worksheet",
            "Target": path.removeprefix("xl/"),
        },
    )
    package.set(rels_path, _bytes(rels))

    workbook_path = "xl/workbook.xml"
    workbook_xml = package.get(workbook_path).decode("utf-8")
    root = ET.fromstring(workbook_xml)
    existing_ids = [
        int(node.get("sheetId", "0") or 0)
        for node in root.findall(f"{{{MAIN}}}sheets/{{{MAIN}}}sheet")
    ]
    sheet_id = max(existing_ids, default=0) + 1
    fragment = (
        f'<sheet name={quoteattr(sheet.title)} sheetId="{sheet_id}" '
        f'r:id="{rid}"/>'
    )
    workbook_xml = _insert_sheet_fragment(workbook_xml, fragment, index)
    package.set(workbook_path, workbook_xml.encode("utf-8"))
    updater._sheets[sheet.title] = path
    sheet._path = path
    sheet._rel_id = rid
    sheet._sheet_id = sheet_id


def _remove_sheet_node(workbook_xml: str, name: str) -> tuple[str, str | None]:
    for match in list(element_pattern("sheet").finditer(workbook_xml)):
        attrs = _attributes(match.group(0))
        if attrs.get("name") != name:
            continue
        # Attribute keys keep their source prefix (usually ``r:id``).
        rid = next((value for key, value in attrs.items() if key.rsplit(":", 1)[-1] == "id"), None)
        return workbook_xml[:match.start()] + workbook_xml[match.end():], rid
    return workbook_xml, None


def remove_sheet(updater, workbook: Workbook, name: str) -> None:
    package = updater._package
    path = updater._sheets.get(name)
    workbook_path = "xl/workbook.xml"
    workbook_xml = package.get(workbook_path).decode("utf-8")
    workbook_xml, rid = _remove_sheet_node(workbook_xml, name)
    package.set(workbook_path, workbook_xml.encode("utf-8"))

    rels_path = "xl/_rels/workbook.xml.rels"
    rels = ET.fromstring(package.get(rels_path))
    for node in list(rels):
        if rid and node.get("Id") == rid:
            rels.remove(node)
    package.set(rels_path, _bytes(rels))

    types = ET.fromstring(package.get("[Content_Types].xml"))
    for node in list(types):
        if node.get("PartName") == "/" + path:
            types.remove(node)
    package.set("[Content_Types].xml", _bytes(types))

    if path:
        package.remove(path)
        package.remove(rels_path_for_part(path))
    updater._sheets.pop(name, None)


def rels_path_for_part(part: str) -> str:
    directory, filename = part.rsplit("/", 1)
    return f"{directory}/_rels/{filename}.rels"


def rename_sheet(updater, old: str, new: str) -> None:
    package = updater._package
    workbook_path = "xl/workbook.xml"
    workbook_xml = package.get(workbook_path).decode("utf-8")
    for match in list(element_pattern("sheet").finditer(workbook_xml)):
        attrs = _attributes(match.group(0))
        if attrs.get("name") != old:
            continue
        fragment = match.group(0)
        fragment = re.sub(
            r'\bname\s*=\s*("[^"]*"|\'[^\']*\')',
            f"name={quoteattr(new)}",
            fragment,
            count=1,
        )
        workbook_xml = workbook_xml[:match.start()] + fragment + workbook_xml[match.end():]
        break
    package.set(workbook_path, workbook_xml.encode("utf-8"))
    path = updater._sheets.pop(old, None)
    if path is not None:
        updater._sheets[new] = path


def _reorder_and_state(workbook: Workbook, updater) -> None:
    path = "xl/workbook.xml"
    original = updater._package.get(path)
    root = ET.fromstring(original)
    sheets_parent = root.find(f"{{{MAIN}}}sheets")
    if sheets_parent is None:
        return
    original_nodes = list(sheets_parent)
    existing = {}
    for node in original_nodes:
        name = node.get("name")
        if name:
            existing[name] = node
        sheets_parent.remove(node)
    next_id = max(
        (int(node.get("sheetId", "0") or 0) for node in existing.values()),
        default=0,
    ) + 1
    worksheet_nodes = []
    for sheet in workbook._sheets:
        node = existing.get(sheet.title)
        if node is None:
            node = ET.Element(_q("sheet"))
            node.set("name", sheet.title)
            node.set("sheetId", str(getattr(sheet, "_sheet_id", next_id)))
            rid = getattr(sheet, "_rel_id", None)
            if rid:
                node.set(f"{{{REL}}}id", rid)
            next_id += 1
        node.set("name", sheet.title)
        node.set("state", sheet.sheet_state)
        worksheet_nodes.append(node)
    # Model worksheets fill the original worksheet positions in model order;
    # chartsheets and other sheet entries keep their own positions.
    titles = {sheet.title for sheet in workbook._sheets}
    pending = iter(worksheet_nodes)
    for node in original_nodes:
        sheets_parent.append(next(pending) if node.get("name") in titles else node)
    for node in pending:
        sheets_parent.append(node)
    view = root.find(f"{{{MAIN}}}bookViews/{{{MAIN}}}workbookView")
    if view is not None and worksheet_nodes:
        # activeTab counts every sheet entry, including chartsheets.
        active = worksheet_nodes[min(workbook._active_index, len(worksheet_nodes) - 1)]
        view.set("activeTab", str(list(sheets_parent).index(active)))
    updater._package.set(path, serialize_like(root, original))


def _sheet_order(updater) -> list[str]:
    root = ET.fromstring(updater._package.get("xl/workbook.xml"))
    return [node.get("name", "") for node in root.findall(f"{{{MAIN}}}sheets/{{{MAIN}}}sheet")]


def _remap_local_names(updater, before: list[str], renamed: dict[str, str]) -> None:
    """Keep localSheetId pointing at the same sheet after removals and reordering.

    Names local to a removed sheet are dropped, as Excel does.
    """
    after = _sheet_order(updater)
    path = "xl/workbook.xml"
    xml = updater._package.get(path).decode("utf-8")

    def remap(match):
        fragment = match[0]
        opening = re.match(r"<[^>]+>", fragment)[0]
        scope = _attributes(opening).get("localSheetId")
        if scope is None:
            return fragment
        index = int(scope)
        title = before[index] if 0 <= index < len(before) else None
        title = renamed.get(title, title)
        if title not in after:
            return ""
        return _set_attr(opening, "localSheetId", after.index(title)) + fragment[len(opening):]

    updated = element_pattern("definedName").sub(remap, xml)
    if not element_pattern("definedName").search(updated):
        updated = replace_element(updated, "definedNames", "")
    if updated != xml:
        updater._package.set(path, updated.encode("utf-8"))


def save_loaded(workbook: Workbook, filename) -> None:
    updater = workbook._updater
    if updater is None:
        raise ValueError("this workbook was not loaded from an XLSX package")
    if workbook._styles_dirty:
        updater._package.set("xl/styles.xml", workbook._styles.to_bytes())
        workbook._styles_dirty = False

    sheet_order = _sheet_order(updater)
    renamed = {sheet._original_title: sheet.title for sheet in workbook._sheets}
    current_titles = {sheet.title for sheet in workbook._sheets}
    keep = current_titles | {sheet._original_title for sheet in workbook._sheets}
    removed_any = False
    for name in list(workbook._loaded_worksheet_names):
        if name not in keep:
            remove_sheet(updater, workbook, name)
            removed_any = True
    if removed_any:
        # calcChain entries of a removed sheet make Excel repair the file.
        updates, removed_parts = {}, []
        updater._recalculate(updates, removed_parts)
        updater._apply_parts(updates, removed_parts)
    workbook._loaded_worksheet_names -= {
        name for name in workbook._loaded_worksheet_names if name not in keep
    }

    added = set()
    for index, sheet in enumerate(workbook._sheets):
        if sheet._original_title not in updater._sheets and sheet.title not in updater._sheets:
            add_sheet(updater, workbook, sheet, index)
            added.add(sheet.title)
        elif (
            sheet.title != sheet._original_title
            and sheet._original_title in updater._sheets
        ):
            rename_sheet(updater, sheet._original_title, sheet.title)

    for sheet in workbook._sheets:
        name = sheet.title
        if name not in added:
            updates = {}
            for row, column in list(sheet._dirty_cells):
                cell = sheet._existing_cell(row, column)
                if cell is None:
                    continue
                updates[cell.coordinate] = _value_for_save(cell, updater._date_1904)
            if updates:
                updater.update_cells(name, updates)
            style_updates = {}
            for ref in sheet._dirty_styles:
                col, row, _, _ = parse_cell_range(ref)
                cell = sheet._existing_cell(row + 1, col + 1)
                if cell is not None:
                    style_updates[ref] = cell._style_id
            if style_updates:
                if workbook._styles_dirty:
                    updater._package.set("xl/styles.xml", workbook._styles.to_bytes())
                    workbook._styles_dirty = False
                patch_cell_styles(updater, name, style_updates)
            if sheet._meta_signature is None or _meta_signature(sheet) != sheet._meta_signature:
                patch_sheet_metadata(updater, sheet, name)
                sheet._meta_signature = _meta_signature(sheet)
            if sheet._dirty_comments:
                comments = {}
                for ref in sheet._dirty_comments:
                    col, row, _, _ = parse_cell_range(ref)
                    cell = sheet._existing_cell(row + 1, col + 1)
                    comments[ref] = cell.comment if cell is not None else None
                updater.set_comments(name, comments)
            if sheet._dirty_hyperlinks:
                raise NotImplementedError(
                    "editing hyperlinks in a loaded workbook is not implemented yet; "
                    "the original hyperlinks are preserved when other cells are edited"
                )
        elif sheet._dirty_comments:
            comments = {}
            for ref in sheet._dirty_comments:
                col, row, _, _ = parse_cell_range(ref)
                cell = sheet._existing_cell(row + 1, col + 1)
                comments[ref] = cell.comment if cell is not None else None
            updater.set_comments(name, comments)

    _reorder_and_state(workbook, updater)
    _remap_local_names(updater, sheet_order, renamed)
    updater.save(filename)
    for sheet in workbook._sheets:
        sheet._original_title = sheet.title
        sheet._meta_signature = _meta_signature(sheet)
        sheet._clear_dirty()
    workbook._loaded_worksheet_names = set(workbook.sheetnames)
    workbook._styles_dirty = False
