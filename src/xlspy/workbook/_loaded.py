"""Save edits back into an existing XLSX/XLSM package.

Only changed worksheet parts and the workbook/style metadata are rewritten.
Untouched parts (drawings, charts, pivots, macros, embedded objects) are copied
byte-for-byte by ``ZipPackage``.
"""

from __future__ import annotations

import datetime as _datetime
import re
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape, quoteattr

from .._metadata import MAIN, REL, PKG_REL, CT, element_pattern, replace_element, serialize_like
from .._xlsx_edit import _attributes, _bytes, _cell_value
from ..formula.sheets import rename_sheet_references
from ..formulas import Formula
from ._drawings import check_removals, remove_drawing_objects
from ._objects import (
    add_objects, check_object_rewrites, has_pending_objects, mark_objects_saved, sync_objects,
)
from ..refs import column_label, parse_cell_range
from ..styles import _coerce_color, _color_key
from ..updater_utils import renamed_chart_parts, replace_or_add_xml_attribute
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
    if getattr(dimension, "bestFit", None) is not None:
        attrs["bestFit"] = "1" if dimension.bestFit else "0"
    if getattr(dimension, "style", None) is not None:
        attrs["s"] = str(dimension.style)
        attrs["customFormat"] = "1"
    return attrs


def _snapshot(obj) -> tuple:
    return tuple((name, getattr(obj, name)) for name in obj.__dataclass_fields__)


def _header_footer_snapshot(footer) -> tuple:
    return (
        footer.differentFirst, footer.differentOddEven,
        footer.scaleWithDoc, footer.alignWithMargins,
        footer.oddHeader.code(), footer.oddFooter.code(),
        footer.evenHeader.code(), footer.evenFooter.code(),
        footer.firstHeader.code(), footer.firstFooter.code(),
    )


def _break_snapshot(container) -> tuple:
    return tuple((item.id, item.min, item.max, item.man, item.pt) for item in container.brk)


def _auto_filter_snapshot(filter_) -> tuple:
    sort = filter_.sortState
    return (
        filter_.ref,
        tuple((column.colId, tuple(column.vals), column.blank) for column in filter_.filterColumn),
        None if sort is None else (
            sort.ref,
            tuple((item.ref, item.descending, item.sortBy) for item in sort.sortCondition),
        ),
    )


def _meta_signature(sheet: Worksheet):
    """A hashable snapshot of the sheet metadata edited through this model."""
    cols = tuple(
        (letter, dim.width, dim.hidden, dim.outlineLevel, dim.collapsed,
         dim.bestFit, dim.autoSize, dim.style)
        for letter, dim in sorted(
            sheet.column_dimensions.items(),
            key=lambda item: parse_cell_range(f"{item[0]}1")[0],
        )
    )
    rows = tuple(
        (row, dim.height, dim.hidden, dim.outlineLevel, dim.collapsed,
         dim.bestFit, dim.style)
        for row, dim in sorted(sheet.row_dimensions.items())
    )
    view = sheet.sheet_view
    pane = None if view.pane is None else (
        view.pane.xSplit, view.pane.ySplit, view.pane.topLeftCell,
        view.pane.activePane, view.pane.state,
    )
    selection = tuple(
        (item.pane, item.activeCell, item.activeCellId, item.sqref)
        for item in view.selection
    )
    properties = sheet.sheet_properties
    return (
        cols, rows, sheet.freeze_panes,
        _auto_filter_snapshot(sheet.auto_filter),
        tuple(sorted(str(value) for value in sheet.merged_cells.ranges)),
        repr(properties.tabColor), properties.codeName,
        properties.filterMode, properties.published, properties.syncHorizontal,
        properties.syncRef, properties.syncVertical, properties.transitionEvaluation,
        properties.transitionEntry, properties.enableFormatConditionsCalculation,
        _snapshot(properties.outlinePr), _snapshot(properties.pageSetUpPr),
        _snapshot(sheet.sheet_format),
        (
            view.windowProtection, view.showFormulas, view.showGridLines,
            view.showRowColHeaders, view.showZeros, view.rightToLeft,
            view.tabSelected, view.showRuler, view.showOutlineSymbols,
            view.defaultGridColor, view.showWhiteSpace, view.view,
            view.topLeftCell, view.colorId, view.zoomScale, view.zoomScaleNormal,
            view.zoomScaleSheetLayoutView, view.zoomScalePageLayoutView,
            view.workbookViewId, pane, selection,
        ),
        _snapshot(sheet.page_setup), _snapshot(sheet.print_options),
        _snapshot(sheet.page_margins), _header_footer_snapshot(sheet.HeaderFooter),
        _snapshot(sheet.protection),
        _break_snapshot(sheet.row_breaks), _break_snapshot(sheet.col_breaks),
        sheet.print_area, sheet.print_title_rows, sheet.print_title_cols,
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
    for name in (
        "showFormulas", "showGridLines", "showRowColHeaders", "showZeros",
        "rightToLeft", "tabSelected", "showRuler", "showOutlineSymbols",
        "defaultGridColor", "showWhiteSpace", "windowProtection",
    ):
        value = getattr(view, name, None)
        if value is not None:
            attrs.append(f'{name}="{int(bool(value))}"')
    for name in ("view", "topLeftCell"):
        value = getattr(view, name, None)
        if value is not None:
            attrs.append(f'{name}="{value}"')
    for name in (
        "colorId", "zoomScale", "zoomScaleNormal", "zoomScaleSheetLayoutView",
        "zoomScalePageLayoutView",
    ):
        value = getattr(view, name, None)
        if value is not None:
            attrs.append(f'{name}="{value}"')

    pane = ""
    if view.pane is not None:
        pane_attrs = []
        if view.pane.xSplit is not None:
            pane_attrs.append(f'xSplit="{view.pane.xSplit:g}"')
        if view.pane.ySplit is not None:
            pane_attrs.append(f'ySplit="{view.pane.ySplit:g}"')
        if view.pane.topLeftCell is not None:
            pane_attrs.append(f'topLeftCell="{view.pane.topLeftCell}"')
        pane_attrs.append(f'activePane="{view.pane.activePane}"')
        pane_attrs.append(f'state="{view.pane.state}"')
        pane = f'<{prefix}pane {" ".join(pane_attrs)}/>'
    elif sheet.freeze_panes:
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

    if view.selection:
        selection = "".join(
            f'<{prefix}selection'
            + (f' pane="{item.pane}"' if item.pane is not None else "")
            + f' activeCell="{item.activeCell}" sqref="{item.sqref}"'
            + (f' activeCellId="{item.activeCellId}"' if item.activeCellId is not None else "")
            + "/>"
            for item in view.selection
        )
    elif sheet.freeze_panes:
        col, row, _, _ = parse_cell_range(sheet.freeze_panes)
        active = "bottomRight" if col and row else ("topRight" if col else "bottomLeft")
        selection = (
            f'<{prefix}selection pane="{active}" '
            f'activeCell="{sheet.freeze_panes}" sqref="{sheet.freeze_panes}"/>'
        )
    else:
        selection = ""
    view_attrs = (" " + " ".join(attrs)) if attrs else ""
    return (
        f'<{prefix}sheetViews><{prefix}sheetView{view_attrs} '
        f'workbookViewId="{view.workbookViewId}">'
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
        if getattr(dimension, "bestFit", None) is not None:
            attrs.append(f'bestFit="{1 if dimension.bestFit else 0}"')
        if getattr(dimension, "autoSize", None) is not None:
            attrs.append(f'autoSize="{1 if dimension.autoSize else 0}"')
        if getattr(dimension, "style", None) is not None:
            attrs.append(f'style="{dimension.style}"')
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
        new_opening = opening
        # The reader loads every row's attributes, so a row the model has no
        # dimension for (cleared, or moved by insert_rows) has none.
        for name in ("ht", "customHeight", "hidden", "outlineLevel", "collapsed"):
            new_opening = _remove_attr(new_opening, name)
        if dimension is None:
            return new_opening + row_xml[len(opening):]
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


def _number_text(value) -> str:
    return f"{value:g}" if isinstance(value, float) else str(value)


def _sheet_format_xml(sheet: Worksheet, prefix: str) -> str:
    fmt = sheet.sheet_format
    attrs = []
    for name in ("baseColWidth", "outlineLevelRow", "outlineLevelCol"):
        value = getattr(fmt, name)
        if value is not None:
            attrs.append(f'{name}="{value}"')
    for name in ("defaultColWidth", "defaultRowHeight"):
        value = getattr(fmt, name)
        if value is not None:
            attrs.append(f'{name}="{_number_text(value)}"')
    for name in ("customHeight", "zeroHeight", "thickTop", "thickBottom"):
        value = getattr(fmt, name)
        if value is not None:
            attrs.append(f'{name}="{int(bool(value))}"')
    if not attrs:
        return ""
    return f'<{prefix}sheetFormatPr {" ".join(attrs)}/>'


def _sheet_pr_xml(sheet: Worksheet, prefix: str) -> str:
    properties = sheet.sheet_properties
    attrs = []
    for name in ("codeName", "syncRef"):
        value = getattr(properties, name)
        if value is not None:
            attrs.append(f"{name}={quoteattr(str(value))}")
    for name in (
        "enableFormatConditionsCalculation", "filterMode", "published",
        "syncHorizontal", "syncVertical", "transitionEvaluation", "transitionEntry",
    ):
        value = getattr(properties, name)
        if value is not None:
            attrs.append(f'{name}="{int(bool(value))}"')
    children = []
    tab = _tab_color_xml(sheet, prefix)
    if tab:
        children.append(tab)
    outline = properties.outlinePr
    outline_attrs = []
    for name in ("applyStyles", "summaryBelow", "summaryRight", "showOutlineSymbols"):
        value = getattr(outline, name)
        if value is not None:
            outline_attrs.append(f'{name}="{int(bool(value))}"')
    if outline_attrs:
        children.append(f'<{prefix}outlinePr {" ".join(outline_attrs)}/>')
    page_attrs = []
    if properties.pageSetUpPr.fitToPage is not None:
        page_attrs.append(f'fitToPage="{int(bool(properties.pageSetUpPr.fitToPage))}"')
    if page_attrs:
        children.append(f'<{prefix}pageSetUpPr {" ".join(page_attrs)}/>')
    if not attrs and not children:
        return ""
    attr_text = (" " + " ".join(attrs)) if attrs else ""
    body = "".join(children)
    return f'<{prefix}sheetPr{attr_text}>{body}</{prefix}sheetPr>'


def _page_setup_xml(sheet: Worksheet, prefix: str) -> str:
    setup = sheet.page_setup
    attrs = []
    for name in (
        "orientation", "paperSize", "scale", "firstPageNumber", "fitToHeight",
        "fitToWidth", "paperHeight", "paperWidth", "pageOrder", "cellComments",
        "errors", "horizontalDpi", "verticalDpi", "copies",
    ):
        value = getattr(setup, name)
        if value is not None:
            attrs.append(f"{name}={quoteattr(_number_text(value))}")
    for name in ("useFirstPageNumber", "usePrinterDefaults", "blackAndWhite", "draft"):
        value = getattr(setup, name)
        if value is not None:
            attrs.append(f'{name}="{int(bool(value))}"')
    if not attrs:
        return ""
    namespace = f' xmlns="{MAIN}"' if prefix else ""
    return f"<pageSetup{namespace} {' '.join(attrs)}/>"


def _print_options_xml(sheet: Worksheet, prefix: str) -> str:
    options = sheet.print_options
    attrs = []
    for name in ("horizontalCentered", "verticalCentered", "headings", "gridLines", "gridLinesSet"):
        value = getattr(options, name)
        if value is not None:
            attrs.append(f'{name}="{int(bool(value))}"')
    if not attrs:
        return ""
    return f'<{prefix}printOptions {" ".join(attrs)}/>'


def _page_margins_xml(sheet: Worksheet, prefix: str) -> str:
    margins = sheet.page_margins
    attrs = " ".join(
        f'{name}="{_number_text(getattr(margins, name))}"'
        for name in ("left", "right", "top", "bottom", "header", "footer")
    )
    return f"<{prefix}pageMargins {attrs}/>"


def _header_footer_xml(sheet: Worksheet, prefix: str) -> str:
    footer = sheet.HeaderFooter
    attrs = []
    for name in ("differentFirst", "differentOddEven", "scaleWithDoc", "alignWithMargins"):
        value = getattr(footer, name)
        if value is not None:
            attrs.append(f'{name}="{int(bool(value))}"')
    children = []
    for name in ("oddHeader", "oddFooter", "evenHeader", "evenFooter", "firstHeader", "firstFooter"):
        code = getattr(footer, name).code()
        if code:
            children.append(f"<{prefix}{name}>{escape(code)}</{prefix}{name}>")
    if not attrs and not children:
        return ""
    attr_text = (" " + " ".join(attrs)) if attrs else ""
    return f'<{prefix}headerFooter{attr_text}>{"".join(children)}</{prefix}headerFooter>'


def _breaks_xml(container, tag: str, prefix: str, default_max: int) -> str:
    if not container.brk:
        return ""
    items = []
    for item in container.brk:
        items.append(
            f'<{prefix}brk id="{item.id}" min="{item.min}" max="{item.max or default_max}"'
            f' man="{1 if item.man else 0}"'
            + (f' pt="{item.pt}"' if item.pt is not None else "")
            + "/>"
        )
    return (
        f'<{prefix}{tag} count="{len(container.brk)}" '
        f'manualBreakCount="{len(container.brk)}">{"".join(items)}</{prefix}{tag}>'
    )


def _protection_xml(sheet: Worksheet, prefix: str) -> str:
    protection = sheet.protection
    if not protection.sheet and protection.password is None:
        return ""
    attrs = [f'sheet="{1 if protection.sheet else 0}"']
    for name in (
        "objects", "scenarios", "formatCells", "formatRows", "formatColumns",
        "insertColumns", "insertRows", "insertHyperlinks", "deleteColumns",
        "deleteRows", "selectLockedCells", "selectUnlockedCells", "sort",
        "autoFilter", "pivotTables",
    ):
        attrs.append(f'{name}="{1 if getattr(protection, name) else 0}"')
    if protection.password is not None:
        attrs.append(f"password={quoteattr(protection.password)}")
    for name in ("algorithmName", "saltValue", "hashValue"):
        value = getattr(protection, name)
        if value is not None:
            attrs.append(f"{name}={quoteattr(str(value))}")
    if protection.spinCount is not None:
        attrs.append(f'spinCount="{protection.spinCount}"')
    return f'<{prefix}sheetProtection {" ".join(attrs)}/>'


def _auto_filter_xml(sheet: Worksheet, prefix: str) -> str:
    filter_ = sheet.auto_filter
    if not filter_.ref and not filter_.filterColumn and filter_.sortState is None:
        return ""
    attrs = f' ref="{_normalize_range(filter_.ref)}"' if filter_.ref else ""
    children = []
    for column in filter_.filterColumn:
        column_attrs = [f'colId="{column.colId}"']
        if column.hiddenButton:
            column_attrs.append('hiddenButton="1"')
        if not column.showButton:
            column_attrs.append('showButton="0"')
        values = "".join(
            f'<{prefix}filter val={quoteattr(str(value))}/>' for value in column.vals
        )
        if values or column.blank:
            blank = ' blank="1"' if column.blank else ""
            values = f'<{prefix}filters{blank}>{values}</{prefix}filters>'
        children.append(f'<{prefix}filterColumn {" ".join(column_attrs)}>{values}</{prefix}filterColumn>')
    sort = filter_.sortState
    if sort is not None:
        sort_attrs = []
        if sort.ref:
            sort_attrs.append(f'ref="{sort.ref}"')
        if sort.columnSort is not None:
            sort_attrs.append(f'columnSort="{int(bool(sort.columnSort))}"')
        if sort.caseSensitive is not None:
            sort_attrs.append(f'caseSensitive="{int(bool(sort.caseSensitive))}"')
        if sort.sortMethod is not None:
            sort_attrs.append(f'sortMethod={quoteattr(sort.sortMethod)}')
        conditions = "".join(
            f'<{prefix}sortCondition ref="{item.ref}"'
            + (f' descending="{int(bool(item.descending))}"' if item.descending is not None else "")
            + (f' sortBy={quoteattr(item.sortBy)}' if item.sortBy is not None else "")
            + "/>"
            for item in sort.sortCondition
        )
        children.append(f'<{prefix}sortState {" ".join(sort_attrs)}>{conditions}</{prefix}sortState>')
    body = "".join(children)
    return f'<{prefix}autoFilter{attrs}>{body}</{prefix}autoFilter>' if body else f'<{prefix}autoFilter{attrs}/>'


def patch_sheet_metadata(updater, sheet: Worksheet, name: str) -> None:
    path, xml = updater._sheet_xml(name)
    prefix = updater._tag_prefix(xml, "sheetData")
    xml = replace_element(
        xml, "sheetPr", _sheet_pr_xml(sheet, prefix),
        before=("dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData"),
    )
    xml = replace_element(
        xml, "sheetViews", _sheet_views_xml(sheet, prefix),
        before=("sheetFormatPr", "cols", "sheetData"),
    )
    xml = replace_element(
        xml, "sheetFormatPr", _sheet_format_xml(sheet, prefix),
        before=("cols", "sheetData"),
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
        xml, "sheetProtection", _protection_xml(sheet, prefix),
        before=("protectedRanges", "scenarios", "autoFilter", "sortState",
                "mergeCells", "conditionalFormatting", "dataValidations",
                "hyperlinks", "printOptions", "pageMargins", "pageSetup",
                "headerFooter", "drawing", "legacyDrawing", "tableParts", "extLst"),
    )
    xml = replace_element(
        xml, "autoFilter", _auto_filter_xml(sheet, prefix),
        before=("sortState", "mergeCells", "conditionalFormatting", "dataValidations",
                "hyperlinks", "printOptions", "pageMargins", "pageSetup",
                "headerFooter", "drawing", "legacyDrawing", "tableParts", "extLst"),
    )
    xml = replace_element(
        xml, "mergeCells", _merge_xml(sheet, prefix),
        before=("conditionalFormatting", "dataValidations", "hyperlinks",
                "printOptions", "pageMargins", "pageSetup", "headerFooter",
                "drawing", "legacyDrawing", "tableParts", "extLst"),
    )
    xml = replace_element(
        xml, "printOptions", _print_options_xml(sheet, prefix),
        before=("pageMargins", "pageSetup", "headerFooter", "rowBreaks",
                "colBreaks", "drawing", "legacyDrawing", "tableParts", "extLst"),
    )
    xml = replace_element(
        xml, "pageMargins", _page_margins_xml(sheet, prefix),
        before=("pageSetup", "headerFooter", "rowBreaks", "colBreaks",
                "drawing", "legacyDrawing", "tableParts", "extLst"),
    )
    xml = replace_element(
        xml, "pageSetup", _page_setup_xml(sheet, prefix),
        before=("headerFooter", "rowBreaks", "colBreaks", "drawing",
                "legacyDrawing", "tableParts", "extLst"),
    )
    xml = replace_element(
        xml, "headerFooter", _header_footer_xml(sheet, prefix),
        before=("rowBreaks", "colBreaks", "drawing", "legacyDrawing",
                "tableParts", "extLst"),
    )
    xml = replace_element(
        xml, "rowBreaks", _breaks_xml(sheet.row_breaks, "rowBreaks", prefix, 16383),
        before=("colBreaks", "drawing", "legacyDrawing", "tableParts", "extLst"),
    )
    xml = replace_element(
        xml, "colBreaks", _breaks_xml(sheet.col_breaks, "colBreaks", prefix, 1048575),
        before=("drawing", "legacyDrawing", "tableParts", "extLst"),
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


def _save_hyperlinks(updater, sheet: Worksheet, name: str) -> None:
    if not sheet._dirty_hyperlinks:
        return
    links = {}
    for ref in sheet._dirty_hyperlinks:
        col, row, _, _ = parse_cell_range(ref)
        cell = sheet._existing_cell(row + 1, col + 1)
        links[ref] = cell.hyperlink if cell is not None else None
    updater.set_hyperlinks(name, links)


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

    def rename_references(match):
        fragment = match[0]
        opening = re.match(r"<[^>]+>", fragment)[0]
        body = fragment[len(opening):]
        return opening + rename_sheet_references(body, old, new)

    # Formulas of defined names hold sheet titles (the cells' were updated in the model).
    workbook_xml = element_pattern("definedName").sub(rename_references, workbook_xml)
    package.set(workbook_path, workbook_xml.encode("utf-8"))
    for part, data in renamed_chart_parts(package, old, new).items():
        package.set(part, data)
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


def _sheet_states(updater) -> dict[str, str]:
    root = ET.fromstring(updater._package.get("xl/workbook.xml"))
    return {
        node.get("name", ""): node.get("state", "visible")
        for node in root.findall(f"{{{MAIN}}}sheets/{{{MAIN}}}sheet")
    }


def _active_tab(updater) -> int:
    root = ET.fromstring(updater._package.get("xl/workbook.xml"))
    view = root.find(f"{{{MAIN}}}bookViews/{{{MAIN}}}workbookView")
    if view is None:
        return 0
    try:
        return int(view.get("activeTab", "0"))
    except ValueError:
        return 0


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


def _workbook_meta_signature(workbook):
    return (
        _snapshot(workbook.security),
        _snapshot(workbook.calculation),
        tuple(_snapshot(view) for view in workbook.views),
    )


def _workbook_protection_xml(workbook, prefix: str) -> str:
    security = workbook.security
    attrs = []
    for name in ("lockStructure", "lockWindows", "lockRevision"):
        value = getattr(security, name)
        if value is not None:
            attrs.append(f'{name}="{int(bool(value))}"')
    for name in (
        "workbookPassword", "revisionsPassword",
        "workbookAlgorithmName", "workbookHashValue", "workbookSaltValue",
        "revisionsAlgorithmName", "revisionsHashValue", "revisionsSaltValue",
    ):
        value = getattr(security, name)
        if value is not None:
            attrs.append(f"{name}={quoteattr(str(value))}")
    for name in ("workbookSpinCount", "revisionsSpinCount"):
        value = getattr(security, name)
        if value is not None:
            attrs.append(f'{name}="{value}"')
    if not attrs:
        return ""
    return f'<{prefix}workbookProtection {" ".join(attrs)}/>'


def _calc_properties_xml(workbook, prefix: str) -> str:
    calc = workbook.calculation
    attrs = []
    for name in ("calcId", "iterateCount", "concurrentManualCount"):
        value = getattr(calc, name)
        if value is not None:
            attrs.append(f'{name}="{value}"')
    if calc.iterateDelta is not None:
        attrs.append(f'iterateDelta="{_number_text(calc.iterateDelta)}"')
    for name in (
        "fullCalcOnLoad", "iterate", "fullPrecision", "calcCompleted",
        "calcOnSave", "concurrentCalc", "forceFullCalc",
    ):
        value = getattr(calc, name)
        if value is not None:
            attrs.append(f'{name}="{int(bool(value))}"')
    for name in ("calcMode", "refMode"):
        value = getattr(calc, name)
        if value is not None:
            attrs.append(f"{name}={quoteattr(str(value))}")
    if not attrs:
        return ""
    return f'<{prefix}calcPr {" ".join(attrs)}/>'


def _book_views_xml(workbook, prefix: str) -> str:
    if not workbook.views:
        return ""
    views = []
    for view in workbook.views:
        attrs = [f"visibility={quoteattr(str(view.visibility))}"]
        for name in (
            "minimized", "showHorizontalScroll", "showVerticalScroll",
            "showSheetTabs", "autoFilterDateGrouping",
        ):
            attrs.append(f'{name}="{int(bool(getattr(view, name)))}"')
        for name in (
            "xWindow", "yWindow", "windowWidth", "windowHeight",
            "tabRatio", "firstSheet", "activeTab",
        ):
            value = getattr(view, name)
            if value is not None:
                attrs.append(f'{name}="{value}"')
        views.append(f'<{prefix}workbookView {" ".join(attrs)}/>')
    return f'<{prefix}bookViews>{"".join(views)}</{prefix}bookViews>'


def patch_workbook_metadata(updater, workbook: Workbook) -> None:
    path = "xl/workbook.xml"
    xml = updater._package.get(path).decode("utf-8")
    prefix = updater._tag_prefix(xml, "sheets")
    xml = replace_element(
        xml, "workbookProtection", _workbook_protection_xml(workbook, prefix),
        before=("bookViews", "sheets", "functionGroups", "externalReferences",
                "definedNames", "calcPr", "oleSize", "customWorkbookViews",
                "pivotCaches", "extLst"),
    )
    xml = replace_element(
        xml, "bookViews", _book_views_xml(workbook, prefix),
        before=("sheets", "functionGroups", "externalReferences", "definedNames",
                "calcPr", "oleSize", "customWorkbookViews", "pivotCaches", "extLst"),
    )
    xml = replace_element(
        xml, "calcPr", _calc_properties_xml(workbook, prefix),
        before=("oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr",
                "smartTagTypes", "webPublishing", "fileRecoveryPr",
                "webPublishObjects", "extLst"),
    )
    updater._package.set(path, xml.encode("utf-8"))


def save_loaded(workbook: Workbook, filename) -> None:
    updater = workbook._updater
    if updater is None:
        raise ValueError("this workbook was not loaded from an XLSX package")
    from ._chartsheets import check_chartsheets

    check_chartsheets(workbook)
    # Refusals first, before the package changes.
    for sheet in workbook._sheets:
        check_removals(sheet)
        check_object_rewrites(sheet)
    if workbook._styles_dirty:
        updater._package.set("xl/styles.xml", workbook._styles.to_bytes())
        workbook._styles_dirty = False

    if (
        getattr(workbook, "_loaded_workbook_meta", None)
        != _workbook_meta_signature(workbook)
    ):
        patch_workbook_metadata(updater, workbook)
        workbook._loaded_workbook_meta = _workbook_meta_signature(workbook)

    sheet_order = _sheet_order(updater)
    original_states = _sheet_states(updater)
    original_active = _active_tab(updater)
    original_loaded_order = [
        name for name in sheet_order if name in workbook._loaded_worksheet_names
    ]
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
            _save_hyperlinks(updater, sheet, name)
        else:
            if sheet._dirty_comments:
                comments = {}
                for ref in sheet._dirty_comments:
                    col, row, _, _ = parse_cell_range(ref)
                    cell = sheet._existing_cell(row + 1, col + 1)
                    comments[ref] = cell.comment if cell is not None else None
                updater.set_comments(name, comments)
            # A new sheet's part has no links yet: write all of them.
            sheet._dirty_hyperlinks.update(
                cell.coordinate for cells in sheet._cells.values() for cell in cells.values()
                if cell._hyperlink is not None
            )
            _save_hyperlinks(updater, sheet, name)

    desired_order = [sheet._original_title for sheet in workbook._sheets]
    state_changed = any(
        sheet.sheet_state != original_states.get(sheet._original_title, "visible")
        for sheet in workbook._sheets
    )
    renamed_any = any(sheet.title != sheet._original_title for sheet in workbook._sheets)
    active_changed = workbook._active_index != getattr(
        workbook, "_loaded_active_index", original_active
    )
    order_changed = desired_order != original_loaded_order
    changed_sheets = added or removed_any or renamed_any or order_changed
    if (
        changed_sheets
        or state_changed
        or active_changed
    ):
        _reorder_and_state(workbook, updater)
    if removed_any or renamed_any or order_changed:
        _remap_local_names(updater, sheet_order, renamed)
    # After the sheets have their final order: print names count it.
    final_order = _sheet_order(updater)
    for sheet in workbook._sheets:
        remove_drawing_objects(updater, sheet)
        sync_objects(updater, workbook, sheet)
        if has_pending_objects(sheet):
            if workbook._styles_dirty:
                updater._package.set("xl/styles.xml", workbook._styles.to_bytes())
                workbook._styles_dirty = False
            add_objects(updater, workbook, sheet, final_order)
    from ._chartsheets import add_chartsheets
    from ._props import save_properties

    add_chartsheets(updater, workbook)
    save_properties(updater, workbook)
    updater.save(filename)
    for sheet in workbook._sheets:
        sheet._original_title = sheet.title
        sheet._meta_signature = _meta_signature(sheet)
        sheet._clear_dirty()
        mark_objects_saved(sheet)
    workbook._loaded_worksheet_names = set(workbook.sheetnames)
    workbook._loaded_active_index = workbook._active_index
    workbook._styles_dirty = False
