"""Load XLSX/XLSM files into the openpyxl-like workbook model."""

from __future__ import annotations

import contextlib
import datetime as _datetime
import functools
import io
import os
import posixpath
import re
import zipfile
from typing import Any, overload

from .._untrusted import SafeZipFile
from xml.etree import ElementTree as ET
from .._safe_xml import fromstring as safe_fromstring, iterparse as safe_iterparse

from .._metadata import MAIN, REL, local_name as _local, qname, read_defined_names
from .._text import xstring_unescape
from ..comments import find_comment_parts, read_comments, rels_path
from ..data_validation import DataValidation, OPERATORS, VALIDATION_TYPES
from ..excel_dates import serial_to_datetime
from ..external_links import read_external_links
from ..hyperlinks import Hyperlink
from ..ignored_errors import FLAGS as IGNORED_ERROR_FLAGS, IgnoredError
from ..reader import ExcelReader
from ..refs import cell_position, column_label, parse_cell_range
from ..styles import (
    AboveAverageRule,
    BlanksRule,
    CellIsRule,
    ColorScaleRule,
    DataBarRule,
    DuplicateRule,
    FormatObject,
    FormulaRule,
    IconSetRule,
    TextRule,
    TimePeriodRule,
    TopBottomRule,
)
from ..updater_utils import parse_shared_strings_xml, resolve_zip_target
from ..xlsx_updater import XlsxUpdater
from ._model import (
    AutoFilter,
    BookView,
    Break,
    Cell,
    Chartsheet,
    ColBreak,
    DefinedNameList,
    DocumentProperties,
    FilterColumn,
    NamedStyleList,
    RowBreak,
    Selection,
    SortCondition,
    SortState,
    Table,
    TableColumn,
    TableStyleInfo,
    Workbook,
    Worksheet,
    _coerce_column,
    _coordinate,
    _normalize_range,
)
from ._loaded import _meta_signature, _workbook_meta_signature
from ._objects import mark_objects_saved, note_unread
from ._print import print_settings_for
from ._styles import XlsxStyleTable, _read_color

_NS_DC = "http://purl.org/dc/elements/1.1/"
_NS_DCTERMS = "http://purl.org/dc/terms/"
_NS_CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"


def _element_xml(element) -> str:
    """The XML of one element, without the text that follows it."""
    tail, element.tail = element.tail, None
    try:
        return ET.tostring(element, encoding="unicode")
    finally:
        element.tail = tail


@overload
def _bool_attr(node, name: str, default: bool) -> bool: ...


@overload
def _bool_attr(node, name: str, default: None = None) -> bool | None: ...


def _bool_attr(node, name: str, default: bool | None = None) -> bool | None:
    value = node.get(name) if node is not None else None
    if value is None:
        return default
    return value.lower() not in {"0", "false", "off"}


# Rows of an XLSX worksheet are loaded by regular expressions when possible
# (workbook/_fast_sheet.py); tests switch this off to compare both paths.
_FAST_SHEET_DATA = True

_Q_C = f"{{{MAIN}}}c"
_Q_F = f"{{{MAIN}}}f"
_Q_V = f"{{{MAIN}}}v"
_Q_IS = f"{{{MAIN}}}is"
_cell_position = cell_position


@functools.lru_cache(maxsize=1024)
def _date_kind(format_code: str | None) -> str | None:
    """Return ``"date"``, ``"time"`` or ``"duration"`` for a number format."""
    if not format_code or format_code == "General":
        return None
    stripped = re.sub(r'"[^"]*"', "", format_code)
    stripped = re.sub(r"\\.", "", stripped)
    # Keep elapsed-time brackets; drop colours, locales and conditions.
    def keep(match):
        inner = match.group(0)[1:-1]
        return match.group(0) if re.fullmatch(r"[hmsHMS]+", inner) else ""

    stripped = re.sub(r"\[[^\]]*\]", keep, stripped)
    lower = stripped.lower()
    if re.search(r"\[[hms]+\]", format_code, re.IGNORECASE):
        return "duration"
    has_date = any(char in lower for char in "yd")
    has_time = any(char in lower for char in "hs")
    if has_date:
        return "date"
    if has_time or (":" in lower and "m" in lower):
        return "time"
    return None


def _number_value(number: float, date_kind: str | None, date_1904: bool):
    """A numeric cell value: a date/time/duration for a date format, else int or float."""
    if date_kind is not None:
        try:
            if date_kind == "duration":
                return _datetime.timedelta(days=number)
            if date_kind == "time":
                seconds = (number % 1.0) * 86400.0
                whole = int(seconds)
                micros = round((seconds - whole) * 1_000_000)
                return _datetime.time(
                    (whole // 3600) % 24,
                    (whole % 3600) // 60,
                    whole % 60,
                    micros,
                )
            return serial_to_datetime(number, date_1904)
        except (ValueError, OverflowError):
            return number
    return int(number) if number.is_integer() else number


def _read_cell_value(cell: ET.Element, style, *, date_1904: bool, shared_strings: list[str],
           rich_text: bool = False):
    kind = cell.get("t", "n")
    if kind == "inlineStr":
        inline = cell.find(_Q_IS)
        if inline is None:
            return None
        if rich_text:
            from ..rich_text import rich_text_from_si

            return rich_text_from_si(inline)
        return xstring_unescape("".join(node.text or "" for node in inline.iter(qname("t"))))
    raw = cell.findtext(_Q_V, "")
    if kind == "s":
        try:
            index = int(raw)
        except ValueError:
            return raw
        return shared_strings[index] if 0 <= index < len(shared_strings) else raw
    if kind == "b":
        return raw == "1"
    if kind == "e":
        return raw
    if kind == "str":
        return xstring_unescape(raw)
    if kind == "d":
        try:
            return _datetime.datetime.fromisoformat(raw) if raw else None
        except ValueError:
            return raw
    if raw == "":
        return None
    try:
        number = float(raw)
    except ValueError:
        return raw
    format_code = style.number_format if style is not None else None
    return _number_value(number, _date_kind(format_code), date_1904)


def _cell_data(cell: ET.Element, style_table: XlsxStyleTable, *, data_only: bool,
               date_1904: bool, shared: dict[str, tuple[str, str]],
               shared_strings: list[str], rich_text: bool = False):
    style_id = int(cell.get("s", "0") or 0)
    style = style_table.style_at(style_id)
    formula = cell.find(_Q_F)
    if formula is not None and not data_only:
        text = formula.text or ""
        kind = formula.get("t")
        ref = cell.get("r")
        if kind == "shared":
            sid = formula.get("si", "")
            if text:
                shared[sid] = (ref or "", text)
            elif sid in shared:
                origin, master = shared[sid]
                if ref:
                    from .._xlsx_read import translate_formula

                    text = translate_formula(master, origin, ref)
            else:
                text = ""
        return None, "f", text, style_id
    value = _read_cell_value(
        cell, style, date_1904=date_1904, shared_strings=shared_strings,
        rich_text=rich_text,
    )
    kind = cell.get("t", "n")
    if kind in {"s", "inlineStr", "str"}:
        data_type = "s"
    elif kind == "b":
        data_type = "b"
    elif kind == "e":
        data_type = "e"
    elif kind == "d":
        data_type = "d"
    else:
        data_type = "n"
    return value, data_type, None, style_id


def _read_properties(package) -> DocumentProperties:
    props = DocumentProperties()
    if not package.has("docProps/core.xml"):
        return props
    try:
        root = safe_fromstring(package.get("docProps/core.xml"))
    except ET.ParseError:
        return props
    mapping = {
        f"{{{_NS_DC}}}creator": "creator",
        f"{{{_NS_DC}}}title": "title",
        f"{{{_NS_DC}}}subject": "subject",
        f"{{{_NS_DC}}}description": "description",
        f"{{{_NS_CP}}}keywords": "keywords",
        f"{{{_NS_CP}}}category": "category",
        f"{{{_NS_CP}}}lastModifiedBy": "lastModifiedBy",
        f"{{{_NS_CP}}}revision": "revision",
        f"{{{_NS_CP}}}contentStatus": "contentStatus",
        f"{{{_NS_DC}}}identifier": "identifier",
        f"{{{_NS_DC}}}language": "language",
        f"{{{_NS_CP}}}version": "version",
        f"{{{_NS_CP}}}lastPrinted": "lastPrinted",
    }
    for node in root:
        attr = mapping.get(node.tag)
        if attr is not None and node.text is not None:
            setattr(props, attr, node.text)
    for node, attr in (
        (root.find(f"{{{_NS_DCTERMS}}}created"), "created"),
        (root.find(f"{{{_NS_DCTERMS}}}modified"), "modified"),
        (root.find(f"{{{_NS_CP}}}lastPrinted"), "lastPrinted"),
    ):
        if node is not None and node.text:
            try:
                text = node.text.replace("Z", "+00:00")
                value = _datetime.datetime.fromisoformat(text)
                if value.tzinfo is not None:
                    value = value.replace(tzinfo=None)
                setattr(props, attr, value)
            except ValueError:
                pass
    return props


_NS_CUSTOM = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"


def _read_custom_properties(package):
    from ._model import CustomProperty, CustomPropertyList

    result = CustomPropertyList()
    if not package.has("docProps/custom.xml"):
        return result
    try:
        root = safe_fromstring(package.get("docProps/custom.xml"))
    except ET.ParseError:
        return result
    for node in root.findall(f"{{{_NS_CUSTOM}}}property"):
        name = node.get("name")
        if not name or len(node) == 0:
            continue
        value_node = node[0]
        kind = _local(value_node.tag)
        text = value_node.text or ""
        try:
            if kind in {"i1", "i2", "i4", "i8", "int", "uint"}:
                value: Any = int(text)
            elif kind in {"r4", "r8", "decimal"}:
                value = float(text)
            elif kind == "bool":
                value = text.lower() in {"1", "true"}
            elif kind == "filetime":
                value = _datetime.datetime.fromisoformat(text.replace("Z", "+00:00"))
            else:
                value = text
        except ValueError:
            value = text
        result.append(CustomProperty(name, value))
    return result


def _read_workbook_metadata(workbook: Workbook, root, package) -> None:
    security = workbook.security
    protection = root.find(f"{{{MAIN}}}workbookProtection")
    if protection is not None:
        for name in ("lockStructure", "lockWindows", "lockRevision"):
            if protection.get(name) is not None:
                setattr(security, name, _bool_attr(protection, name, None))
        for name in (
            "workbookPassword", "revisionsPassword",
            "workbookAlgorithmName", "workbookHashValue", "workbookSaltValue",
            "revisionsAlgorithmName", "revisionsHashValue", "revisionsSaltValue",
        ):
            if protection.get(name) is not None:
                setattr(security, name, protection.get(name))
        for name in ("workbookSpinCount", "revisionsSpinCount"):
            if protection.get(name) is not None:
                setattr(security, name, int(protection.get(name)))

    calculation = workbook.calculation
    calc = root.find(f"{{{MAIN}}}calcPr")
    if calc is not None:
        for name in ("calcId", "iterateCount", "concurrentManualCount"):
            if calc.get(name) is not None:
                setattr(calculation, name, int(calc.get(name)))
        for name in (
            "fullCalcOnLoad", "iterate", "fullPrecision", "calcCompleted",
            "calcOnSave", "concurrentCalc", "forceFullCalc",
        ):
            if calc.get(name) is not None:
                setattr(calculation, name, _bool_attr(calc, name, None))
        if calc.get("iterateDelta") is not None:
            calculation.iterateDelta = float(calc.get("iterateDelta"))
        for name in ("calcMode", "refMode"):
            if calc.get(name) is not None:
                setattr(calculation, name, calc.get(name))

    views = []
    for node in root.findall(f"{{{MAIN}}}bookViews/{{{MAIN}}}workbookView"):
        view = BookView()
        if node.get("visibility") is not None:
            view.visibility = node.get("visibility")
        for name in (
            "minimized", "showHorizontalScroll", "showVerticalScroll",
            "showSheetTabs", "autoFilterDateGrouping",
        ):
            if node.get(name) is not None:
                setattr(view, name, _bool_attr(node, name, None))
        for name in (
            "xWindow", "yWindow", "windowWidth", "windowHeight",
            "tabRatio", "firstSheet", "activeTab",
        ):
            if node.get(name) is not None:
                setattr(view, name, int(node.get(name)))
        views.append(view)
    if views:
        workbook.views = views
    workbook.custom_doc_props = _read_custom_properties(package)
    workbook.external_links = read_external_links(package, binary=False)


def _read_sheet_relationships(package, sheet_path: str) -> dict[str, tuple[str, str | None, str]]:
    path = rels_path(sheet_path)
    if not package.has(path):
        return {}
    root = safe_fromstring(package.get(path))
    result = {}
    for node in root:
        rid = node.get("Id")
        target = node.get("Target")
        if rid and target:
            result[rid] = (
                target, node.get("TargetMode"), node.get("Type", "")
            )
    return result


def _read_page_setup(sheet: Worksheet, node) -> None:
    """Read every ``<pageSetup>`` attribute into the model."""
    setup = sheet.page_setup
    if node.get("orientation") in ("portrait", "landscape"):
        setup.orientation = node.get("orientation")

    def number(name, low, high):
        try:
            value = int(node.get(name))
        except (TypeError, ValueError):
            return None
        return value if low <= value <= high else None

    setup.paperSize = number("paperSize", 1, 255)
    setup.scale = number("scale", 10, 400)
    setup.fitToWidth = number("fitToWidth", 0, 32767)
    setup.fitToHeight = number("fitToHeight", 0, 32767)
    setup.firstPageNumber = number("firstPageNumber", 0, 32767)
    setup.horizontalDpi = number("horizontalDpi", 0, 100000)
    setup.verticalDpi = number("verticalDpi", 0, 100000)
    setup.copies = number("copies", 0, 32767)
    setup.useFirstPageNumber = _bool_attr(node, "useFirstPageNumber", None)
    setup.usePrinterDefaults = _bool_attr(node, "usePrinterDefaults", None)
    setup.blackAndWhite = _bool_attr(node, "blackAndWhite", None)
    setup.draft = _bool_attr(node, "draft", None)
    setup.paperHeight = node.get("paperHeight")
    setup.paperWidth = node.get("paperWidth")
    setup.pageOrder = node.get("pageOrder")
    setup.cellComments = node.get("cellComments")
    setup.errors = node.get("errors")


_OXML_VALIDATION_TYPES = {value: key for key, value in VALIDATION_TYPES.items()}
_OXML_OPERATORS = {value: key for key, value in OPERATORS.items()}


def _read_data_validations(sheet: Worksheet, node) -> None:
    for child in node.findall(qname("dataValidation")):
        formula1 = child.findtext(qname("formula1"))
        if formula1 is None:
            note_unread(sheet, "data_validations")
            continue
        kind = _OXML_VALIDATION_TYPES.get(child.get("type", "custom"), "custom")
        operator = (
            _OXML_OPERATORS.get(child.get("operator"))
            if child.get("operator") is not None else None
        )
        prompt = child.get("prompt")
        error = child.get("error")
        show_input = _bool_attr(child, "showInputMessage", False)
        show_error = _bool_attr(child, "showErrorMessage", False)
        try:
            rule = DataValidation(
                type=kind,
                sqref=child.get("sqref", "A1"),
                formula1=formula1,
                formula2=child.findtext(qname("formula2")),
                operator=operator,
                allow_blank=_bool_attr(child, "allowBlank", True),
                show_dropdown=not _bool_attr(child, "showDropDown", False),
                # openpyxl stores text even when its display flag is off; the
                # constructor rejects that combination, so set it back after.
                show_input_message=show_input or bool(prompt),
                prompt_title=child.get("promptTitle"),
                prompt=prompt,
                show_error_message=show_error or bool(error),
                error_title=child.get("errorTitle"),
                error=error,
                error_style=child.get("errorStyle", "stop"),
                ime_mode=child.get("imeMode"),
            )
        except (TypeError, ValueError):
            note_unread(sheet, "data_validations")
            continue
        if prompt and not show_input:
            object.__setattr__(rule, "show_input_message", False)
        if error and not show_error:
            object.__setattr__(rule, "show_error_message", False)
        sheet._data_validations.append(rule)


def _parse_table(xml: bytes) -> Table:
    root = safe_fromstring(xml)
    attrs = root.attrib
    table = Table(
        id=int(attrs.get("id", "1") or 1),
        displayName=attrs.get("displayName"),
        ref=attrs.get("ref"),
        name=attrs.get("name"),
        comment=attrs.get("comment"),
        tableType=attrs.get("tableType"),
        headerRowCount=int(attrs.get("headerRowCount", "1") or 1),
        insertRow=_bool_attr(root, "insertRow", None),
        insertRowShift=_bool_attr(root, "insertRowShift", None),
        totalsRowCount=(
            int(attrs["totalsRowCount"]) if attrs.get("totalsRowCount") is not None else None
        ),
        totalsRowShown=_bool_attr(root, "totalsRowShown", None),
        published=_bool_attr(root, "published", None),
    )
    columns = root.find(qname("tableColumns"))
    if columns is not None:
        for node in columns.findall(qname("tableColumn")):
            table.tableColumns.append(
                TableColumn(
                    id=int(node.get("id")) if node.get("id") is not None else None,
                    name=node.get("name"),
                    totalsRowFunction=node.get("totalsRowFunction"),
                    totalsRowLabel=node.get("totalsRowLabel"),
                    calculatedColumnFormula=node.findtext(qname("calculatedColumnFormula")),
                    totalsRowFormula=node.findtext(qname("totalsRowFormula")),
                )
            )
    style = root.find(qname("tableStyleInfo"))
    if style is not None:
        table.tableStyleInfo = TableStyleInfo(
            name=style.get("name"),
            showFirstColumn=_bool_attr(style, "showFirstColumn", None),
            showLastColumn=_bool_attr(style, "showLastColumn", None),
            showRowStripes=_bool_attr(style, "showRowStripes", None),
            showColumnStripes=_bool_attr(style, "showColumnStripes", None),
        )
    auto_filter = root.find(qname("autoFilter"))
    if auto_filter is not None:
        table.autoFilter = AutoFilter(auto_filter.get("ref"))
    return table


def _read_table_parts(sheet, package, sheet_path, relationships, table_parts) -> None:
    for rid in table_parts:
        target, _mode, _kind = relationships.get(rid, (None, None, ""))
        if not target:
            continue
        part = resolve_zip_target(target, posixpath.dirname(sheet_path))
        if not package.has(part):
            continue
        try:
            sheet._tables.append(_parse_table(package.get(part)))
        except (ET.ParseError, TypeError, ValueError):
            continue


def _cfvo(node) -> FormatObject:
    kind, value = node.get("type", "num"), node.get("val")
    if value is not None and kind in ("num", "percent", "percentile"):
        # A number, as openpyxl reads it; anything else is a formula.
        try:
            number = float(value)
        except ValueError:
            pass
        else:
            value = int(number) if number.is_integer() else number
    return FormatObject(kind, value)


def _cf_rule(sqref: str, node, style_table):
    rule_type = node.get("type")
    priority = int(node.get("priority", "1") or 1)
    stop = _bool_attr(node, "stopIfTrue", None)
    dxf_id = node.get("dxfId")
    dxf = style_table.dxf_at(int(dxf_id)) if dxf_id is not None else None
    formulas = [item.text or "" for item in node.findall(qname("formula"))]

    if rule_type == "cellIs":
        operator = node.get("operator", "equal")
        return CellIsRule(sqref, operator, formulas, dxf, stop, priority)
    if rule_type == "expression":
        return FormulaRule(sqref, formulas, dxf, stop, priority)
    if rule_type == "colorScale":
        scale = node.find(qname("colorScale"))
        if scale is None:
            return None
        cfvos = [_cfvo(item) for item in scale.findall(qname("cfvo"))]
        colors = [_read_color(item) for item in scale.findall(qname("color"))]
        solid: list[Any] = [color for color in colors if color is not None]
        if len(cfvos) < 2 or len(solid) != len(colors):
            return None
        if len(cfvos) == 2:
            return ColorScaleRule(sqref, cfvos[0], None, cfvos[1], solid, priority)
        return ColorScaleRule(sqref, cfvos[0], cfvos[1], cfvos[2], solid, priority)
    if rule_type == "dataBar":
        bar = node.find(qname("dataBar"))
        if bar is None:
            return None
        cfvos = [_cfvo(item) for item in bar.findall(qname("cfvo"))]
        color = _read_color(bar.find(qname("color")))
        if len(cfvos) < 2 or color is None:
            return None
        return DataBarRule(
            sqref, cfvos[0], cfvos[1], color,
            show_value=_bool_attr(bar, "showValue", None),
            min_length=(
                int(bar.get("minLength")) if bar.get("minLength") is not None else None
            ),
            max_length=(
                int(bar.get("maxLength")) if bar.get("maxLength") is not None else None
            ),
            priority=priority,
        )
    if rule_type == "iconSet":
        icon_set = node.find(qname("iconSet"))
        if icon_set is None:
            return None
        cfvos = [_cfvo(item) for item in icon_set.findall(qname("cfvo"))]
        values: list[Any] = []
        for item in cfvos:
            value: Any = item.value  # a number, a formula, or None
            try:
                values.append(float(value))
            except (TypeError, ValueError):
                values.append(value)
        if not values:
            return None
        return IconSetRule(
            sqref,
            icon_set.get("iconSet", "3TrafficLights1"),
            values,
            cfvo_type=cfvos[0].type if cfvos else "percent",
            show_value=_bool_attr(icon_set, "showValue", None),
            percent=_bool_attr(icon_set, "percent", None),
            reverse=_bool_attr(icon_set, "reverse", None),
            stop_if_true=stop,
            priority=priority,
        )
    if rule_type == "top10":
        return TopBottomRule(
            sqref,
            rank=int(node.get("rank", "10") or 10),
            percent=bool(_bool_attr(node, "percent", False)),
            bottom=bool(_bool_attr(node, "bottom", False)),
            dxf=dxf,
            stop_if_true=stop,
            priority=priority,
        )
    if rule_type == "aboveAverage":
        std_dev = node.get("stdDev")
        return AboveAverageRule(
            sqref,
            below=bool(_bool_attr(node, "below", False)),
            equal_average=bool(_bool_attr(node, "equalAverage", False)),
            std_dev=int(std_dev) if std_dev is not None else None,
            dxf=dxf,
            stop_if_true=stop,
            priority=priority,
        )
    if rule_type in {"containsText", "notContains", "beginsWith", "endsWith"}:
        operator = "contains" if rule_type == "containsText" else rule_type
        return TextRule(
            sqref,
            operator,
            node.get("text", ""),
            dxf,
            stop,
            priority,
        )
    if rule_type in {"duplicateValues", "uniqueValues"}:
        return DuplicateRule(
            sqref,
            unique=rule_type == "uniqueValues",
            dxf=dxf,
            stop_if_true=stop,
            priority=priority,
        )
    if rule_type == "timePeriod":
        return TimePeriodRule(
            sqref,
            node.get("timePeriod", "today"),
            dxf,
            stop,
            priority,
        )
    if rule_type in {"containsBlanks", "notContainsBlanks", "containsErrors", "notContainsErrors"}:
        kind = {
            "containsBlanks": "blanks",
            "notContainsBlanks": "no_blanks",
            "containsErrors": "errors",
            "notContainsErrors": "no_errors",
        }[rule_type]
        return BlanksRule(sqref, kind, dxf, stop, priority)
    return None


def _read_conditional_formatting(sheet, node, style_table) -> None:
    sqref = node.get("sqref")
    if not sqref:
        return
    for child in node.findall(qname("cfRule")):
        try:
            rule = _cf_rule(sqref, child, style_table)
        except (TypeError, ValueError):
            rule = None
        if rule is None:
            note_unread(sheet, "conditional_formatting")
        else:
            sheet._conditional_formats.add(sqref, rule)


_PRINT_NAME_PART = re.compile(r"(?:'(?:[^']|'')*'|[^,'!]+)!([^,]+)")


def _read_print_names(sheet: Worksheet, names: dict[str, str]) -> None:
    """Print area and print titles from the sheet's local defined names."""
    area = names.get("_xlnm.Print_Area")
    if area:
        parts = [part.replace("$", "") for part in _PRINT_NAME_PART.findall(area)]
        try:
            for part in parts:
                parse_cell_range(part)
        except ValueError:
            parts = []
        if parts:
            sheet.print_area = ",".join(parts)
    titles = names.get("_xlnm.Print_Titles")
    for part in _PRINT_NAME_PART.findall(titles or ""):
        part = part.replace("$", "")
        if re.fullmatch(r"\d+:\d+", part):
            sheet.print_title_rows = part
        elif re.fullmatch(r"[A-Za-z]{1,3}:[A-Za-z]{1,3}", part):
            sheet.print_title_cols = part.upper()


def _apply_row_attributes(sheet: Worksheet, row_number: int, attrs) -> None:
    """Row height, visibility, outline and style from a ``<row>`` element's attributes."""
    # A group's summary row carries collapsed="1" alone; it has to
    # be loaded too, or saving row metadata would drop the flag.
    if not ("ht" in attrs or "hidden" in attrs or "outlineLevel" in attrs
            or "collapsed" in attrs):
        return
    dimension = sheet.row_dimensions[row_number]
    if "ht" in attrs:
        dimension.height = float(attrs["ht"])
    dimension.hidden = attrs.get("hidden", "0").lower() in {"1", "true"}
    dimension.outlineLevel = int(attrs.get("outlineLevel", "0") or 0)
    dimension.collapsed = attrs.get("collapsed", "0").lower() in {"1", "true"}
    best_fit = attrs.get("bestFit")
    dimension.bestFit = None if best_fit is None else best_fit.lower() not in {"0", "false", "off"}
    if attrs.get("s") is not None:
        dimension.style = int(attrs["s"])


def _load_row(sheet: Worksheet, node: ET.Element, next_row: int, style_table: XlsxStyleTable,
              *, data_only: bool, date_1904: bool, shared, shared_strings: list[str],
              rich_text: bool):
    """Load one ``<row>`` element into the sheet; return ``(row, max_row, max_column)``."""
    row_number = int(node.get("r", next_row) or next_row)
    max_row = row_number
    max_column = 0
    _apply_row_attributes(sheet, row_number, node.attrib)
    next_column = 0
    for cell in node:
        if cell.tag != _Q_C:
            continue
        ref = cell.get("r")
        if ref:
            col, row = _cell_position(ref)
        else:
            col, row = next_column, row_number - 1
        next_column = col + 1
        value, data_type, formula_text, style_id = _cell_data(
            cell, style_table, data_only=data_only,
            date_1904=date_1904, shared=shared,
            shared_strings=shared_strings, rich_text=rich_text,
        )
        if value is None and formula_text is None and style_id == 0:
            continue
        sheet._set_cell_raw(
            row + 1, col + 1, value,
            style_id=style_id, data_type=data_type,
            formula_text=formula_text,
        )
        max_row = max(max_row, row + 1)
        max_column = max(max_column, col + 1)
    return row_number, max_row, max_column


def _read_core_sheet(sheet: Worksheet, style_table: XlsxStyleTable,
                     *, data_only: bool, date_1904: bool, package,
                     sheet_path: str, shared_strings: list[str],
                     rich_text: bool = False):
    shared_formulas: dict[str, tuple[str, str]] = {}
    hyperlinks: list[dict] = []
    table_parts: list[str] = []
    next_row = 1
    max_row = 0
    max_column = 0
    fast = None
    if _FAST_SHEET_DATA:
        from ._fast_sheet import read_sheet_data

        fast = read_sheet_data(
            sheet, package, sheet_path, style_table, data_only=data_only,
            date_1904=date_1904, shared_strings=shared_strings, rich_text=rich_text,
        )
    if fast is None:
        source = package.open_entry(sheet_path)
    else:
        max_row, max_column, remainder = fast
        source = contextlib.nullcontext(io.BytesIO(remainder))
    with source as stream:
        stack = []
        for event, node in safe_iterparse(stream, events=("start", "end")):
            if event == "start":
                stack.append(node)
                continue
            tag = _local(node.tag)
            if tag == "row" and len(stack) >= 2 and _local(stack[-2].tag) == "sheetData":
                row_number, row_max, column_max = _load_row(
                    sheet, node, next_row, style_table, data_only=data_only,
                    date_1904=date_1904, shared=shared_formulas,
                    shared_strings=shared_strings, rich_text=rich_text,
                )
                next_row = row_number + 1
                max_row = max(max_row, row_max)
                max_column = max(max_column, column_max)
                stack[-2].remove(node)
                node.clear()
                stack.pop()
                continue
            if tag == "dimension" and node.get("ref"):
                try:
                    c1, r1, c2, r2 = parse_cell_range(node.get("ref"))
                    max_row = max(max_row, r2 + 1)
                    max_column = max(max_column, c2 + 1)
                except ValueError:
                    pass
            elif tag == "mergeCell" and node.get("ref"):
                try:
                    sheet._merged_cells.add(_normalize_range(node.get("ref")))
                except ValueError:
                    pass
            elif tag == "autoFilter":
                filter_ = AutoFilter(_normalize_range(node.get("ref")) if node.get("ref") else None)
                for child in node.findall(qname("filterColumn")):
                    column = FilterColumn(
                        colId=int(child.get("colId", "0") or 0),
                        hiddenButton=_bool_attr(child, "hiddenButton", False),
                        showButton=_bool_attr(child, "showButton", True),
                    )
                    filters = child.find(qname("filters"))
                    if filters is not None:
                        column.vals = [
                            item.get("val")
                            for item in filters.findall(qname("filter"))
                            if item.get("val") is not None
                        ]
                        column.blank = _bool_attr(filters, "blank", False)
                    # Any other criterion is kept verbatim: the model does not interpret it.
                    column.criteria = [
                        _element_xml(item) for item in child if item.tag != qname("filters")
                    ]
                    filter_.filterColumn.append(column)
                sort = node.find(qname("sortState"))
                if sort is not None:
                    state = SortState(
                        ref=sort.get("ref"),
                        columnSort=_bool_attr(sort, "columnSort", None),
                        caseSensitive=_bool_attr(sort, "caseSensitive", None),
                        sortMethod=sort.get("sortMethod"),
                    )
                    for condition in sort.findall(qname("sortCondition")):
                        dxf_id = condition.get("dxfId")
                        state.sortCondition.append(
                            SortCondition(
                                ref=condition.get("ref"),
                                descending=_bool_attr(condition, "descending", None),
                                sortBy=condition.get("sortBy"),
                                customList=condition.get("customList"),
                                dxfId=int(dxf_id) if dxf_id is not None else None,
                                iconSet=condition.get("iconSet"),
                                iconId=(
                                    int(condition.get("iconId"))
                                    if condition.get("iconId") is not None else None
                                ),
                            )
                        )
                    filter_.sortState = state
                sheet._auto_filter = filter_
            elif tag == "sheetPr":
                props = sheet.sheet_properties
                tab = node.find(qname("tabColor"))
                if tab is not None:
                    try:
                        props.tabColor = _read_color(tab)
                    except (TypeError, ValueError):
                        pass
                for attr, field in (
                    ("codeName", "codeName"),
                    ("filterMode", "filterMode"),
                    ("published", "published"),
                    ("syncHorizontal", "syncHorizontal"),
                    ("syncRef", "syncRef"),
                    ("syncVertical", "syncVertical"),
                    ("transitionEvaluation", "transitionEvaluation"),
                    ("transitionEntry", "transitionEntry"),
                    ("enableFormatConditionsCalculation", "enableFormatConditionsCalculation"),
                ):
                    if node.get(attr) is not None:
                        value = node.get(attr)
                        setattr(props, field, value if field in {"codeName", "syncRef"} else _bool_attr(node, attr, None))
                outline = node.find(qname("outlinePr"))
                if outline is not None:
                    for name in ("applyStyles", "summaryBelow", "summaryRight", "showOutlineSymbols"):
                        setattr(props.outlinePr, name, _bool_attr(outline, name, None))
                setup = node.find(qname("pageSetUpPr"))
                if setup is not None:
                    props.pageSetUpPr.fitToPage = _bool_attr(setup, "fitToPage", None)
                    if hasattr(props.pageSetUpPr, "autoPageBreaks"):
                        props.pageSetUpPr.autoPageBreaks = _bool_attr(setup, "autoPageBreaks", None)
            elif tag == "sheetFormatPr":
                fmt = sheet.sheet_format
                for name in ("baseColWidth", "outlineLevelRow", "outlineLevelCol"):
                    if node.get(name) is not None:
                        setattr(fmt, name, int(node.get(name)))
                for name in ("defaultColWidth", "defaultRowHeight"):
                    if node.get(name) is not None:
                        setattr(fmt, name, float(node.get(name)))
                for name in ("customHeight", "zeroHeight", "thickTop", "thickBottom"):
                    if node.get(name) is not None:
                        setattr(fmt, name, _bool_attr(node, name, None))
            elif tag == "pageSetup":
                _read_page_setup(sheet, node)
            elif tag == "printOptions":
                options = sheet.print_options
                for name in ("horizontalCentered", "verticalCentered", "headings", "gridLines", "gridLinesSet"):
                    if node.get(name) is not None:
                        setattr(options, name, _bool_attr(node, name, None))
            elif tag == "headerFooter":
                footer = sheet.HeaderFooter
                for name in ("differentFirst", "differentOddEven", "scaleWithDoc", "alignWithMargins"):
                    if node.get(name) is not None:
                        setattr(footer, name, _bool_attr(node, name, None))
            elif tag in ("oddHeader", "oddFooter", "evenHeader", "evenFooter", "firstHeader", "firstFooter") \
                    and len(stack) >= 2 and _local(stack[-2].tag) == "headerFooter":
                getattr(sheet.HeaderFooter, tag).load_code(node.text)
            elif tag == "sheetView":
                view = sheet._sheet_view
                for name in (
                    "showFormulas", "showGridLines", "showRowColHeaders", "showZeros",
                    "rightToLeft", "tabSelected", "showRuler", "showOutlineSymbols",
                    "defaultGridColor", "showWhiteSpace", "windowProtection",
                ):
                    if node.get(name) is not None:
                        setattr(view, name, _bool_attr(node, name, None))
                view.view = node.get("view")
                view.topLeftCell = node.get("topLeftCell")
                for name in ("colorId", "zoomScale", "zoomScaleNormal",
                             "zoomScaleSheetLayoutView", "zoomScalePageLayoutView",
                             "workbookViewId"):
                    if node.get(name) is not None:
                        setattr(view, name, int(node.get(name)))
            elif tag == "pane":
                from ._model import Pane

                pane = Pane(
                    xSplit=float(node.get("xSplit")) if node.get("xSplit") else None,
                    ySplit=float(node.get("ySplit")) if node.get("ySplit") else None,
                    topLeftCell=node.get("topLeftCell"),
                    activePane=node.get("activePane", "topLeft"),
                    state=node.get("state", "split"),
                )
                sheet._sheet_view.pane = pane
                if pane.state == "frozen" and pane.topLeftCell:
                    sheet._freeze_panes = pane.topLeftCell
            elif tag == "selection":
                sheet._sheet_view.selection.append(
                    Selection(
                        pane=node.get("pane"),
                        activeCell=node.get("activeCell", "A1"),
                        activeCellId=(
                            int(node.get("activeCellId"))
                            if node.get("activeCellId") is not None else None
                        ),
                        sqref=node.get("sqref", "A1"),
                    )
                )
            elif tag == "cols":
                for col in node.findall(qname("col")):
                    minimum = int(col.get("min", "1"))
                    maximum = int(col.get("max", str(minimum)))
                    width = col.get("width")
                    hidden = col.get("hidden", "0").lower() in {"1", "true"}
                    outline = int(col.get("outlineLevel", "0") or 0)
                    collapsed = col.get("collapsed", "0").lower() in {"1", "true"}
                    for index in range(minimum, maximum + 1):
                        dimension = sheet.column_dimensions[column_label(index - 1)]
                        dimension.min = minimum
                        dimension.max = maximum
                        if width is not None:
                            dimension.width = float(width)
                        dimension.hidden = hidden
                        dimension.outlineLevel = outline
                        dimension.collapsed = collapsed
                        dimension.bestFit = _bool_attr(col, "bestFit", None)
                        dimension.autoSize = _bool_attr(col, "autoSize", None)
                        if col.get("style") is not None:
                            dimension.style = int(col.get("style"))
            elif tag == "ignoredErrors":
                for item in node.findall(qname("ignoredError")):
                    if not item.get("sqref"):
                        continue
                    try:
                        sheet.ignored_errors.append(IgnoredError(
                            item.get("sqref"),
                            **{name: _bool_attr(item, name, None) for name in IGNORED_ERROR_FLAGS},
                        ))
                    except ValueError:
                        continue  # a range this library cannot parse; the check stays as saved
            elif tag == "rowBreaks":
                sheet.row_breaks = RowBreak([
                    Break(
                        id=int(item.get("id", "0") or 0),
                        min=int(item.get("min", "0") or 0),
                        max=int(item.get("max", "16383") or 16383),
                        man=_bool_attr(item, "man", True),
                        pt=int(item.get("pt")) if item.get("pt") is not None else None,
                    )
                    for item in node.findall(qname("brk"))
                ])
            elif tag == "colBreaks":
                sheet.col_breaks = ColBreak([
                    Break(
                        id=int(item.get("id", "0") or 0),
                        min=int(item.get("min", "0") or 0),
                        max=int(item.get("max", "1048575") or 1048575),
                        man=_bool_attr(item, "man", True),
                        pt=int(item.get("pt")) if item.get("pt") is not None else None,
                    )
                    for item in node.findall(qname("brk"))
                ])
            elif tag == "hyperlink":
                hyperlinks.append(dict(node.attrib))
            elif tag == "tablePart":
                rid = node.get(f"{{{REL}}}id")
                if rid:
                    table_parts.append(rid)
            elif tag == "dataValidations":
                _read_data_validations(sheet, node)
            elif tag == "conditionalFormatting":
                _read_conditional_formatting(sheet, node, style_table)
            elif tag == "sheetProtection":
                protection = sheet.protection
                for name in (
                    "sheet", "objects", "scenarios", "formatCells", "formatRows",
                    "formatColumns", "insertColumns", "insertRows", "insertHyperlinks",
                    "deleteColumns", "deleteRows", "selectLockedCells", "sort",
                    "autoFilter", "pivotTables", "selectUnlockedCells",
                ):
                    if node.get(name) is not None:
                        setattr(protection, name, _bool_attr(node, name, None))
                protection.set_password(node.get("password"), already_hashed=True)
                protection.algorithmName = node.get("algorithmName")
                protection.saltValue = node.get("saltValue")
                protection.hashValue = node.get("hashValue")
                if node.get("spinCount") is not None:
                    protection.spinCount = int(node.get("spinCount"))
            elif tag == "pageMargins":
                for name in ("left", "right", "top", "bottom", "header", "footer"):
                    value = node.get(name)
                    if value is not None:
                        setattr(sheet.page_margins, name, float(value))
            if stack:
                stack.pop()

    sheet._max_row = max_row or sheet._max_row
    sheet._max_column = max_column or sheet._max_column
    for merged in sheet.merged_cells.ranges:
        _, _, c2, r2 = parse_cell_range(merged)
        sheet._max_row = max(sheet._max_row, r2 + 1)
        sheet._max_column = max(sheet._max_column, c2 + 1)
    if sheet._max_row or sheet._max_column:
        sheet._max_row = max(sheet._max_row, 1)
        sheet._max_column = max(sheet._max_column, 1)

    # Hyperlinks, table parts and drawings need the worksheet relationships.
    rels = _read_sheet_relationships(package, sheet_path)
    _read_table_parts(sheet, package, sheet_path, rels, table_parts)
    from ..chart_reader import read_drawing

    for rel_target, _mode, kind in rels.values():
        if kind.endswith("/drawing"):
            drawing_part = resolve_zip_target(rel_target, posixpath.dirname(sheet_path))
            read_drawing(package, sheet, drawing_part)
    for attrs in hyperlinks:
        ref = attrs.get("ref")
        if not ref:
            continue
        rid = attrs.get(f"{{{REL}}}id")
        location = attrs.get("location")
        target: str | None = None
        if rid and rid in rels:
            raw_target, mode, _kind = rels[rid]
            if mode == "External":
                target = raw_target
        if target is None and location:
            target = "#" + location
        if target is None:
            continue
        try:
            col, row, _, _ = parse_cell_range(ref.split()[0])
            cell = sheet._existing_cell(row + 1, col + 1)
            if cell is None:
                cell = sheet._set_cell_raw(row + 1, col + 1, None)
            cell._hyperlink = Hyperlink(
                target,
                text=attrs.get("display"),
                tooltip=attrs.get("tooltip"),
            )
        except (TypeError, ValueError):
            continue

    # Classic notes are stored in a sibling comments part.
    try:
        _, _, parts = find_comment_parts(package.get, package.names(), sheet_path)
        if "comments" in parts:
            for ref, comment in read_comments(package.get(parts["comments"][0])).items():
                col, row, _, _ = parse_cell_range(ref)
                cell = sheet._existing_cell(row + 1, col + 1)
                if cell is None:
                    cell = sheet._set_cell_raw(row + 1, col + 1, None)
                cell._comment = comment
    except (KeyError, ValueError, ET.ParseError):
        pass


def _read_only_bounds(path: str, sheet_path: str) -> tuple[int, int, int, int] | None:
    """Return ``(min_col, min_row, max_col, max_row)`` from a worksheet
    ``dimension`` ref, or None when the sheet is unsized (as openpyxl)."""
    if not sheet_path.endswith(".xml"):
        return None
    try:
        with SafeZipFile(path) as archive:
            with archive.open(sheet_path) as stream:
                for _event, node in safe_iterparse(stream, events=("start",)):
                    if _local(node.tag) == "dimension" and node.get("ref"):
                        try:
                            c1, r1, c2, r2 = parse_cell_range(node.get("ref"))
                        except (TypeError, ValueError):
                            return None
                        return c1 + 1, r1 + 1, c2 + 1, r2 + 1
                    if _local(node.tag) == "sheetData":
                        return None
    except (KeyError, ET.ParseError, zipfile.BadZipFile):
        return None
    return None


class ReadOnlyCell:
    """A cell value streamed from a read-only worksheet (as openpyxl's)."""

    __slots__ = ("row", "column", "value", "number_format")

    def __init__(self, row: int, column: int, value, number_format="General"):
        self.row = row
        self.column = column
        self.value = value
        self.number_format = number_format

    @property
    def coordinate(self) -> str:
        return _coordinate(self.row, self.column)

    @property
    def column_letter(self) -> str:
        return column_label(self.column - 1)

    @property
    def data_type(self) -> str:
        if isinstance(self.value, str) and self.value.startswith("="):
            return "f"
        return Cell._type_for(self.value)

    @property
    def is_date(self) -> bool:
        return isinstance(self.value, (_datetime.date, _datetime.time, _datetime.timedelta))

    def __eq__(self, other):
        return (isinstance(other, ReadOnlyCell)
                and (self.row, self.column, self.value) == (other.row, other.column, other.value))

    def __repr__(self):
        return f"<ReadOnlyCell {self.coordinate}>"


class EmptyCell:
    """Filler for cells absent from a read-only worksheet (openpyxl EMPTY_CELL)."""

    __slots__ = ()
    value = None
    data_type = "n"
    is_date = False

    def __repr__(self):
        return "<EmptyCell>"


EMPTY_CELL = EmptyCell()


class ReadOnlyWorksheet(Worksheet):
    """Streams rows on demand with openpyxl's read-only semantics.

    Rows and columns use worksheet coordinates: rows missing from the file are
    yielded as empty rows and columns start at ``min_col`` (default 1). Bounds
    come from the stored dimension; without one, ``max_row`` and
    ``max_column`` are None and rows are as wide as their last cell.
    """

    def __init__(self, workbook, title: str, *, reader: ExcelReader, data_only: bool,
                 index: int | None = None):
        super().__init__(workbook, title, index=index)
        self._reader = reader
        self._data_only = data_only
        self._loaded = True
        self._min_row = self._min_column = 1
        # Unknown until the bounds are read; the class docstring says so.
        self._max_row = self._max_column = None  # type: ignore[assignment]

    def _set_bounds(self, bounds):
        if bounds is not None:
            self._min_column, self._min_row, self._max_column, self._max_row = bounds

    @property
    def min_row(self) -> int:
        return self._min_row

    @property
    def min_column(self) -> int:
        return self._min_column

    @property
    def max_row(self):
        return self._max_row

    @property
    def max_column(self):
        return self._max_column

    @property
    def dimensions(self) -> str:
        return self.calculate_dimension()

    def calculate_dimension(self, force: bool = False) -> str:
        if not (self._max_row and self._max_column):
            if not force:
                raise ValueError("Worksheet is unsized, use calculate_dimension(force=True)")
            self._calculate_dimension()
        return (f"{_coordinate(self._min_row, self._min_column)}:"
                f"{_coordinate(self._max_row, self._max_column)}")

    def _calculate_dimension(self):
        max_row = max_column = 0
        for row_number, cells in self._indexed_rows():
            if cells:
                max_row = row_number
                max_column = max(max_column, cells[-1][0] + 1)
        self._max_row, self._max_column = max_row or 1, max_column or 1

    def reset_dimensions(self):
        """Ignore a wrong stored dimension, as openpyxl's reset_dimensions()."""
        self._max_row = self._max_column = None  # type: ignore[assignment]  # unknown until the bounds are read

    def _indexed_rows(self):
        if self._reader._is_xlsb:
            from .._xlsb_read import iter_xlsb_indexed_rows as iter_xlsb_rows
            return iter_xlsb_rows(self._reader, self.title, data_only=self._data_only)
        from .._xlsx_read import iter_xlsx_indexed_rows as iter_xlsx_rows
        return iter_xlsx_rows(self._reader, self.title, self._data_only)

    def _cells_by_row(self, min_col, min_row, max_col, max_row, values_only):
        filler = None if values_only else EMPTY_CELL
        empty_row = (filler,) * (max_col + 1 - min_col) if max_col is not None else ()
        counter = min_row
        last = 1
        for row_number, cells in self._indexed_rows():
            last = row_number
            if max_row is not None and row_number > max_row:
                break
            # Rows missing from the file are yielded as empty rows.
            for _ in range(counter, row_number):
                counter += 1
                yield empty_row
            if counter <= row_number:
                counter += 1
                yield self._row(row_number, cells, min_col, max_col, values_only)
        # As openpyxl: pad up to max_row only when the file has rows past it.
        if max_row is not None and max_row < last:
            for _ in range(counter, max_row + 1):
                yield empty_row

    def _row(self, row_number, cells, min_col, max_col, values_only):
        if not cells and not max_col:
            return ()
        last = max_col or cells[-1][0] + 1
        if last < min_col:
            return ()
        row: list[Any] = [None if values_only else EMPTY_CELL] * (last + 1 - min_col)
        for column, value in cells:
            column += 1
            if min_col <= column <= last:
                row[column - min_col] = value if values_only else ReadOnlyCell(row_number, column, value)
        return tuple(row)

    def iter_rows(self, min_row=None, max_row=None, min_col=None, max_col=None,
                  values_only: bool = False):
        min_col = _coerce_column(min_col) if min_col else 1
        min_row = int(min_row) if min_row else 1
        max_col = (_coerce_column(max_col) if max_col else None) or self._max_column
        max_row = int(max_row) if max_row else self._max_row
        return self._cells_by_row(min_col, min_row, max_col, max_row, values_only)

    def iter_cols(self, *args, **kwargs):
        raise AttributeError("iter_cols is not available in read-only mode")

    def cell(self, row, column, value=None):
        if value is not None:
            raise AttributeError("read-only worksheets cannot be modified")
        column = _coerce_column(column)
        for cells in self._cells_by_row(column, int(row), column, int(row), False):
            if cells:
                return cells[0]
        return EMPTY_CELL

    def __getitem__(self, key):
        if isinstance(key, int):
            return next(self.iter_rows(min_row=key, max_row=key), ())
        if isinstance(key, str) and ":" not in key:
            col, row, _, _ = parse_cell_range(key)
            return self.cell(row + 1, col + 1)
        c1, r1, c2, r2 = parse_cell_range(key)
        return tuple(self.iter_rows(min_row=r1 + 1, max_row=r2 + 1, min_col=c1 + 1, max_col=c2 + 1))

    def __setitem__(self, key, value):
        raise AttributeError("read-only worksheets cannot be modified")

    def append(self, iterable):
        raise AttributeError("read-only worksheets cannot be modified")


def load_workbook(filename, *, read_only: bool = False, data_only: bool = False,
                  keep_vba: bool = False, rich_text: bool = False,
                  keep_links: bool = True) -> Workbook:
    """Load an XLSX/XLSM workbook, or stream an XLSB/XLSX with ``read_only``.

    The fast ``ExcelReader`` remains the best choice for raw value extraction.
    This function adds coordinates, styles, merges, dimensions and editing.
    """
    path = os.fspath(filename)
    lower = path.lower()
    if read_only:
        reader = ExcelReader(path)
        reader.open(path)
        workbook = Workbook()
        workbook._sheets.clear()
        workbook._read_only = True
        workbook._loaded = True
        workbook._source_path = path
        for index, name in enumerate(reader.get_sheet_names()):
            sheet = ReadOnlyWorksheet(
                workbook, name, reader=reader, data_only=data_only, index=index
            )
            if reader._is_xlsb:
                from .._xlsb_read import sheet_dimension
                sheet._set_bounds(sheet_dimension(reader, name))
            else:
                rid = reader._worksheet_name_to_id.get(name)
                location = reader._worksheet_id_to_location.get(rid) if rid else None
                if location:
                    sheet._set_bounds(_read_only_bounds(path, reader._resolve_path(location)))
            workbook._sheets.append(sheet)
        if workbook._sheets:
            workbook._active_index = 0
        return workbook

    if lower.endswith(".xlsb"):
        from ._xlsb_reader import load_workbook_xlsb

        return load_workbook_xlsb(path, data_only=data_only)

    updater = XlsxUpdater(path)
    try:
        workbook = _load_xlsx_workbook(
            updater, path, data_only=data_only, rich_text=rich_text
        )
        if keep_vba:
            # As openpyxl: the source package, from which a new workbook
            # saved as .xlsm takes its VBA project. A loaded workbook keeps
            # its own whether or not this is set.
            import io

            with open(path, "rb") as source:
                workbook.vba_archive = SafeZipFile(io.BytesIO(source.read()))
        return workbook
    except BaseException:
        # A failed load (or an interrupt) must not leave the updater's temporary package behind.
        updater.close()
        raise


def _load_xlsx_workbook(updater, path: str, *, data_only: bool, rich_text: bool) -> Workbook:
    package = updater._package
    workbook = Workbook()
    workbook._sheets.clear()
    workbook._loaded = True
    workbook._updater = updater
    workbook._source_path = path
    workbook._loaded_worksheet_names = set()
    workbook._properties = _read_properties(package)

    workbook_xml = package.get("xl/workbook.xml").decode("utf-8")
    workbook_root = safe_fromstring(workbook_xml)
    _read_workbook_metadata(workbook, workbook_root, package)
    workbook._loaded_workbook_meta = _workbook_meta_signature(workbook)
    rels_root = safe_fromstring(package.get("xl/_rels/workbook.xml.rels"))
    relationships = {}
    for node in rels_root:
        rid = node.get("Id")
        target = node.get("Target")
        if rid and target:
            relationships[rid] = (
                node.get("Type", ""),
                resolve_zip_target(target, "xl"),
            )
    active_tab = 0
    view = workbook_root.find(f"{{{MAIN}}}bookViews/{{{MAIN}}}workbookView")
    if view is not None:
        try:
            active_tab = int(view.get("activeTab", "0"))
        except ValueError:
            active_tab = 0

    worksheet_entries = []
    chartsheet_entries = []
    sheet_nodes = workbook_root.findall(f"{{{MAIN}}}sheets/{{{MAIN}}}sheet")
    # localSheetId counts every sheet, chartsheets included.
    print_names: dict[str, dict[str, str]] = {}
    for node in workbook_root.findall(f"{{{MAIN}}}definedNames/{{{MAIN}}}definedName"):
        if node.get("name") not in ("_xlnm.Print_Area", "_xlnm.Print_Titles"):
            continue
        try:
            owner = sheet_nodes[int(node.get("localSheetId"))].get("name")
        except (TypeError, ValueError, IndexError):
            continue
        print_names.setdefault(owner, {})[node.get("name")] = node.text or ""
    for sheet_node in sheet_nodes:
        name = sheet_node.get("name")
        rid = sheet_node.get(f"{{{REL}}}id")
        if not name or not rid or rid not in relationships:
            continue
        rel_type, target = relationships[rid]
        if rel_type.endswith("/worksheet"):
            worksheet_entries.append((name, target, sheet_node.get("state", "visible")))
        elif rel_type.endswith("/chartsheet"):
            chartsheet_entries.append((name, target, sheet_node.get("state", "visible")))

    style_table = XlsxStyleTable.from_xml(
        package.get("xl/styles.xml") if package.has("xl/styles.xml") else None
    )
    workbook._styles = style_table
    named_styles: list[Any] = list(style_table.named_style_defs.values())
    if not named_styles:
        named_styles = list(style_table.named_styles) or ["Normal"]
    workbook.named_styles = NamedStyleList(named_styles)
    workbook._loaded_worksheet_names = {name for name, _path, _state in worksheet_entries}

    shared_strings: list = []
    if package.has("xl/sharedStrings.xml"):
        if rich_text:
            from ..rich_text import parse_rich_shared_strings

            shared_strings = parse_rich_shared_strings(package.get("xl/sharedStrings.xml"))
        else:
            shared_strings = [
                xstring_unescape(value)
                for value in parse_shared_strings_xml(package.get("xl/sharedStrings.xml"))
            ]

    for index, (name, sheet_path, state) in enumerate(worksheet_entries):
        sheet = Worksheet(workbook, name, index=index)
        sheet._loaded = True
        sheet._path = sheet_path
        sheet._sheet_state = state
        _read_core_sheet(
            sheet,
            style_table,
            data_only=data_only,
            date_1904=updater._date_1904,
            package=package,
            sheet_path=sheet_path,
            shared_strings=shared_strings,
            rich_text=rich_text,
        )
        _read_print_names(sheet, print_names.get(name, {}))
        sheet._meta_signature = _meta_signature(sheet)
        sheet._saved_print_settings = print_settings_for(sheet)
        mark_objects_saved(sheet)
        workbook._sheets.append(sheet)

    if not workbook._sheets and not chartsheet_entries:
        raise ValueError("No worksheets were found in workbook.xml")
    from ..chart_reader import read_chartsheet

    for index, (name, part, state) in enumerate(chartsheet_entries):
        chartsheet = Chartsheet(workbook, name, index=index)
        chartsheet.sheet_state = state
        chartsheet._part = part
        read_chartsheet(package, part, name, chartsheet)
        workbook._chartsheets.append(chartsheet)
    if not workbook._sheets:
        workbook._sheets = []
        workbook._active_index = 0
    else:
        workbook._active_index = min(max(active_tab, 0), len(workbook._sheets) - 1)
    workbook._loaded_active_index = workbook._active_index
    from ._props import remember as _remember_properties
    _remember_properties(workbook)
    try:
        # localSheetId counts every <sheet>, including chartsheets.
        all_sheets = [node.get("name", "") for node in workbook_root.findall(f"{{{MAIN}}}sheets/{{{MAIN}}}sheet")]
        workbook._defined_names = DefinedNameList(
            workbook, read_defined_names(workbook_xml, all_sheets)
        )
    except (TypeError, ValueError, IndexError, ET.ParseError):
        workbook._defined_names = DefinedNameList(workbook)
    return workbook
