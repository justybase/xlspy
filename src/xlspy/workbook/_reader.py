"""Load XLSX/XLSM files into the openpyxl-like workbook model."""

from __future__ import annotations

import datetime as _datetime
import os
import re
import zipfile
from xml.etree import ElementTree as ET

from .._metadata import MAIN, REL, PKG_REL, read_defined_names
from .._text import xstring_unescape
from ..comments import find_comment_parts, read_comments, rels_path
from ..excel_dates import serial_to_datetime
from ..hyperlinks import Hyperlink
from ..reader import ExcelReader
from ..refs import column_label, parse_cell_range
from ..updater_utils import parse_relationships_xml, parse_shared_strings_xml, resolve_zip_target
from ..xlsx_updater import XlsxUpdater
from ._model import (
    AutoFilter,
    Cell,
    DocumentProperties,
    MultiCellRange,
    SheetProperties,
    SheetView,
    Workbook,
    Worksheet,
    _coerce_column,
    _coordinate,
    _normalize_range,
)
from ._loaded import _meta_signature
from ._styles import XlsxStyleTable, _read_color

_NS_DC = "http://purl.org/dc/elements/1.1/"
_NS_DCTERMS = "http://purl.org/dc/terms/"
_NS_CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
_NS_APP = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"


def _q(tag: str) -> str:
    return f"{{{MAIN}}}{tag}"


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _text(node, path: str) -> str | None:
    found = node.find(path) if node is not None else None
    return found.text if found is not None else None


def _bool_attr(node, name: str, default: bool | None = None) -> bool | None:
    value = node.get(name) if node is not None else None
    if value is None:
        return default
    return value.lower() not in {"0", "false", "off"}


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


def _value(cell: ET.Element, style, *, date_1904: bool, shared_strings: list[str],
           rich_text: bool = False):
    kind = cell.get("t", "n")
    if kind == "inlineStr":
        inline = cell.find(_q("is"))
        if inline is None:
            return None
        if rich_text:
            from ..rich_text import rich_text_from_si

            return rich_text_from_si(inline)
        return xstring_unescape("".join(node.text or "" for node in inline.iter(_q("t"))))
    raw = cell.findtext(_q("v"), "")
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
    date_kind = _date_kind(format_code)
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


def _cell_data(cell: ET.Element, style_table: XlsxStyleTable, *, data_only: bool,
               date_1904: bool, shared: dict[str, tuple[str, str]],
               shared_strings: list[str], rich_text: bool = False):
    style_id = int(cell.get("s", "0") or 0)
    style = style_table.style_at(style_id)
    formula = cell.find(_q("f"))
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
    value = _value(
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
        root = ET.fromstring(package.get("docProps/core.xml"))
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
    }
    for node in root:
        attr = mapping.get(node.tag)
        if attr is not None and node.text is not None:
            setattr(props, attr, node.text)
    for node, attr in (
        (root.find(f"{{{_NS_DCTERMS}}}created"), "created"),
        (root.find(f"{{{_NS_DCTERMS}}}modified"), "modified"),
    ):
        if node is not None and node.text:
            try:
                text = node.text.replace("Z", "+00:00")
                setattr(props, attr, _datetime.datetime.fromisoformat(text))
            except ValueError:
                pass
    return props


def _read_sheet_relationships(package, sheet_path: str) -> dict[str, tuple[str, str | None]]:
    path = rels_path(sheet_path)
    if not package.has(path):
        return {}
    root = ET.fromstring(package.get(path))
    result = {}
    for node in root:
        rid = node.get("Id")
        target = node.get("Target")
        if rid and target:
            result[rid] = (target, node.get("TargetMode"))
    return result


def _read_core_sheet(sheet: Worksheet, style_table: XlsxStyleTable,
                     *, data_only: bool, date_1904: bool, package,
                     sheet_path: str, shared_strings: list[str],
                     rich_text: bool = False):
    shared_formulas: dict[str, tuple[str, str]] = {}
    hyperlinks: list[dict] = []
    next_row = 1
    max_row = 0
    max_column = 0
    with package.open_entry(sheet_path) as stream:
        stack = []
        for event, node in ET.iterparse(stream, events=("start", "end")):
            if event == "start":
                stack.append(node)
                continue
            tag = _local(node.tag)
            if tag == "row" and len(stack) >= 2 and _local(stack[-2].tag) == "sheetData":
                row_number = int(node.get("r", next_row) or next_row)
                next_row = row_number + 1
                max_row = max(max_row, row_number)
                attrs = node.attrib
                if "ht" in attrs or "hidden" in attrs or "outlineLevel" in attrs:
                    dimension = sheet.row_dimensions[row_number]
                    if "ht" in attrs:
                        dimension.height = float(attrs["ht"])
                    dimension.hidden = attrs.get("hidden", "0").lower() in {"1", "true"}
                    dimension.outlineLevel = int(attrs.get("outlineLevel", "0") or 0)
                    dimension.collapsed = attrs.get("collapsed", "0").lower() in {"1", "true"}
                next_column = 0
                for cell in node:
                    if cell.tag != _q("c"):
                        continue
                    ref = cell.get("r")
                    if ref:
                        col, row, _, _ = parse_cell_range(ref)
                        col += 0
                    else:
                        col, row = next_column, row_number - 1
                        ref = _coordinate(row + 1, col + 1)
                    next_column = col + 1
                    value, data_type, formula_text, style_id = _cell_data(
                        cell, style_table, data_only=data_only,
                        date_1904=date_1904, shared=shared_formulas,
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
            elif tag == "autoFilter" and node.get("ref"):
                sheet._auto_filter = AutoFilter(_normalize_range(node.get("ref")))
            elif tag == "sheetPr":
                tab = node.find(_q("tabColor"))
                if tab is not None:
                    try:
                        sheet._sheet_properties.tabColor = _read_color(tab)
                    except (TypeError, ValueError):
                        pass
            elif tag == "sheetView":
                view = sheet._sheet_view
                view.showGridLines = _bool_attr(node, "showGridLines", True)
                view.showRowColHeaders = _bool_attr(node, "showRowColHeaders", True)
                view.zoomScale = int(node.get("zoomScale", "100")) if node.get("zoomScale") else None
                view.zoomScaleNormal = int(node.get("zoomScaleNormal", "100")) if node.get("zoomScaleNormal") else None
                view.tabSelected = _bool_attr(node, "tabSelected", False)
                view.rightToLeft = _bool_attr(node, "rightToLeft", False)
            elif tag == "pane" and node.get("state") == "frozen" and node.get("topLeftCell"):
                sheet._freeze_panes = node.get("topLeftCell")
            elif tag == "cols":
                for col in node.findall(_q("col")):
                    minimum = int(col.get("min", "1"))
                    maximum = int(col.get("max", str(minimum)))
                    width = col.get("width")
                    hidden = col.get("hidden", "0").lower() in {"1", "true"}
                    outline = int(col.get("outlineLevel", "0") or 0)
                    collapsed = col.get("collapsed", "0").lower() in {"1", "true"}
                    for index in range(minimum, maximum + 1):
                        dimension = sheet.column_dimensions[column_label(index - 1)]
                        if width is not None:
                            dimension.width = float(width)
                        dimension.hidden = hidden
                        dimension.outlineLevel = outline
                        dimension.collapsed = collapsed
            elif tag == "hyperlink":
                hyperlinks.append(dict(node.attrib))
            elif tag == "sheetProtection":
                from .._metadata import SheetProtection

                try:
                    sheet._sheet_protection = SheetProtection(
                        enabled=node.get("sheet", "1").lower() not in {"0", "false"},
                        select_locked_cells=node.get("selectLockedCells", "1").lower() in {"0", "false"},
                        select_unlocked_cells=node.get("selectUnlockedCells", "1").lower() in {"0", "false"},
                    )
                except (TypeError, ValueError):
                    pass
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

    # Hyperlinks need the worksheet relationships.
    rels = _read_sheet_relationships(package, sheet_path)
    for attrs in hyperlinks:
        ref = attrs.get("ref")
        if not ref:
            continue
        rid = attrs.get(f"{{{REL}}}id")
        location = attrs.get("location")
        target = None
        if rid and rid in rels:
            raw_target, mode = rels[rid]
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
        with zipfile.ZipFile(path) as archive:
            with archive.open(sheet_path) as stream:
                for _event, node in ET.iterparse(stream, events=("start",)):
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

    __slots__ = ("row", "column", "value")

    def __init__(self, row: int, column: int, value):
        self.row = row
        self.column = column
        self.value = value

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
        self._max_row = self._max_column = None

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
        self._max_row = self._max_column = None

    def _indexed_rows(self):
        if self._reader._is_xlsb:
            from .._xlsb_read import iter_indexed_rows
            return iter_indexed_rows(self._reader, self.title, data_only=self._data_only)
        from .._xlsx_read import iter_indexed_rows
        return iter_indexed_rows(self._reader, self.title, self._data_only)

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
        row = [None if values_only else EMPTY_CELL] * (last + 1 - min_col)
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
        raise NotImplementedError(
            "XLSB does not have an openpyxl-like workbook model yet; "
            "use ExcelReader for values or XlsbUpdater for sparse edits"
        )

    updater = XlsxUpdater(path)
    package = updater._package
    workbook = Workbook()
    workbook._sheets.clear()
    workbook._loaded = True
    workbook._updater = updater
    workbook._source_path = path
    workbook._loaded_worksheet_names = set()
    workbook._properties = _read_properties(package)

    workbook_xml = package.get("xl/workbook.xml").decode("utf-8")
    workbook_root = ET.fromstring(workbook_xml)
    rels_root = ET.fromstring(package.get("xl/_rels/workbook.xml.rels"))
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
    for sheet_node in workbook_root.findall(f"{{{MAIN}}}sheets/{{{MAIN}}}sheet"):
        name = sheet_node.get("name")
        rid = sheet_node.get(f"{{{REL}}}id")
        if not name or not rid or rid not in relationships:
            continue
        rel_type, target = relationships[rid]
        if not rel_type.endswith("/worksheet"):
            continue
        worksheet_entries.append((name, target, sheet_node.get("state", "visible")))

    style_table = XlsxStyleTable.from_xml(
        package.get("xl/styles.xml") if package.has("xl/styles.xml") else None
    )
    workbook._styles = style_table
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
        sheet._meta_signature = _meta_signature(sheet)
        workbook._sheets.append(sheet)

    if not workbook._sheets:
        raise ValueError("No worksheets were found in workbook.xml")
    workbook._active_index = min(max(active_tab, 0), len(workbook._sheets) - 1)
    try:
        # localSheetId counts every <sheet>, including chartsheets.
        all_sheets = [node.get("name", "") for node in workbook_root.findall(f"{{{MAIN}}}sheets/{{{MAIN}}}sheet")]
        workbook._defined_names = read_defined_names(workbook_xml, all_sheets)
    except (TypeError, ValueError, IndexError, ET.ParseError):
        workbook._defined_names = []
    return workbook
