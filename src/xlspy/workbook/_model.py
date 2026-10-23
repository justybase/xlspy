"""Openpyxl-like workbook model.

This module deliberately contains no XLSX parsing or serialisation code.  The
fast path (``XlsxWriter``/``XlsbWriter``/``ExcelReader``) does not import it.
"""

from __future__ import annotations

from collections.abc import Iterable, Iterator, Mapping
from dataclasses import dataclass, field, replace
import datetime as _datetime
from decimal import Decimal
import math
import numbers
import re
from typing import Any

from ..charts import Chart
from ..comments import Comment
from ..data_validation import DataValidation
from ..formats import DEFAULT_DATE_FORMAT, DEFAULT_DATETIME_FORMAT
from ..formulas import Formula
from ..hyperlinks import Hyperlink
from ..images import Image
from ..refs import column_label, parse_cell_range
from ..rich_text import CellRichText
from ..styles import (
    Alignment,
    Border,
    CellIsRule,
    Color,
    ColorScaleRule,
    DataBarRule,
    EXTRA_CF_RULES,
    ExcelTable,
    Font,
    FormulaRule,
    GradientFill,
    IconSetRule,
    PatternFill,
    Protection,
    Style,
)
from . import _shift
from ._print import (
    HeaderFooter,
    PageMargins,
    PageSetup,
    PageSetupProperties,
    PrintOptions,
)
from ._styles import XlsxStyleTable

_CELL_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?([1-9][0-9]*)$")


def _coerce_column(column: int | str) -> int:
    if isinstance(column, str):
        match = _CELL_RE.match(column)
        if match:
            column = match.group(1)
        if not re.fullmatch(r"[A-Za-z]{1,3}", column):
            raise ValueError(f"invalid column {column!r}")
        value = 0
        for char in column.upper():
            value = value * 26 + ord(char) - 64
        return value
    if not isinstance(column, int) or isinstance(column, bool):
        raise TypeError("column must be a 1-based int or a column letter")
    if not 1 <= column <= 16384:
        raise ValueError("column must be 1..16384")
    return column


def _coerce_row(row: int) -> int:
    if not isinstance(row, int) or isinstance(row, bool):
        raise TypeError("row must be a 1-based int")
    if not 1 <= row <= 1048576:
        raise ValueError("row must be 1..1048576")
    return row


def _coordinate(row: int, column: int) -> str:
    return f"{column_label(column - 1)}{row}"


def _normalize_range(reference: str) -> str:
    if not isinstance(reference, str):
        raise TypeError("a range must be an A1 string")
    first, _, last = reference.partition(":")
    if not last:
        last = first
    c1, r1, c2, r2 = parse_cell_range(f"{first}:{last}")
    return f"{_coordinate(r1 + 1, c1 + 1)}:{_coordinate(r2 + 1, c2 + 1)}"


class CellRange(str):
    """An A1 range string with openpyxl-style boundary properties."""

    def __new__(cls, value: str):
        normalized = _normalize_range(value)
        return str.__new__(cls, normalized)

    @property
    def min_col(self):
        return parse_cell_range(self)[0] + 1

    @property
    def min_row(self):
        return parse_cell_range(self)[1] + 1

    @property
    def max_col(self):
        return parse_cell_range(self)[2] + 1

    @property
    def max_row(self):
        return parse_cell_range(self)[3] + 1

    def __contains__(self, item):
        if not isinstance(item, str):
            return False
        try:
            col, row, _, _ = parse_cell_range(item)
        except (TypeError, ValueError):
            return False
        return (
            self.min_col - 1 <= col <= self.max_col - 1
            and self.min_row - 1 <= row <= self.max_row - 1
        )


class MultiCellRange:
    def __init__(self, ranges: Iterable[str] | None = None):
        self.ranges: list[CellRange] = []
        if ranges is not None:
            for value in ranges:
                self.add(value)

    def add(self, value: str) -> None:
        cell_range = CellRange(value)
        if cell_range not in self.ranges:
            self.ranges.append(cell_range)

    def remove(self, value: str) -> None:
        cell_range = CellRange(value)
        for index, existing in enumerate(self.ranges):
            if existing == cell_range:
                del self.ranges[index]
                return
        raise ValueError(f"range {value!r} is not merged")

    def __contains__(self, item) -> bool:
        if isinstance(item, str) and ":" not in item:
            return any(item in cell_range for cell_range in self.ranges)
        try:
            cell_range = CellRange(item)
        except (TypeError, ValueError):
            return False
        return cell_range in self.ranges

    def __iter__(self):
        return iter(self.ranges)

    def __len__(self):
        return len(self.ranges)

    def __bool__(self):
        return bool(self.ranges)

    def __str__(self):
        return " ".join(self.ranges)

    def __repr__(self):
        return f"MultiCellRange({list(self.ranges)!r})"


@dataclass
class DocumentProperties:
    creator: str = "xlspy"
    title: str | None = None
    subject: str | None = None
    description: str | None = None
    keywords: str | None = None
    category: str | None = None
    lastModifiedBy: str | None = None
    created: _datetime.datetime = field(default_factory=_datetime.datetime.now)
    modified: _datetime.datetime = field(default_factory=_datetime.datetime.now)
    revision: str | None = None
    contentStatus: str | None = None
    identifier: str | None = None
    language: str | None = None
    version: str | None = None


@dataclass
class SheetProperties:
    tabColor: Color | str | None = None
    codeName: str | None = None
    pageSetUpPr: PageSetupProperties = field(default_factory=PageSetupProperties)


@dataclass
class SheetView:
    showGridLines: bool | None = None
    showRowColHeaders: bool | None = None
    zoomScale: int | None = None
    zoomScaleNormal: int | None = None
    tabSelected: bool | None = None
    rightToLeft: bool | None = None


@dataclass
class AutoFilter:
    ref: str | None = None

    def add_filter_column(self, col_id, vals, blank=False):
        raise NotImplementedError("AutoFilter column filters are not implemented yet")

    def add_sort_condition(self, ref, descending=False):
        raise NotImplementedError("AutoFilter sort conditions are not implemented yet")


_CF_RULES = (CellIsRule, FormulaRule, ColorScaleRule, DataBarRule, IconSetRule, *EXTRA_CF_RULES)


def _expect(value, expected, what: str, hint: str):
    """Type check for the ``add_*`` methods; openpyxl objects get a hint."""
    if not isinstance(value, expected):
        module = type(value).__module__ or ""
        extra = f"; use {hint} instead of the openpyxl class" if module.startswith("openpyxl") else ""
        raise TypeError(f"{what} must be {hint}, not {type(value).__name__}{extra}")


class ConditionalFormattingList:
    """``ws.conditional_formatting``: the sheet's rules, in the order added."""

    def __init__(self):
        self._rules: list = []

    def add(self, range_string: str, rule) -> None:
        """Apply ``rule`` to ``range_string`` (its own ``sqref`` is replaced)."""
        _expect(rule, _CF_RULES, "rule", "an xlspy conditional formatting rule (CellIsRule, ...)")
        self._rules.append(replace(rule, sqref=_normalize_range(range_string)))

    def __iter__(self):
        return iter(self._rules)

    def __len__(self):
        return len(self._rules)

    def __bool__(self):
        return bool(self._rules)


class Dimension:
    def __init__(self, index: str | int):
        self.index = index
        self.width: float | None = None
        self.height: float | None = None
        self.hidden: bool = False
        self.bestFit: bool | None = None
        self.outlineLevel: int = 0
        self.collapsed: bool = False
        self.style: int | None = None

    @property
    def customWidth(self) -> bool:
        return self.width is not None

    @property
    def customHeight(self) -> bool:
        return self.height is not None


class ColumnDimension(Dimension):
    pass


class RowDimension(Dimension):
    pass


class DimensionHolder(dict):
    def __init__(self, factory, transform):
        super().__init__()
        self._factory = factory
        self._transform = transform

    def __getitem__(self, key):
        key = self._transform(key)
        if key not in self:
            super().__setitem__(key, self._factory(key))
        return super().__getitem__(key)

    def __setitem__(self, key, value):
        super().__setitem__(self._transform(key), value)

    def get(self, key, default=None):
        return super().get(self._transform(key), default)

    def __contains__(self, key):
        return super().__contains__(self._transform(key))


def _column_key(key: str | int) -> str:
    if isinstance(key, int) and not isinstance(key, bool):
        if not 1 <= key <= 16384:
            raise ValueError("column index must be 1..16384")
        return column_label(key - 1)
    if not isinstance(key, str):
        raise TypeError("column key must be a letter or 1-based int")
    return _coerce_column(key) and column_label(_coerce_column(key) - 1)


def _row_key(key: int) -> int:
    return _coerce_row(key)


class Cell:
    __slots__ = (
        "_sheet", "_row", "_column", "_value", "_data_type",
        "_style_id", "_comment", "_hyperlink", "_dirty", "_formula",
    )

    def __init__(self, sheet, row: int, column: int, value=None, style_id: int = 0):
        self._sheet = sheet
        self._row = _coerce_row(row)
        self._column = _coerce_column(column)
        self._value = None
        self._formula: Formula | None = None
        self._data_type = "n"
        self._style_id = int(style_id)
        self._comment: Comment | None = None
        self._hyperlink: Hyperlink | None = None
        self._dirty = False
        if value is not None:
            self.value = value
            self._dirty = False

    # -- identity -----------------------------------------------------
    @property
    def row(self) -> int:
        return self._row

    @property
    def column(self) -> int:
        return self._column

    @property
    def column_letter(self) -> str:
        return column_label(self._column - 1)

    @property
    def coordinate(self) -> str:
        return _coordinate(self._row, self._column)

    @property
    def parent(self):
        return self._sheet

    def __repr__(self):
        return f"<Cell {self._sheet.title!r}.{self.coordinate}>"

    # -- value --------------------------------------------------------
    @property
    def value(self):
        if isinstance(self._formula, Formula):
            return "=" + self._formula.text
        if isinstance(self._formula, str):
            return "=" + self._formula
        return self._value

    @value.setter
    def value(self, value):
        if isinstance(value, Cell):
            value = value.value
        if isinstance(value, Formula):
            self._formula = value
            self._value = None
            self._data_type = "f"
        elif isinstance(value, str) and value.startswith("="):
            self._formula = Formula(value[1:])
            self._value = None
            self._data_type = "f"
        else:
            self._formula = None
            if isinstance(value, Hyperlink):
                self._hyperlink = value
                self._value = value.display_text
                self._data_type = "s"
                self._sheet._dirty_hyperlinks.add(self.coordinate)
            else:
                self._hyperlink = None
                self._value = value
                self._data_type = self._type_for(value)
        if isinstance(value, _datetime.time) and self.number_format == "General":
            self.number_format = "h:mm:ss"
        elif isinstance(value, _datetime.timedelta) and self.number_format == "General":
            self.number_format = "[h]:mm:ss"
        elif isinstance(value, _datetime.date) and self.number_format == "General":
            # The writers' defaults; a General date would show its serial.
            self.number_format = (
                DEFAULT_DATETIME_FORMAT if isinstance(value, _datetime.datetime)
                else DEFAULT_DATE_FORMAT
            )
        self._mark_dirty()

    @property
    def formula(self):
        return self._formula

    @property
    def data_type(self) -> str:
        return self._data_type

    @property
    def internal_value(self):
        return self._value

    @property
    def hyperlink(self):
        return self._hyperlink

    @hyperlink.setter
    def hyperlink(self, value):
        if value is not None and not isinstance(value, Hyperlink):
            raise TypeError("hyperlink must be a xlspy Hyperlink or None")
        self._hyperlink = value
        self._sheet._dirty_hyperlinks.add(self.coordinate)
        self._sheet._mark_dirty("hyperlinks")

    @property
    def comment(self):
        return self._comment

    @comment.setter
    def comment(self, value):
        if value is not None and not isinstance(value, Comment):
            raise TypeError("comment must be a xlspy Comment or None")
        self._comment = value
        self._sheet._dirty_comments.add(self.coordinate)
        self._sheet._mark_dirty("comments")

    @staticmethod
    def _type_for(value) -> str:
        if value is None:
            return "n"
        if isinstance(value, bool):
            return "b"
        if isinstance(value, (int, float, Decimal)) and not isinstance(value, bool):
            return "n"
        if isinstance(value, (_datetime.datetime, _datetime.date, _datetime.time, _datetime.timedelta)):
            return "d"
        if isinstance(value, (str, CellRichText)):
            return "s"
        if isinstance(value, bytes):
            return "s"
        raise TypeError(
            f"unsupported cell value {type(value).__name__}; use a number, string, "
            "date/time, bool, Formula, Hyperlink or None"
        )

    # -- style --------------------------------------------------------
    def _style(self) -> Style:
        return self._sheet._workbook._styles.style_at(self._style_id)

    def _set_style(self, style: Style) -> None:
        self._style_id = self._sheet._workbook._styles.add(style)
        self._sheet._workbook._styles_dirty = True
        self._sheet._dirty_styles.add(self.coordinate)
        self._sheet._mark_dirty("styles")

    @property
    def font(self) -> Font | None:
        return self._style().font

    @font.setter
    def font(self, value: Font | None):
        if value is not None and not isinstance(value, Font):
            raise TypeError("font must be a xlspy Font or None")
        self._set_style(replace(self._style(), font=value, bold=None, italic=None,
                                underline=None, font_color=None))

    @property
    def fill(self):
        return self._style().fill

    @fill.setter
    def fill(self, value):
        if value is not None and not isinstance(value, (str, PatternFill, GradientFill)):
            raise TypeError("fill must be a color, PatternFill, GradientFill or None")
        self._set_style(replace(self._style(), fill=value, pattern=None))

    @property
    def border(self):
        return self._style().border

    @border.setter
    def border(self, value):
        if value is not None and not isinstance(value, Border):
            raise TypeError("border must be a xlspy Border or None")
        self._set_style(replace(self._style(), border=value))

    @property
    def alignment(self):
        return self._style().alignment

    @alignment.setter
    def alignment(self, value):
        if value is not None and not isinstance(value, Alignment):
            raise TypeError("alignment must be a xlspy Alignment or None")
        self._set_style(replace(self._style(), alignment=value,
                                horizontal=None, vertical=None, wrap=None))

    @property
    def protection(self):
        return self._style().protection

    @protection.setter
    def protection(self, value):
        if value is not None and not isinstance(value, Protection):
            raise TypeError("protection must be a xlspy Protection or None")
        self._set_style(replace(self._style(), protection=value))

    @property
    def number_format(self) -> str:
        fmt = self._style().number_format
        return "General" if fmt is None else fmt

    @number_format.setter
    def number_format(self, value: str):
        if not isinstance(value, str):
            raise TypeError("number_format must be a string")
        self._set_style(replace(
            self._style(),
            number_format=None if value == "General" else value,
        ))

    @property
    def style(self):
        return self._style()

    @style.setter
    def style(self, value):
        if isinstance(value, str):
            raise NotImplementedError("named styles are not implemented yet")
        if not isinstance(value, Style):
            raise TypeError("style must be a xlspy Style")
        self._set_style(value)

    def _load_value(self, value, data_type: str, formula_text: str | None = None):
        """Install a parsed value without importing/validating arbitrary formulas."""
        self._value = value
        self._formula = formula_text
        self._data_type = data_type
        self._dirty = False

    def _mark_dirty(self):
        self._dirty = True
        self._sheet._dirty_cells.add((self._row, self._column))
        self._sheet._mark_dirty()

    def _clear_dirty(self):
        self._dirty = False

    # -- convenience --------------------------------------------------
    def offset(self, row=0, column=0):
        return self._sheet.cell(self._row + row, self._column + column)


class MergedCell(Cell):
    """A cell inside a merged range other than its top-left anchor."""

    @Cell.value.setter
    def value(self, value):
        raise AttributeError("read-only cell (merged range)")


class Worksheet:
    ORIENTATION_PORTRAIT = "portrait"
    ORIENTATION_LANDSCAPE = "landscape"
    PAPERSIZE_LETTER = 1
    PAPERSIZE_LEGAL = 5
    PAPERSIZE_A3 = 8
    PAPERSIZE_A4 = 9

    def __init__(self, workbook, title: str, *, index: int | None = None):
        self._workbook = workbook
        self._title = title
        self._index = index
        self._cells: dict[int, dict[int, Cell]] = {}
        self._max_row = 0
        self._max_column = 0
        self._merged_cells = MultiCellRange()
        self._column_dimensions = DimensionHolder(ColumnDimension, _column_key)
        self._row_dimensions = DimensionHolder(RowDimension, _row_key)
        self._freeze_panes: str | None = None
        self._auto_filter = AutoFilter()
        self._sheet_state = "visible"
        self._sheet_properties = SheetProperties()
        self._sheet_view = SheetView(showGridLines=True)
        self._dirty = False
        self._dirty_cells: set[tuple[int, int]] = set()
        self._dirty_styles: set[str] = set()
        self._dirty_comments: set[str] = set()
        self._dirty_hyperlinks: set[str] = set()
        self._dirty_meta: set[str] = set()
        self._loaded = False
        self._original_title = title
        self._meta_signature = None
        self._path: str | None = None
        self._row_heights: dict[int, float] = {}
        self._hidden_rows: set[int] = set()
        self._data_validations: list = []
        self._tables: list = []
        self._charts: list = []
        self._images: list = []
        self._conditional_formats = ConditionalFormattingList()
        # For a loaded workbook: how many objects of each list a save already
        # put into the package (see _objects.pending_objects).
        self._saved_object_counts: dict[str, int] = {}
        # The print setup a loaded sheet already has in the package, so saving
        # does not rewrite a sheet whose print settings were not edited.
        self._saved_print_settings = None
        self._print_settings = None
        self.page_setup = PageSetup()
        self.page_margins = PageMargins()
        self.print_options = PrintOptions()
        self.HeaderFooter = HeaderFooter()
        self._print_area: str | None = None
        self._print_title_rows: str | None = None
        self._print_title_cols: str | None = None
        self._sheet_protection = None
        self._comments: dict[str, Comment] = {}
        self._legacy_drawing = None

    def __repr__(self):
        return f"<Worksheet {self._title!r}>"

    @property
    def title(self) -> str:
        return self._title

    @title.setter
    def title(self, value: str):
        if not isinstance(value, str) or not value:
            raise ValueError("sheet title must be a non-empty string")
        if len(value) > 31 or any(char in value for char in "[]:*?/\\"):
            raise ValueError(f"invalid sheet title {value!r}")
        if value != self._title:
            if value in self._workbook.sheetnames:
                raise ValueError(f"worksheet {value!r} already exists")
            self._workbook._rename_sheet(self, value)
            self._title = value
            self._mark_dirty("title")

    @property
    def sheet_state(self) -> str:
        return self._sheet_state

    @sheet_state.setter
    def sheet_state(self, value: str):
        if value not in {"visible", "hidden", "veryHidden"}:
            raise ValueError("sheet_state must be visible, hidden or veryHidden")
        self._sheet_state = value
        self._mark_dirty("sheet_state")

    @property
    def merged_cells(self) -> MultiCellRange:
        return self._merged_cells

    @property
    def column_dimensions(self) -> DimensionHolder:
        return self._column_dimensions

    @property
    def row_dimensions(self) -> DimensionHolder:
        return self._row_dimensions

    @property
    def freeze_panes(self) -> str | None:
        return self._freeze_panes

    @freeze_panes.setter
    def freeze_panes(self, value: str | None):
        if value == "":
            value = None
        if value is not None:
            if not isinstance(value, str):
                raise TypeError("freeze_panes must be an A1 cell or None")
            col, row, c2, r2 = parse_cell_range(value)
            if (col, row) != (c2, r2):
                raise ValueError("freeze_panes must be a single cell")
            value = _coordinate(row + 1, col + 1)
        self._freeze_panes = value
        self._mark_dirty("freeze_panes")

    @property
    def auto_filter(self) -> AutoFilter:
        return self._auto_filter

    @auto_filter.setter
    def auto_filter(self, value):
        if value is None:
            self._auto_filter = AutoFilter()
        elif isinstance(value, AutoFilter):
            self._auto_filter = value
        elif isinstance(value, str):
            self._auto_filter = AutoFilter(_normalize_range(value))
        else:
            raise TypeError("auto_filter must be an AutoFilter, A1 range or None")
        self._mark_dirty("auto_filter")

    @property
    def sheet_properties(self) -> SheetProperties:
        return self._sheet_properties

    @property
    def sheet_view(self) -> SheetView:
        return self._sheet_view

    @property
    def max_row(self) -> int:
        return self._max_row

    @property
    def max_column(self) -> int:
        return self._max_column

    @property
    def min_row(self) -> int:
        return 1

    @property
    def min_column(self) -> int:
        return 1

    @property
    def dimensions(self) -> str:
        if self._max_row == 0 or self._max_column == 0:
            return "A1:A1"
        return f"A1:{_coordinate(self._max_row, self._max_column)}"

    def calculate_dimension(self, force: bool = False) -> str:
        return self.dimensions

    @property
    def rows(self):
        return self.iter_rows()

    @property
    def columns(self):
        return self.iter_cols()

    @property
    def values(self):
        return self.iter_rows(values_only=True)

    @property
    def data_validations(self):
        return self._data_validations

    @property
    def tables(self):
        return self._tables

    @property
    def charts(self):
        return self._charts

    @property
    def images(self):
        return self._images

    @property
    def conditional_formatting(self) -> ConditionalFormattingList:
        return self._conditional_formats

    # -- printing (openpyxl names) --------------------------------------
    @property
    def oddHeader(self):
        return self.HeaderFooter.oddHeader

    @property
    def oddFooter(self):
        return self.HeaderFooter.oddFooter

    @property
    def print_area(self) -> str | None:
        return self._print_area

    @print_area.setter
    def print_area(self, value):
        """An A1 range, several separated by commas, or a list of ranges."""
        if value is not None and not isinstance(value, str):
            value = ",".join(value)
        self._print_area = value

    @property
    def print_title_rows(self) -> str | None:
        """Rows repeated at the top of every printed page, such as ``"1:2"``."""
        return self._print_title_rows

    @print_title_rows.setter
    def print_title_rows(self, value: str | None):
        self._print_title_rows = value

    @property
    def print_title_cols(self) -> str | None:
        """Columns repeated at the left of every printed page, such as ``"A:B"``."""
        return self._print_title_cols

    @print_title_cols.setter
    def print_title_cols(self, value: str | None):
        self._print_title_cols = value

    # -- inserting, deleting and moving (openpyxl semantics, see _shift) --
    def insert_rows(self, idx: int, amount: int = 1) -> None:
        """Insert ``amount`` empty rows before row ``idx``."""
        _shift.shift_rows(self, idx, amount, delete=False)

    def delete_rows(self, idx: int, amount: int = 1) -> None:
        """Delete ``amount`` rows starting at row ``idx``; the rows below move up."""
        _shift.shift_rows(self, idx, amount, delete=True)

    def insert_cols(self, idx: int, amount: int = 1) -> None:
        """Insert ``amount`` empty columns before column ``idx`` (1-based)."""
        _shift.shift_cols(self, idx, amount, delete=False)

    def delete_cols(self, idx: int, amount: int = 1) -> None:
        """Delete ``amount`` columns starting at column ``idx``; the rest move left."""
        _shift.shift_cols(self, idx, amount, delete=True)

    def move_range(self, cell_range, rows: int = 0, cols: int = 0, translate: bool = False) -> None:
        """Move the cells of ``cell_range``; cells at the destination are replaced.
        ``translate=True`` shifts the relative references of moved formulas."""
        _shift.move_range(self, cell_range, rows, cols, translate)

    # -- objects (openpyxl names, xlspy types) ------------------------
    def add_chart(self, chart: Chart, anchor: str | None = None) -> None:
        _expect(chart, Chart, "chart", "an xlspy Chart")
        if anchor is not None:
            chart = replace(chart, anchor=anchor)
        self._charts.append(chart)

    def add_image(self, img: Image, anchor: str | None = None) -> None:
        _expect(img, Image, "image", "an xlspy Image")
        if anchor is not None:
            # The bytes already read: a file object source is used up.
            img = replace(img, path=img.data, anchor=anchor)
        self._images.append(img)

    def add_table(self, table: ExcelTable) -> None:
        _expect(table, ExcelTable, "table", "an xlspy ExcelTable")
        self._tables.append(table)

    def add_data_validation(self, data_validation: DataValidation) -> None:
        _expect(data_validation, DataValidation, "data_validation", "an xlspy DataValidation")
        self._data_validations.append(data_validation)

    # -- cells --------------------------------------------------------
    def cell(self, row: int, column: int | str, value=None) -> Cell:
        row = _coerce_row(row)
        column = _coerce_column(column)
        if self._is_merged_subcell(row, column):
            cell = MergedCell(self, row, column, style_id=0)
            if value is not None:
                raise AttributeError("read-only cell (merged range)")
            return cell
        cell = self._cells.get(row, {}).get(column)
        if cell is None:
            cell = Cell(self, row, column, style_id=0)
            self._cells.setdefault(row, {})[column] = cell
            self._max_row = max(self._max_row, row)
            self._max_column = max(self._max_column, column)
        if value is not None:
            cell.value = value
        return cell

    def _set_cell(self, row: int, column: int, value=None, style_id: int = 0,
                  *, mark: bool = False) -> Cell:
        cell = Cell(self, row, column, style_id=style_id)
        if value is not None:
            cell.value = value
            cell._dirty = False
        self._cells.setdefault(row, {})[column] = cell
        self._max_row = max(self._max_row, row)
        self._max_column = max(self._max_column, column)
        if mark:
            cell._mark_dirty()
        return cell

    def _set_cell_raw(self, row: int, column: int, value=None, style_id: int = 0,
                      data_type: str = "n", formula_text: str | None = None) -> Cell:
        cell = Cell(self, row, column, style_id=style_id)
        cell._load_value(value, data_type, formula_text)
        self._cells.setdefault(row, {})[column] = cell
        if value is not None or formula_text is not None or style_id or data_type != "n":
            self._max_row = max(self._max_row, row)
            self._max_column = max(self._max_column, column)
        return cell

    def _existing_cell(self, row: int, column: int) -> Cell | None:
        return self._cells.get(row, {}).get(column)

    def __getitem__(self, key):
        if isinstance(key, str):
            if ":" in key:
                c1, r1, c2, r2 = parse_cell_range(key)
                return tuple(
                    tuple(self.cell(row + 1, col + 1) for col in range(c1, c2 + 1))
                    for row in range(r1, r2 + 1)
                )
            if _CELL_RE.match(key):
                col, row, _, _ = parse_cell_range(key)
                return self.cell(row + 1, col + 1)
        if isinstance(key, int) and not isinstance(key, bool):
            row = _coerce_row(key)
            return tuple(self.cell(row, col) for col in range(1, self._max_column + 1))
        raise TypeError("worksheet indices must be A1 references or row numbers")

    def __setitem__(self, key, value):
        if isinstance(key, str) and ":" in key:
            c1, r1, c2, r2 = parse_cell_range(key)
            if isinstance(value, (tuple, list)):
                values = list(value)
                if len(values) == r2 - r1 + 1 and all(
                    isinstance(row, (tuple, list)) for row in values
                ):
                    for row_offset, row_values in enumerate(values):
                        for col_offset, item in enumerate(row_values):
                            self.cell(r1 + 1 + row_offset, c1 + 1 + col_offset, item)
                    return
            for row in range(r1, r2 + 1):
                for col in range(c1, c2 + 1):
                    self.cell(row + 1, col + 1, value)
            return
        if isinstance(key, str) and _CELL_RE.match(key):
            col, row, _, _ = parse_cell_range(key)
            self.cell(row + 1, col + 1, value)
            return
        raise TypeError("worksheet indices must be A1 references")

    def append(self, iterable: Iterable[Any]):
        if isinstance(iterable, (str, bytes, bytearray)):
            raise TypeError("a row must be an iterable of cells, not text")
        values = list(iterable)
        row = self._max_row + 1
        for column, value in enumerate(values, 1):
            if value is not None:
                self.cell(row, column, value)
            else:
                # An explicit empty cell still expands the used range, matching
                # openpyxl when a row is appended with a trailing None.
                self._max_column = max(self._max_column, column)
        if values:
            self._max_row = max(self._max_row, row)
        return row

    def iter_rows(self, min_row=None, max_row=None, min_col=None, max_col=None,
                  values_only: bool = False):
        min_row = 1 if min_row is None else _coerce_row(min_row)
        max_row = self._max_row if max_row is None else _coerce_row(max_row)
        min_col = 1 if min_col is None else _coerce_column(min_col)
        max_col = self._max_column if max_col is None else _coerce_column(max_col)
        for row in range(min_row, max_row + 1):
            if values_only:
                yield tuple(
                    (self._cells.get(row, {}).get(col).value
                     if self._cells.get(row, {}).get(col) is not None else None)
                    for col in range(min_col, max_col + 1)
                )
            else:
                yield tuple(self.cell(row, col) for col in range(min_col, max_col + 1))

    def iter_cols(self, min_col=None, max_col=None, min_row=None, max_row=None,
                  values_only: bool = False):
        min_col = 1 if min_col is None else _coerce_column(min_col)
        max_col = self._max_column if max_col is None else _coerce_column(max_col)
        min_row = 1 if min_row is None else _coerce_row(min_row)
        max_row = self._max_row if max_row is None else _coerce_row(max_row)
        for col in range(min_col, max_col + 1):
            if values_only:
                yield tuple(
                    (self._cells.get(row, {}).get(col).value
                     if self._cells.get(row, {}).get(col) is not None else None)
                    for row in range(min_row, max_row + 1)
                )
            else:
                yield tuple(self.cell(row, col) for row in range(min_row, max_row + 1))

    # -- merges -------------------------------------------------------
    def merge_cells(self, range_string: str | None = None, start_row: int | None = None,
                    start_column: int | None = None, end_row: int | None = None,
                    end_column: int | None = None):
        if range_string is None:
            if None in (start_row, start_column, end_row, end_column):
                raise TypeError("merge_cells needs a range or all four boundaries")
            range_string = (
                f"{_coordinate(start_row, _coerce_column(start_column))}:"
                f"{_coordinate(end_row, _coerce_column(end_column))}"
            )
        reference = _normalize_range(range_string)
        for existing in self._merged_cells.ranges:
            if self._ranges_overlap(existing, reference):
                raise ValueError(f"merge range {reference!r} overlaps {existing!r}")
        self._merged_cells.add(reference)
        self._mark_dirty("merges")

    def unmerge_cells(self, range_string: str | None = None, start_row: int | None = None,
                      start_column: int | None = None, end_row: int | None = None,
                      end_column: int | None = None):
        if range_string is None:
            if None in (start_row, start_column, end_row, end_column):
                raise TypeError("unmerge_cells needs a range or all four boundaries")
            range_string = (
                f"{_coordinate(start_row, _coerce_column(start_column))}:"
                f"{_coordinate(end_row, _coerce_column(end_column))}"
            )
        self._merged_cells.remove(range_string)
        self._mark_dirty("merges")

    @staticmethod
    def _ranges_overlap(a: str, b: str) -> bool:
        c1, r1, c2, r2 = parse_cell_range(a)
        d1, s1, d2, s2 = parse_cell_range(b)
        return not (c2 < d1 or d2 < c1 or r2 < s1 or s2 < r1)

    def _is_merged_subcell(self, row: int, column: int) -> bool:
        if not self._merged_cells:
            return False
        coordinate = _coordinate(row, column)
        for cell_range in self._merged_cells.ranges:
            if coordinate in cell_range:
                c1, r1, _, _ = parse_cell_range(cell_range)
                return (row, column) != (r1 + 1, c1 + 1)
        return False

    # -- dirty tracking ----------------------------------------------
    def _mark_dirty(self, *parts: str):
        self._dirty = True
        for part in parts:
            self._dirty_meta.add(part)

    def _clear_dirty(self):
        self._dirty = False
        self._dirty_cells.clear()
        self._dirty_styles.clear()
        self._dirty_comments.clear()
        self._dirty_hyperlinks.clear()
        self._dirty_meta.clear()
        for row in self._cells.values():
            for cell in row.values():
                cell._clear_dirty()


class WriteOnlyWorksheet(Worksheet):
    """Streaming-friendly sheet that stores rows rather than Cell objects."""

    def __init__(self, workbook, title: str, *, index: int | None = None):
        super().__init__(workbook, title, index=index)
        self._rows: list[list[Any]] = []

    def append(self, iterable: Iterable[Any]):
        if isinstance(iterable, (str, bytes, bytearray)):
            raise TypeError("a row must be an iterable of cells, not text")
        self._rows.append(list(iterable))
        self._max_row = len(self._rows)
        self._max_column = max((len(row) for row in self._rows), default=0)
        self._mark_dirty("rows")
        return self._max_row

    def cell(self, row, column, value=None):
        raise ValueError("cannot access a cell in a write-only worksheet; use append()")

    def iter_rows(self, *args, **kwargs):
        for row in self._rows:
            if kwargs.get("values_only"):
                yield tuple(row)
            else:
                raise ValueError("write-only worksheets only support values_only rows")


class Workbook:
    def __init__(self, write_only: bool = False, iso_dates: bool = False):
        self.write_only = bool(write_only)
        self.iso_dates = bool(iso_dates)
        self._sheets: list[Worksheet] = []
        self._active_index = 0
        self._styles = XlsxStyleTable.empty()
        self._styles_dirty = False
        self._properties = DocumentProperties()
        self._defined_names: list = []
        self._loaded = False
        self._read_only = False
        self._updater = None
        self._source_path: str | None = None
        self._sheet_visibility: dict[str, str] = {}
        self._loaded_worksheet_names: set[str] = set()
        self._mark_loaded = False
        if not self.write_only:
            self.create_sheet("Sheet")

    def __repr__(self):
        return f"<Workbook {self.sheetnames!r}>"

    @property
    def sheetnames(self) -> list[str]:
        return [sheet.title for sheet in self._sheets]

    @property
    def worksheets(self) -> list[Worksheet]:
        return list(self._sheets)

    @property
    def active(self) -> Worksheet:
        if not self._sheets:
            self.create_sheet()
        return self._sheets[self._active_index]

    @active.setter
    def active(self, value):
        if isinstance(value, int) and not isinstance(value, bool):
            self._active_index = value
        elif value in self._sheets:
            self._active_index = self._sheets.index(value)
        else:
            raise ValueError("active must be a sheet index or Worksheet")

    def __getitem__(self, key) -> Worksheet:
        if isinstance(key, int) and not isinstance(key, bool):
            return self._sheets[key]
        for sheet in self._sheets:
            if sheet.title == key:
                return sheet
        raise KeyError(f"Worksheet {key!r} does not exist")

    def get_sheet_by_name(self, name: str) -> Worksheet:
        return self[name]

    def index(self, worksheet: Worksheet) -> int:
        return self._sheets.index(worksheet)

    def create_sheet(self, title: str | None = None, index: int | None = None) -> Worksheet:
        if len(self._sheets) >= 1 and title is None:
            title = f"Sheet{len(self._sheets) + 1}"
        elif title is None:
            title = "Sheet"
        if title in self.sheetnames:
            raise ValueError(f"worksheet {title!r} already exists")
        cls = WriteOnlyWorksheet if self.write_only else Worksheet
        sheet = cls(self, title, index=index)
        if index is None:
            self._sheets.append(sheet)
        else:
            self._sheets.insert(index, sheet)
            for position, item in enumerate(self._sheets):
                item._index = position
        sheet._loaded = self._loaded
        return sheet

    def remove(self, worksheet: Worksheet):
        if worksheet not in self._sheets:
            raise ValueError("worksheet is not part of this workbook")
        self._sheets.remove(worksheet)
        if self._active_index >= len(self._sheets):
            self._active_index = max(0, len(self._sheets) - 1)
        if self._loaded:
            self._mark_loaded = True

    def copy_worksheet(self, from_worksheet: Worksheet) -> Worksheet:
        if self.write_only:
            raise TypeError("copy_worksheet is not available in write-only mode")
        title = from_worksheet.title
        base = title
        counter = 1
        while title in self.sheetnames:
            title = f"{base}{counter}"
            counter += 1
        target = self.create_sheet(title)
        for row_number, row in from_worksheet._cells.items():
            for column, cell in row.items():
                new_cell = target.cell(row_number, column, cell.value)
                new_cell._style_id = cell._style_id
                new_cell._comment = cell.comment
        target._merged_cells = MultiCellRange(from_worksheet._merged_cells.ranges)
        target._column_dimensions.update(from_worksheet._column_dimensions)
        target._row_dimensions.update(from_worksheet._row_dimensions)
        target._freeze_panes = from_worksheet.freeze_panes
        target._auto_filter = AutoFilter(from_worksheet.auto_filter.ref)
        target._sheet_state = from_worksheet.sheet_state
        target._sheet_properties = replace(from_worksheet.sheet_properties)
        target._sheet_view = replace(from_worksheet.sheet_view)
        target._mark_dirty("copy")
        return target

    @property
    def properties(self) -> DocumentProperties:
        return self._properties

    @properties.setter
    def properties(self, value: DocumentProperties):
        if not isinstance(value, DocumentProperties):
            raise TypeError("properties must be DocumentProperties")
        self._properties = value

    @property
    def defined_names(self):
        return self._defined_names

    def _rename_sheet(self, sheet: Worksheet, title: str):
        if self._loaded:
            self._mark_loaded = True

    def save(self, filename):
        if self._read_only:
            raise ValueError("a read-only workbook cannot be saved")
        from ._writer import save_workbook
        save_workbook(self, filename)

    def close(self):
        if self._updater is not None:
            self._updater.close()
            self._updater = None
        for sheet in self._sheets:
            reader = getattr(sheet, "_reader", None)
            if reader is not None:
                reader.close()
                sheet._reader = None
