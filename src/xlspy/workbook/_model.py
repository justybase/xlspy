"""Openpyxl-like workbook model.

This module deliberately contains no XLSX parsing or serialisation code.  The
fast path (``XlsxWriter``/``XlsbWriter``/``ExcelReader``) does not import it.
"""

from __future__ import annotations

import copy

from collections.abc import Iterable, Mapping
from dataclasses import dataclass, field, replace
import datetime as _datetime
from decimal import Decimal
import re
from typing import Any

from ..charts import Chart
from ..comments import Comment
from ..data_validation import DataValidation
from ..formats import DEFAULT_DATE_FORMAT, DEFAULT_DATETIME_FORMAT
from ..formulas import Formula
from ..hyperlinks import Hyperlink
from ..images import Image
from .._metadata import DefinedName
from ..refs import check_sheet_name, column_label, parse_cell_range, quote_sheet_name
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
    NamedStyle,
    NamedStyleList,
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


def _normalize_sqref(reference: str) -> str:
    """A range or space-separated ranges, each normalized."""
    if not isinstance(reference, str):
        raise TypeError("a range must be an A1 string")
    parts = reference.split()
    if not parts:
        raise ValueError("a range must not be empty")
    return " ".join(_normalize_range(part) for part in parts)


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

    @property
    def coord(self) -> str:
        return str(self)

    @property
    def bounds(self):
        return (self.min_col, self.min_row, self.max_col, self.max_row)

    @property
    def size(self):
        return {
            "columns": self.max_col - self.min_col + 1,
            "rows": self.max_row - self.min_row + 1,
        }

    def shift(self, col_shift: int = 0, row_shift: int = 0):
        min_col = self.min_col + col_shift
        min_row = self.min_row + row_shift
        max_col = self.max_col + col_shift
        max_row = self.max_row + row_shift
        if min_col < 1 or min_row < 1 or max_col > 16384 or max_row > 1048576:
            raise ValueError(
                f"Invalid shift value: col_shift={col_shift}, row_shift={row_shift}"
            )
        return CellRange(f"{column_label(min_col - 1)}{min_row}:{column_label(max_col - 1)}{max_row}")

    def expand(self, right: int = 0, down: int = 0, left: int = 0, up: int = 0):
        min_col = self.min_col - left
        min_row = self.min_row - up
        max_col = self.max_col + right
        max_row = self.max_row + down
        if min_col < 1 or min_row < 1 or max_col > 16384 or max_row > 1048576:
            raise ValueError("expanded range is outside the worksheet")
        return CellRange(f"{column_label(min_col - 1)}{min_row}:{column_label(max_col - 1)}{max_row}")

    def shrink(self, right: int = 0, bottom: int = 0, left: int = 0, top: int = 0):
        return self.expand(right=-right, down=-bottom, left=-left, up=-top)

    def union(self, other):
        other = other if isinstance(other, CellRange) else CellRange(str(other))
        return CellRange(
            f"{column_label(min(self.min_col, other.min_col) - 1)}"
            f"{min(self.min_row, other.min_row)}:"
            f"{column_label(max(self.max_col, other.max_col) - 1)}"
            f"{max(self.max_row, other.max_row)}"
        )

    def intersection(self, other):
        other = other if isinstance(other, CellRange) else CellRange(str(other))
        min_col = max(self.min_col, other.min_col)
        min_row = max(self.min_row, other.min_row)
        max_col = min(self.max_col, other.max_col)
        max_row = min(self.max_row, other.max_row)
        if min_col > max_col or min_row > max_row:
            raise ValueError(f"Range {self} doesn't intersect {other}")
        return CellRange(f"{column_label(min_col - 1)}{min_row}:{column_label(max_col - 1)}{max_row}")

    def isdisjoint(self, other) -> bool:
        other = other if isinstance(other, CellRange) else CellRange(str(other))
        return (
            self.max_col < other.min_col
            or other.max_col < self.min_col
            or self.max_row < other.min_row
            or other.max_row < self.min_row
        )

    def issubset(self, other) -> bool:
        other = other if isinstance(other, CellRange) else CellRange(str(other))
        return (
            other.min_col <= self.min_col <= self.max_col <= other.max_col
            and other.min_row <= self.min_row <= self.max_row <= other.max_row
        )

    def issuperset(self, other) -> bool:
        other = other if isinstance(other, CellRange) else CellRange(str(other))
        return other.issubset(self)


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
class WorkbookSecurity:
    workbookPassword: str | None = None
    workbookPasswordCharacterSet: str | None = None
    revisionsPassword: str | None = None
    revisionsPasswordCharacterSet: str | None = None
    lockStructure: bool | None = None
    lockWindows: bool | None = None
    lockRevision: bool | None = None
    revisionsAlgorithmName: str | None = None
    revisionsHashValue: str | None = None
    revisionsSaltValue: str | None = None
    revisionsSpinCount: int | None = None
    workbookAlgorithmName: str | None = None
    workbookHashValue: str | None = None
    workbookSaltValue: str | None = None
    workbookSpinCount: int | None = None


@dataclass
class CalcProperties:
    calcId: int = 124519
    calcMode: str | None = None
    fullCalcOnLoad: bool = True
    refMode: str | None = None
    iterate: bool | None = None
    iterateCount: int | None = None
    iterateDelta: float | None = None
    fullPrecision: bool | None = None
    calcCompleted: bool | None = None
    calcOnSave: bool | None = None
    concurrentCalc: bool | None = None
    concurrentManualCount: int | None = None
    forceFullCalc: bool | None = None


@dataclass
class BookView:
    visibility: str = "visible"
    minimized: bool = False
    showHorizontalScroll: bool = True
    showVerticalScroll: bool = True
    showSheetTabs: bool = True
    xWindow: int | None = None
    yWindow: int | None = None
    windowWidth: int | None = None
    windowHeight: int | None = None
    tabRatio: int = 600
    firstSheet: int = 0
    activeTab: int = 0
    autoFilterDateGrouping: bool = True


class CustomPropertyList(list):
    """A list of custom document properties, keyed by ``name`` for convenience."""

    @property
    def names(self):
        return [item.name for item in self]


class DefinedNameList:
    """A list/dict hybrid of :class:`DefinedName` entries."""

    def __init__(self, workbook=None, definitions=None):
        self._workbook = workbook
        self._items: list[DefinedName] = []
        for definition in definitions or []:
            self.add(definition, stage=False)

    def __iter__(self):
        return iter(self._items)

    def __len__(self):
        return len(self._items)

    def __bool__(self):
        return bool(self._items)

    def __repr__(self):
        return f"DefinedNameList({self._items!r})"

    def __getitem__(self, key):
        if isinstance(key, int) and not isinstance(key, bool):
            return self._items[key]
        if isinstance(key, str):
            folded = key.casefold()
            for definition in self._items:
                if definition.name.casefold() == folded and definition.scope is None:
                    return definition
            for definition in self._items:
                if definition.name.casefold() == folded:
                    return definition
            raise KeyError(key)
        raise TypeError("defined name keys must be a name or an index")

    def __contains__(self, key):
        try:
            self[key]
        except (KeyError, TypeError):
            return False
        return True

    def __setitem__(self, key, value):
        self.add(value)

    @property
    def names(self):
        return [definition.name for definition in self._items]

    def add(self, definition, *, stage: bool = True):
        if not isinstance(definition, DefinedName):
            raise TypeError("defined names must be DefinedName instances")
        self._items = [
            item for item in self._items
            if not (item.name.casefold() == definition.name.casefold()
                    and item.scope == definition.scope)
        ]
        self._items.append(definition)
        if stage and self._workbook is not None:
            if self._workbook._updater is not None:
                self._workbook._updater.set_defined_name(definition)
            else:
                self._workbook._defined_names_dirty = True

    append = add

    def remove(self, name, scope=None):
        if not isinstance(name, str):
            raise TypeError("defined name must be a string")
        folded = name.casefold()
        target = None
        for definition in self._items:
            if definition.name.casefold() == folded and (
                scope is None or definition.scope == scope
            ):
                target = definition
                break
        if target is None:
            raise KeyError(name)
        self._items.remove(target)
        if self._workbook is not None:
            if self._workbook._updater is not None:
                self._workbook._updater.remove_defined_name(target.name, target.scope)
            else:
                self._workbook._defined_names_dirty = True
        return target


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
    lastPrinted: _datetime.datetime | None = None


@dataclass
class Outline:
    applyStyles: bool | None = None
    summaryBelow: bool | None = None
    summaryRight: bool | None = None
    showOutlineSymbols: bool | None = None


@dataclass
class SheetProperties:
    tabColor: Color | str | None = None
    codeName: str | None = None
    pageSetUpPr: PageSetupProperties = field(default_factory=PageSetupProperties)
    outlinePr: Outline = field(default_factory=Outline)
    enableFormatConditionsCalculation: bool | None = None
    filterMode: bool | None = None
    published: bool | None = None
    syncHorizontal: bool | None = None
    syncRef: str | None = None
    syncVertical: bool | None = None
    transitionEvaluation: bool | None = None
    transitionEntry: bool | None = None


@dataclass
class SheetFormatProperties:
    baseColWidth: int = 8
    defaultColWidth: float | None = None
    defaultRowHeight: float = 15
    customHeight: bool | None = None
    zeroHeight: bool | None = None
    thickTop: bool | None = None
    thickBottom: bool | None = None
    outlineLevelRow: int | None = None
    outlineLevelCol: int | None = None


@dataclass
class Pane:
    xSplit: float | int | None = None
    ySplit: float | int | None = None
    topLeftCell: str | None = None
    activePane: str = "topLeft"
    state: str = "split"


@dataclass
class Selection:
    pane: str | None = None
    activeCell: str = "A1"
    activeCellId: int | None = None
    sqref: str = "A1"


@dataclass
class SheetView:
    windowProtection: bool | None = None
    showFormulas: bool | None = None
    showGridLines: bool | None = None
    showRowColHeaders: bool | None = None
    showZeros: bool | None = None
    rightToLeft: bool | None = None
    tabSelected: bool | None = None
    showRuler: bool | None = None
    showOutlineSymbols: bool | None = None
    defaultGridColor: bool | None = None
    showWhiteSpace: bool | None = None
    view: str | None = None
    topLeftCell: str | None = None
    colorId: int | None = None
    zoomScale: int | None = None
    zoomScaleNormal: int | None = None
    zoomScaleSheetLayoutView: int | None = None
    zoomScalePageLayoutView: int | None = None
    zoomToFit: bool | None = None
    workbookViewId: int = 0
    selection: list[Selection] = field(default_factory=list)
    pane: Pane | None = None


@dataclass
class FilterColumn:
    colId: int = 0
    hiddenButton: bool = False
    showButton: bool = True
    vals: list[str] = field(default_factory=list)
    blank: bool = False


@dataclass
class SortCondition:
    ref: str | None = None
    descending: bool | None = None
    sortBy: str | None = None
    customList: str | None = None
    dxfId: int | None = None
    iconSet: str | None = None
    iconId: int | None = None


@dataclass
class SortState:
    ref: str | None = None
    columnSort: bool | None = None
    caseSensitive: bool | None = None
    sortMethod: str | None = None
    sortCondition: list[SortCondition] = field(default_factory=list)


@dataclass
class AutoFilter:
    ref: str | None = None
    filterColumn: list[FilterColumn] = field(default_factory=list)
    sortState: SortState | None = None

    def add_filter_column(self, col_id, vals, blank=False):
        self.filterColumn.append(
            FilterColumn(colId=col_id, vals=list(vals or []), blank=bool(blank))
        )

    def add_sort_condition(self, ref, descending=False):
        if self.sortState is None:
            self.sortState = SortState(ref=self.ref)
        self.sortState.sortCondition.append(
            SortCondition(ref=ref, descending=bool(descending))
        )


@dataclass
class Break:
    id: int = 0
    min: int = 0
    max: int = 16383
    man: bool = True
    pt: int | None = None


class RowBreak:
    def __init__(self, brk=()):
        self.brk = list(brk)

    @property
    def count(self):
        return len(self.brk)

    @property
    def manualBreakCount(self):
        return len(self.brk)

    def append(self, brk=None):
        if not isinstance(brk, Break):
            brk = Break(id=self.count + 1)
        self.brk.append(brk)

    def __len__(self):
        return len(self.brk)

    def __bool__(self):
        return bool(self.brk)


class ColBreak(RowBreak):
    pass


@dataclass
class SheetProtection:
    sheet: bool = False
    objects: bool = False
    scenarios: bool = False
    formatCells: bool = True
    formatRows: bool = True
    formatColumns: bool = True
    insertColumns: bool = True
    insertRows: bool = True
    insertHyperlinks: bool = True
    deleteColumns: bool = True
    deleteRows: bool = True
    selectLockedCells: bool = False
    selectUnlockedCells: bool = False
    sort: bool = True
    autoFilter: bool = True
    pivotTables: bool = True
    password: str | None = None
    algorithmName: str | None = None
    saltValue: str | None = None
    spinCount: int | None = None
    hashValue: str | None = None

    def __setattr__(self, name, value):
        # As in openpyxl, a password set directly is stored as its legacy hash.
        if name == "password" and value is not None:
            from .._metadata import password_hash

            value = password_hash(value)
        object.__setattr__(self, name, value)

    def set_password(self, value: str | None = "", already_hashed: bool = False) -> None:
        if value is not None and not already_hashed:
            from .._metadata import password_hash

            value = password_hash(value)
        object.__setattr__(self, "password", value)

    def enable(self) -> None:
        self.sheet = True

    def disable(self) -> None:
        self.sheet = False

    def __bool__(self) -> bool:
        return self.sheet


@dataclass
class CustomProperty:
    name: str
    value: object


@dataclass
class DocumentSecurity:
    lockStructure: bool = False
    lockWindows: bool = False
    lockRevision: bool = False


_CF_RULES = (CellIsRule, FormulaRule, ColorScaleRule, DataBarRule, IconSetRule, *EXTRA_CF_RULES)


def _expect(value, expected, what: str, hint: str):
    """Type check for the ``add_*`` methods; openpyxl objects get a hint."""
    if not isinstance(value, expected):
        module = type(value).__module__ or ""
        extra = f"; use {hint} instead of the openpyxl class" if module.startswith("openpyxl") else ""
        raise TypeError(f"{what} must be {hint}, not {type(value).__name__}{extra}")


@dataclass
class ConditionalFormatting:
    sqref: str
    rules: list = field(default_factory=list)
    pivot: bool | None = None


class ConditionalFormattingList:
    """``ws.conditional_formatting``: rules grouped by range, as openpyxl."""

    def __init__(self):
        self._items: list[ConditionalFormatting] = []

    def add(self, range_string: str, rule):
        """Apply ``rule`` to ``range_string`` (its own ``sqref`` is replaced).
        Returns the rule as stored."""
        _expect(rule, _CF_RULES, "rule", "an xlspy conditional formatting rule (CellIsRule, ...)")
        normalized = _normalize_sqref(range_string)
        stored = replace(rule, sqref=normalized)
        for item in self._items:
            if item.sqref == normalized:
                item.rules.append(stored)
                return stored
        self._items.append(ConditionalFormatting(normalized, [stored]))
        return stored

    def __iter__(self):
        return iter(self._items)

    def __len__(self):
        return len(self._items)

    def __bool__(self):
        return bool(self._items)

    def __getitem__(self, key):
        if isinstance(key, int) and not isinstance(key, bool):
            return self._items[key]
        normalized = _normalize_sqref(key)
        for item in self._items:
            if item.sqref == normalized:
                return item.rules
        raise KeyError(key)

    def all_rules(self):
        return [rule for item in self._items for rule in item.rules]

    def remove(self, rule_or_range) -> None:
        if isinstance(rule_or_range, str):
            normalized = _normalize_sqref(rule_or_range)
            self._items = [item for item in self._items if item.sqref != normalized]
            return
        for item in list(self._items):
            item.rules = [rule for rule in item.rules if rule is not rule_or_range]
            if not item.rules:
                self._items.remove(item)


class TableColumn:
    def __init__(
        self,
        id=None,
        name=None,
        totalsRowFunction=None,
        totalsRowLabel=None,
        calculatedColumnFormula=None,
        totalsRowFormula=None,
        xmlColumnPr=None,
    ):
        self.id = id
        self.name = name
        self.totalsRowFunction = totalsRowFunction
        self.totalsRowLabel = totalsRowLabel
        self.calculatedColumnFormula = calculatedColumnFormula
        self.totalsRowFormula = totalsRowFormula
        self.xmlColumnPr = xmlColumnPr


class TableStyleInfo:
    def __init__(
        self,
        name=None,
        showFirstColumn=None,
        showLastColumn=None,
        showRowStripes=None,
        showColumnStripes=None,
    ):
        self.name = name
        self.showFirstColumn = showFirstColumn
        self.showLastColumn = showLastColumn
        self.showRowStripes = showRowStripes
        self.showColumnStripes = showColumnStripes


class Table:
    def __init__(
        self,
        id=1,
        displayName=None,
        ref=None,
        name=None,
        comment=None,
        tableType=None,
        headerRowCount=1,
        insertRow=None,
        insertRowShift=None,
        totalsRowCount=None,
        totalsRowShown=None,
        published=None,
        autoFilter=None,
        sortState=None,
        tableColumns=(),
        tableStyleInfo=None,
    ):
        self.id = id
        self.displayName = displayName
        self.ref = ref
        self.name = name or displayName
        self.comment = comment
        self.tableType = tableType
        self.headerRowCount = headerRowCount
        self.insertRow = insertRow
        self.insertRowShift = insertRowShift
        self.totalsRowCount = totalsRowCount
        self.totalsRowShown = totalsRowShown
        self.published = published
        self.autoFilter = autoFilter
        self.sortState = sortState
        self.tableColumns = list(tableColumns)
        self.tableStyleInfo = tableStyleInfo

    @property
    def display_name(self):
        return self.displayName

    @property
    def header_row(self):
        return bool(self.headerRowCount)

    @property
    def totals_row(self):
        return bool(self.totalsRowCount)

    @property
    def columns(self):
        return [column.name for column in self.tableColumns]

    @property
    def style_name(self):
        return None if self.tableStyleInfo is None else self.tableStyleInfo.name


class TableList(list):
    """A list of tables that also accepts a table name as a lookup key."""

    def __getitem__(self, key):
        if isinstance(key, int) and not isinstance(key, bool):
            return super().__getitem__(key)
        for table in self:
            if key in {getattr(table, "displayName", None), getattr(table, "name", None)}:
                return table
        raise KeyError(key)

    def __contains__(self, key):
        if isinstance(key, str):
            return any(
                key in {getattr(table, "displayName", None), getattr(table, "name", None)}
                for table in self
            )
        return super().__contains__(key)

    def add(self, table) -> None:
        self.append(table)

    def remove(self, table_or_name) -> None:
        if isinstance(table_or_name, str):
            for table in list(self):
                if table_or_name in {
                    getattr(table, "displayName", None), getattr(table, "name", None)
                }:
                    super().remove(table)
                    return
            raise KeyError(table_or_name)
        super().remove(table_or_name)


class DataValidationList(list):
    def add(self, rule) -> None:
        self.append(rule)

    def remove(self, rule) -> None:
        super().remove(rule)


class Dimension:
    def __init__(self, index: str | int):
        self.index = index
        self.min: int | None = None
        self.max: int | None = None
        self.width: float | None = None
        self.height: float | None = None
        self.hidden: bool = False
        self.bestFit: bool | None = None
        self.autoSize: bool | None = None
        self.outlineLevel: int = 0
        self.collapsed: bool = False
        self.style: int | None = None

    @property
    def customWidth(self) -> bool:
        return self.width is not None

    @property
    def customHeight(self) -> bool:
        return self.height is not None

    @property
    def outline_level(self) -> int:
        return self.outlineLevel

    @outline_level.setter
    def outline_level(self, value: int):
        self.outlineLevel = value


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

    def group(self, start, end=None, outline_level: int = 1, hidden: bool = False):
        """Group consecutive rows or columns, as openpyxl does."""
        if end is None:
            end = start
        if self._factory is ColumnDimension:
            first = _coerce_column(start)
            last = _coerce_column(end)
            for index in range(first, last + 1):
                dimension = self[column_label(index - 1)]
                dimension.outlineLevel = outline_level
                dimension.hidden = hidden
                dimension.min = first
                dimension.max = last
        else:
            for index in range(int(start), int(end) + 1):
                dimension = self[index]
                dimension.outlineLevel = outline_level
                dimension.hidden = hidden


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


_new_cell = object.__new__
_PLAIN_TYPES = {int: "n", float: "n", bool: "b", str: "s"}


class Cell:
    __slots__ = (
        "_sheet", "_row", "_column", "_value", "_data_type",
        "_style_id", "_comment", "_hyperlink", "_dirty", "_formula",
        "_named_style",
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
        self._named_style: str | None = None
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
        plain = _PLAIN_TYPES.get(type(value))
        if plain is not None and not (plain == "s" and value.startswith("=")):
            # Exact int/float/bool/str: no formula, hyperlink or date handling.
            self._formula = None
            self._hyperlink = None
            self._value = value
            self._data_type = plain
            self._dirty = True
            sheet = self._sheet
            sheet._dirty_cells.add((self._row, self._column))
            sheet._dirty = True
            return
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
    def is_date(self) -> bool:
        return isinstance(
            self._value,
            (_datetime.datetime, _datetime.date, _datetime.time, _datetime.timedelta),
        )

    @property
    def encoding(self) -> str:
        return "utf-8"

    @property
    def base_date(self):
        return self._sheet._workbook.epoch

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
        workbook = self._sheet._workbook
        styles = workbook._styles
        cache = getattr(styles, "_number_format_ids", None)
        if cache is None:
            cache = styles._number_format_ids = {}
        key = (self._style_id, value)
        style_id = cache.get(key)
        if style_id is None:
            style_id = cache[key] = styles.add(replace(
                self._style(),
                number_format=None if value == "General" else value,
            ))
        self._style_id = style_id
        workbook._styles_dirty = True
        self._sheet._dirty_styles.add(self.coordinate)
        self._sheet._mark_dirty("styles")

    @property
    def style(self):
        return self._style()

    @style.setter
    def style(self, value):
        if isinstance(value, str):
            named = None
            for entry in self._sheet._workbook.named_styles:
                if (entry if isinstance(entry, str) else entry.name) == value:
                    named = entry
                    break
            if named is None:
                raise NotImplementedError(
                    "named styles are not implemented yet; "
                    f"{value!r} is not defined, add it with add_named_style()"
                )
            if isinstance(named, NamedStyle):
                current = self._style()
                self._set_style(Style(
                    font=named.font if named.font is not None else current.font,
                    fill=named.fill if named.fill is not None else current.fill,
                    border=named.border if named.border is not None else current.border,
                    alignment=named.alignment if named.alignment is not None else current.alignment,
                    protection=named.protection if named.protection is not None else current.protection,
                    number_format=(
                        named.number_format if named.number_format is not None
                        else current.number_format
                    ),
                ))
            self._named_style = value
            return
        if not isinstance(value, Style):
            raise TypeError("style must be a xlspy Style or a named style")
        self._named_style = None
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


# Slot descriptors, in the order the C loaders fill them (see c_ext/_c_core.c).
CELL_SLOTS = tuple(Cell.__dict__[name] for name in (
    "_sheet", "_row", "_column", "_value", "_formula", "_data_type",
    "_style_id", "_comment", "_hyperlink", "_named_style", "_dirty",
))


class MergedCell(Cell):
    """A cell inside a merged range other than its top-left anchor."""

    @Cell.value.setter
    def value(self, value):
        raise AttributeError("read-only cell (merged range)")


class WriteOnlyCell:
    """A value with optional style, used with ``WriteOnlyWorksheet.append``."""

    def __init__(self, worksheet, value=None):
        self.parent = worksheet
        self.value = value
        self.font = None
        self.fill = None
        self.border = None
        self.alignment = None
        self.protection = None
        self.number_format = "General"
        self.comment = None
        self.hyperlink = None

    def __repr__(self):
        return f"<WriteOnlyCell {self.value!r}>"


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
        self._data_validations = DataValidationList()
        self._tables = TableList()
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
        self.sheet_format = SheetFormatProperties()
        self.protection = SheetProtection()
        self.row_breaks = RowBreak()
        self.col_breaks = ColBreak()
        self._print_area: str | None = None
        self._print_title_rows: str | None = None
        self._print_title_cols: str | None = None
        self._comments: dict[str, Comment] = {}
        self._legacy_drawing = None

    def __repr__(self):
        return f"<Worksheet {self._title!r}>"

    @property
    def title(self) -> str:
        return self._title

    @title.setter
    def title(self, value: str):
        if value != self._title:
            self._workbook._check_title(value, self)
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
        if value is None:
            self._sheet_view.pane = None
            self._sheet_view.selection = []
        else:
            col, row, _, _ = parse_cell_range(value)
            if col and row:
                active = "bottomRight"
            elif col:
                active = "topRight"
            else:
                active = "bottomLeft"
            self._sheet_view.pane = Pane(
                xSplit=col if col else None,
                ySplit=row if row else None,
                topLeftCell=value,
                activePane=active,
                state="frozen",
            )
            self._sheet_view.selection = [
                Selection(pane=active, activeCell=value, sqref=value)
            ]
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
    def active_cell(self) -> str | None:
        if not self._sheet_view.selection:
            return None
        return self._sheet_view.selection[0].activeCell

    @property
    def selected_cell(self) -> str | None:
        if not self._sheet_view.selection:
            return None
        return self._sheet_view.selection[0].sqref

    @property
    def show_gridlines(self) -> bool | None:
        return self._sheet_view.showGridLines

    @show_gridlines.setter
    def show_gridlines(self, value: bool | None):
        self._sheet_view.showGridLines = value
        self._mark_dirty("sheet_view")

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
    def evenHeader(self):
        return self.HeaderFooter.evenHeader

    @property
    def evenFooter(self):
        return self.HeaderFooter.evenFooter

    @property
    def firstHeader(self):
        return self.HeaderFooter.firstHeader

    @property
    def firstFooter(self):
        return self.HeaderFooter.firstFooter

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

    @property
    def print_titles(self) -> str:
        """The print titles as openpyxl formats them, e.g. ``'Sheet'!$1:$2``."""
        title = quote_sheet_name(self.title)
        values = []
        if self._print_title_rows:
            values.append("$" + self._print_title_rows.replace(":", ":$"))
        if self._print_title_cols:
            values.append("$" + self._print_title_cols.replace(":", ":$"))
        return ",".join(f"{title}!{value}" for value in values)

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

    def add_table(self, table) -> None:
        _expect(table, (ExcelTable, Table), "table", "an xlspy ExcelTable or Table")
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
        # Loader hot path: bounds are checked inline and the cell is filled
        # directly, which is what Cell.__init__ + _load_value would do.
        if not (1 <= row <= 1048576 and 1 <= column <= 16384):
            _coerce_row(row)
            _coerce_column(column)
        cell = _new_cell(Cell)
        cell._sheet = self
        cell._row = row
        cell._column = column
        cell._value = value
        cell._formula = formula_text
        cell._data_type = data_type
        cell._style_id = style_id
        cell._comment = None
        cell._hyperlink = None
        cell._named_style = None
        cell._dirty = False
        row_cells = self._cells.get(row)
        if row_cells is None:
            row_cells = self._cells[row] = {}
        row_cells[column] = cell
        if value is not None or formula_text is not None or style_id or data_type != "n":
            if row > self._max_row:
                self._max_row = row
            if column > self._max_column:
                self._max_column = column
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
        if isinstance(iterable, Mapping):
            row = self._max_row + 1
            for key, value in iterable.items():
                if isinstance(key, str):
                    column = _coerce_column(key)
                elif isinstance(key, int) and not isinstance(key, bool):
                    if not 1 <= key <= 16384:
                        raise ValueError("column index must be 1..16384")
                    column = key
                else:
                    raise TypeError("dict keys must be column letters or 1-based ints")
                if isinstance(value, WriteOnlyCell):
                    value = value.value
                self.cell(row, column, value)
            if iterable:
                self._max_row = max(self._max_row, row)
                self._max_column = max(
                    self._max_column,
                    max(_coerce_column(key) for key in iterable),
                )
            return row
        if isinstance(iterable, (str, bytes, bytearray)):
            raise TypeError("a row must be an iterable of cells, not text")
        values = [
            value.value if isinstance(value, WriteOnlyCell) else value
            for value in iterable
        ]
        row = self._max_row + 1
        if (values and row <= 1048576 and len(values) <= 16384
                and not self._merged_cells and row not in self._cells):
            self._append_new_row(row, values)
            return row
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

    def _append_new_row(self, row: int, values: list) -> None:
        """append() into a row with no cells yet and no merged ranges.

        Plain int/float/bool/str values are stored as ``cell(row, column,
        value)`` would store them, without its checks; any other value goes
        through ``cell()``.
        """
        row_cells = self._cells[row] = {}
        dirty = self._dirty_cells
        plain_types = _PLAIN_TYPES
        date_styles = None
        dated = False
        for column, value in enumerate(values, 1):
            if value is None:
                continue
            kind = type(value)
            plain = plain_types.get(kind)
            style_id = 0
            if plain is None:
                # A date or datetime gets the writers' default format, as the
                # value setter gives it; reuse the style id it already made.
                if kind is _datetime.datetime or kind is _datetime.date:
                    if date_styles is None:
                        date_styles = self._default_date_styles()
                    style_id = date_styles.get(kind)
                if not style_id:
                    self.cell(row, column, value)
                    continue
                plain = "d"
                dated = True
                self._dirty_styles.add(_coordinate(row, column))
            elif plain == "s" and value.startswith("="):
                self.cell(row, column, value)
                continue
            cell = _new_cell(Cell)
            cell._sheet = self
            cell._row = row
            cell._column = column
            cell._value = value
            cell._formula = None
            cell._data_type = plain
            cell._style_id = style_id
            cell._comment = None
            cell._hyperlink = None
            cell._named_style = None
            cell._dirty = True
            row_cells[column] = cell
            dirty.add((row, column))
        if dated:
            self._workbook._styles_dirty = True
            self._dirty_meta.add("styles")
        if row_cells:
            self._dirty = True
        else:
            del self._cells[row]
        # As for cell() and the None branch of append(): every position counts.
        if row > self._max_row:
            self._max_row = row
        if len(values) > self._max_column:
            self._max_column = len(values)

    def _default_date_styles(self) -> dict:
        """Style ids a new date/datetime cell gets from the value setter, when known.

        Taken from the number-format memo of ``Cell.number_format``; a type
        missing here goes through ``cell()`` once, which fills the memo.
        """
        styles = self._workbook._styles
        if styles.style_at(0).number_format not in (None, "General"):
            return {}
        cache = getattr(styles, "_number_format_ids", None) or {}
        found = {}
        for kind, fmt in ((_datetime.datetime, DEFAULT_DATETIME_FORMAT),
                          (_datetime.date, DEFAULT_DATE_FORMAT)):
            style_id = cache.get((0, fmt))
            if style_id:
                found[kind] = style_id
        return found

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
        if isinstance(iterable, Mapping):
            entries = []
            for key, value in iterable.items():
                if isinstance(key, str):
                    column = _coerce_column(key)
                elif isinstance(key, int) and not isinstance(key, bool):
                    if not 1 <= key <= 16384:
                        raise ValueError("column index must be 1..16384")
                    column = key
                else:
                    raise TypeError("dict keys must be column letters or 1-based ints")
                entries.append((column, value.value if isinstance(value, WriteOnlyCell) else value))
            row = [None] * (max((column for column, _ in entries), default=0))
            for column, value in entries:
                row[column - 1] = value
            self._rows.append(row)
        else:
            if isinstance(iterable, (str, bytes, bytearray)):
                raise TypeError("a row must be an iterable of cells, not text")
            self._rows.append([
                value.value if isinstance(value, WriteOnlyCell) else value
                for value in iterable
            ])
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


class Chartsheet:
    """A chartsheet: a sheet that shows one chart.

    A new chartsheet's chart reads its ranges from ``source`` (a worksheet or
    its title; by default the workbook's first worksheet).
    """

    def __init__(self, workbook, title: str, index: int | None = None):
        self._workbook = workbook
        self._title = title  # the caller checks the name
        self.sheet_state = "visible"
        self._index = index
        self._charts: list = []
        self._part: str | None = None  # the package part, once in a file
        self._source = None

    def __repr__(self):
        return f"<Chartsheet {self.title!r}>"

    @property
    def title(self) -> str:
        return self._title

    @title.setter
    def title(self, value: str):
        if value == self._title:
            return
        if self._part is not None:
            raise NotImplementedError("renaming a loaded chartsheet is not supported")
        self._workbook._check_title(value, self)
        self._title = value

    @property
    def charts(self):
        return self._charts

    def add_chart(self, chart, source: "Worksheet | str | None" = None) -> None:
        if self._part is None:
            _expect(chart, Chart, "chart", "an xlspy Chart")
            if self._charts:
                raise ValueError("a chartsheet holds one chart")
            if source is not None:
                self._source = source
        self._charts.append(chart)


class Workbook:
    def __init__(self, write_only: bool = False, iso_dates: bool = False):
        self.write_only = bool(write_only)
        self.iso_dates = bool(iso_dates)
        self._sheets: list[Worksheet] = []
        self._chartsheets: list[Chartsheet] = []
        self._active_index = 0
        self._styles = XlsxStyleTable.empty()
        self._styles_dirty = False
        self._properties = DocumentProperties()
        self._defined_names = DefinedNameList(self)
        self._defined_names_dirty = False
        self.template = False
        self.epoch = _datetime.datetime(1899, 12, 30)
        self.vba_archive = None
        self.security = WorkbookSecurity()
        self.calculation = CalcProperties()
        self.views = [BookView()]
        self.custom_doc_props = CustomPropertyList()
        self.named_styles = NamedStyleList(["Normal"])
        self._format = "xlsx"
        self._loaded = False
        self._loaded_active_index = 0
        self._loaded_workbook_meta = None
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

    def _check_title(self, title: str, sheet=None) -> None:
        """Raise unless ``title`` can name ``sheet`` (default: a new sheet).

        Excel's naming rules apply, and no other worksheet or chartsheet may
        have the name in any letter case.
        """
        check_sheet_name(title)
        folded = title.casefold()
        for other in (*self._sheets, *self._chartsheets):
            if other is not sheet and other.title.casefold() == folded:
                raise ValueError(f"a sheet named {other.title!r} already exists")

    @property
    def read_only(self) -> bool:
        return self._read_only

    @read_only.setter
    def read_only(self, value: bool):
        self._read_only = bool(value)

    @property
    def worksheets(self) -> list[Worksheet]:
        return list(self._sheets)

    @property
    def chartsheets(self) -> list[Chartsheet]:
        return list(self._chartsheets)

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
        self._check_title(title)
        cls =WriteOnlyWorksheet if self.write_only else Worksheet
        sheet = cls(self, title, index=index)
        if index is None:
            self._sheets.append(sheet)
        else:
            self._sheets.insert(index, sheet)
            for position, item in enumerate(self._sheets):
                item._index = position
        sheet._loaded = self._loaded
        return sheet

    def create_chartsheet(self, title: str | None = None, index: int | None = None) -> Chartsheet:
        """Add a chartsheet after the other sheets (XLSX only), as openpyxl does."""
        if self.write_only:
            raise TypeError("create_chartsheet is not available in write-only mode")
        if title is None:
            taken = {sheet.title.casefold() for sheet in (*self._sheets, *self._chartsheets)}
            number = 1
            while f"chart{number}" in taken:
                number += 1
            title = f"Chart{number}"
        self._check_title(title)
        count = len(self._sheets) + len(self._chartsheets)
        if index is not None and index < count:
            raise NotImplementedError("new chartsheets are added after the other sheets")
        sheet = Chartsheet(self, title, index=count)
        self._chartsheets.append(sheet)
        return sheet

    def move_sheet(self, sheet: Worksheet | str, offset: int = 0) -> None:
        """Move a sheet ``offset`` positions (negative: towards the front), as openpyxl does."""
        if isinstance(sheet, str):
            sheet = self[sheet]
        if sheet not in self._sheets:
            raise ValueError("worksheet is not part of this workbook")
        active = self._sheets[self._active_index] if self._sheets else None
        index = self._sheets.index(sheet)
        self._sheets.remove(sheet)
        self._sheets.insert(max(0, min(index + offset, len(self._sheets))), sheet)
        for position, item in enumerate(self._sheets):
            item._index = position
        if active is not None:
            self._active_index = self._sheets.index(active)

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
        taken = {sheet.title.casefold() for sheet in (*self._sheets, *self._chartsheets)}
        while title.casefold() in taken:
            title = f"{base}{counter}"
            counter += 1
        target = self.create_sheet(title)
        for row_number, row in from_worksheet._cells.items():
            for column, cell in row.items():
                new_cell = target.cell(row_number, column, cell.value)
                new_cell._style_id = cell._style_id
                # Own copies, as openpyxl makes: editing one sheet leaves the other alone.
                new_cell._comment = copy.copy(cell.comment)
                new_cell._hyperlink = copy.copy(cell.hyperlink)
                new_cell._named_style = cell._named_style
        target._merged_cells = MultiCellRange(from_worksheet._merged_cells.ranges)
        for name in ("_column_dimensions", "_row_dimensions"):
            getattr(target, name).update(copy.deepcopy(dict(getattr(from_worksheet, name))))
        target._row_heights = dict(from_worksheet._row_heights)
        target._hidden_rows = set(from_worksheet._hidden_rows)
        target._freeze_panes = from_worksheet.freeze_panes
        target._auto_filter = copy.deepcopy(from_worksheet.auto_filter)
        target._sheet_state = from_worksheet.sheet_state
        target._sheet_properties = copy.deepcopy(from_worksheet.sheet_properties)
        target._sheet_view = copy.deepcopy(from_worksheet.sheet_view)
        for name in ("page_setup", "page_margins", "print_options", "HeaderFooter",
                     "sheet_format", "row_breaks", "col_breaks"):
            setattr(target, name, copy.deepcopy(getattr(from_worksheet, name)))
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

    def add_named_style(self, style) -> None:
        if not isinstance(style, NamedStyle):
            raise TypeError("style must be a NamedStyle")
        self._styles.add_named_style(style)
        self.named_styles.add(style)
        self._styles_dirty = True

    @property
    def defined_names(self) -> DefinedNameList:
        if not isinstance(self._defined_names, DefinedNameList):
            self._defined_names = DefinedNameList(self, self._defined_names)
        return self._defined_names

    def _rename_sheet(self, sheet: Worksheet, title: str):
        """Point formulas, validations, conditional formats, internal links and
        defined names at the new title, as Excel does (loaded charts: on save)."""
        from ..formula.sheets import rename_sheet_references

        old = sheet.title

        def renamed(text):
            if not isinstance(text, str) or "!" not in text:
                return text
            return rename_sheet_references(text, old, title)

        for other in self._sheets:
            for row in other._cells.values():
                for cell in row.values():
                    link = cell._hyperlink
                    if link is not None and link.is_internal:
                        target = "#" + renamed(link.location)
                        if target != link.target:
                            cell.hyperlink = replace(link, target=target)
                    formula = cell._formula
                    if formula is None:
                        continue
                    text = formula.text if isinstance(formula, Formula) else formula
                    updated = renamed(text)
                    if updated != text:
                        cell.value = (
                            replace(formula, text=updated) if isinstance(formula, Formula)
                            else "=" + updated
                        )
            for index, rule in enumerate(other._data_validations):
                values = {name: renamed(getattr(rule, name)) for name in ("formula1", "formula2")}
                if any(values[name] != getattr(rule, name) for name in values):
                    other._data_validations[index] = replace(rule, **values)
            for item in other.conditional_formatting:
                for index, rule in enumerate(item.rules):
                    formulas = getattr(rule, "formula", None)
                    if formulas:
                        updated = tuple(renamed(text) for text in formulas)
                        if updated != tuple(formulas):
                            item.rules[index] = replace(rule, formula=updated)
        names = self.defined_names
        names._items = [
            DefinedName(
                item.name, renamed(item.formula), title if item.scope == old else item.scope,
                item.localSheetId, item.hidden, item.comment,
            ) if renamed(item.formula) != item.formula or item.scope == old else item
            for item in names._items
        ]
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
