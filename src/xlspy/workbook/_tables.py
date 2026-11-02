"""Worksheet lists: conditional formatting, tables and data validations."""

from __future__ import annotations
from collections.abc import Iterable
from dataclasses import dataclass, field, replace
from typing import TYPE_CHECKING
from ..styles import (
    CellIsRule,
    ColorScaleRule,
    DataBarRule,
    EXTRA_CF_RULES,
    FormulaRule,
    IconSetRule,
)
from ._coords import _normalize_sqref

if TYPE_CHECKING:
    from ._properties import AutoFilter, SortState


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
    """A column of a ``Table``: its name and totals row settings."""
    def __init__(
        self,
        id: int | None = None,
        name: str | None = None,
        totalsRowFunction: str | None = None,
        totalsRowLabel: str | None = None,
        calculatedColumnFormula: str | None = None,
        totalsRowFormula: str | None = None,
        xmlColumnPr: object = None,
    ) -> None:
        self.id = id
        self.name = name
        self.totalsRowFunction = totalsRowFunction
        self.totalsRowLabel = totalsRowLabel
        self.calculatedColumnFormula = calculatedColumnFormula
        self.totalsRowFormula = totalsRowFormula
        self.xmlColumnPr = xmlColumnPr


class TableStyleInfo:
    """A table's style name and banding options."""
    def __init__(
        self,
        name: str | None = None,
        showFirstColumn: bool | None = None,
        showLastColumn: bool | None = None,
        showRowStripes: bool | None = None,
        showColumnStripes: bool | None = None,
    ) -> None:
        self.name = name
        self.showFirstColumn = showFirstColumn
        self.showLastColumn = showLastColumn
        self.showRowStripes = showRowStripes
        self.showColumnStripes = showColumnStripes


class Table:
    """An openpyxl-style table: a named range with a header row, columns and a style.

    Add it with ``Worksheet.add_table``. ``ExcelTable`` is the equivalent for
    the writers.
    """
    def __init__(
        self,
        id: int = 1,
        displayName: str | None = None,
        ref: str | None = None,
        name: str | None = None,
        comment: str | None = None,
        tableType: str | None = None,
        headerRowCount: int = 1,
        insertRow: bool | None = None,
        insertRowShift: bool | None = None,
        totalsRowCount: int | None = None,
        totalsRowShown: bool | None = None,
        published: bool | None = None,
        autoFilter: AutoFilter | None = None,
        sortState: SortState | None = None,
        tableColumns: Iterable[TableColumn] = (),
        tableStyleInfo: TableStyleInfo | None = None,
    ) -> None:
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
    def display_name(self) -> str | None:
        """The table name shown in Excel; the same as ``displayName``."""
        return self.displayName

    @property
    def header_row(self) -> bool:
        """True when the table has a header row."""
        return bool(self.headerRowCount)

    @property
    def totals_row(self) -> bool:
        """True when the table has a totals row."""
        return bool(self.totalsRowCount)

    @property
    def columns(self) -> list[str | None]:
        """The column names."""
        return [column.name for column in self.tableColumns]

    @property
    def style_name(self) -> str | None:
        """The table style name, or None."""
        return None if self.tableStyleInfo is None else self.tableStyleInfo.name


class TableList(list):
    """The tables of a worksheet: a list that also works like openpyxl's name-keyed dict.

    Iterating gives the tables themselves (openpyxl iterates the names; use
    ``keys()`` for those). ``tables["Sales"]``, ``"Sales" in tables``,
    ``del tables["Sales"]``, ``get``, ``keys``, ``values`` and ``items`` take or
    give table names as openpyxl does; an int index still works like a list.
    """

    @staticmethod
    def _names(table: object) -> set[object]:
        # A Table has displayName and name; the writers' ExcelTable has display_name.
        return {getattr(table, attr, None) for attr in ("displayName", "name", "display_name")}

    @staticmethod
    def _name(table: object) -> str:
        return next(name for attr in ("displayName", "display_name", "name")
                    if isinstance(name := getattr(table, attr, None), str))

    def _index(self, name: str) -> int:
        for index, table in enumerate(self):
            if name in self._names(table):
                return index
        raise KeyError(name)

    def __getitem__(self, key):
        if isinstance(key, str):
            return super().__getitem__(self._index(key))
        return super().__getitem__(key)

    def __delitem__(self, key) -> None:
        super().__delitem__(self._index(key) if isinstance(key, str) else key)

    def __contains__(self, key: object) -> bool:
        if isinstance(key, str):
            return any(key in self._names(table) for table in self)
        return super().__contains__(key)

    def get(self, name: str, default: object = None) -> object:
        """The table called ``name``, or ``default`` when there is none."""
        return self[name] if name in self else default

    def keys(self) -> list[str]:
        """The table names, in order."""
        return [self._name(table) for table in self]

    def values(self) -> list[object]:
        """The tables, in order."""
        return list(self)

    def items(self) -> list[tuple[str, str]]:
        """``(name, ref)`` for each table, as openpyxl's ``TableList.items()``."""
        return [(name, table.ref) for name, table in zip(self.keys(), self)]

    def add(self, table) -> None:
        """Add a table (the same as ``append``)."""
        self.append(table)

    def remove(self, table_or_name) -> None:
        """Remove a table, given as the object or its name; raises KeyError for an unknown name."""
        if isinstance(table_or_name, str):
            del self[table_or_name]
        else:
            super().remove(table_or_name)


class DataValidationList(list):
    def add(self, rule) -> None:
        self.append(rule)

    def remove(self, rule) -> None:
        super().remove(rule)
