"""Worksheet lists: conditional formatting, tables and data validations."""

from __future__ import annotations
from dataclasses import dataclass, field, replace
from ..styles import (
    CellIsRule,
    ColorScaleRule,
    DataBarRule,
    EXTRA_CF_RULES,
    FormulaRule,
    IconSetRule,
)
from ._coords import _normalize_sqref


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
