"""Conditional-formatting rule models and the rule-type tables."""

from __future__ import annotations
from dataclasses import dataclass
import warnings
from collections.abc import Callable
from typing import Any
from ..refs import parse_cell_range
from ._attrs import Color, _CELL_IS_OPERATORS, _CFVO_TYPES, _FORMAT_OBJECT_TYPES, _ICON_SETS, _coerce_color
from ._style import DifferentialStyle


def _normalize_formula(formula: str | list[str] | tuple[str, ...]) -> tuple[str, ...]:
    if isinstance(formula, str):
        return (formula,)
    if isinstance(formula, (list, tuple)):
        items = tuple(formula)
        for item in items:
            if not isinstance(item, str):
                raise TypeError("formula entries must be strings")
        return items
    raise TypeError("formula must be a string or a list of strings")


def _validate_priority(priority: int | None) -> None:
    if priority is None:
        return
    if not isinstance(priority, int) or isinstance(priority, bool):
        raise TypeError("priority must be a positive int or None")
    if priority <= 0:
        raise ValueError("priority must be a positive int")


def _validate_sqref(sqref: str) -> None:
    """A range or space-separated ranges ("A1:A5 C1:C5"), as Excel writes them."""
    if not isinstance(sqref, str) or not sqref.split():
        raise ValueError(f"invalid sqref {sqref!r}")
    for part in sqref.split():
        parse_cell_range(part)


@dataclass(frozen=True, slots=True)
class Rule:
    """A generic conditional-formatting rule, as openpyxl exposes it."""

    type: str
    dxf: DifferentialStyle | None = None
    priority: int | None = None
    stop_if_true: bool | None = None


@dataclass(frozen=True, slots=True)
class CellIsRule:
    """A ``cellIs`` conditional-formatting rule."""

    sqref: str
    operator: str
    formula: list[str] | tuple[str, ...] | str
    dxf: DifferentialStyle | None = None
    stop_if_true: bool | None = None
    priority: int | None = None

    def __post_init__(self):
        _validate_sqref(self.sqref)
        if self.operator not in _CELL_IS_OPERATORS:
            raise ValueError(f"invalid operator {self.operator!r}")
        object.__setattr__(self, "formula", _normalize_formula(self.formula))
        if self.dxf is not None and not isinstance(self.dxf, DifferentialStyle):
            raise TypeError("dxf must be a DifferentialStyle or None")
        if self.stop_if_true is not None and not isinstance(self.stop_if_true, bool):
            raise TypeError("stop_if_true must be bool or None")
        _validate_priority(self.priority)


@dataclass(frozen=True, slots=True)
class FormulaRule:
    """An ``expression`` conditional-formatting rule."""

    sqref: str
    formula: list[str] | tuple[str, ...] | str
    dxf: DifferentialStyle | None = None
    stop_if_true: bool | None = None
    priority: int | None = None

    def __post_init__(self):
        _validate_sqref(self.sqref)
        object.__setattr__(self, "formula", _normalize_formula(self.formula))
        if self.dxf is not None and not isinstance(self.dxf, DifferentialStyle):
            raise TypeError("dxf must be a DifferentialStyle or None")
        if self.stop_if_true is not None and not isinstance(self.stop_if_true, bool):
            raise TypeError("stop_if_true must be bool or None")
        _validate_priority(self.priority)


@dataclass(frozen=True, slots=True)
class FormatObject:
    """A ``cfvo`` data-bar / color-scale endpoint."""

    type: str
    value: str | float | None = None

    def __post_init__(self):
        if self.type not in _FORMAT_OBJECT_TYPES:
            raise ValueError(f"invalid format object type {self.type!r}")
        if (
            self.value is not None
            and not isinstance(self.value, (str, int, float))
        ):
            raise TypeError("value must be a string, number, or None")


@dataclass(frozen=True, slots=True)
class ColorScaleRule:
    """A 2-color (``midpoint`` None) or 3-color scale rule."""

    sqref: str
    minimum: FormatObject | None
    midpoint: FormatObject | None
    maximum: FormatObject | None
    colors: list[Color | str] | tuple[Color | str, ...]
    priority: int | None = None
    def __post_init__(self):
        _validate_sqref(self.sqref)
        for name in ("minimum", "midpoint", "maximum"):
            value = getattr(self, name)
            if value is not None and not isinstance(value, FormatObject):
                raise TypeError(f"{name} must be a FormatObject or None")
        if not isinstance(self.colors, (list, tuple)):
            raise TypeError("colors must be a list of Color or color strings")
        colors = tuple(_coerce_color(item, "colors") for item in self.colors)
        expected = 2 if self.midpoint is None else 3
        if len(colors) != expected:
            raise ValueError(
                f"expected {expected} colors for a "
                f"{'2' if expected == 2 else '3'}-color scale"
            )
        object.__setattr__(self, "colors", colors)
        _validate_priority(self.priority)


@dataclass(frozen=True, slots=True)
class DataBarRule:
    """A ``dataBar`` conditional-formatting rule."""

    sqref: str
    minimum: FormatObject | None
    maximum: FormatObject | None
    color: Color | str
    show_value: bool | None = None
    min_length: int | None = None
    max_length: int | None = None
    priority: int | None = None

    def __post_init__(self):
        _validate_sqref(self.sqref)
        for name in ("minimum", "maximum"):
            value = getattr(self, name)
            if value is not None and not isinstance(value, FormatObject):
                raise TypeError(f"{name} must be a FormatObject or None")
        object.__setattr__(self, "color", _coerce_color(self.color, "color"))
        if self.color is None:
            raise ValueError("color is required")
        if self.show_value is not None and not isinstance(self.show_value, bool):
            raise TypeError("show_value must be bool or None")
        for name in ("min_length", "max_length"):
            value = getattr(self, name)
            if value is not None:
                if not isinstance(value, int) or isinstance(value, bool):
                    raise TypeError(f"{name} must be an int or None")
                if value < 0:
                    raise ValueError(f"{name} must be >= 0")
        _validate_priority(self.priority)


@dataclass(frozen=True, slots=True, init=False)
class IconSetRule:
    """An ``iconSet`` conditional-formatting rule (openpyxl-like).

    ``cfvo_type`` is the type of the thresholds in ``values``; ``type`` is
    the rule type, ``"iconSet"``, as in openpyxl.
    """

    sqref: str
    icon_style: str
    values: list[float | str] | tuple[float | str, ...]
    cfvo_type: str = "percent"
    show_value: bool | None = None
    percent: bool | None = None
    reverse: bool | None = None
    stop_if_true: bool | None = None
    priority: int | None = None

    def __init__(self, sqref: str, icon_style: str, values: list[float | str] | tuple[float | str, ...],
                 cfvo_type: str = "percent", show_value: bool | None = None,
                 percent: bool | None = None, reverse: bool | None = None,
                 stop_if_true: bool | None = None, priority: int | None = None, *,
                 type: str | None = None) -> None:
        if type is not None and type != "iconSet":
            warnings.warn(
                "IconSetRule(type=...) is deprecated; pass cfvo_type= "
                "(type is the rule type, 'iconSet', as in openpyxl)",
                DeprecationWarning, stacklevel=2,
            )
            cfvo_type = type
        for name, value in (
            ("sqref", sqref), ("icon_style", icon_style), ("values", values),
            ("cfvo_type", cfvo_type), ("show_value", show_value), ("percent", percent),
            ("reverse", reverse), ("stop_if_true", stop_if_true), ("priority", priority),
        ):
            object.__setattr__(self, name, value)
        self.__post_init__()

    def __post_init__(self):
        _validate_sqref(self.sqref)
        if self.icon_style not in _ICON_SETS:
            raise ValueError(f"invalid icon set {self.icon_style!r}")
        if not isinstance(self.values, (list, tuple)):
            raise TypeError("values must be a list of numbers")
        values = tuple(self.values)
        if not values:
            raise ValueError("an icon set rule needs at least one value")
        for value in values:
            if isinstance(value, bool) or not isinstance(value, (int, float, str)):
                raise TypeError("icon set values must be numbers or formulas")
        if self.cfvo_type not in _CFVO_TYPES:
            raise ValueError(f"invalid icon set value type {self.cfvo_type!r}")
        if self.cfvo_type != "formula":
            for value in values:
                if isinstance(value, str):
                    raise ValueError(
                        "string values require cfvo_type='formula'"
                    )
        elif any(not isinstance(value, str) for value in values):
            raise ValueError("cfvo_type='formula' requires string values")
        for attr in ("show_value", "percent", "reverse", "stop_if_true"):
            value = getattr(self, attr)
            if value is not None and not isinstance(value, bool):
                raise TypeError(f"{attr} must be bool or None")
        object.__setattr__(self, "values", values)
        _validate_priority(self.priority)


def _check_dxf_rule(rule) -> None:
    """Checks shared by the rules that apply a ``DifferentialStyle``."""
    _validate_sqref(rule.sqref)
    if rule.dxf is not None and not isinstance(rule.dxf, DifferentialStyle):
        raise TypeError("dxf must be a DifferentialStyle or None")
    if rule.stop_if_true is not None and not isinstance(rule.stop_if_true, bool):
        raise TypeError("stop_if_true must be bool or None")
    _validate_priority(rule.priority)


def _check_flag(value, name: str) -> None:
    if not isinstance(value, bool):
        raise TypeError(f"{name} must be a bool")


@dataclass(frozen=True, slots=True)
class TopBottomRule:
    """Top or bottom ``rank`` values (``percent=True``: ``rank`` percent) (XLSX)."""

    sqref: str
    rank: int = 10
    percent: bool = False
    bottom: bool = False
    dxf: DifferentialStyle | None = None
    stop_if_true: bool | None = None
    priority: int | None = None

    def __post_init__(self):
        _check_dxf_rule(self)
        if isinstance(self.rank, bool) or not isinstance(self.rank, int):
            raise TypeError("rank must be an int")
        if not 1 <= self.rank <= (100 if self.percent else 1000):
            raise ValueError("rank must be 1-1000 (1-100 with percent=True)")
        _check_flag(self.percent, "percent")
        _check_flag(self.bottom, "bottom")


@dataclass(frozen=True, slots=True)
class AboveAverageRule:
    """Values above (or ``below``) the range's average; ``std_dev`` moves the
    threshold by that many standard deviations (XLSX)."""

    sqref: str
    below: bool = False
    equal_average: bool = False
    std_dev: int | None = None
    dxf: DifferentialStyle | None = None
    stop_if_true: bool | None = None
    priority: int | None = None

    def __post_init__(self):
        _check_dxf_rule(self)
        _check_flag(self.below, "below")
        _check_flag(self.equal_average, "equal_average")
        if self.std_dev is not None:
            if isinstance(self.std_dev, bool) or not isinstance(self.std_dev, int):
                raise TypeError("std_dev must be an int or None")
            if not 1 <= self.std_dev <= 3:
                raise ValueError("std_dev must be 1, 2 or 3")


TEXT_OPERATORS = frozenset({"contains", "notContains", "beginsWith", "endsWith"})


@dataclass(frozen=True, slots=True)
class TextRule:
    """Cells whose text ``contains``, ``notContains``, ``beginsWith`` or
    ``endsWith`` ``text`` (case-insensitive, as in Excel) (XLSX)."""

    sqref: str
    operator: str
    text: str
    dxf: DifferentialStyle | None = None
    stop_if_true: bool | None = None
    priority: int | None = None

    def __post_init__(self):
        _check_dxf_rule(self)
        if self.operator not in TEXT_OPERATORS:
            raise ValueError(f"operator must be one of {sorted(TEXT_OPERATORS)}")
        if not isinstance(self.text, str) or not self.text:
            raise ValueError("text must be a non-empty string")
        if len(self.text) > 255:
            raise ValueError("text must be at most 255 characters")


@dataclass(frozen=True, slots=True)
class DuplicateRule:
    """Duplicate values in the range, or ``unique=True`` for values that occur once (XLSX)."""

    sqref: str
    unique: bool = False
    dxf: DifferentialStyle | None = None
    stop_if_true: bool | None = None
    priority: int | None = None

    def __post_init__(self):
        _check_dxf_rule(self)
        _check_flag(self.unique, "unique")


TIME_PERIODS = frozenset({
    "today", "yesterday", "tomorrow", "last7Days", "thisWeek", "lastWeek",
    "nextWeek", "thisMonth", "lastMonth", "nextMonth",
})


@dataclass(frozen=True, slots=True)
class TimePeriodRule:
    """Dates in a period relative to today, such as ``"last7Days"`` (XLSX)."""

    sqref: str
    period: str
    dxf: DifferentialStyle | None = None
    stop_if_true: bool | None = None
    priority: int | None = None

    def __post_init__(self):
        _check_dxf_rule(self)
        if self.period not in TIME_PERIODS:
            raise ValueError(f"period must be one of {sorted(TIME_PERIODS)}")


BLANK_KINDS = frozenset({"blanks", "no_blanks", "errors", "no_errors"})


@dataclass(frozen=True, slots=True)
class BlanksRule:
    """Blank (``"blanks"``/``"no_blanks"``) or error (``"errors"``/``"no_errors"``) cells (XLSX)."""

    sqref: str
    kind: str = "blanks"
    dxf: DifferentialStyle | None = None
    stop_if_true: bool | None = None
    priority: int | None = None

    def __post_init__(self):
        _check_dxf_rule(self)
        if self.kind not in BLANK_KINDS:
            raise ValueError(f"kind must be one of {sorted(BLANK_KINDS)}")


# Rules XlsxWriter writes beyond the original five.
# Formulas Excel itself stores with these rules, for the range's top-left
# cell ({c}); Excel evaluates them relative to each cell of the range.
CF_TEXT_RULES = {
    "contains": ("containsText", "containsText", 'NOT(ISERROR(SEARCH({t},{c})))'),
    "notContains": ("notContainsText", "notContains", 'ISERROR(SEARCH({t},{c}))'),
    "beginsWith": ("beginsWith", "beginsWith", 'LEFT({c},LEN({t}))={t}'),
    "endsWith": ("endsWith", "endsWith", 'RIGHT({c},LEN({t}))={t}'),
}
CF_TIME_PERIOD_FORMULAS = {
    "today": "FLOOR({c},1)=TODAY()",
    "yesterday": "FLOOR({c},1)=TODAY()-1",
    "tomorrow": "FLOOR({c},1)=TODAY()+1",
    "last7Days": "AND(TODAY()-FLOOR({c},1)<=6,FLOOR({c},1)<=TODAY())",
    "thisWeek": "AND(TODAY()-ROUNDDOWN({c},0)<=WEEKDAY(TODAY())-1,ROUNDDOWN({c},0)-TODAY()<=7-WEEKDAY(TODAY()))",
    "lastWeek": "AND(TODAY()-ROUNDDOWN({c},0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN({c},0)<(WEEKDAY(TODAY())+7))",
    "nextWeek": "AND(ROUNDDOWN({c},0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN({c},0)-TODAY()<(15-WEEKDAY(TODAY())))",
    "thisMonth": "AND(MONTH({c})=MONTH(TODAY()),YEAR({c})=YEAR(TODAY()))",
    "lastMonth": "AND(MONTH({c})=MONTH(EDATE(TODAY(),0-1)),YEAR({c})=YEAR(EDATE(TODAY(),0-1)))",
    "nextMonth": "AND(MONTH({c})=MONTH(EDATE(TODAY(),0+1)),YEAR({c})=YEAR(EDATE(TODAY(),0+1)))",
}
CF_BLANK_RULES = {
    "blanks": ("containsBlanks", "LEN(TRIM({c}))=0"),
    "no_blanks": ("notContainsBlanks", "LEN(TRIM({c}))>0"),
    "errors": ("containsErrors", "ISERROR({c})"),
    "no_errors": ("notContainsErrors", "NOT(ISERROR({c}))"),
}
EXTRA_CF_RULES = (
    TopBottomRule, AboveAverageRule, TextRule, DuplicateRule, TimePeriodRule, BlanksRule,
)


# openpyxl exposes the OOXML rule type on every conditional-formatting rule.
# The local rule classes use descriptive field names, so the aliases are added
# here without changing their constructors.
def _text_rule_type(rule: Any) -> str:
    return "containsText" if rule.operator == "contains" else rule.operator


def _duplicate_rule_type(rule: Any) -> str:
    return "uniqueValues" if rule.unique else "duplicateValues"


def _blanks_rule_type(rule: Any) -> str:
    types: dict[str, str] = {
        "blanks": "containsBlanks",
        "no_blanks": "notContainsBlanks",
        "errors": "containsErrors",
        "no_errors": "notContainsErrors",
    }
    return types[rule.kind]


def _fixed_rule_type(name: str) -> Callable[[Any], str]:
    def rule_type(rule: Any) -> str:
        return name
    return rule_type


_RULE_TYPES = (
    (CellIsRule, _fixed_rule_type("cellIs")),
    (FormulaRule, _fixed_rule_type("expression")),
    (ColorScaleRule, _fixed_rule_type("colorScale")),
    (DataBarRule, _fixed_rule_type("dataBar")),
    (IconSetRule, _fixed_rule_type("iconSet")),
    (TopBottomRule, _fixed_rule_type("top10")),
    (AboveAverageRule, _fixed_rule_type("aboveAverage")),
    (TextRule, _text_rule_type),
    (DuplicateRule, _duplicate_rule_type),
    (TimePeriodRule, _fixed_rule_type("timePeriod")),
    (BlanksRule, _blanks_rule_type),
)


for _rule_class, _rule_type in _RULE_TYPES:
    _rule_class_any: Any = _rule_class  # the rule classes share no base for these
    _rule_class_any.rule_type = property(_rule_type)
    _rule_class_any.type = property(_rule_type)


del _rule_class, _rule_type, _rule_class_any
