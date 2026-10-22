"""Data validation rules for XLSX worksheets (dropdown lists and input checks)."""

from __future__ import annotations

import datetime
from dataclasses import dataclass
from xml.sax.saxutils import escape

from .excel_dates import datetime_to_serial
from .refs import parse_cell_range

# User-facing type names and their OOXML spelling.
VALIDATION_TYPES = {
    "list": "list",
    "whole": "whole",
    "decimal": "decimal",
    "date": "date",
    "time": "time",
    "text_length": "textLength",
    "custom": "custom",
}
OPERATORS = {
    "between": "between",
    "not_between": "notBetween",
    "equal": "equal",
    "not_equal": "notEqual",
    "greater_than": "greaterThan",
    "less_than": "lessThan",
    "greater_than_or_equal": "greaterThanOrEqual",
    "less_than_or_equal": "lessThanOrEqual",
}
ERROR_STYLES = frozenset({"stop", "warning", "information"})

_TWO_OPERAND_OPERATORS = frozenset({"between", "not_between"})
_FORMULA_TYPES = frozenset({"list", "custom"})
_MAX_LIST_LENGTH = 255
_MAX_PROMPT_TITLE = 32
_MAX_PROMPT = 255
_MAX_ERROR_TITLE = 32
_MAX_ERROR = 255


def _check_text(value, kind: str, limit: int) -> None:
    if value is None:
        return
    if not isinstance(value, str):
        raise TypeError(f"{kind} must be a string or None")
    if len(value) > limit:
        raise ValueError(f"{kind} must be at most {limit} characters")


def _number_text(number: float) -> str:
    return str(int(number)) if float(number).is_integer() else repr(float(number))


def _has_top_level_comma(text: str) -> bool:
    """True when ``text`` has a comma outside function parentheses and string literals.

    ``Yes,No`` is a literal list missing its quotes; ``OFFSET($B$1,0,0,5,1)``
    is a dynamic list, whose commas separate function arguments.
    """
    depth = 0
    quote = None  # '"' inside a string literal, "'" inside a quoted sheet name
    for char in text:
        if quote is not None:
            if char == quote:
                quote = None
        elif char in "\"'":
            quote = char
        elif char == "(":
            depth += 1
        elif char == ")":
            depth -= 1
        elif char == "," and depth <= 0:
            return True
    return False


def _operand_text(value, kind: str) -> str:
    """Render a validation operand as formula text."""
    if isinstance(value, bool):
        raise TypeError(f"{kind} must be a number, date, time, string or formula")
    if isinstance(value, datetime.time):
        seconds = value.hour * 3600 + value.minute * 60 + value.second
        return _number_text(seconds / 86400)
    if isinstance(value, (datetime.datetime, datetime.date)):
        return _number_text(datetime_to_serial(value))
    if isinstance(value, (int, float)):
        return _number_text(value)
    if isinstance(value, str):
        if not value:
            raise ValueError(f"{kind} must not be empty")
        return value.removeprefix("=")
    raise TypeError(f"{kind} must be a number, date, time, string or formula")


def _normalize_sqref(sqref) -> str:
    if isinstance(sqref, str):
        parts = sqref.split()
    elif isinstance(sqref, (list, tuple)) and all(isinstance(part, str) for part in sqref):
        parts = list(sqref)
    else:
        raise TypeError("sqref must be an A1 range or a list of A1 ranges")
    if not parts:
        raise ValueError("sqref must name at least one cell range")
    for part in parts:
        parse_cell_range(part)
    return " ".join(parts)


@dataclass(frozen=True, slots=True)
class DataValidation:
    """A validation rule applied to one or more cell ranges.

    ``type`` is one of ``list``, ``whole``, ``decimal``, ``date``, ``time``,
    ``text_length`` or ``custom``. For ``list`` the ``formula1`` is either a
    literal list such as ``'"New,Active,Closed"'`` or a range such as
    ``"Lists!$A$2:$A$9"``. For ``custom`` it is a formula that must evaluate to
    true for valid input. ``operator`` applies to the numeric, date, time and
    text-length types only, and defaults to ``between``.

    ``show_dropdown=False`` hides the arrow of a list rule. ``allow_blank``
    accepts empty cells. Prompts and error messages are shown when their
    ``show_*`` flag is true.
    """

    type: str
    sqref: str | tuple[str, ...] | list[str]
    formula1: object = None
    formula2: object = None
    operator: str | None = None
    allow_blank: bool = True
    show_dropdown: bool = True
    show_input_message: bool = False
    prompt_title: str | None = None
    prompt: str | None = None
    show_error_message: bool = True
    error_title: str | None = None
    error: str | None = None
    error_style: str = "stop"

    def __post_init__(self):
        if self.type not in VALIDATION_TYPES:
            raise ValueError(
                f"invalid validation type {self.type!r}; "
                f"expected one of {sorted(VALIDATION_TYPES)}"
            )
        object.__setattr__(self, "sqref", _normalize_sqref(self.sqref))
        if self.error_style not in ERROR_STYLES:
            raise ValueError(
                f"invalid error_style {self.error_style!r}; "
                f"expected one of {sorted(ERROR_STYLES)}"
            )
        self._check_messages()
        self._check_flags()
        self._check_operands()

    def _check_messages(self) -> None:
        _check_text(self.prompt_title, "prompt_title", _MAX_PROMPT_TITLE)
        _check_text(self.prompt, "prompt", _MAX_PROMPT)
        _check_text(self.error_title, "error_title", _MAX_ERROR_TITLE)
        _check_text(self.error, "error", _MAX_ERROR)
        if self.prompt is not None and not self.show_input_message:
            raise ValueError("prompt text is set but show_input_message is False")
        if self.error is not None and not self.show_error_message:
            raise ValueError("error text is set but show_error_message is False")

    def _check_flags(self) -> None:
        for name in ("allow_blank", "show_dropdown", "show_input_message", "show_error_message"):
            if not isinstance(getattr(self, name), bool):
                raise TypeError(f"{name} must be a bool")
        if not self.show_dropdown and self.type != "list":
            raise ValueError("show_dropdown only applies to list validation")

    def _check_operands(self) -> None:
        kind = self.type
        if kind in _FORMULA_TYPES:
            if self.operator is not None:
                raise ValueError(f"operator does not apply to {kind} validation")
            if self.formula2 is not None:
                raise ValueError(f"formula2 does not apply to {kind} validation")
        else:
            operator = self.operator if self.operator is not None else "between"
            if operator not in OPERATORS:
                raise ValueError(
                    f"invalid operator {operator!r}; expected one of {sorted(OPERATORS)}"
                )
            object.__setattr__(self, "operator", operator)
            if operator in _TWO_OPERAND_OPERATORS:
                if self.formula2 is None:
                    raise ValueError(f"{operator} validation needs formula2")
            elif self.formula2 is not None:
                raise ValueError("formula2 is only used with between or not_between")
        if self.formula1 is None:
            raise ValueError(f"{kind} validation needs formula1")
        text1 = _operand_text(self.formula1, "formula1")
        object.__setattr__(self, "formula1", text1)
        if self.formula2 is not None:
            object.__setattr__(self, "formula2", _operand_text(self.formula2, "formula2"))
        if kind == "list" and not text1.startswith('"') and _has_top_level_comma(text1):
            raise ValueError(
                "a literal list must be wrapped in double quotes; a range or "
                "formula has commas only inside function arguments"
            )
        if kind == "list" and text1.startswith('"'):
            if len(text1) < 2 or not text1.endswith('"'):
                raise ValueError("a literal list must be wrapped in double quotes")
            if len(text1) - 2 > _MAX_LIST_LENGTH:
                raise ValueError(f"a literal list may be at most {_MAX_LIST_LENGTH} characters")

    @property
    def ooxml_type(self) -> str:
        return VALIDATION_TYPES[self.type]


def _attr(value: str) -> str:
    return escape(value, {'"': "&quot;", "\n": "&#10;"})


def normalize_validations(validations) -> list[DataValidation]:
    """The ``data_validations`` option of ``write_sheet`` as a list."""
    if not validations:
        return []
    for rule in validations:
        if not isinstance(rule, DataValidation):
            raise TypeError("data_validations must be DataValidation instances")
    return list(validations)


def data_validation_xml(rule: DataValidation) -> str:
    """One ``dataValidation`` element; attributes that match the OOXML defaults are omitted."""
    attrs = [f'type="{rule.ooxml_type}"']
    if rule.error_style != "stop":
        attrs.append(f'errorStyle="{rule.error_style}"')
    if rule.operator is not None and rule.operator != "between":
        attrs.append(f'operator="{OPERATORS[rule.operator]}"')
    if rule.allow_blank:
        attrs.append('allowBlank="1"')
    if rule.show_input_message:
        attrs.append('showInputMessage="1"')
    if rule.show_error_message:
        attrs.append('showErrorMessage="1"')
    if not rule.show_dropdown:
        # OOXML spells this flag inversely: showDropDown="1" hides the arrow.
        attrs.append('showDropDown="1"')
    if rule.error_title:
        attrs.append(f'errorTitle="{_attr(rule.error_title)}"')
    if rule.error:
        attrs.append(f'error="{_attr(rule.error)}"')
    if rule.prompt_title:
        attrs.append(f'promptTitle="{_attr(rule.prompt_title)}"')
    if rule.prompt:
        attrs.append(f'prompt="{_attr(rule.prompt)}"')
    attrs.append(f'sqref="{rule.sqref}"')
    inner = f"<formula1>{escape(rule.formula1)}</formula1>"
    if rule.formula2 is not None:
        inner += f"<formula2>{escape(rule.formula2)}</formula2>"
    return f"<dataValidation {' '.join(attrs)}>{inner}</dataValidation>"
