"""Writer input coercion, matching XlsxWriter's ``strings_to_*`` and
``nan_inf_to_errors`` options.

The defaults are all off: a string is written as a string, and a non-finite
number raises ``ValueError``. Turning an option on changes only the cells it
matches, at write time, before either the native or the Python row encoder
sees the row.
"""

from __future__ import annotations

from collections.abc import Iterable, Iterator
from dataclasses import dataclass
from decimal import Decimal
import math
import re
from typing import Any, cast

from ._errors import CellError
from ._scalars import numpy_value
from .formulas import Formula
from .hyperlinks import Hyperlink

# A decimal number with an optional sign and exponent. Deliberately stricter
# than ``float()``: "nan", "inf", "0x10" and "1_000" stay text.
_COERCION_NUMBER_RE = re.compile(r"^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$")
# What XlsxWriter treats as a URL by default.
_URL_RE = re.compile(r"^(?:https?|ftp)://|^mailto:", re.IGNORECASE)


@dataclass(frozen=True, slots=True)
class Coercion:
    """The four input options, as one immutable value per writer."""

    strings_to_formulas: bool = False
    strings_to_numbers: bool = False
    strings_to_urls: bool = False
    nan_inf_to_errors: bool = False

    @property
    def active(self) -> bool:
        """Whether writing a row needs a pass at all."""
        return self.strings_to_formulas or self.strings_to_numbers or self.strings_to_urls or self.nan_inf_to_errors


def check_options(
    strings_to_formulas: object,
    strings_to_numbers: object,
    strings_to_urls: object,
    nan_inf_to_errors: object,
) -> Coercion:
    """Validate the constructor keywords and build their ``Coercion``."""
    names = (
        ("strings_to_formulas", strings_to_formulas),
        ("strings_to_numbers", strings_to_numbers),
        ("strings_to_urls", strings_to_urls),
        ("nan_inf_to_errors", nan_inf_to_errors),
    )
    for name, value in names:
        if not isinstance(value, bool):
            raise TypeError(f"{name} must be a bool")
    return Coercion(
        strings_to_formulas=cast(bool, strings_to_formulas),
        strings_to_numbers=cast(bool, strings_to_numbers),
        strings_to_urls=cast(bool, strings_to_urls),
        nan_inf_to_errors=cast(bool, nan_inf_to_errors),
    )


def _coerce_scalar(value: Any, coercion: Coercion) -> Any:
    if isinstance(value, CellError):
        return value
    value = numpy_value(value)
    if isinstance(value, (float, Decimal)) and coercion.nan_inf_to_errors:
        number = float(value)
        if not math.isfinite(number):
            return CellError("error:" + ("#NUM!" if math.isnan(number) else "#DIV/0!"))
        return value
    if not isinstance(value, str):
        return value
    if coercion.strings_to_formulas and value.startswith("="):
        try:
            return Formula(value)
        except (TypeError, ValueError):
            return value
    if coercion.strings_to_numbers and _COERCION_NUMBER_RE.match(value):
        try:
            if "." in value or "e" in value.lower():
                return float(value)
            return int(value)
        except ValueError:
            return value
    if coercion.strings_to_urls and _URL_RE.match(value):
        return Hyperlink(value, text=value)
    return value


def coerce_value(value: Any, coercion: Coercion) -> Any:
    """Apply the writer's options to one cell value."""
    if not coercion.active:
        return value
    # A (value, number format) cell keeps its format; only the value changes.
    if isinstance(value, tuple) and len(value) == 2 and isinstance(value[1], str):
        inner = _coerce_scalar(value[0], coercion)
        return value if inner is value[0] else (inner, value[1])
    return _coerce_scalar(value, coercion)


def coerce_rows(rows: Iterable[list[Any]], coercion: Coercion) -> Iterator[list[Any]]:
    """Yield each row with the writer's input options applied."""
    if not coercion.active:
        yield from rows
        return
    for row in rows:
        yield [coerce_value(cell, coercion) for cell in row]
