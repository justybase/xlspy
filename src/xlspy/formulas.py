"""Excel formula cells.

A ``Formula`` is written as a formula, never as text. xlspy does not evaluate
formulas: XLSX stores normalized formula text, and XLSB compiles supported
expressions to BIFF12 tokens. Excel recalculates on open.
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass
from decimal import Decimal

from .refs import parse_cell_range

MAX_FORMULA_LENGTH = 8192

# Functions added after Excel 2007. The file format stores them with an
# ``_xlfn.`` prefix; without it Excel reports #NAME?. Only the prefix is
# added, the formula otherwise stays exactly as written.
_XLFN_FUNCTIONS = frozenset({
    "AGGREGATE", "BETA.DIST", "BETA.INV", "BINOM.DIST", "BINOM.INV",
    "CEILING.MATH", "CEILING.PRECISE", "CHISQ.DIST", "CHISQ.DIST.RT",
    "CHISQ.INV", "CHISQ.INV.RT", "CHISQ.TEST", "CONCAT", "COVARIANCE.P",
    "COVARIANCE.S", "DAYS", "ERF.PRECISE", "ERFC.PRECISE", "EXPON.DIST",
    "F.DIST", "F.DIST.RT", "F.INV", "F.INV.RT", "F.TEST", "FILTER",
    "FLOOR.MATH", "FLOOR.PRECISE", "GAMMA", "GAMMA.DIST", "GAMMA.INV",
    "GAMMALN.PRECISE", "HYPGEOM.DIST", "IFNA", "IFS", "ISOWEEKNUM",
    "LOGNORM.DIST", "LOGNORM.INV", "MAXIFS", "MINIFS", "MODE.MULT",
    "MODE.SNGL", "NEGBINOM.DIST", "NORM.DIST", "NORM.INV", "NORM.S.DIST",
    "NORM.S.INV", "PERCENTILE.EXC", "PERCENTILE.INC", "PERCENTRANK.EXC",
    "PERCENTRANK.INC", "POISSON.DIST", "QUARTILE.EXC", "QUARTILE.INC",
    "RANK.AVG", "RANK.EQ", "SORT", "SORTBY", "STDEV.P", "STDEV.S", "SWITCH",
    "T.DIST", "T.DIST.2T", "T.DIST.RT", "T.INV", "T.INV.2T", "T.TEST",
    "TEXTAFTER", "TEXTBEFORE", "TEXTJOIN", "TEXTSPLIT", "UNIQUE", "VAR.P",
    "VAR.S", "WEIBULL.DIST", "XLOOKUP", "XMATCH", "XOR",
    "SEQUENCE", "RANDARRAY", "TAKE", "DROP", "CHOOSECOLS", "CHOOSEROWS",
    "VSTACK", "HSTACK", "TOCOL", "TOROW", "WRAPCOLS", "WRAPROWS", "EXPAND",
})

# Functions whose arguments need ``_xlpm.`` parameter names or cell-level
# lambda metadata. They are refused rather than written as #NAME? formulas.
_UNSUPPORTED_FUNCTIONS = frozenset({
    "LET", "LAMBDA", "MAP", "REDUCE", "SCAN", "BYROW", "BYCOL", "MAKEARRAY",
})

_IDENT_RE = re.compile(r"[A-Za-z_][A-Za-z0-9_.]*")
_CONTROL_RE = re.compile(r"[\x00-\x1f\x7f]")


def _scan_tokens(text: str):
    """Yield ``(kind, value, start, end)`` for identifiers and structure.

    String literals and quoted sheet names are yielded as ``"literal"`` so the
    prefixing and parenthesis checks never look inside them.
    """
    i = 0
    length = len(text)
    while i < length:
        char = text[i]
        if char in "\"'":
            end = i + 1
            while True:
                if end >= length:
                    kind = "string" if char == '"' else "sheet"
                    raise ValueError(f"unterminated {kind} in formula {text!r}")
                if text[end] == char:
                    if end + 1 < length and text[end + 1] == char:
                        end += 2
                        continue
                    break
                end += 1
            yield "literal", text[i:end + 1], i, end + 1
            i = end + 1
            continue
        match = _IDENT_RE.match(text, i) if (char.isalpha() or char == "_") else None
        if match is not None:
            yield "ident", match.group(), i, match.end()
            i = match.end()
            continue
        yield "char", char, i, i + 1
        i += 1


def _prefix_future_functions(text: str) -> str:
    tokens = list(_scan_tokens(text))
    pieces = []
    cursor = 0
    for index, (kind, value, start, end) in enumerate(tokens):
        if kind != "ident":
            continue
        upper = value.upper()
        if upper in _UNSUPPORTED_FUNCTIONS:
            following = tokens[index + 1] if index + 1 < len(tokens) else None
            if following is not None and following[1] == "(":
                raise ValueError(
                    f"{upper} needs _xlpm parameter names, which xlspy does not write"
                )
        if upper not in _XLFN_FUNCTIONS:
            continue
        following = tokens[index + 1] if index + 1 < len(tokens) else None
        if following is None or following[1] != "(":
            continue
        pieces.append(text[cursor:start])
        pieces.append("_xlfn." + value)
        cursor = end
    pieces.append(text[cursor:])
    return "".join(pieces)


def _check_balance(text: str) -> None:
    depth = 0
    for kind, value, _start, _end in _scan_tokens(text):
        if kind != "char":
            continue
        if value == "(":
            depth += 1
        elif value == ")":
            depth -= 1
            if depth < 0:
                break
    if depth != 0:
        raise ValueError(f"unbalanced parentheses in formula {text!r}")


def normalize_formula(text: str) -> str:
    """Validate formula text and return the form stored in the file."""
    if not isinstance(text, str):
        raise TypeError("formula text must be a string such as 'SUM(A1:A3)'")
    body = text.strip()
    if body.startswith("="):
        body = body[1:].lstrip()
    if not body:
        raise ValueError("formula text is empty")
    if len(body) > MAX_FORMULA_LENGTH:
        raise ValueError(f"formula is longer than {MAX_FORMULA_LENGTH} characters")
    if _CONTROL_RE.search(body):
        raise ValueError("formula text must not contain control characters")
    _check_balance(body)
    return _prefix_future_functions(body)


@dataclass(frozen=True, slots=True)
class Formula:
    """A formula cell.

    ``text`` is the formula without its leading ``=``. ``value`` is an optional
    cached result, written so that readers that do not recalculate still see a
    number. ``array_ref`` (for example ``"B2:B5"``) writes a legacy
    array (CSE) formula over that range. Dynamic-array spill formulas are not
    written; see the documentation for the supported forms.
    """

    text: str
    value: object = None
    array_ref: str | None = None

    def __post_init__(self):
        object.__setattr__(self, "text", normalize_formula(self.text))
        if self.value is not None and (
            isinstance(self.value, bool)
            or not isinstance(self.value, (int, float, str, Decimal))
        ):
            raise TypeError("formula cached value must be a number, a string or None (not a bool)")
        # A cached <v> must be a finite xsd:double; NaN and infinities are not.
        if isinstance(self.value, (float, Decimal)) and not math.isfinite(self.value):
            raise ValueError("formula cached value must be finite (not NaN or infinity)")
        if self.array_ref is not None:
            if not isinstance(self.array_ref, str):
                raise TypeError("array_ref must be an A1 range or None")
            parse_cell_range(self.array_ref)
