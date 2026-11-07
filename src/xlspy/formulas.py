"""Excel formula cells.

A ``Formula`` is written as a formula, never as text. xlspy does not evaluate
formulas: XLSX stores normalized formula text, and XLSB compiles supported
expressions to BIFF12 tokens. Excel recalculates on open.
"""

from __future__ import annotations
from collections.abc import Iterator

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
    "ARRAYTOTEXT",
    "LET",
    "LAMBDA",
    "MAP",
    "REDUCE",
    "SCAN",
    "BYROW",
    "BYCOL",
    "MAKEARRAY",
    "BITAND",
    "BITLSHIFT",
    "BITOR",
    "BITRSHIFT",
    "BITXOR",
    "ECMA.CEILING",
    "ENCODEURL",
    "FILTERXML",
    "FORECAST.LINEAR",
    "FORMULATEXT",
    "GAUSS",
    "ISFORMULA",
    "ISO.CEILING",
    "NETWORKDAYS.INTL",
    "NUMBERVALUE",
    "PDURATION",
    "PHI",
    "RRI",
    "SHEET",
    "SHEETS",
    "SKEW.P",
    "UNICHAR",
    "UNICODE",
    "VALUETOTEXT",
    "WEBSERVICE",
    "WORKDAY.INTL",
    "Z.TEST",
    # Missing before: Excel stores these with the prefix too, and without it the
    # cell shows #NAME?. Names checked against the list XlsxWriter writes (MS-XLSX).
    "ACOT", "ACOTH", "ARABIC", "BASE", "BINOM.DIST.RANGE", "COMBINA",
    "CONFIDENCE.NORM", "CONFIDENCE.T", "COT", "COTH", "CSC", "CSCH", "DECIMAL",
    "FORECAST.ETS", "FORECAST.ETS.CONFINT", "FORECAST.ETS.SEASONALITY",
    "FORECAST.ETS.STAT", "IMAGE", "IMCOSH", "IMCOT", "IMCSC", "IMCSCH", "IMSEC",
    "IMSECH", "IMSINH", "IMTAN", "ISOMITTED", "MUNIT", "PERMUTATIONA",
    "QUERYSTRING", "SEC", "SECH",
})

# Functions stored in the _xlws namespace: "_xlfn._xlws.SORT", not "_xlfn.SORT".
_XLWS_FUNCTIONS = frozenset({"FILTER", "SORT"})


def future_function_name(name: str) -> str:
    """The stored name of a post-2007 function: ``_xlfn.CONCAT``, ``_xlfn._xlws.SORT``.

    Accepts the name with or without a prefix already in place.
    """
    bare = name.removeprefix("_XLFN.").removeprefix("_xlfn.")
    bare = bare.removeprefix("_xlws.").removeprefix("_XLWS.")
    if bare.upper() in _XLWS_FUNCTIONS:
        return "_xlfn._xlws." + bare
    return "_xlfn." + bare


_IDENT_RE = re.compile(r"[A-Za-z_][A-Za-z0-9_.]*")
_CONTROL_RE = re.compile(r"[\x00-\x1f\x7f]")


def _scan_tokens(text: str) -> Iterator[tuple[str, str, int, int]]:
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
        if upper not in _XLFN_FUNCTIONS:
            continue
        following = tokens[index + 1] if index + 1 < len(tokens) else None
        if following is None or following[1] != "(":
            continue
        pieces.append(text[cursor:start])
        pieces.append(future_function_name(value))
        cursor = end
    pieces.append(text[cursor:])
    return "".join(pieces)


def _matching_paren(tokens: list[tuple[str, str, int, int]], open_index: int) -> int | None:
    depth = 0
    for index in range(open_index, len(tokens)):
        kind, value, _start, _end = tokens[index]
        if kind != "char":
            continue
        if value == "(":
            depth += 1
        elif value == ")":
            depth -= 1
            if depth == 0:
                return index
    return None


def _split_arguments(
    tokens: list[tuple[str, str, int, int]], start: int, end: int
) -> list[tuple[int, int]]:
    """Token ranges of the top-level arguments between ``start`` and ``end``."""
    ranges = []
    depth = 0
    argument_start = start
    for index in range(start, end):
        kind, value, _start, _end = tokens[index]
        if kind != "char":
            continue
        if value in "([{":
            depth += 1
        elif value in ")]}":
            depth -= 1
        elif value == "," and depth == 0:
            ranges.append((argument_start, index))
            argument_start = index + 1
    ranges.append((argument_start, end))
    return ranges


def _prefix_lambda_parameters(text: str) -> str:
    """Prefix the names bound by ``LET`` and ``LAMBDA`` with ``_xlpm.``.

    Excel stores ``LET(x,1,x+1)`` as ``_xlfn.LET(_xlpm.x,1,_xlpm.x+1)``: the
    binding and every use of it. ``LET`` is sequential (a value sees the names
    bound before it); a ``LAMBDA`` body sees all of its parameters. Nested calls
    and a lambda call (``LAMBDA(x,x*2)(3)``) are handled.
    """
    lowered = text.lower()
    if "let" not in lowered and "lambda" not in lowered:
        return text
    tokens = list(_scan_tokens(text))

    def significant(first: int, last: int) -> list[tuple[str, str, int, int]]:
        return [
            token for token in tokens[first:last]
            if not (token[0] == "char" and token[1].isspace())
        ]

    def following_significant(index: int, last: int) -> str:
        for next_index in range(index + 1, last):
            kind, value, _start, _end = tokens[next_index]
            if not (kind == "char" and value.isspace()):
                return value
        return ""

    def is_call(index: int, end: int) -> int | None:
        if index + 1 >= end or tokens[index + 1][1] != "(":
            return None
        close = _matching_paren(tokens, index + 1)
        return close if close is not None and close < end else None

    def render(start: int, end: int, names: frozenset[str]) -> str:
        if start >= end:
            return ""
        pieces = []
        cursor = tokens[start][2]
        index = start
        bracket_depth = 0
        previous = ""
        while index < end:
            kind, value, token_start, token_end = tokens[index]
            if kind == "char":
                if value == "[":
                    bracket_depth += 1
                elif value == "]":
                    bracket_depth = max(0, bracket_depth - 1)
                if not value.isspace():
                    previous = value
                index += 1
                continue
            upper = value.upper()
            close = None
            if kind == "ident" and upper in ("LET", "LAMBDA"):
                close = is_call(index, end)
            if close is not None:
                pieces.append(text[cursor:token_end])
                pieces.append(render_call(index, close, names))
                cursor = tokens[close][3]
                index = close + 1
                previous = ")"
                continue
            if (
                kind == "ident"
                and bracket_depth == 0
                and value.casefold() in names
                and not value.lower().startswith("_xlpm.")
                and previous not in ("!", ".")
                and following_significant(index, end) not in ("!", "[", ":")
            ):
                pieces.append(text[cursor:token_start])
                pieces.append("_xlpm.")
                cursor = token_start
            previous = value
            index += 1
        pieces.append(text[cursor:tokens[end - 1][3]])
        return "".join(pieces)

    def render_call(index: int, close: int, names: frozenset[str]) -> str:
        upper = tokens[index][1].upper()
        arguments = _split_arguments(tokens, index + 2, close)
        rendered = []
        if upper == "LET":
            scope = set(names)
            last = len(arguments) - 1
            for argument_index, (first, last_token) in enumerate(arguments):
                tokens_here = significant(first, last_token)
                if (
                    argument_index < last and argument_index % 2 == 0
                    and len(tokens_here) == 1 and tokens_here[0][0] == "ident"
                ):
                    name = tokens_here[0][1]
                    rendered.append(
                        name if name.lower().startswith("_xlpm.") else "_xlpm." + name
                    )
                    continue
                rendered.append(render(first, last_token, frozenset(scope)))
                if argument_index < last and argument_index % 2 == 1:
                    name_token = significant(*arguments[argument_index - 1])
                    if len(name_token) == 1 and name_token[0][0] == "ident":
                        scope.add(name_token[0][1].casefold())
        else:
            bindings = set()
            for argument_index, (first, last_token) in enumerate(arguments[:-1]):
                tokens_here = significant(first, last_token)
                if len(tokens_here) == 1 and tokens_here[0][0] == "ident":
                    name = tokens_here[0][1]
                    bindings.add(name.casefold())
                    rendered.append(
                        name if name.lower().startswith("_xlpm.") else "_xlpm." + name
                    )
                else:
                    rendered.append(render(first, last_token, names))
            if arguments:
                rendered.append(render(*arguments[-1], names | frozenset(bindings)))
        return "(" + ",".join(rendered) + ")"

    return render(0, len(tokens), frozenset())


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
    return _prefix_future_functions(_prefix_lambda_parameters(body))


@dataclass(frozen=True, slots=True)
class ArrayFormula:
    """A legacy CSE array formula that fills ``ref``.

    The same object openpyxl's ``openpyxl.worksheet.formula.ArrayFormula``
    represents: Excel stores the formula once, in the top-left cell of the
    range, and fills the other cells from it. ``text`` is the formula without
    its leading ``=``. Assign it to a cell to write an array formula:

    ```python
    from xlspy import ArrayFormula, Workbook

    workbook = Workbook()
    sheet = workbook.active
    sheet["A1"] = 1
    sheet["A2"] = 2
    sheet["B1"] = ArrayFormula("B1:B2", "A1:A2*2")
    workbook.save("array.xlsx")
    ```

    A dynamic-array formula is written with ``Formula(..., spill_ref=...)``;
    ``ArrayFormula`` is the static CSE form.
    """

    ref: str
    text: str = ""

    def __post_init__(self) -> None:
        if not isinstance(self.ref, str):
            raise TypeError("array formula ref must be an A1 range")
        parse_cell_range(self.ref)
        if not isinstance(self.text, str):
            raise TypeError("array formula text must be a string")
        body = self.text.strip()
        if body.startswith("="):
            body = body[1:].lstrip()
        object.__setattr__(self, "text", normalize_formula(body) if body else "")

    def __iter__(self) -> Iterator[tuple[str, str]]:
        yield "t", "array"
        if self.ref:
            yield "ref", self.ref


@dataclass(frozen=True, slots=True)
class DataTableFormula:
    """A data-table formula (``What-If Analysis > Data Table``).

    The fields follow openpyxl's ``DataTableFormula`` and the OOXML
    ``<f t="dataTable">`` attributes: ``r1``/``r2`` are the input cells,
    ``dt2D`` marks a two-dimensional table, ``dtr`` a row-oriented one, and
    ``del1``/``del2`` are Excel's "delete" flags. A data-table formula has no
    text and no cached value of its own.
    """

    ref: str
    ca: bool = False
    dt2D: bool = False
    dtr: bool = False
    r1: str | None = None
    r2: str | None = None
    del1: bool = False
    del2: bool = False

    def __post_init__(self) -> None:
        if not isinstance(self.ref, str):
            raise TypeError("data-table formula ref must be an A1 range")
        parse_cell_range(self.ref)
        for name in ("r1", "r2"):
            value = getattr(self, name)
            if value is not None and not isinstance(value, str):
                raise TypeError(f"data-table formula {name} must be a cell reference or None")
        for name in ("ca", "dt2D", "dtr", "del1", "del2"):
            if not isinstance(getattr(self, name), bool):
                raise TypeError(f"data-table formula {name} must be a bool")

    def __iter__(self) -> Iterator[tuple[str, object]]:
        yield "t", "dataTable"
        for name in ("ref", "dt2D", "dtr", "r1", "r2", "del1", "del2", "ca"):
            value = getattr(self, name)
            if value:
                yield name, value


@dataclass(frozen=True, slots=True)
class Formula:
    """A formula cell.

    ``text`` is the formula without its leading ``=``. ``value`` is an optional
    cached result, written so that readers that do not recalculate still see a
    number. ``array_ref`` (for example ``"B2:B5"``) writes a legacy
    array (CSE) formula over that range. ``spill_ref`` writes a dynamic-array
    formula over the range: Excel stores the anchor as an array formula and the
    file carries the dynamic-array cell metadata.

    Examples:
        Put a formula in a row, with an optional cached result:

        ```python
        from xlspy import Formula, XlsxWriter

        rows = [
            ["A", "B", "Total"],
            [2, 3, Formula("A2+B2", value=5)],
            [4, 6, Formula("A3+B3")],
        ]
        with XlsxWriter("formulas.xlsx") as workbook:
            workbook.add_sheet("Data")
            workbook.write_sheet(rows)
        ```

        Formula text has no leading ``=``. A plain string is always written as
        text, so exported data cannot inject a formula.

    See Also:
        `Hyperlink` for links, and the guide section on XLSB formula limits.
    """

    text: str
    value: object = None
    array_ref: str | None = None
    spill_ref: str | None = None

    def __post_init__(self) -> None:
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
        if self.spill_ref is not None:
            if not isinstance(self.spill_ref, str):
                raise TypeError("spill_ref must be an A1 range or None")
            if self.array_ref is not None:
                raise ValueError("a formula takes array_ref or spill_ref, not both")
            parse_cell_range(self.spill_ref)
