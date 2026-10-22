"""Automatic column widths: the text a cell shows and how wide it is.

Widths are in Excel's column-width units (the width of the digit ``0`` in the
default font, Calibri 11). The per-character widths below were measured with
Excel's own AutoFit (``tools/calibrate_column_widths.py``); a character not
listed counts as one digit, East Asian wide characters as two.

Excel renders dates, month names, booleans and separators in the reader's
locale, which is unknown when the file is written, so those estimates take a
long rendering (``poniedziałek``, ``październik``, ``PRAWDA``).
"""

from __future__ import annotations

import datetime
import math
import re
import unicodedata
from decimal import Decimal

from .formulas import Formula
from .hyperlinks import Hyperlink

DEFAULT_WIDTH = 8.43
MAX_WIDTH = 255.0
# The file's column width also holds Excel's 5-pixel cell padding: 5/7 of a
# digit in Calibri 11 (so the default 8.43 is stored as 9.140625).
_CELL_PADDING = 5 / 7
# AutoFit leaves a little room after the text; this also absorbs rounding.
_PADDING = 1.0
# The autofilter (or table) drop-down button covers the end of a header cell.
FILTER_BUTTON = 2.0
_BOLD = 1.03

_WIDTH_GROUPS = (
    ("il'j,ł", 0.47),
    (" ", 0.47),
    (".:;I", 0.55),
    ("!()-[]`{}fJ", 0.63),
    ("rt°", 0.70),
    ('"/\\szśźż', 0.78),
    ("cxLćŁ„”", 0.86),
    ("?FSZagkvyą|ŚŹŻ", 0.94),
    ("0123456789#$*+<=>^_~EPTYeęĘé€£§±×–", 1.02),
    ("BCKRXbdhnopquĆńóüß", 1.10),
    ("AVĄ", 1.18),
    ("DGH", 1.25),
    ("NOUŃÓÜΩ", 1.33),
    ("&Qж…", 1.41),
    ("%w", 1.49),
    ("mЖ", 1.65),
    ("M", 1.72),
    ("@W", 1.80),
    ("—", 1.88),
)
_WIDTHS = {char: width for chars, width in _WIDTH_GROUPS for char in chars}


def _char_width(char: str) -> float:
    width = _WIDTHS.get(char)
    if width is not None:
        return width
    if unicodedata.east_asian_width(char) in ("W", "F"):
        return 2.04
    if unicodedata.combining(char):
        return 0.0
    return 1.02


def text_width(text: str) -> float:
    """Width of ``text`` on one line in the default font, without padding."""
    widths = _WIDTHS
    total = 0.0
    for char in text:
        width = widths.get(char)
        total += width if width is not None else _char_width(char)
    return total


# --- what a cell shows -----------------------------------------------------

_QUOTED = re.compile(r'"[^"]*"')
_BRACKETS = re.compile(r"\[[^\]]*\]")
_DATE_TOKENS = re.compile(
    r"(yyyy|yy|mmmmm|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|am/pm|a/p|\.0+)",
    re.IGNORECASE,
)
# Long renderings: Excel uses the reader's language for names.
_DATE_SAMPLES = {
    "yyyy": "2024", "yy": "24", "mmmmm": "W", "mmmm": "październik", "mmm": "Sep",
    "mm": "12", "m": "12", "dddd": "poniedziałek", "ddd": "Wed", "dd": "30",
    "d": "30", "hh": "23", "h": "23", "ss": "59", "s": "59", "am/pm": "PM",
    "a/p": "P",
}
_SHORT_DATE = "30.09.2024"           # numFmt 14, the writer's default for dates
_SHORT_DATETIME = "30.09.2024 13:45"  # numFmt 22, the default for datetimes
_BOOLEAN = "PRAWDA"                 # TRUE in the longest common language


def _sections(number_format: str) -> list[str]:
    sections, current, quoted = [], [], False
    for char in number_format:
        if char == '"':
            quoted = not quoted
        if char == ";" and not quoted:
            sections.append("".join(current))
            current = []
        else:
            current.append(char)
    sections.append("".join(current))
    return sections


def _literal_text(part: str) -> str:
    """Format text outside placeholders as shown: quotes, escapes, ``_x``, ``*x``."""
    out, i = [], 0
    while i < len(part):
        char = part[i]
        if char == '"':
            end = part.find('"', i + 1)
            end = len(part) if end < 0 else end
            out.append(part[i + 1:end])
            i = end + 1
            continue
        if char in "\\_" and i + 1 < len(part):
            out.append(part[i + 1])  # ``_x`` leaves the width of x
            i += 2
            continue
        if char == "*":
            i += 2  # a fill character takes no width of its own
            continue
        out.append(char)
        i += 1
    return "".join(out)


def _currency(match: re.Match) -> str:
    inner = match.group(0)[1:-1]
    if inner.startswith("$"):
        return '"' + inner[1:].split("-", 1)[0] + '"'
    return ""


def _is_date_format(section: str) -> bool:
    plain = _QUOTED.sub("", _BRACKETS.sub("", section)).lower()
    return any(token in plain for token in ("y", "d", "h", "s")) or (
        "m" in plain and "0" not in plain and "#" not in plain
    )


def _date_text(section: str) -> str:
    def sample(match: re.Match) -> str:
        token = match.group(0)
        if token.startswith("."):
            return "." + "0" * (len(token) - 1)
        return _DATE_SAMPLES[token.lower()]

    pieces = re.split(r'("[^"]*")', section)
    out = []
    for piece in pieces:
        if piece.startswith('"'):
            out.append(piece[1:-1])
        else:
            out.append(_DATE_TOKENS.sub(sample, _literal_text(piece)))
    return "".join(out)


def _scientific(value: float) -> str:
    """General's scientific form: up to 6 significant digits, zeros trimmed."""
    mantissa, power = f"{value:.5E}".split("E")
    if "." in mantissa:
        mantissa = mantissa.rstrip("0").rstrip(".")
    return f"{mantissa}E{power[0]}{abs(int(power)):02d}"


def _general_number(value: float) -> str:
    """Excel's General format: at most 11 characters, else scientific."""
    if value == 0:
        return "0"
    if isinstance(value, int) or float(value).is_integer():
        text = str(int(value))
        if len(text) <= 11:
            return text
        return _scientific(value)
    magnitude = abs(value)
    if 1e-9 <= magnitude < 1e11:
        whole = len(str(int(magnitude))) + (1 if value < 0 else 0)
        decimals = max(0, 10 - whole)
        text = f"{value:.{decimals}f}".rstrip("0").rstrip(".")
        if text not in ("0", "-0"):
            return text
    return _scientific(value)


def _number_text(value: float, number_format: str) -> str:
    sections = _sections(number_format)
    section = sections[0]
    if value < 0 and len(sections) > 1 and sections[1]:
        section, value = sections[1], -value
    section = _BRACKETS.sub(_currency, section)
    if section.strip().lower() in ("", "general"):
        return _general_number(value)
    if "@" in section and not re.search(r"[0#?]", _QUOTED.sub("", section)):
        return _literal_text(section.replace("@", _general_number(value)))
    if _is_date_format(section):
        return _date_text(section)
    plain = _QUOTED.sub("", section)
    value *= 100 ** plain.count("%")
    match = re.search(r"[0#?][0#?,.]*(?:[eE][+-][0#]+)?", section)
    if match is None:
        return _literal_text(section)
    placeholder = match.group(0)
    head, tail = section[:match.start()], section[match.end():]
    exponent = re.search(r"[eE][+-]([0#]+)$", placeholder)
    if exponent:
        mantissa = placeholder[:exponent.start()]
        decimals = len(mantissa.split(".", 1)[1]) if "." in mantissa else 0
        number = f"{abs(value):.{decimals}E}"
        digits, power = number.split("E")
        number = f"{digits}E{power[0]}{abs(int(power)):0{len(exponent.group(1))}d}"
    else:
        integer_part, _, fraction = placeholder.partition(".")
        decimals = sum(fraction.count(c) for c in "0#?")
        thousands = "," in integer_part.strip(",")
        number = format(abs(value), f"{',' if thousands else ''}.{decimals}f")
        min_digits = integer_part.count("0")
        whole, dot, rest = number.partition(".")
        if len(whole.replace(",", "")) < min_digits:
            whole = whole.zfill(min_digits)
        number = whole + dot + rest
    sign = "-" if value < 0 and len(sections) < 2 else ""
    return _literal_text(head) + sign + number + _literal_text(tail)


def cell_text(value, number_format: str | None = None) -> str | None:
    """The text a cell shows in Excel, approximately; ``None`` for no value."""
    if isinstance(value, Hyperlink):
        value = value.display_text
    elif isinstance(value, Formula):
        value = value.value
    if value is None:
        return None
    if isinstance(value, str):
        return value
    if isinstance(value, bool):
        return _BOOLEAN
    if isinstance(value, (datetime.date, datetime.datetime)):
        if number_format:
            section = _BRACKETS.sub(_currency, _sections(number_format)[0])
            return _date_text(section)
        return _SHORT_DATETIME if isinstance(value, datetime.datetime) else _SHORT_DATE
    if isinstance(value, (int, float, Decimal)):
        number = float(value) if isinstance(value, Decimal) else value
        if isinstance(number, float) and not math.isfinite(number):
            return "NaN" if number != number else "-∞" if number < 0 else "∞"
        if number_format:
            try:
                return _number_text(number, number_format)
            except (ValueError, OverflowError):
                return _general_number(number)
        return _general_number(number)
    return str(value)


def cell_width(value, number_format: str | None = None, *, bold: bool = False,
               size: float | None = None, wrap: bool = False) -> float:
    """Column width that shows ``value`` whole (0.0 when it shows nothing)."""
    text = cell_text(value, number_format)
    if not text:
        return 0.0
    if wrap and "\n" in text:
        width = max(text_width(line) for line in text.split("\n"))
    else:
        width = text_width(text.replace("\n", ""))
    if bold:
        width *= _BOLD
    if size:
        width *= size / 11.0
    return width + _PADDING


def file_width(width: float) -> float:
    """A width in characters (Excel's ``ColumnWidth``) as stored in the file."""
    return width + _CELL_PADDING


def _style_props(style, cache: dict):
    """(number_format, bold, size, wrap) of a Style, cached by identity."""
    if style is None:
        return None, False, None, False
    cached = cache.get(id(style))
    if cached is not None and cached[0] is style:
        return cached[1]
    font, alignment = style.font, style.alignment
    props = (
        style.number_format,
        bool(style.bold or (font is not None and font.bold)),
        font.size if font is not None else None,
        bool(style.wrap or (alignment is not None and alignment.wrap_text)),
    )
    cache[id(style)] = (style, props)
    return props


def estimate_column_widths(rows, max_cols: int, styles_at=None, filter_cells=()) -> list[float]:
    """Column widths that show the sampled ``rows`` whole.

    ``styles_at(row_index, row_length)`` gives the row's ``{column: Style}``
    (header and range styles); ``(value, format)`` and ``(value, Style)``
    cells add their own. ``filter_cells`` holds ``(row, first_col,
    last_col)`` header spans whose cells carry a filter drop-down button.
    Row 1 is bold by default. Widths never go below Excel's default.
    """
    from .formats import _is_formatted_cell
    from .styles import merge_styles

    widths = [0.0] * max_cols
    cache: dict = {}
    for row_idx, row in enumerate(rows):
        row_styles = styles_at(row_idx, len(row)) if styles_at is not None else None
        for col_idx, cell in enumerate(row):
            if col_idx >= max_cols:
                break
            if cell is None:
                continue
            style = row_styles.get(col_idx) if row_styles else None
            number_format = None
            if _is_formatted_cell(cell):
                cell, overlay = cell
                if isinstance(overlay, str):
                    number_format = overlay
                else:
                    style = merge_styles(style, overlay)
            style_format, bold, size, wrap = _style_props(style, cache)
            width = cell_width(
                cell, number_format or style_format,
                bold=bold or (row_idx == 0 and style is None), size=size, wrap=wrap,
            )
            if width and any(r == row_idx and c1 <= col_idx <= c2 for r, c1, c2 in filter_cells):
                width += FILTER_BUTTON
            if width > widths[col_idx]:
                widths[col_idx] = width
    return [min(MAX_WIDTH, max(DEFAULT_WIDTH, width)) for width in widths]
