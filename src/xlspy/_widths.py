"""Automatic column widths: the text a cell shows and how wide it is.

Widths are in Excel's column-width units: the width of the digit ``0`` in the
default font, Calibri 11, at 96 DPI (7 pixels). The per-character widths are
Calibri's pixel advances at that size, as Excel measures text for AutoFit on a
display at 100% scaling. At other scalings Excel measures at the scaled size,
where hinting rounds the letters differently; the 96-DPI widths come out a few
percent wider there, never narrower.

Excel renders dates, month names, booleans and separators in the reader's
locale, which is unknown when the file is written, so those estimates take a
long rendering (``poniedziałek``, ``październik``, ``PRAWDA``).
"""

from __future__ import annotations

import datetime
import functools
import math
import re
import unicodedata
from decimal import ROUND_HALF_UP, Decimal

from . import _calibri_metrics as _calibri
from ._scalars import numpy_value
from .formulas import Formula
from .hyperlinks import Hyperlink
from .formats import _is_formatted_cell
from .styles import merge_styles

DEFAULT_WIDTH = 8.43
MAX_WIDTH = 255.0
# The file's column width also holds Excel's 5-pixel cell padding: 5/7 of a
# digit in Calibri 11 (so the default 8.43 is stored as 9.140625).
_CELL_PADDING = 5 / 7
# Pixels of the digit 0 in Calibri 11 at 96 DPI: the column-width unit.
_DIGIT_PX = 7
_DEFAULT_SIZE = 11.0
# One pixel more than Excel's AutoFit: a cut-off text (or ### for a number) is
# worse than a column a pixel too wide, and other fonts or display scalings
# measure a little differently.
_SPARE_PX = 1
# The autofilter (or table) drop-down button covers the end of a header cell.
FILTER_BUTTON = 2.0

_DESIGN = {False: dict(zip(_calibri.CHARS, _calibri.REGULAR)), True: dict(zip(_calibri.CHARS, _calibri.BOLD))}


def _pixel_size(size: float | None) -> int:
    """Calibri's em in pixels at 96 DPI for a size in points (11 pt: 15)."""
    return max(1, math.floor((size or _DEFAULT_SIZE) * 4 / 3 + 0.5))


@functools.lru_cache(maxsize=64)
def _advances(bold: bool, pixels: int) -> dict[str, int]:
    """Pixel advance of each character as GDI draws Calibri ``pixels`` high."""
    units = _calibri.UNITS
    advances = {
        char: (design * pixels * 2 + units) // (2 * units) for char, design in _DESIGN[bold].items()
    }
    plus, minus = _calibri.HINTING.get((bold, pixels), ("", ""))
    for char in plus:
        advances[char] += 1
    for char in minus:
        advances[char] -= 1
    return advances


def _char_pixels(char: str, bold: bool, pixels: int, advances: dict[str, int]) -> int:
    """A character Calibri lacks: East Asian text comes from a fallback font."""
    if unicodedata.east_asian_width(char) in ("W", "F"):
        return (pixels * (17 if bold else 16) + 14) // 15
    if unicodedata.combining(char) or unicodedata.category(char) in ("Cc", "Cf"):
        return 0
    return advances["0"]


def _sum_pixels(text: str, advances: dict[str, int], bold: bool, pixels: int) -> int:
    try:
        return sum(map(advances.__getitem__, text))
    except KeyError:
        total = 0
        for char in text:
            advance = advances.get(char)
            total += advance if advance is not None else _char_pixels(char, bold, pixels, advances)
        return total


def text_pixels(text: str, bold: bool = False, size: float | None = None) -> int:
    """Width of ``text`` on one line in Calibri at 96 DPI, in pixels."""
    pixels = _pixel_size(size)
    return _sum_pixels(text, _advances(bold, pixels), bold, pixels)


def text_width(text: str, bold: bool = False, size: float | None = None) -> float:
    """Width of ``text`` on one line in digits of the default font, without padding."""
    return text_pixels(text, bold, size) / _DIGIT_PX


def _autofit_pixels(text_px: int, size: float | None) -> int:
    """The column Excel's AutoFit makes for text ``text_px`` pixels wide.

    Measured on a display at 100% scaling (6,144 cells, 8 to 36 pt, regular
    and bold, all matched): a margin that grows by 2 pixels every 8 pixels of
    font size (2 at 11 pt), and long text gets 1/28 more room.
    """
    margin = 2 * math.ceil(_pixel_size(size) / 8) - 2
    return max(text_px + margin, text_px * 29 // 28 + margin - 2) + _SPARE_PX


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
# numFmt 14 and 22 follow the reader's locale: 2024-09-30, 30.09.2024 or (widest,
# with slashes) 12/31/2024.
_SHORT_DATE = "12/31/2024"           # numFmt 14, the writer's default for dates
_SHORT_DATETIME = "12/31/2024 23:45"  # numFmt 22, the default for datetimes
_BOOLEAN = "PRAWDA"                 # TRUE in the longest common language


def _sections(number_format: str) -> list[str]:
    sections: list[str] = []
    current: list[str] = []
    quoted = False
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


def _currency(match: re.Match[str]) -> str:
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
    def sample(match: re.Match[str]) -> str:
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


def _decimal(value: float) -> Decimal:
    """The number as Excel sees it: the shortest decimal that round-trips."""
    return Decimal(repr(float(value))) if not isinstance(value, int) else Decimal(value)


def _scientific(value: float) -> str:
    """General's scientific form: up to 6 significant digits, zeros trimmed."""
    mantissa, power = f"{_decimal(value):.5E}".split("E")
    if "." in mantissa:
        mantissa = mantissa.rstrip("0").rstrip(".")
    return f"{mantissa}E{power[0]}{abs(int(power)):02d}"


def _general_number(value: float) -> str:
    """Excel's General format: 11 characters besides the sign, rounded half up.

    Scientific instead when the number does not fit, or when the fixed form
    rounds it to fewer than the 6 significant digits scientific shows.
    """
    if value == 0:
        return "0"
    sign = "-" if value < 0 else ""
    magnitude = abs(_decimal(value))
    if magnitude == magnitude.to_integral_value():
        digits = str(int(magnitude))
        return sign + digits if len(digits) <= 11 else _scientific(value)
    if Decimal("1e-9") <= magnitude < Decimal("1e11"):
        whole = int(magnitude)
        decimals = max(0, 10 - len(str(whole)))
        rounded = magnitude.quantize(Decimal(1).scaleb(-decimals), rounding=ROUND_HALF_UP)
        fixed = f"{rounded:f}"
        if "." in fixed:
            fixed = fixed.rstrip("0").rstrip(".")
        # Significant digits the fixed form has room for.
        room = decimals + len(str(whole)) if whole else decimals - (-magnitude.adjusted() - 1)
        if fixed != "0" and (rounded == magnitude or room >= 6):
            return sign + fixed
    return _scientific(value)


def _formatted_number_text(value: float, number_format: str) -> str:
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
    # A negative number that rounds to zero shows no sign.
    sign = "-" if value < 0 and len(sections) < 2 and number.strip("0.,") else ""
    return _literal_text(head) + sign + number + _literal_text(tail)


def _shown(value, number_format: str | None) -> tuple[str | None, str | None]:
    """(the text a cell shows, the text AutoFit makes room for)."""
    value = numpy_value(value)  # written as the Python value it holds
    if isinstance(value, Hyperlink):
        value = value.display_text
    elif isinstance(value, Formula):
        value = value.value
    if value is None:
        return None, None
    if isinstance(value, str):
        return value, value
    if isinstance(value, bool):
        return _BOOLEAN, _BOOLEAN
    if isinstance(value, (datetime.date, datetime.datetime)):
        if number_format:
            section = _BRACKETS.sub(_currency, _sections(number_format)[0])
            text = _date_text(section)
        else:
            text = _SHORT_DATETIME if isinstance(value, datetime.datetime) else _SHORT_DATE
        return text, text
    if isinstance(value, (int, float, Decimal)):
        number = float(value) if isinstance(value, Decimal) else value
        if isinstance(number, float) and not math.isfinite(number):
            text = "NaN" if number != number else "-∞" if number < 0 else "∞"
            return text, text
        if number_format and number_format.strip().lower() != "general":
            try:
                text = _formatted_number_text(number, number_format)
                return text, text
            except (ValueError, OverflowError):
                pass
        text = _general_number(number)
        # AutoFit keeps room for any General rendering of the number: the
        # decimal point as wide as a digit and a negative exponent as E+.
        room = text.replace(".", "0").replace("E-", "E+")
        return text, room
    text = str(value)
    return text, text


def _general_room(number: float) -> str | None:
    """``_shown(number, None)[1]`` for an int or float, faster for common values."""
    if number == 0:
        return "0"
    if number == number and number not in (math.inf, -math.inf):
        text = repr(number)
        # A short exact repr is what General shows (1.5, -42, 0.25).
        digits = len(text) - (number < 0)
        if digits <= 11 and "e" not in text and "n" not in text:
            if text.endswith(".0"):
                text = text[:-2]
            return text.replace(".", "0")
    return _shown(number, None)[1]


def cell_text(value, number_format: str | None = None) -> str | None:
    """The text a cell shows in Excel, approximately; ``None`` for no value."""
    return _shown(value, number_format)[0]


def cell_width(value, number_format: str | None = None, *, bold: bool = False,
               size: float | None = None, wrap: bool = False) -> float:
    """Column width that shows ``value`` whole (0.0 when it shows nothing).

    Excel's own AutoFit on a display at 100% scaling, for Calibri: other
    scalings and fonts come out close, usually a little wider.
    """
    text = _shown(value, number_format)[1]
    if not text:
        return 0.0
    if wrap and "\n" in text:
        text_px = max(text_pixels(line, bold, size) for line in text.split("\n"))
    else:
        text_px = text_pixels(text.replace("\n", ""), bold, size)
    return _autofit_pixels(text_px, size) / _DIGIT_PX


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


class ColumnWidthMeter:
    """Automatic column widths, measured as the rows stream past.

    ``styles_at(row_index, row_length)`` gives the row's ``{column: Style}``
    (header and range styles); ``(value, format)`` and ``(value, Style)``
    cells add their own. Row 1 is bold by default. ``rows`` is how many rows
    are measured (``None``: all). ``button_rows`` are the rows that may hold
    filter drop-down buttons: their cells are also kept one by one, so a
    button widens only the cell under it.
    """

    def __init__(self, styles_at=None, rows: int | None = 1000, button_rows=frozenset()):
        self.limit = math.inf if rows is None else rows
        self._styles_at = styles_at
        self._button_rows = frozenset(button_rows)
        self._best: list[int] = []  # column pixels
        self._cells: dict[int, dict[int, int]] = {}  # button row -> {column: pixels}
        self._style_cache: dict = {}
        self._fonts: dict = {}  # (bold, size) -> (advances, pixels, margin)
        self._constant: dict = {}  # (type, number format, bold, size) -> pixels of dates and booleans

    def add(self, row_idx: int, row) -> None:
        """Measure one row (the caller checks ``row_idx < meter.limit``)."""
        styles_at = self._styles_at
        row_styles = styles_at(row_idx, len(row)) if styles_at is not None else None
        best = self._best
        if len(row) > len(best):
            best.extend([0] * (len(row) - len(best)))
        kept = self._cells.setdefault(row_idx, {}) if row_idx in self._button_rows else None
        style_cache, fonts, constant = self._style_cache, self._fonts, self._constant
        for col_idx, cell in enumerate(row):
            if cell is None:
                continue
            style = row_styles.get(col_idx) if row_styles else None
            number_format = None
            if type(cell) is tuple and _is_formatted_cell(cell):
                cell, overlay = cell
                if isinstance(overlay, str):
                    number_format = overlay
                else:
                    style = merge_styles(style, overlay)
            style_format, bold, size, wrap = _style_props(style, style_cache)
            number_format = number_format or style_format
            bold = bold or (row_idx == 0 and style is None)
            font = fonts.get((bold, size))
            if font is None:
                pixels = _pixel_size(size)
                font = fonts[(bold, size)] = (_advances(bold, pixels), pixels, 2 * math.ceil(pixels / 8) - 2)
            advances, pixels, margin = font
            kind = type(cell)
            if kind is str:
                text = cell
            elif kind is bool or kind is datetime.date or kind is datetime.datetime:
                key = (kind, number_format, bold, size)
                text_px = constant.get(key)
                if text_px is None:
                    text_px = constant[key] = _sum_pixels(_shown(cell, number_format)[1] or "", advances, bold, pixels)
                text = None
            elif (kind is int or kind is float) and (number_format is None or number_format == "General"):
                text = _general_room(cell)
            else:
                text = _shown(cell, number_format)[1]
                if text is None:  # a formula without a cached value
                    continue
            if text is not None:
                if not text:
                    continue
                if "\n" in text:
                    if wrap:
                        text_px = max(_sum_pixels(line, advances, bold, pixels) for line in text.split("\n"))
                    else:
                        text_px = _sum_pixels(text.replace("\n", ""), advances, bold, pixels)
                else:
                    text_px = _sum_pixels(text, advances, bold, pixels)
            if not text_px:
                continue
            width = max(text_px + margin, text_px * 29 // 28 + margin - 2) + _SPARE_PX
            if width > best[col_idx]:
                best[col_idx] = width
            if kept is not None:
                kept[col_idx] = width

    def widths(self, max_cols: int, filter_cells=(), padding: float = 0.0) -> list[float]:
        """Widths of the first ``max_cols`` columns in characters, never below
        Excel's default. ``filter_cells`` holds ``(row, first_col, last_col)``
        header spans whose cells carry a filter drop-down button; ``padding``
        (characters) is added to every measured column."""
        best = (self._best + [0] * max_cols)[:max_cols]
        button = round(FILTER_BUTTON * _DIGIT_PX)
        for r, c1, c2 in filter_cells:
            kept = self._cells.get(r)
            for col in range(max(c1, 0), min(c2 + 1, max_cols)):
                if kept is not None:
                    width = kept.get(col)
                elif r < self.limit and r not in self._button_rows:
                    width = best[col]  # the cell was not kept: assume it is the widest
                else:
                    width = None
                if width:
                    best[col] = max(best[col], width + button)
        return [
            min(MAX_WIDTH, max(DEFAULT_WIDTH, px / _DIGIT_PX + padding)) if px else DEFAULT_WIDTH
            for px in best
        ]


def estimate_column_widths(rows, max_cols: int, styles_at=None, filter_cells=(),
                           padding: float = 0.0) -> list[float]:
    """Column widths that show ``rows`` whole (see ``ColumnWidthMeter``)."""
    meter = ColumnWidthMeter(styles_at, None, {r for r, _c1, _c2 in filter_cells})
    for row_idx, row in enumerate(rows):
        meter.add(row_idx, row[:max_cols] if len(row) > max_cols else row)
    return meter.widths(max_cols, filter_cells, padding)
