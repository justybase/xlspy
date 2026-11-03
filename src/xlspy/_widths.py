"""Automatic column widths: the text a cell shows and how wide it is.

Widths are in Excel's column-width units: the width of the digit ``0`` in the
default font, Calibri 11, at 96 DPI (7 pixels). A text is as wide as the sum
of its characters' pixel advances in the cell's font (``_fonts``), as Excel
measures text for AutoFit on a display at 100% scaling. At other scalings
Excel measures at the scaled size, where hinting rounds the letters
differently; the 96-DPI widths come out a few percent wider there, never
narrower.

Excel renders dates, month names, booleans and separators in the reader's
locale, which is unknown when the file is written, so those estimates take a
long rendering (``poniedziałek``, ``październik``, ``PRAWDA``).
"""

from __future__ import annotations

import datetime
import math
import re
import unicodedata
from decimal import ROUND_HALF_UP, Decimal

from . import _fonts
from ._accel import WidthMeter as _NativeWidthMeter
from ._scalars import plain_value
from .formulas import Formula
from .hyperlinks import Hyperlink
from .formats import _is_formatted_cell
from .rich_text import CellRichText, TextBlock
from .styles import Alignment, merge_styles

DEFAULT_WIDTH = 8.43
MAX_WIDTH = 255.0
# The file's column width also holds Excel's 5-pixel cell padding: 5/7 of a
# digit in Calibri 11 (so the default 8.43 is stored as 9.140625).
_CELL_PADDING = 5 / 7
# Pixels of the digit 0 in Calibri 11 at 96 DPI: the column-width unit.
_DIGIT_PX = 7
_DEFAULT_SIZE = 11.0
_DEFAULT_FAMILY = _fonts.DEFAULT_FAMILY
# One pixel more than Excel's AutoFit: a cut-off text (or ### for a number) is
# worse than a column a pixel too wide, and other fonts or display scalings
# measure a little differently.
_SPARE_PX = 1
# The autofilter (or table) drop-down button covers the end of a header cell:
# Excel's AutoFit adds 16 pixels for it at any font size (measured at 100%).
_FILTER_BUTTON_PX = 16
FILTER_BUTTON = _FILTER_BUTTON_PX / _DIGIT_PX  # in characters
# Each level of an alignment indent: 9 pixels at any font size (Excel sizes
# it in the Normal style's font). Centred and distributed text is indented
# from both sides, so AutoFit makes room for the indent twice.
_INDENT_PX = 9
_TWO_SIDED_INDENT = frozenset(("center", "distributed"))


def _pixel_size(size: float | None) -> int:
    """A font's em in pixels at 96 DPI for a size in points (11 pt: 15)."""
    return max(1, math.floor((size or _DEFAULT_SIZE) * 4 / 3 + 0.5))


def _advances(bold: bool, pixels: int, family: str = _DEFAULT_FAMILY, italic: bool = False) -> dict[str, int]:
    """Pixel advance of each character as GDI draws the font ``pixels`` high."""
    return _fonts.advances(family, bold, italic, pixels)


def _char_pixels(char: str, bold: bool, pixels: int, advances: dict[str, int]) -> int:
    """A character the font's table lacks: an emoji from Segoe UI Emoji, East
    Asian text from a CJK font, marks and controls take no room."""
    fallback = _fonts.fallback_pixels(char, pixels)
    if fallback is not None:
        return fallback
    east_asian = unicodedata.east_asian_width(char)
    if east_asian in ("W", "F"):
        return (pixels * (17 if bold else 16) + 14) // 15
    if east_asian == "H":  # half-width katakana and Hangul: about 8/15 of the em
        return (pixels * 8 + 14) // 15
    if unicodedata.combining(char) or unicodedata.category(char) in ("Cc", "Cf", "Mn", "Me"):
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


def text_pixels(text: str, bold: bool = False, size: float | None = None, *,
                font: str | None = None, italic: bool = False) -> int:
    """Width of ``text`` on one line at 96 DPI, in pixels, in the font named
    ``font`` (Calibri by default)."""
    pixels = _pixel_size(size)
    return _sum_pixels(text, _advances(bold, pixels, _fonts.family_key(font), italic), bold, pixels)


def text_width(text: str, bold: bool = False, size: float | None = None, *,
               font: str | None = None, italic: bool = False) -> float:
    """Width of ``text`` on one line in digits of the default font, without padding."""
    return text_pixels(text, bold, size, font=font, italic=italic) / _DIGIT_PX


def _column_pixels(text_px: int, margin: int) -> int:
    """The column AutoFit makes for ``text_px`` (slant and indent included),
    plus the spare pixel: the font's margin, and long text gets 1/28 more room."""
    return max(text_px + margin, text_px * 29 // 28 + margin - 2) + _SPARE_PX


def _autofit_pixels(text_px: int, size: float | None, font: str | None = None,
                    bold: bool = False, italic: bool = False) -> int:
    """The column Excel's AutoFit makes for text ``text_px`` pixels wide.

    Measured on a display at 100% scaling (Calibri: 11,786 cells, 8 to 36 pt;
    14 fonts in 4 styles: 3,696 cells, 8 to 40 pixels): the margin grows with
    the font's digit width (2 pixels at Calibri 11, ``_fonts.margin``), italic
    text gets room for its slant, and long text 1/28 more room.
    """
    family, pixels = _fonts.family_key(font), _pixel_size(size)
    slant = _fonts.italic_pixels(family, bold, italic, pixels)
    return _column_pixels(text_px + slant, _fonts.margin(_advances(bold, pixels, family, italic)))


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


_NUMBER_PLACEHOLDER = re.compile(r"[0#?][0#?,.]*(?:[eE][+-][0#]+)?")
_EXPONENT = re.compile(r"[eE][+-]([0#]+)$")


def _format_plan(number_format: str, measure) -> tuple | None:
    """``_formatted_number_text`` compiled for the native meter: what each
    section shows, with its literal text measured by ``measure(text)``.

    ``(sections, positive, negative or None)``; a section is ``(tag, a, b, c,
    d, e, f, g)``: 0 General, 1 constant text (a: pixels), 2 ``@`` text (a:
    pixels of the literal, b: General renderings), 3 a number (head and tail
    pixels, decimals, thousands, min digits, percent signs, exponent digits or
    -1). None when the native meter leaves the format to Python.
    """
    sections = _sections(number_format)
    chosen = [sections[0]]
    if len(sections) > 1 and sections[1]:
        chosen.append(sections[1])
    plans = []
    for section in chosen:
        section = _BRACKETS.sub(_currency, section)
        if section.strip().lower() in ("", "general"):
            plans.append((0, 0, 0, 0, 0, 0, 0, 0))
            continue
        if "@" in section and not re.search(r"[0#?]", _QUOTED.sub("", section)):
            # Every @ shows the General text; a NUL marks where, unless the
            # literal rules eat it (``*@``), which Python then measures.
            literal = _literal_text(section.replace("@", "\0"))
            if literal.count("\0") != section.count("@"):
                return None
            plans.append((2, measure(literal.replace("\0", "")), section.count("@"), 0, 0, 0, 0, 0))
            continue
        if _is_date_format(section):
            plans.append((1, measure(_date_text(section)), 0, 0, 0, 0, 0, 0))
            continue
        match = _NUMBER_PLACEHOLDER.search(section)
        if match is None:
            plans.append((1, measure(_literal_text(section)), 0, 0, 0, 0, 0, 0))
            continue
        placeholder = match.group(0)
        head = measure(_literal_text(section[:match.start()]))
        tail = measure(_literal_text(section[match.end():]))
        percent = _QUOTED.sub("", section).count("%")
        exponent = _EXPONENT.search(placeholder)
        if exponent:
            mantissa = placeholder[:exponent.start()]
            decimals = len(mantissa.split(".", 1)[1]) if "." in mantissa else 0
            plans.append((3, head, tail, decimals, 0, 0, percent, len(exponent.group(1))))
        else:
            integer_part, _, fraction = placeholder.partition(".")
            decimals = sum(fraction.count(c) for c in "0#?")
            thousands = "," in integer_part.strip(",")
            plans.append((3, head, tail, decimals, int(thousands), integer_part.count("0"), percent, -1))
    if any(None in plan for plan in plans):
        return None
    return len(sections) > 1, plans[0], plans[1] if len(plans) > 1 else None


def _shown(value, number_format: str | None) -> tuple[str | None, str | None]:
    """(the text a cell shows, the text AutoFit makes room for)."""
    value = plain_value(value)  # written as the Python value it holds (NaT, NA: empty)
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
               size: float | None = None, wrap: bool = False, font: str | None = None,
               italic: bool = False) -> float:
    """Column width that shows ``value`` whole (0.0 when it shows nothing).

    Excel's own AutoFit on a display at 100% scaling, for the fonts with
    metrics in ``_fonts`` (Calibri by default); other fonts are measured as a
    widened Calibri, and other scalings come out close, usually a little wider.
    """
    text = _shown(value, number_format)[1]
    if not text:
        return 0.0
    if wrap and "\n" in text:
        text_px = max(text_pixels(line, bold, size, font=font, italic=italic) for line in text.split("\n"))
    else:
        text_px = text_pixels(text.replace("\n", ""), bold, size, font=font, italic=italic)
    return _autofit_pixels(text_px, size, font, bold, italic) / _DIGIT_PX


def file_width(width: float) -> float:
    """A width in characters (Excel's ``ColumnWidth``) as stored in the file."""
    return width + _CELL_PADDING


def _indent_pixels(alignment: Alignment | None) -> int:
    """The room an alignment's indent takes in a column, in pixels."""
    if alignment is None or not alignment.indent or alignment.indent < 0:
        return 0
    sides = 2 if alignment.horizontal in _TWO_SIDED_INDENT else 1
    return sides * _INDENT_PX * alignment.indent


_PLAIN_PROPS = (None, False, None, False, 0, _DEFAULT_FAMILY, False)
_HEADER_PROPS = (None, True, None, False, 0, _DEFAULT_FAMILY, False)


def _style_props(style, cache: dict):
    """(number_format, bold, size, wrap, indent pixels, font family, italic) of
    a Style, cached by identity."""
    if style is None:
        return _PLAIN_PROPS
    cached = cache.get(id(style))
    if cached is not None and cached[0] is style:
        return cached[1]
    font, alignment = style.font, style.alignment
    props = (
        style.number_format,
        bool(style.bold or (font is not None and font.bold)),
        font.size if font is not None else None,
        bool(style.wrap or (alignment is not None and alignment.wrap_text)),
        _indent_pixels(alignment),
        _fonts.family_key(font.name if font is not None else None),
        bool(style.italic or (font is not None and font.italic)),
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

    With the native extension, ``fast_add(row, row_styles)`` measures a row
    that is neither the first row nor in ``slow_rows`` (the writers' hot path).
    """

    def __init__(self, styles_at=None, rows: int | None = 1000, button_rows=frozenset()):
        self.limit = math.inf if rows is None else rows
        self._styles_at = styles_at
        self._button_rows = frozenset(button_rows)
        self._best: list[int] = []  # column pixels
        self._cells: dict[int, dict[int, int]] = {}  # button row -> {column: pixels}
        self._style_cache: dict = {}
        self._fonts: dict = {}  # (family, bold, italic, size) -> (advances, pixels, margin, slant)
        self._constant: dict[tuple, int] = {}  # (type, number format, font) -> pixels of dates and booleans
        self._native = None
        self.fast_add = None
        self.slow_rows = self._button_rows | {0}
        if _NativeWidthMeter is not None:
            # Style props -> native kind; 0 plain, 1 bold (an unstyled header cell).
            self._kinds: dict = {_PLAIN_PROPS: 0, _HEADER_PROPS: 1}
            self._native = _NativeWidthMeter(
                self._native_fallback, self._kind_of, [bool, datetime.date, datetime.datetime],
                self._overlay_kind)
            for props, kind in self._kinds.items():
                self._native.add_kind(kind, self._kind_spec(props))
            self.fast_add = self._native.add

    def add(self, row_idx: int, row, row_styles=None) -> None:
        """Measure one row (the caller checks ``row_idx < meter.limit``).
        ``row_styles``: the row's ``{column: Style}``, from ``styles_at`` when None."""
        if row_styles is None:
            styles_at = self._styles_at
            row_styles = styles_at(row_idx, len(row)) if styles_at is not None else None
        native = self._native
        if native is not None and (not row_styles or type(row_styles) is dict):
            kept = native.add(row, row_styles or None, row_idx == 0, row_idx in self._button_rows)
            if kept is not None:
                self._cells[row_idx] = {col: width for col, width in enumerate(kept) if width}
            return
        best = self._best
        if len(row) > len(best):
            best.extend([0] * (len(row) - len(best)))
        kept = self._cells.setdefault(row_idx, {}) if row_idx in self._button_rows else None
        cell_pixels = self._cell_pixels
        for col_idx, cell in enumerate(row):
            if cell is None:
                continue
            width = cell_pixels(row_idx, cell, row_styles.get(col_idx) if row_styles else None)
            if not width:
                continue
            if width > best[col_idx]:
                best[col_idx] = width
            if kept is not None:
                kept[col_idx] = width

    def _font(self, family: str, bold: bool, italic: bool, size) -> tuple:
        """(advances, pixels, margin, slant pixels) of a font at ``size`` points."""
        key = (family, bold, italic, size)
        font = self._fonts.get(key)
        if font is None:
            pixels = _pixel_size(size)
            advances = _advances(bold, pixels, family, italic)
            font = self._fonts[key] = (advances, pixels, _fonts.margin(advances),
                                       _fonts.italic_pixels(family, bold, italic, pixels))
        return font

    def _rich_pixels(self, rich, family: str, bold: bool, italic: bool, size, wrap: bool) -> int:
        """Pixels of a rich text, the room for italic runs included.

        A plain ``str`` run is drawn in the cell's font; a ``TextBlock`` in
        its own font, whose unset fields are Excel's default (Calibri 11, not
        bold or italic), not the cell's. Excel's AutoFit gives each italic run
        half the room an italic cell gets (``_fonts.italic_pixels``); the
        cell's font sets the margin. Measured at 8 to 26 pt.
        """
        widest = line = slant = 0
        for run in rich:
            if isinstance(run, TextBlock):
                font, text = run.font, run.text
                run_bold, run_italic = bool(font.b), bool(font.i)
                advances, pixels, _margin, run_slant = self._font(
                    _fonts.family_key(font.rFont), run_bold, run_italic, font.sz)
            else:
                text, run_bold, run_italic = str(run), bold, italic
                advances, pixels, _margin, run_slant = self._font(family, bold, italic, size)
            if run_italic and text:
                slant += run_slant // 2
            parts = text.split("\n") if wrap else [text.replace("\n", "")]
            for index, part in enumerate(parts):
                if index:
                    widest, line = max(widest, line), 0
                line += _sum_pixels(part, advances, run_bold, pixels)
        text_px = max(widest, line)
        return text_px + slant if text_px else 0

    def _cell_pixels(self, row_idx: int, cell, style) -> int:
        """The column width in pixels one cell needs (0: none)."""
        number_format = None
        if type(cell) is tuple and _is_formatted_cell(cell):
            cell, overlay = cell
            if isinstance(overlay, str):
                number_format = overlay
            else:
                style = merge_styles(style, overlay)
        style_format, bold, size, wrap, indent, family, italic = _style_props(style, self._style_cache)
        number_format = number_format or style_format
        bold = bold or (row_idx == 0 and style is None)
        advances, pixels, margin, slant = self._font(family, bold, italic, size)
        kind = type(cell)
        if kind is str:
            text = cell
        elif kind is CellRichText:
            # The runs carry their own italic room; an italic cell adds half its own.
            text_px = self._rich_pixels(cell, family, bold, italic, size, wrap)
            slant //= 2
            text = None
        elif kind is bool or kind is datetime.date or kind is datetime.datetime:
            key = (kind, number_format, family, bold, italic, size)
            known = self._constant.get(key)
            if known is None:
                known = self._constant[key] = _sum_pixels(
                    _shown(cell, number_format)[1] or "", advances, bold, pixels)
            text_px = known
            text = None
        elif (kind is int or kind is float) and (number_format is None or number_format == "General"):
            text = _general_room(cell)
        else:
            text = _shown(cell, number_format)[1]
            if text is None:  # a formula without a cached value
                return 0
        if text is not None:
            if not text:
                return 0
            if "\n" in text:
                if wrap:
                    text_px = max(_sum_pixels(line, advances, bold, pixels) for line in text.split("\n"))
                else:
                    text_px = _sum_pixels(text.replace("\n", ""), advances, bold, pixels)
            else:
                text_px = _sum_pixels(text, advances, bold, pixels)
        if not text_px:
            return 0
        return _column_pixels(text_px + indent + slant, margin)

    # --- the native meter (rows past the first, see autofit.rs) ---

    def _native_fallback(self, cell, style, header: bool) -> int:
        """A cell the native meter does not measure itself."""
        return self._cell_pixels(0 if header else -1, cell, style)

    def _kind_of(self, style):
        """(kind id, its spec if the native meter has not seen it) of a Style."""
        return self._kind_for(_style_props(style, self._style_cache))

    def _overlay_kind(self, style, overlay, header: bool):
        """The kind of a ``(value, format)`` or ``(value, Style)`` cell under a
        column style, as ``_cell_pixels`` applies it."""
        if isinstance(overlay, str):
            number_format, bold, size, wrap, indent, family, italic = _style_props(style, self._style_cache)
            props = (overlay or number_format, bold or (header and style is None), size, wrap, indent, family, italic)
        else:
            props = _style_props(merge_styles(style, overlay), self._style_cache)
        return self._kind_for(props)

    def _kind_for(self, props):
        kind = self._kinds.get(props)
        if kind is not None:
            return kind, None
        kind = self._kinds[props] = len(self._kinds)
        return kind, self._kind_spec(props)

    def _kind_spec(self, props) -> tuple:
        number_format, bold, size, wrap, indent, family, italic = props
        advances, pixels, margin, slant = self._font(family, bold, italic, size)
        general = not number_format or number_format.strip().lower() == "general"

        def char_px(char: str) -> int:
            return _char_pixels(char, bold, pixels, advances)

        def measure(text: str) -> int | None:
            # Text on several lines is measured line by line: left to Python.
            return None if "\n" in text else _sum_pixels(text, advances, bold, pixels)

        plan = None
        if not general:
            try:
                plan = _format_plan(number_format, measure)
            except (ValueError, KeyError, IndexError):
                plan = None
        return advances, margin, general, wrap, char_px, plan, indent + slant

    def widths(self, max_cols: int, filter_cells=(), padding: float = 0.0) -> list[float]:
        """Widths of the first ``max_cols`` columns in characters, never below
        Excel's default. ``filter_cells`` holds ``(row, first_col, last_col)``
        header spans whose cells carry a filter drop-down button; ``padding``
        (characters) is added to every measured column."""
        best = (self._best + [0] * max_cols)[:max_cols]
        if self._native is not None:
            for col, px in enumerate(self._native.best()[:max_cols]):
                if px > best[col]:
                    best[col] = px
        button = _FILTER_BUTTON_PX
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
