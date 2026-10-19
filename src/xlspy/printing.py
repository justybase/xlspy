"""Page setup for printing an XLSX worksheet."""

from __future__ import annotations

from dataclasses import dataclass

from .refs import column_label, parse_cell_range, quote_sheet_name

ORIENTATIONS = frozenset({"portrait", "landscape"})

# Paper size names and the OOXML ``paperSize`` codes.
PAPER_SIZES = {
    "letter": 1,
    "tabloid": 3,
    "ledger": 4,
    "legal": 5,
    "statement": 6,
    "executive": 7,
    "a3": 8,
    "a4": 9,
    "a4_small": 10,
    "a5": 11,
    "b4": 12,
    "b5": 13,
    "folio": 14,
    "quarto": 15,
    "envelope_10": 20,
    "envelope_dl": 27,
    "envelope_c5": 28,
    "envelope_monarch": 37,
}
_MAX_HEADER_LENGTH = 255
_MAX_BREAK_INDEX = 1048576


@dataclass(frozen=True, slots=True)
class PageMargins:
    """Page margins in inches. Defaults match the historical xlspy layout."""

    left: float = 0.7
    right: float = 0.7
    top: float = 0.75
    bottom: float = 0.75
    header: float = 0.3
    footer: float = 0.3

    def __post_init__(self):
        for name in ("left", "right", "top", "bottom", "header", "footer"):
            value = getattr(self, name)
            if isinstance(value, bool) or not isinstance(value, (int, float)):
                raise TypeError(f"margin {name} must be a number of inches")
            if value < 0 or value > 50:
                raise ValueError(f"margin {name} must be between 0 and 50 inches")


def _paper_code(value) -> int:
    if isinstance(value, bool):
        raise TypeError("paper_size must be a name such as 'A4' or an int code")
    if isinstance(value, int):
        if not 1 <= value <= 255:
            raise ValueError("paper_size code must be between 1 and 255")
        return value
    if isinstance(value, str):
        key = value.lower().replace(" ", "_").replace("-", "_")
        if key in PAPER_SIZES:
            return PAPER_SIZES[key]
        raise ValueError(f"unknown paper_size {value!r}; expected one of {sorted(PAPER_SIZES)}")
    raise TypeError("paper_size must be a name such as 'A4' or an int code")


def _check_optional_int(value, name: str, minimum: int, maximum: int) -> None:
    if value is None:
        return
    if isinstance(value, bool) or not isinstance(value, int):
        raise TypeError(f"{name} must be an int or None")
    if not minimum <= value <= maximum:
        raise ValueError(f"{name} must be between {minimum} and {maximum}")


def _check_header(value, name: str) -> None:
    if value is None:
        return
    if not isinstance(value, str):
        raise TypeError(f"{name} must be a string or None")
    if len(value) > _MAX_HEADER_LENGTH:
        raise ValueError(f"{name} must be at most {_MAX_HEADER_LENGTH} characters")


def _check_breaks(values, name: str, limit: int) -> tuple[int, ...] | None:
    if values is None:
        return None
    result = []
    for value in values:
        if isinstance(value, bool) or not isinstance(value, int) or not 1 <= value < limit:
            raise ValueError(f"{name} entries must be positive ints below {limit}")
        result.append(value)
    return tuple(sorted(set(result)))


def _normalize_print_area(value: str) -> str:
    if not isinstance(value, str) or not value:
        raise TypeError("print_area must be an A1 range such as 'A1:H100'")
    for part in value.split(","):
        parse_cell_range(part)
    return value


def _normalize_repeat_rows(value: str) -> str:
    if not isinstance(value, str):
        raise TypeError("repeat_rows must be a row range such as '1:2'")
    first, _, last = value.partition(":")
    if not first.isdigit() or not last.isdigit() or int(first) < 1 or int(first) > int(last):
        raise ValueError(f"repeat_rows must look like '1:2', got {value!r}")
    return value


def _normalize_repeat_columns(value: str) -> str:
    if not isinstance(value, str):
        raise TypeError("repeat_columns must be a column range such as 'A:B'")
    first, _, last = value.partition(":")
    if not first.isalpha() or not last.isalpha():
        raise ValueError(f"repeat_columns must look like 'A:B', got {value!r}")
    parse_cell_range(f"{first.upper()}1:{last.upper()}1")
    return f"{first.upper()}:{last.upper()}"


@dataclass(frozen=True, slots=True)
class PrintSettings:
    """Print layout for one worksheet.

    ``fit_to_width`` and ``fit_to_height`` switch the sheet to fit-to-page
    scaling; ``fit_to_height=0`` lets the page count follow the content.
    Otherwise ``scale`` (10-400 percent) applies. ``header`` and ``footer``
    are raw Excel header codes such as ``"&LReport&RPage &P of &N"``; write
    a literal ampersand as ``&&``. ``row_breaks`` and ``col_breaks`` name the
    row or column after which a manual page break is placed (1-based).
    """

    orientation: str | None = None
    paper_size: str | int | None = None
    margins: PageMargins | None = None
    print_area: str | None = None
    repeat_rows: str | None = None
    repeat_columns: str | None = None
    fit_to_width: int | None = None
    fit_to_height: int | None = None
    scale: int | None = None
    center_horizontally: bool = False
    center_vertically: bool = False
    header: str | None = None
    footer: str | None = None
    row_breaks: tuple[int, ...] | list[int] | None = None
    col_breaks: tuple[int, ...] | list[int] | None = None

    def __post_init__(self):
        if self.orientation is not None and self.orientation not in ORIENTATIONS:
            raise ValueError(f"orientation must be 'portrait' or 'landscape', got {self.orientation!r}")
        if self.paper_size is not None:
            _paper_code(self.paper_size)
        if self.margins is not None and not isinstance(self.margins, PageMargins):
            raise TypeError("margins must be a PageMargins or None")
        if self.print_area is not None:
            object.__setattr__(self, "print_area", _normalize_print_area(self.print_area))
        if self.repeat_rows is not None:
            object.__setattr__(self, "repeat_rows", _normalize_repeat_rows(self.repeat_rows))
        if self.repeat_columns is not None:
            object.__setattr__(
                self, "repeat_columns", _normalize_repeat_columns(self.repeat_columns),
            )
        _check_optional_int(self.fit_to_width, "fit_to_width", 0, 32767)
        _check_optional_int(self.fit_to_height, "fit_to_height", 0, 32767)
        _check_optional_int(self.scale, "scale", 10, 400)
        if self.scale is not None and self.fit_to_page:
            raise ValueError("scale cannot be combined with fit_to_width or fit_to_height")
        for name in ("center_horizontally", "center_vertically"):
            if not isinstance(getattr(self, name), bool):
                raise TypeError(f"{name} must be a bool")
        _check_header(self.header, "header")
        _check_header(self.footer, "footer")
        object.__setattr__(self, "row_breaks", _check_breaks(self.row_breaks, "row_breaks", _MAX_BREAK_INDEX))
        object.__setattr__(self, "col_breaks", _check_breaks(self.col_breaks, "col_breaks", 16384))

    @property
    def fit_to_page(self) -> bool:
        return self.fit_to_width is not None or self.fit_to_height is not None


def print_area_name(sheet_name: str, area: str) -> str:
    """Defined-name value for a print area, one absolute reference per part."""
    parts = []
    for part in area.split(","):
        first, _, last = part.partition(":")
        first_col, first_row, last_col, last_row = parse_cell_range(part)
        cell = f"${column_label(first_col)}${first_row + 1}"
        end = f"${column_label(last_col)}${last_row + 1}"
        parts.append(f"{quote_sheet_name(sheet_name)}!{cell}" + (f":{end}" if last else ""))
    return ",".join(parts)


def print_titles_name(sheet_name: str, settings: PrintSettings) -> str | None:
    """Defined-name value for repeating rows and columns, or None."""
    parts = []
    quoted = quote_sheet_name(sheet_name)
    if settings.repeat_rows is not None:
        first, _, last = settings.repeat_rows.partition(":")
        parts.append(f"{quoted}!${first}:${last}")
    if settings.repeat_columns is not None:
        first, _, last = settings.repeat_columns.partition(":")
        parts.append(f"{quoted}!${first}:${last}")
    return ",".join(parts) if parts else None
