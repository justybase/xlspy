"""Cell, merged cell and write-only cell."""

from __future__ import annotations
from ..exceptions import UnsupportedFeatureError
from dataclasses import replace
import datetime as _datetime
from decimal import Decimal
from typing import TYPE_CHECKING, Any
from ..comments import Comment
from ..formats import DEFAULT_DATE_FORMAT, DEFAULT_DATETIME_FORMAT
from ..formulas import ArrayFormula, DataTableFormula, Formula
from ..hyperlinks import Hyperlink
from ..refs import column_label
from ..rich_text import CellRichText
from ..styles import (
    Alignment,
    Border,
    Font,
    GradientFill,
    NamedStyle,
    PatternFill,
    Protection,
    Style,
)
from ._coords import _coerce_column, _coerce_row, _coordinate
from .. import _accel

if TYPE_CHECKING:
    from ._worksheet import Worksheet


_new_cell = object.__new__
_CELL_TYPE_KINDS = {int: "n", float: "n", bool: "b", str: "s"}


class Cell:
    """One cell of a worksheet: its value, formula, style, note and hyperlink.

    Get cells from the worksheet (``ws["A1"]``, ``Worksheet.cell``); they are not
    created directly.
    """
    __slots__ = (
        "_sheet", "_row", "_column", "_value", "_data_type",
        "_style_id", "_comment", "_hyperlink", "_formula",
        "_named_style",
    )

    def __init__(self, sheet: Worksheet, row: int, column: int, value: Any = None, style_id: int = 0) -> None:
        self._sheet = sheet
        self._row = _coerce_row(row)
        self._column = _coerce_column(column)
        self._value: Any = None
        self._formula: Formula | ArrayFormula | DataTableFormula | str | None = None
        self._data_type = "n"
        self._style_id = int(style_id)
        self._comment: Comment | None = None
        self._hyperlink: Hyperlink | None = None
        self._named_style: str | None = None
        if value is not None:
            self.value = value

    # -- identity -----------------------------------------------------
    @property
    def row(self) -> int:
        """The 1-based row number."""
        return self._row

    @property
    def column(self) -> int:
        """The 1-based column number."""
        return self._column

    @property
    def column_letter(self) -> str:
        """The column letter, such as ``"C"``."""
        return column_label(self._column - 1)

    @property
    def col_idx(self) -> int:
        """The 1-based column number, as openpyxl names it."""
        return self._column

    @property
    def coordinate(self) -> str:
        """The cell reference, such as ``"C5"``."""
        return _coordinate(self._row, self._column)

    @property
    def parent(self) -> Worksheet:
        """The worksheet that holds the cell."""
        return self._sheet

    def __repr__(self) -> str:
        return f"<Cell {self._sheet.title!r}.{self.coordinate}>"

    # -- value --------------------------------------------------------
    @property
    def value(self) -> Any:
        """The cell value; a scalar formula reads back as ``"=..."`` text.

        An ``ArrayFormula`` or ``DataTableFormula`` reads back as that object, as
        in openpyxl. Assigning text that starts with ``=`` or a Formula sets a
        formula; an ``ArrayFormula`` or ``DataTableFormula`` sets that formula
        object, and a Hyperlink sets the link and shows its text; any other value
        keeps the cell's link, as in Excel (``hyperlink = None`` removes it). A
        date, time or timedelta gets a default number format when the cell's
        format is ``General``.
        """
        if isinstance(self._formula, Formula):
            return "=" + self._formula.text
        if isinstance(self._formula, (ArrayFormula, DataTableFormula)):
            # openpyxl returns these objects from Cell.value; their text is in
            # ArrayFormula.text, and DataTableFormula has no text at all.
            return self._formula
        if isinstance(self._formula, str):
            return "=" + self._formula
        return self._value

    @value.setter
    def value(self, value: Any) -> None:
        plain = _CELL_TYPE_KINDS.get(type(value))
        if plain is not None and not (plain == "s" and value.startswith("=")):
            # Exact int/float/bool/str: no formula, hyperlink or date handling.
            self._formula = None
            self._value = value
            self._data_type = plain
            sheet = self._sheet
            if sheet._workbook._loaded:
                sheet._dirty_cells.add((self._row, self._column))
            sheet._dirty = True
            return
        if isinstance(value, Cell):
            value = value.value
        if isinstance(value, (ArrayFormula, DataTableFormula)):
            self._formula = value
            self._value = None
            self._data_type = "f"
        elif isinstance(value, Formula):
            self._formula = value
            self._value = None
            self._data_type = "f"
        elif isinstance(value, str) and value.startswith("="):
            self._formula = Formula(value[1:])
            self._value = None
            self._data_type = "f"
        else:
            self._formula = None
            if isinstance(value, Hyperlink):
                self._hyperlink = value
                self._value = value.display_text
                self._data_type = "s"
                self._sheet._dirty_hyperlinks.add(self.coordinate)
            else:
                self._value = value
                self._data_type = self._type_for(value)
        if isinstance(value, _datetime.time) and self.number_format == "General":
            self.number_format = "h:mm:ss"
        elif isinstance(value, _datetime.timedelta) and self.number_format == "General":
            self.number_format = "[h]:mm:ss"
        elif isinstance(value, _datetime.date) and self.number_format == "General":
            # The writers' defaults; a General date would show its serial.
            self.number_format = (
                DEFAULT_DATETIME_FORMAT if isinstance(value, _datetime.datetime)
                else DEFAULT_DATE_FORMAT
            )
        self._mark_dirty()

    @property
    def formula(self) -> Formula | ArrayFormula | DataTableFormula | str | None:
        """The formula object or text without its leading ``=``, or None."""
        return self._formula

    @property
    def data_type(self) -> str:
        """openpyxl's type code: ``n`` number or empty, ``s`` text, ``b`` boolean,
        ``d`` date or time, ``f`` formula, ``e`` error.
        """
        return self._data_type

    @property
    def internal_value(self) -> Any:
        """The stored value; None for a formula cell that holds no cached result."""
        return self._value

    @property
    def is_date(self) -> bool:
        """True when the value is a date, datetime, time or timedelta."""
        return isinstance(
            self._value,
            (_datetime.datetime, _datetime.date, _datetime.time, _datetime.timedelta),
        )

    @property
    def encoding(self) -> str:
        """Always ``"utf-8"``, for openpyxl compatibility."""
        return "utf-8"

    @property
    def base_date(self) -> _datetime.datetime:
        """The workbook's ``epoch``, the base date of serial date numbers (1899-12-30)."""
        return self._sheet._workbook.epoch

    @property
    def hyperlink(self) -> Hyperlink | None:
        """The cell's Hyperlink, or None."""
        return self._hyperlink

    @hyperlink.setter
    def hyperlink(self, value: Hyperlink | None) -> None:
        if value is not None and not isinstance(value, Hyperlink):
            raise TypeError("hyperlink must be a xlspy Hyperlink or None")
        self._hyperlink = value
        self._sheet._dirty_hyperlinks.add(self.coordinate)
        self._sheet._mark_dirty("hyperlinks")

    @property
    def comment(self) -> Comment | None:
        """The cell's note (a Comment), or None."""
        return self._comment

    @comment.setter
    def comment(self, value: Comment | None) -> None:
        if value is not None and not isinstance(value, Comment):
            raise TypeError("comment must be a xlspy Comment or None")
        self._comment = value
        self._sheet._dirty_comments.add(self.coordinate)
        self._sheet._mark_dirty("comments")

    @staticmethod
    def _type_for(value) -> str:
        if value is None:
            return "n"
        if isinstance(value, bool):
            return "b"
        if isinstance(value, (int, float, Decimal)) and not isinstance(value, bool):
            return "n"
        if isinstance(value, (_datetime.datetime, _datetime.date, _datetime.time, _datetime.timedelta)):
            return "d"
        if isinstance(value, (str, CellRichText)):
            return "s"
        if isinstance(value, bytes):
            return "s"
        raise TypeError(
            f"unsupported cell value {type(value).__name__}; use a number, string, "
            "date/time, bool, Formula, Hyperlink or None"
        )

    # -- style --------------------------------------------------------
    @property
    def style_id(self) -> int:
        """The index of the cell's style in the workbook's style table."""
        return self._style_id

    @property
    def has_style(self) -> bool:
        """True when the cell has a style other than the workbook default."""
        return bool(self._style_id)

    def _style(self) -> Style:
        return self._sheet._workbook._styles.style_at(self._style_id)

    def _set_style(self, style: Style) -> None:
        self._style_id = self._sheet._workbook._styles.add(style)
        self._sheet._workbook._styles_dirty = True
        self._sheet._dirty_styles.add(self.coordinate)
        self._sheet._mark_dirty("styles")

    @property
    def font(self) -> Font | None:
        """The cell font, or None. Assigning one replaces the font of the cell's style."""
        return self._style().font

    @font.setter
    def font(self, value: Font | None) -> None:
        if value is not None and not isinstance(value, Font):
            raise TypeError("font must be a xlspy Font or None")
        self._set_style(replace(self._style(), font=value, bold=None, italic=None,
                                underline=None, font_color=None))

    @property
    def fill(self) -> str | PatternFill | GradientFill | None:
        """The cell fill: a colour, a PatternFill or a GradientFill, or None."""
        return self._style().fill

    @fill.setter
    def fill(self, value: str | PatternFill | GradientFill | None) -> None:
        if value is not None and not isinstance(value, (str, PatternFill, GradientFill)):
            raise TypeError("fill must be a color, PatternFill, GradientFill or None")
        self._set_style(replace(self._style(), fill=value, pattern=None))

    @property
    def border(self) -> Border | None:
        """The cell border, or None."""
        return self._style().border

    @border.setter
    def border(self, value: Border | None) -> None:
        if value is not None and not isinstance(value, Border):
            raise TypeError("border must be a xlspy Border or None")
        self._set_style(replace(self._style(), border=value))

    @property
    def alignment(self) -> Alignment | None:
        """The cell alignment, or None."""
        return self._style().alignment

    @alignment.setter
    def alignment(self, value: Alignment | None) -> None:
        if value is not None and not isinstance(value, Alignment):
            raise TypeError("alignment must be a xlspy Alignment or None")
        self._set_style(replace(self._style(), alignment=value,
                                horizontal=None, vertical=None, wrap=None))

    @property
    def protection(self) -> Protection | None:
        """The cell protection (locked, hidden), or None."""
        return self._style().protection

    @protection.setter
    def protection(self, value: Protection | None) -> None:
        if value is not None and not isinstance(value, Protection):
            raise TypeError("protection must be a xlspy Protection or None")
        self._set_style(replace(self._style(), protection=value))

    @property
    def number_format(self) -> str:
        """The number format code; ``"General"`` when none is set."""
        fmt = self._style().number_format
        return "General" if fmt is None else fmt

    @number_format.setter
    def number_format(self, value: str) -> None:
        if not isinstance(value, str):
            raise TypeError("number_format must be a string")
        workbook = self._sheet._workbook
        styles = workbook._styles
        cache = styles._number_format_ids
        key = (self._style_id, value)
        style_id = cache.get(key)
        if style_id is None:
            style_id = cache[key] = styles.add(replace(
                self._style(),
                number_format=None if value == "General" else value,
            ))
        self._style_id = style_id
        workbook._styles_dirty = True
        self._sheet._dirty_styles.add(self.coordinate)
        self._sheet._mark_dirty("styles")

    @property
    def style(self) -> Style:
        """The cell's Style.

        Assign a Style, or the name of a named style registered with
        ``Workbook.add_named_style``.
        """
        return self._style()

    @style.setter
    def style(self, value: Style | NamedStyle | str) -> None:
        if isinstance(value, str):
            named = None
            for entry in self._sheet._workbook.named_styles:
                if (entry if isinstance(entry, str) else entry.name) == value:
                    named = entry
                    break
            if named is None:
                raise UnsupportedFeatureError(
                    "named styles are not implemented yet; "
                    f"{value!r} is not defined, add it with add_named_style()"
                )
            if isinstance(named, NamedStyle):
                current = self._style()
                self._set_style(Style(
                    font=named.font if named.font is not None else current.font,
                    fill=named.fill if named.fill is not None else current.fill,
                    border=named.border if named.border is not None else current.border,
                    alignment=named.alignment if named.alignment is not None else current.alignment,
                    protection=named.protection if named.protection is not None else current.protection,
                    number_format=(
                        named.number_format if named.number_format is not None
                        else current.number_format
                    ),
                ))
            self._named_style = value
            return
        if not isinstance(value, Style):
            raise TypeError("style must be a xlspy Style or a named style")
        self._named_style = None
        self._set_style(value)

    def _load_value(self, value, data_type: str, formula_text: str | None = None):
        """Install a parsed value without importing/validating arbitrary formulas."""
        self._value = value
        self._formula = formula_text
        self._data_type = data_type

    def _mark_dirty(self):
        if self._sheet._workbook._loaded:
            self._sheet._dirty_cells.add((self._row, self._column))
        self._sheet._mark_dirty()

    # -- convenience --------------------------------------------------
    def offset(self, row: int = 0, column: int = 0) -> Cell:
        """The cell ``row`` rows below and ``column`` columns right of this one.

        Negative offsets go up and left.
        """
        return self._sheet.cell(self._row + row, self._column + column)


# Slot descriptors, in the order the native loaders fill them (see rust/xlspy_core/src/model.rs).
CELL_SLOTS = tuple(Cell.__dict__[name] for name in (
    "_sheet", "_row", "_column", "_value", "_formula", "_data_type",
    "_style_id", "_comment", "_hyperlink", "_named_style",
))
# The model's hot loops, or None: _append_new_row and workbook/_writer.py then do them.
NATIVE_CELLS = None if _accel.ModelCells is None else _accel.ModelCells(Cell, CELL_SLOTS)


class MergedCell(Cell):
    """A cell inside a merged range other than its top-left anchor."""

    @Cell.value.setter  # type: ignore[attr-defined]  # mypy does not model a property setter on the class
    def value(self, value: Any) -> None:
        raise AttributeError("read-only cell (merged range)")


class WriteOnlyCell:
    """A value with optional style, used with ``WriteOnlyWorksheet.append``."""

    def __init__(self, worksheet: Any, value: Any = None) -> None:
        self.parent = worksheet
        self.value = value
        self.font = None
        self.fill = None
        self.border = None
        self.alignment = None
        self.protection = None
        self.number_format = "General"
        self.comment = None
        self.hyperlink = None

    def __repr__(self) -> str:
        return f"<WriteOnlyCell {self.value!r}>"
