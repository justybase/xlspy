"""Worksheet, write-only worksheet and chartsheet."""

from __future__ import annotations
from ..exceptions import UnsupportedFeatureError
from collections.abc import Iterable, Iterator, Mapping
from dataclasses import replace
import datetime as _datetime
from typing import TYPE_CHECKING, Any, Literal, overload
from ..charts import Chart
from ..comments import Comment
from ..data_validation import DataValidation
from ..formats import DEFAULT_DATE_FORMAT, DEFAULT_DATETIME_FORMAT
from ..images import Image
from ..refs import parse_cell_range, quote_sheet_name
from ..styles import (
    ExcelTable,
)
from . import _shift
from ._print import (
    HeaderFooter,
    WorksheetPageMargins,
    PageSetup,
    PrintOptions,
)
from ._cell import Cell, MergedCell, WriteOnlyCell, _CELL_TYPE_KINDS, _new_cell
from ._coords import MultiCellRange, _CELL_RE, _coerce_column, _coerce_row, _coordinate, _normalize_range
from ._dimensions import ColumnDimension, DimensionHolder, RowDimension, _column_key, _row_key
from ._properties import AutoFilter, ColBreak, Pane, RowBreak, Selection, SheetFormatProperties, SheetProperties, SheetView, WorksheetProtection, _as_model_page_margins, _as_model_protection
from ._tables import ConditionalFormattingList, DataValidationList, Table, TableList, _expect

if TYPE_CHECKING:
    from .._metadata import SheetProtection
    from ..printing import PageMargins
    from ._print import HeaderFooterItem
    from ._workbook import Workbook


class Worksheet:
    ORIENTATION_PORTRAIT = "portrait"
    ORIENTATION_LANDSCAPE = "landscape"
    PAPERSIZE_LETTER = 1
    PAPERSIZE_LEGAL = 5
    PAPERSIZE_A3 = 8
    PAPERSIZE_A4 = 9

    # Set by the XLSB loader (workbook/_xlsb_reader.py) on loaded sheets only.
    _cf_extensions: Any
    _saved_filter_records: Any
    _reader: Any
    _saved_print_settings: Any

    def __init__(self, workbook: Workbook, title: str, *, index: int | None = None) -> None:
        self._workbook = workbook
        self._title = title
        self._index = index
        self._cells: dict[int, dict[int, Cell]] = {}
        self._max_row = 0
        self._max_column = 0
        self._merged_cells = MultiCellRange()
        self._column_dimensions = DimensionHolder(ColumnDimension, _column_key)
        self._row_dimensions = DimensionHolder(RowDimension, _row_key)
        self._freeze_panes: str | None = None
        self._auto_filter = AutoFilter()
        self._sheet_state = "visible"
        self._sheet_properties = SheetProperties()
        self._sheet_view = SheetView(showGridLines=True)
        self._dirty = False
        self._dirty_cells: set[tuple[int, int]] = set()
        self._dirty_styles: set[str] = set()
        self._dirty_comments: set[str] = set()
        self._dirty_hyperlinks: set[str] = set()
        self._dirty_meta: set[str] = set()
        self._loaded = False
        self._original_title = title
        self._meta_signature = None
        self._path: str | None = None
        self._row_heights: dict[int, float] = {}
        self._hidden_rows: set[int] = set()
        self._data_validations = DataValidationList()
        self._tables = TableList()
        self._charts: list[Chart] = []
        self._images: list[Image] = []
        self._conditional_formats = ConditionalFormattingList()
        # For a loaded workbook: how many objects of each list a save already
        # put into the package (see _objects.pending_objects).
        self._saved_object_counts: dict[str, int] = {}
        # The print setup a loaded sheet already has in the package, so saving
        # does not rewrite a sheet whose print settings were not edited.
        self._saved_print_settings = None
        self._print_settings = None
        self.page_setup = PageSetup()
        self.page_margins = WorksheetPageMargins()
        self.print_options = PrintOptions()
        self.HeaderFooter = HeaderFooter()
        self.sheet_format = SheetFormatProperties()
        self.protection = WorksheetProtection()
        self.row_breaks = RowBreak()
        self.col_breaks = ColBreak()
        self.ignored_errors: list = []
        self._print_area: str | None = None
        self._print_title_rows: str | None = None
        self._print_title_cols: str | None = None
        self._comments: dict[str, Comment] = {}
        self._legacy_drawing = None

    def __repr__(self) -> str:
        return f"<Worksheet {self._title!r}>"

    @property
    def title(self) -> str:
        return self._title

    @title.setter
    def title(self, value: str) -> None:
        if value != self._title:
            self._workbook._check_title(value, self)
            self._workbook._rename_sheet(self, value)
            self._title = value
            self._mark_dirty("title")

    @property
    def sheet_state(self) -> str:
        return self._sheet_state

    @sheet_state.setter
    def sheet_state(self, value: str) -> None:
        if value not in {"visible", "hidden", "veryHidden"}:
            raise ValueError("sheet_state must be visible, hidden or veryHidden")
        self._sheet_state = value
        self._mark_dirty("sheet_state")

    @property
    def merged_cells(self) -> MultiCellRange:
        return self._merged_cells

    @property
    def column_dimensions(self) -> DimensionHolder:
        return self._column_dimensions

    @property
    def row_dimensions(self) -> DimensionHolder:
        return self._row_dimensions

    @property
    def freeze_panes(self) -> str | None:
        return self._freeze_panes

    @freeze_panes.setter
    def freeze_panes(self, value: str | None) -> None:
        if value == "":
            value = None
        if value is not None:
            if not isinstance(value, str):
                raise TypeError("freeze_panes must be an A1 cell or None")
            col, row, c2, r2 = parse_cell_range(value)
            if (col, row) != (c2, r2):
                raise ValueError("freeze_panes must be a single cell")
            value = _coordinate(row + 1, col + 1)
        self._freeze_panes = value
        if value is None:
            self._sheet_view.pane = None
            self._sheet_view.selection = []
        else:
            col, row, _, _ = parse_cell_range(value)
            if col and row:
                active = "bottomRight"
            elif col:
                active = "topRight"
            else:
                active = "bottomLeft"
            self._sheet_view.pane = Pane(
                xSplit=col if col else None,
                ySplit=row if row else None,
                topLeftCell=value,
                activePane=active,
                state="frozen",
            )
            self._sheet_view.selection = [
                Selection(pane=active, activeCell=value, sqref=value)
            ]
        self._mark_dirty("freeze_panes")

    @property
    def auto_filter(self) -> AutoFilter:
        return self._auto_filter

    @auto_filter.setter
    def auto_filter(self, value: AutoFilter | str | None) -> None:
        if value is None:
            self._auto_filter = AutoFilter()
        elif isinstance(value, AutoFilter):
            self._auto_filter = value
        elif isinstance(value, str):
            self._auto_filter = AutoFilter(_normalize_range(value))
        else:
            raise TypeError("auto_filter must be an AutoFilter, A1 range or None")
        self._mark_dirty("auto_filter")

    @property
    def sheet_properties(self) -> SheetProperties:
        return self._sheet_properties

    @property
    def sheet_view(self) -> SheetView:
        return self._sheet_view

    @property
    def active_cell(self) -> str | None:
        if not self._sheet_view.selection:
            return None
        return self._sheet_view.selection[0].activeCell

    @property
    def selected_cell(self) -> str | None:
        if not self._sheet_view.selection:
            return None
        return self._sheet_view.selection[0].sqref

    @property
    def show_gridlines(self) -> bool | None:
        return self._sheet_view.showGridLines

    @show_gridlines.setter
    def show_gridlines(self, value: bool | None) -> None:
        self._sheet_view.showGridLines = value
        self._mark_dirty("sheet_view")

    @property
    def protection(self) -> WorksheetProtection:
        return self._protection

    @protection.setter
    def protection(self, value: WorksheetProtection | SheetProtection | None) -> None:
        self._protection = _as_model_protection(value)

    @property
    def page_margins(self) -> WorksheetPageMargins:
        return self._page_margins

    @page_margins.setter
    def page_margins(self, value: WorksheetPageMargins | PageMargins) -> None:
        self._page_margins = _as_model_page_margins(value)

    @property
    def max_row(self) -> int:
        return self._max_row

    @property
    def max_column(self) -> int:
        return self._max_column

    @property
    def min_row(self) -> int:
        return 1

    @property
    def min_column(self) -> int:
        return 1

    @property
    def dimensions(self) -> str:
        if self._max_row == 0 or self._max_column == 0:
            return "A1:A1"
        return f"A1:{_coordinate(self._max_row, self._max_column)}"

    def calculate_dimension(self, force: bool = False) -> str:
        return self.dimensions

    @property
    def rows(self) -> Iterator[tuple[Cell, ...]]:
        return self.iter_rows()

    @property
    def columns(self) -> Iterator[tuple[Cell, ...]]:
        return self.iter_cols()

    @property
    def values(self) -> Iterator[tuple[Any, ...]]:
        return self.iter_rows(values_only=True)

    @property
    def data_validations(self) -> DataValidationList:
        return self._data_validations

    @property
    def tables(self) -> TableList:
        return self._tables

    @property
    def charts(self) -> list[Chart]:
        return self._charts

    @property
    def images(self) -> list[Image]:
        return self._images

    @property
    def conditional_formatting(self) -> ConditionalFormattingList:
        return self._conditional_formats

    # -- printing (openpyxl names) --------------------------------------
    @property
    def oddHeader(self) -> HeaderFooterItem:
        return self.HeaderFooter.oddHeader

    @property
    def oddFooter(self) -> HeaderFooterItem:
        return self.HeaderFooter.oddFooter

    @property
    def evenHeader(self) -> HeaderFooterItem:
        return self.HeaderFooter.evenHeader

    @property
    def evenFooter(self) -> HeaderFooterItem:
        return self.HeaderFooter.evenFooter

    @property
    def firstHeader(self) -> HeaderFooterItem:
        return self.HeaderFooter.firstHeader

    @property
    def firstFooter(self) -> HeaderFooterItem:
        return self.HeaderFooter.firstFooter

    @property
    def print_area(self) -> str | None:
        return self._print_area

    @print_area.setter
    def print_area(self, value: str | Iterable[str] | None) -> None:
        """An A1 range, several separated by commas, or a list of ranges."""
        if value is not None and not isinstance(value, str):
            value = ",".join(value)
        self._print_area = value

    @property
    def print_title_rows(self) -> str | None:
        """Rows repeated at the top of every printed page, such as ``"1:2"``."""
        return self._print_title_rows

    @print_title_rows.setter
    def print_title_rows(self, value: str | None) -> None:
        self._print_title_rows = value

    @property
    def print_title_cols(self) -> str | None:
        """Columns repeated at the left of every printed page, such as ``"A:B"``."""
        return self._print_title_cols

    @print_title_cols.setter
    def print_title_cols(self, value: str | None) -> None:
        self._print_title_cols = value

    @property
    def print_titles(self) -> str:
        """The print titles as openpyxl formats them, e.g. ``'Sheet'!$1:$2``."""
        title = quote_sheet_name(self.title)
        values = []
        if self._print_title_rows:
            values.append("$" + self._print_title_rows.replace(":", ":$"))
        if self._print_title_cols:
            values.append("$" + self._print_title_cols.replace(":", ":$"))
        return ",".join(f"{title}!{value}" for value in values)

    # -- inserting, deleting and moving (openpyxl semantics, see _shift) --
    def insert_rows(self, idx: int, amount: int = 1) -> None:
        """Insert ``amount`` empty rows before row ``idx``."""
        _shift.shift_rows(self, idx, amount, delete=False)

    def delete_rows(self, idx: int, amount: int = 1) -> None:
        """Delete ``amount`` rows starting at row ``idx``; the rows below move up."""
        _shift.shift_rows(self, idx, amount, delete=True)

    def insert_cols(self, idx: int, amount: int = 1) -> None:
        """Insert ``amount`` empty columns before column ``idx`` (1-based)."""
        _shift.shift_cols(self, idx, amount, delete=False)

    def delete_cols(self, idx: int, amount: int = 1) -> None:
        """Delete ``amount`` columns starting at column ``idx``; the rest move left."""
        _shift.shift_cols(self, idx, amount, delete=True)

    def move_range(self, cell_range: str, rows: int = 0, cols: int = 0, translate: bool = False) -> None:
        """Move the cells of ``cell_range``; cells at the destination are replaced.
        ``translate=True`` shifts the relative references of moved formulas."""
        _shift.move_range(self, cell_range, rows, cols, translate)

    # -- objects (openpyxl names, xlspy types) ------------------------
    def add_chart(self, chart: Chart, anchor: str | None = None) -> None:
        _expect(chart, Chart, "chart", "an xlspy Chart")
        if anchor is not None:
            chart = replace(chart, anchor=anchor)
        self._charts.append(chart)

    def add_image(self, img: Image, anchor: str | None = None) -> None:
        _expect(img, Image, "image", "an xlspy Image")
        if anchor is not None:
            # The bytes already read: a file object source is used up.
            img = replace(img, path=img.data, anchor=anchor)
        self._images.append(img)

    def add_table(self, table: ExcelTable | Table) -> None:
        _expect(table, (ExcelTable, Table), "table", "an xlspy ExcelTable or Table")
        self._tables.append(table)

    def add_data_validation(self, data_validation: DataValidation) -> None:
        _expect(data_validation, DataValidation, "data_validation", "an xlspy DataValidation")
        self._data_validations.append(data_validation)

    # -- cells --------------------------------------------------------
    def cell(self, row: int, column: int | str, value: Any = None) -> Cell:
        row = _coerce_row(row)
        column = _coerce_column(column)
        if self._is_merged_subcell(row, column):
            merged = MergedCell(self, row, column, style_id=0)
            if value is not None:
                raise AttributeError("read-only cell (merged range)")
            return merged
        cell = self._cells.get(row, {}).get(column)
        if cell is None:
            cell = Cell(self, row, column, style_id=0)
            self._cells.setdefault(row, {})[column] = cell
            self._max_row = max(self._max_row, row)
            self._max_column = max(self._max_column, column)
        if value is not None:
            cell.value = value
        return cell

    def _set_cell_raw(self, row: int, column: int, value: Any = None, style_id: int = 0,
                      data_type: str = "n", formula_text: str | None = None) -> Cell:
        # Loader hot path: bounds are checked inline and the cell is filled
        # directly, which is what Cell.__init__ + _load_value would do.
        if not (1 <= row <= 1048576 and 1 <= column <= 16384):
            _coerce_row(row)
            _coerce_column(column)
        cell = _new_cell(Cell)
        cell._sheet = self
        cell._row = row
        cell._column = column
        cell._value = value
        cell._formula = formula_text
        cell._data_type = data_type
        cell._style_id = style_id
        cell._comment = None
        cell._hyperlink = None
        cell._named_style = None
        cell._dirty = False
        row_cells = self._cells.get(row)
        if row_cells is None:
            row_cells = self._cells[row] = {}
        row_cells[column] = cell
        if value is not None or formula_text is not None or style_id or data_type != "n":
            if row > self._max_row:
                self._max_row = row
            if column > self._max_column:
                self._max_column = column
        return cell

    def _existing_cell(self, row: int, column: int) -> Cell | None:
        return self._cells.get(row, {}).get(column)

    def __getitem__(self, key: str | int) -> Any:
        if isinstance(key, str):
            if ":" in key:
                c1, r1, c2, r2 = parse_cell_range(key)
                return tuple(
                    tuple(self.cell(row + 1, col + 1) for col in range(c1, c2 + 1))
                    for row in range(r1, r2 + 1)
                )
            if _CELL_RE.match(key):
                col, row, _, _ = parse_cell_range(key)
                return self.cell(row + 1, col + 1)
        if isinstance(key, int) and not isinstance(key, bool):
            row = _coerce_row(key)
            return tuple(self.cell(row, col) for col in range(1, self._max_column + 1))
        raise TypeError("worksheet indices must be A1 references or row numbers")

    def __setitem__(self, key: str, value: Any) -> None:
        if isinstance(key, str) and ":" in key:
            c1, r1, c2, r2 = parse_cell_range(key)
            if isinstance(value, (tuple, list)):
                values = list(value)
                if len(values) == r2 - r1 + 1 and all(
                    isinstance(row, (tuple, list)) for row in values
                ):
                    for row_offset, row_values in enumerate(values):
                        for col_offset, item in enumerate(row_values):
                            self.cell(r1 + 1 + row_offset, c1 + 1 + col_offset, item)
                    return
            for row in range(r1, r2 + 1):
                for col in range(c1, c2 + 1):
                    self.cell(row + 1, col + 1, value)
            return
        if isinstance(key, str) and _CELL_RE.match(key):
            col, row, _, _ = parse_cell_range(key)
            self.cell(row + 1, col + 1, value)
            return
        raise TypeError("worksheet indices must be A1 references")

    def append(self, iterable: Iterable[Any]) -> int:
        if isinstance(iterable, Mapping):
            row = self._max_row + 1
            for key, value in iterable.items():
                if isinstance(key, str):
                    column = _coerce_column(key)
                elif isinstance(key, int) and not isinstance(key, bool):
                    if not 1 <= key <= 16384:
                        raise ValueError("column index must be 1..16384")
                    column = key
                else:
                    raise TypeError("dict keys must be column letters or 1-based ints")
                if isinstance(value, WriteOnlyCell):
                    value = value.value
                self.cell(row, column, value)
            if iterable:
                self._max_row = max(self._max_row, row)
                self._max_column = max(
                    self._max_column,
                    max(_coerce_column(key) for key in iterable),
                )
            return row
        if isinstance(iterable, (str, bytes, bytearray)):
            raise TypeError("a row must be an iterable of cells, not text")
        values = [
            value.value if isinstance(value, WriteOnlyCell) else value
            for value in iterable
        ]
        row = self._max_row + 1
        if (values and row <= 1048576 and len(values) <= 16384
                and not self._merged_cells and row not in self._cells):
            self._append_new_row(row, values)
            return row
        for column, value in enumerate(values, 1):
            if value is not None:
                self.cell(row, column, value)
            else:
                # An explicit empty cell still expands the used range, matching
                # openpyxl when a row is appended with a trailing None.
                self._max_column = max(self._max_column, column)
        if values:
            self._max_row = max(self._max_row, row)
        return row

    def _append_new_row(self, row: int, values: list) -> None:
        """append() into a row with no cells yet and no merged ranges.

        Plain int/float/bool/str values are stored as ``cell(row, column,
        value)`` would store them, without its checks; any other value goes
        through ``cell()``.
        """
        row_cells = self._cells[row] = {}
        dirty = self._dirty_cells
        plain_types = _CELL_TYPE_KINDS
        date_styles = None
        dated = False
        for column, value in enumerate(values, 1):
            if value is None:
                continue
            kind = type(value)
            plain = plain_types.get(kind)
            style_id = 0
            if plain is None:
                # A date or datetime gets the writers' default format, as the
                # value setter gives it; reuse the style id it already made.
                if kind is _datetime.datetime or kind is _datetime.date:
                    if date_styles is None:
                        date_styles = self._default_date_styles()
                    style_id = date_styles.get(kind) or 0
                if not style_id:
                    self.cell(row, column, value)
                    continue
                plain = "d"
                dated = True
                self._dirty_styles.add(_coordinate(row, column))
            elif plain == "s" and value.startswith("="):
                self.cell(row, column, value)
                continue
            cell = _new_cell(Cell)
            cell._sheet = self
            cell._row = row
            cell._column = column
            cell._value = value
            cell._formula = None
            cell._data_type = plain
            cell._style_id = style_id
            cell._comment = None
            cell._hyperlink = None
            cell._named_style = None
            cell._dirty = True
            row_cells[column] = cell
            dirty.add((row, column))
        if dated:
            self._workbook._styles_dirty = True
            self._dirty_meta.add("styles")
        if row_cells:
            self._dirty = True
        else:
            del self._cells[row]
        # As for cell() and the None branch of append(): every position counts.
        if row > self._max_row:
            self._max_row = row
        if len(values) > self._max_column:
            self._max_column = len(values)

    def _default_date_styles(self) -> dict:
        """Style ids a new date/datetime cell gets from the value setter, when known.

        Taken from the number-format memo of ``Cell.number_format``; a type
        missing here goes through ``cell()`` once, which fills the memo.
        """
        styles = self._workbook._styles
        if styles.style_at(0).number_format not in (None, "General"):
            return {}
        cache = styles._number_format_ids
        found = {}
        for kind, fmt in ((_datetime.datetime, DEFAULT_DATETIME_FORMAT),
                          (_datetime.date, DEFAULT_DATE_FORMAT)):
            style_id = cache.get((0, fmt))
            if style_id:
                found[kind] = style_id
        return found

    @overload
    def iter_rows(self, min_row: int | None = None, max_row: int | None = None, min_col: int | None = None,
                  max_col: int | None = None, values_only: Literal[False] = False) -> Iterator[tuple[Cell, ...]]: ...
    @overload
    def iter_rows(self, min_row: int | None = None, max_row: int | None = None, min_col: int | None = None,
                  max_col: int | None = None, *, values_only: Literal[True]) -> Iterator[tuple[Any, ...]]: ...
    @overload
    def iter_rows(self, min_row: int | None = None, max_row: int | None = None, min_col: int | None = None,
                  max_col: int | None = None, values_only: bool = False) -> Iterator[tuple[Any, ...]]: ...
    def iter_rows(self, min_row: int | None = None, max_row: int | None = None, min_col: int | None = None,
                  max_col: int | None = None, values_only: bool = False) -> Iterator[tuple[Any, ...]]:
        min_row = 1 if min_row is None else _coerce_row(min_row)
        max_row = self._max_row if max_row is None else _coerce_row(max_row)
        min_col = 1 if min_col is None else _coerce_column(min_col)
        max_col = self._max_column if max_col is None else _coerce_column(max_col)
        for row in range(min_row, max_row + 1):
            if values_only:
                yield tuple(self._value_at(row, col) for col in range(min_col, max_col + 1))
            else:
                yield tuple(self.cell(row, col) for col in range(min_col, max_col + 1))

    def _value_at(self, row: int, col: int) -> Any:
        cell = self._cells.get(row, {}).get(col)
        return None if cell is None else cell.value

    @overload
    def iter_cols(self, min_col: int | None = None, max_col: int | None = None, min_row: int | None = None,
                  max_row: int | None = None, values_only: Literal[False] = False) -> Iterator[tuple[Cell, ...]]: ...
    @overload
    def iter_cols(self, min_col: int | None = None, max_col: int | None = None, min_row: int | None = None,
                  max_row: int | None = None, *, values_only: Literal[True]) -> Iterator[tuple[Any, ...]]: ...
    @overload
    def iter_cols(self, min_col: int | None = None, max_col: int | None = None, min_row: int | None = None,
                  max_row: int | None = None, values_only: bool = False) -> Iterator[tuple[Any, ...]]: ...
    def iter_cols(self, min_col: int | None = None, max_col: int | None = None, min_row: int | None = None,
                  max_row: int | None = None, values_only: bool = False) -> Iterator[tuple[Any, ...]]:
        min_col = 1 if min_col is None else _coerce_column(min_col)
        max_col = self._max_column if max_col is None else _coerce_column(max_col)
        min_row = 1 if min_row is None else _coerce_row(min_row)
        max_row = self._max_row if max_row is None else _coerce_row(max_row)
        for col in range(min_col, max_col + 1):
            if values_only:
                yield tuple(self._value_at(row, col) for row in range(min_row, max_row + 1))
            else:
                yield tuple(self.cell(row, col) for row in range(min_row, max_row + 1))

    # -- merges -------------------------------------------------------
    def merge_cells(self, range_string: str | None = None, start_row: int | None = None,
                    start_column: int | None = None, end_row: int | None = None,
                    end_column: int | None = None) -> None:
        if range_string is None:
            if start_row is None or start_column is None or end_row is None or end_column is None:
                raise TypeError("merge_cells needs a range or all four boundaries")
            range_string = (
                f"{_coordinate(start_row, _coerce_column(start_column))}:"
                f"{_coordinate(end_row, _coerce_column(end_column))}"
            )
        reference = _normalize_range(range_string)
        for existing in self._merged_cells.ranges:
            if self._ranges_overlap(existing, reference):
                raise ValueError(f"merge range {reference!r} overlaps {existing!r}")
        self._merged_cells.add(reference)
        self._mark_dirty("merges")

    def unmerge_cells(self, range_string: str | None = None, start_row: int | None = None,
                      start_column: int | None = None, end_row: int | None = None,
                      end_column: int | None = None) -> None:
        if range_string is None:
            if start_row is None or start_column is None or end_row is None or end_column is None:
                raise TypeError("unmerge_cells needs a range or all four boundaries")
            range_string = (
                f"{_coordinate(start_row, _coerce_column(start_column))}:"
                f"{_coordinate(end_row, _coerce_column(end_column))}"
            )
        self._merged_cells.remove(range_string)
        self._mark_dirty("merges")

    @staticmethod
    def _ranges_overlap(a: str, b: str) -> bool:
        c1, r1, c2, r2 = parse_cell_range(a)
        d1, s1, d2, s2 = parse_cell_range(b)
        return not (c2 < d1 or d2 < c1 or r2 < s1 or s2 < r1)

    def _is_merged_subcell(self, row: int, column: int) -> bool:
        if not self._merged_cells:
            return False
        coordinate = _coordinate(row, column)
        for cell_range in self._merged_cells.ranges:
            if coordinate in cell_range:
                c1, r1, _, _ = parse_cell_range(cell_range)
                return (row, column) != (r1 + 1, c1 + 1)
        return False

    # -- dirty tracking ----------------------------------------------
    def _mark_dirty(self, *parts: str) -> None:
        self._dirty = True
        for part in parts:
            self._dirty_meta.add(part)

    def _clear_dirty(self) -> None:
        self._dirty = False
        self._dirty_cells.clear()
        self._dirty_styles.clear()
        self._dirty_comments.clear()
        self._dirty_hyperlinks.clear()
        self._dirty_meta.clear()
        for row in self._cells.values():
            for cell in row.values():
                cell._clear_dirty()


class WriteOnlyWorksheet(Worksheet):
    """Streaming-friendly sheet that stores rows rather than Cell objects."""

    def __init__(self, workbook: Workbook, title: str, *, index: int | None = None) -> None:
        super().__init__(workbook, title, index=index)
        self._rows: list[list[Any]] = []

    def append(self, iterable: Iterable[Any]) -> int:
        if isinstance(iterable, Mapping):
            entries = []
            for key, value in iterable.items():
                if isinstance(key, str):
                    column = _coerce_column(key)
                elif isinstance(key, int) and not isinstance(key, bool):
                    if not 1 <= key <= 16384:
                        raise ValueError("column index must be 1..16384")
                    column = key
                else:
                    raise TypeError("dict keys must be column letters or 1-based ints")
                entries.append((column, value.value if isinstance(value, WriteOnlyCell) else value))
            row = [None] * (max((column for column, _ in entries), default=0))
            for column, value in entries:
                row[column - 1] = value
            self._rows.append(row)
        else:
            if isinstance(iterable, (str, bytes, bytearray)):
                raise TypeError("a row must be an iterable of cells, not text")
            self._rows.append([
                value.value if isinstance(value, WriteOnlyCell) else value
                for value in iterable
            ])
        self._max_row = len(self._rows)
        self._max_column = max((len(row) for row in self._rows), default=0)
        self._mark_dirty("rows")
        return self._max_row

    def cell(self, row: int, column: int | str, value: Any = None) -> Cell:
        raise ValueError("cannot access a cell in a write-only worksheet; use append()")

    def iter_rows(self, *args: Any, **kwargs: Any) -> Iterator[tuple[Any, ...]]:
        for row in self._rows:
            if kwargs.get("values_only"):
                yield tuple(row)
            else:
                raise ValueError("write-only worksheets only support values_only rows")


class Chartsheet:
    """A chartsheet: a sheet that shows one chart.

    A new chartsheet's chart reads its ranges from ``source`` (a worksheet or
    its title; by default the workbook's first worksheet).
    """

    def __init__(self, workbook: Workbook, title: str, index: int | None = None) -> None:
        self._workbook = workbook
        self._title = title  # the caller checks the name
        self.sheet_state = "visible"
        self._index = index
        self._charts: list[Chart] = []
        self._part: str | None = None  # the package part, once in a file
        self._source: Worksheet | str | None = None

    def __repr__(self) -> str:
        return f"<Chartsheet {self.title!r}>"

    @property
    def title(self) -> str:
        return self._title

    @title.setter
    def title(self, value: str) -> None:
        if value == self._title:
            return
        if self._part is not None:
            raise UnsupportedFeatureError("renaming a loaded chartsheet is not supported")
        self._workbook._check_title(value, self)
        self._title = value

    @property
    def charts(self) -> list[Chart]:
        return self._charts

    def add_chart(self, chart: Chart, source: Worksheet | str | None = None) -> None:
        if self._part is None:
            _expect(chart, Chart, "chart", "an xlspy Chart")
            if self._charts:
                raise ValueError("a chartsheet holds one chart")
            if source is not None:
                self._source = source
        self._charts.append(chart)
