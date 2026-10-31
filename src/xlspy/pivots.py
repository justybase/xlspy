"""Pivot tables: a compact-layout pivot over a table written in the same call.

The XLSX writer creates a pivot table from an ``ExcelTable`` written by the
same ``write_sheet`` call, so the rows stream past once and the values are
aggregated as they go.  The pivot renders its own cells, writes a pivot cache
without cached records (``saveData="0"``) and sets ``refreshOnLoad="1"``, so
Excel rebuilds the report from the source table when the file opens.

Only creation is supported, and only in XLSX.  The compact layout accepts one
row field, at most one column field, optional page (filter) fields and any
number of value fields.  A loaded workbook keeps its pivot tables as untouched
parts, as before.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal
import math
import numbers
import re
import xml.sax.saxutils as saxutils

from ._extlst import X14, ext, ext_list
from .formats import _unwrap_cell
from .formulas import Formula
from .hyperlinks import Hyperlink
from .refs import column_label, parse_cell_range
from .styles import ExcelTable, table_column_names

_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_PLAIN_TYPES = (str, bool, int, float, Decimal, date, datetime)
_PIVOT_EXT_LIST = ext_list([ext(
    "{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}",
    '<x14:pivotTableDefinition hideValuesRow="1"/>',
    ("x14", X14),
)])
_FUNCTIONS = {"sum", "count", "average", "min", "max"}
# The dataField ``subtotal`` values; "sum" is the default and is not written.
_SUBTOTALS = {"count": "count", "average": "average", "min": "min", "max": "max"}
_FUNCTION_LABELS = {
    "sum": "Sum of",
    "count": "Count of",
    "average": "Average of",
    "min": "Min of",
    "max": "Max of",
}
_NAME_RE = re.compile(r"[A-Za-z_][A-Za-z0-9_.]*")


def _attr(value: str) -> str:
    return saxutils.escape(str(value), {'"': "&quot;"})


def _number_text(value) -> str:
    if isinstance(value, float):
        if math.isnan(value):
            return "NaN"
        if math.isinf(value):
            return "INF" if value > 0 else "-INF"
        if value.is_integer():
            return str(int(value))
        return repr(value)
    return str(value)


class _Blank:
    """The cache's missing value; Excel shows it as ``(blank)``."""

    def __repr__(self) -> str:  # pragma: no cover - debugging aid
        return "(blank)"


_BLANK = _Blank()
_NO_COLUMN = object()


def _cell_value(cell):
    """The value a pivot reads from a written cell (formats and links unwrapped)."""
    cell = _unwrap_cell(cell)
    if cell is None or isinstance(cell, _PLAIN_TYPES):
        return cell
    if isinstance(cell, Hyperlink):
        return cell.display_text
    if isinstance(cell, Formula):
        return cell.value
    # numpy scalars from a DataFrame are numbers to the sheet they are written to.
    if isinstance(cell, numbers.Integral):
        return int(cell)
    if isinstance(cell, numbers.Real):
        return float(cell)
    return str(cell)


def _key(value):
    """The grouping key: blanks collapse to one item, NaN follows the sheet."""
    if value is None or value == "":
        return _BLANK
    if isinstance(value, float) and math.isnan(value):
        return "NaN"
    return value


def _sort_key(value):
    """Excel-like ascending order; blanks last and text case-insensitive."""
    if value is _BLANK:
        return (3, "")
    if isinstance(value, bool):
        return (0, int(value))
    if isinstance(value, (int, float, Decimal)):
        return (0, float(value))
    if isinstance(value, (datetime, date)):
        return (1, value.isoformat())
    return (2, str(value).casefold())


def _label(value):
    if value is _BLANK:
        return "(blank)"
    return value


def _shared_item_xml(value) -> str:
    if value is _BLANK:
        return "<m/>"
    if isinstance(value, bool):
        return f'<b v="{1 if value else 0}"/>'
    if isinstance(value, datetime):
        return f'<d v="{value.isoformat()}"/>'
    if isinstance(value, date):
        return f'<d v="{value.isoformat()}T00:00:00"/>'
    if isinstance(value, Decimal):
        return f'<n v="{_attr(value)}"/>'
    if isinstance(value, (int, float)):
        return f'<n v="{_number_text(value)}"/>'
    return f'<s v="{_attr(value)}"/>'


@dataclass(frozen=True, slots=True)
class PivotValue:
    """One value field: ``field`` summarised with ``function``.

    ``name`` is the caption shown in the pivot (default: ``"Sum of Field"``).
    ``number_format`` is an Excel format string applied to the rendered cells
    and to the data field.
    """

    field: str
    function: str = "sum"
    name: str | None = None
    number_format: str | None = None

    def __post_init__(self):
        if not isinstance(self.field, str) or not self.field:
            raise ValueError("pivot value field must be a non-empty string")
        if not isinstance(self.function, str):
            raise TypeError("pivot value function must be a string")
        function = self.function.casefold()
        if function not in _FUNCTIONS:
            raise ValueError(
                f"pivot value function must be one of {sorted(_FUNCTIONS)}, "
                f"got {self.function!r}"
            )
        object.__setattr__(self, "function", function)
        for name, value in (("name", self.name), ("number_format", self.number_format)):
            if value is not None and (not isinstance(value, str) or not value):
                raise ValueError(f"pivot value {name} must be a non-empty string or None")

    @property
    def caption(self) -> str:
        if self.name is not None:
            return self.name
        return f"{_FUNCTION_LABELS[self.function]} {self.field}"


def _axis(value, name: str) -> tuple[str, ...]:
    if value is None:
        return ()
    if isinstance(value, str):
        items = (value,)
    elif isinstance(value, (list, tuple)):
        items = tuple(value)
    else:
        raise TypeError(f"pivot {name} must be a field name or a sequence of names")
    for item in items:
        if not isinstance(item, str) or not item:
            raise ValueError(f"pivot {name} entries must be non-empty strings")
    return items


def _value_spec(value) -> PivotValue:
    if isinstance(value, PivotValue):
        return value
    if isinstance(value, str):
        return PivotValue(value)
    if isinstance(value, (list, tuple)):
        if not 1 <= len(value) <= 4:
            raise ValueError(
                "a pivot value tuple must be (field[, function[, name[, format]]])"
            )
        return PivotValue(*value)
    raise TypeError(
        "pivot values must be PivotValue instances, field names or "
        "(field, function, ...) tuples"
    )


def _values(value) -> tuple[PivotValue, ...]:
    if value is None:
        raise ValueError("a pivot table needs at least one value field")
    if isinstance(value, (PivotValue, str)):
        items = (value,)
    elif isinstance(value, (list, tuple)):
        items = tuple(value)
    else:
        raise TypeError("pivot values must be a PivotValue or a sequence")
    specs = tuple(_value_spec(item) for item in items)
    if not specs:
        raise ValueError("a pivot table needs at least one value field")
    return specs


def _split_location(location: str) -> tuple[str | None, str]:
    """``"Summary!A3"`` -> ``("Summary", "A3")``; a plain cell -> ``(None, "A3")``."""
    if not isinstance(location, str):
        raise TypeError("pivot location must be an A1 cell")
    sheet = None
    cell = location
    if "!" in location:
        sheet_part, _, cell = location.rpartition("!")
        sheet = sheet_part
        if len(sheet) >= 2 and sheet[0] == sheet[-1] == "'":
            sheet = sheet[1:-1].replace("''", "'")
        if not sheet:
            raise ValueError(f"invalid pivot location {location!r}")
    first_col, first_row, last_col, last_row = parse_cell_range(cell)
    if (first_col, first_row) != (last_col, last_row):
        raise ValueError(f"pivot location must be a single cell, got {location!r}")
    return sheet, cell


@dataclass(frozen=True, slots=True)
class PivotTable:
    """A compact pivot table over the table ``source``.

    ``source`` is an ``ExcelTable`` written by the same ``write_sheet`` call
    (or its display name).  ``location`` is the top-left cell of the pivot; it
    may name the sheet (``"Summary!A3"``) instead of using ``sheet``.  When
    ``sheet`` is omitted the pivot is placed on the source sheet, below the
    table.  Otherwise the target sheet must be written after the source sheet,
    or left empty and saved as is.
    """

    source: object
    rows: str | tuple[str, ...]
    values: object
    location: str = "A3"
    columns: str | tuple[str, ...] = ()
    filters: str | tuple[str, ...] = ()
    sheet: str | None = None
    name: str | None = None
    style: str = "PivotStyleLight16"
    data_caption: str = "Values"

    def __post_init__(self):
        rows = _axis(self.rows, "rows")
        if len(rows) != 1:
            raise ValueError(
                "the compact pivot layout needs exactly one row field"
            )
        columns = _axis(self.columns, "columns")
        if len(columns) > 1:
            raise ValueError(
                "the compact pivot layout accepts at most one column field"
            )
        filters = _axis(self.filters, "filters")
        if len(set(rows) & set(columns)) or len((set(rows) | set(columns)) & set(filters)):
            raise ValueError("a pivot field cannot be on two axes at once")
        object.__setattr__(self, "rows", rows)
        object.__setattr__(self, "columns", columns)
        object.__setattr__(self, "filters", filters)
        object.__setattr__(self, "values", _values(self.values))

        source = self.source
        if isinstance(source, ExcelTable):
            object.__setattr__(self, "source", source.display_name)
        elif not isinstance(source, str) or not source:
            raise TypeError(
                "pivot source must be an ExcelTable or a table display name"
            )

        location_sheet, cell = _split_location(self.location)
        object.__setattr__(self, "location", cell)
        if self.sheet is not None and (not isinstance(self.sheet, str) or not self.sheet):
            raise ValueError("pivot sheet must be a non-empty string or None")
        if location_sheet is not None:
            if self.sheet is not None and self.sheet.casefold() != location_sheet.casefold():
                raise ValueError(
                    "pivot location names one sheet but sheet= names another"
                )
            object.__setattr__(self, "sheet", location_sheet)
        if self.name is not None:
            if not isinstance(self.name, str) or _NAME_RE.fullmatch(self.name) is None:
                raise ValueError(
                    f"invalid pivot table name {self.name!r}; use letters, digits, "
                    "underscores and periods, starting with a letter or underscore"
                )
        for attr, value in (("style", self.style), ("data_caption", self.data_caption)):
            if not isinstance(value, str) or not value:
                raise ValueError(f"pivot {attr} must be a non-empty string")

    @property
    def table_name(self) -> str:
        return self.source if isinstance(self.source, str) else self.source.display_name


@dataclass
class _PivotPlan:
    pivot: PivotTable
    table: ExcelTable
    target_sheet: str
    name: str
    start_col: int
    start_row: int


def normalize_pivots(pivots, tables, used_names: set[str], current_sheet: str) -> list[_PivotPlan]:
    """Validate pivots against this call's tables and reserve their names."""
    if not pivots:
        return []
    if not isinstance(pivots, (list, tuple)):
        raise TypeError("pivots must be a list of PivotTable objects or None")
    by_name = {table.display_name.casefold(): table for table in tables}
    names = set(used_names)
    plans = []
    for pivot in pivots:
        if not isinstance(pivot, PivotTable):
            raise TypeError("pivots must be PivotTable instances")
        table = by_name.get(pivot.table_name.casefold())
        if table is None:
            raise ValueError(
                f"pivot source table {pivot.table_name!r} is not a table written "
                "to this sheet in this call"
            )
        if not table.header_row:
            raise ValueError(
                f"pivot source table {table.display_name!r} needs a header row"
            )
        name = pivot.name
        if name is None:
            index = len(names) + 1
            while f"pivottable{index}" in names:
                index += 1
            name = f"PivotTable{index}"
        if name.casefold() in names:
            raise ValueError(f"duplicate pivot table name {name!r}")
        names.add(name.casefold())
        first_col, first_row, _last_col, _last_row = parse_cell_range(pivot.location)
        plans.append(_PivotPlan(
            pivot=pivot,
            table=table,
            target_sheet=pivot.sheet or current_sheet,
            name=name,
            start_col=first_col,
            start_row=first_row,
        ))
    used_names.update(names)
    return plans


class _Metric:
    """One value field over one group."""

    __slots__ = ("records", "count", "count_num", "total", "minimum", "maximum")

    def __init__(self):
        self.records = 0
        self.count = 0
        self.count_num = 0
        self.total = 0.0
        self.minimum = None
        self.maximum = None

    def add(self, value) -> None:
        self.records += 1
        if value is None or value == "":
            return
        self.count += 1
        if isinstance(value, bool) or not isinstance(value, (int, float, Decimal)):
            return
        number = float(value)
        if math.isnan(number):
            return
        self.total += number
        self.count_num += 1
        if self.minimum is None or number < self.minimum:
            self.minimum = number
        if self.maximum is None or number > self.maximum:
            self.maximum = number

    def result(self, function: str):
        if function == "count":
            return self.count if self.records else None
        if not self.count_num:
            return None
        if function == "sum":
            value = self.total
            return int(value) if value.is_integer() else value
        if function == "average":
            return self.total / self.count_num
        if function == "min":
            value = self.minimum
            return int(value) if value.is_integer() else value
        value = self.maximum
        return int(value) if value.is_integer() else value


def _new_metrics(size: int) -> list[_Metric]:
    return [_Metric() for _ in range(size)]


class PivotFeeder:
    """Aggregate the source rows as they stream past, then render the pivot.

    The source is a table: its first row holds the field names and every row
    after it is one cache record.  ``feed`` is called for each sheet row; the
    accumulator starts once the header row passes and stops at the table's
    last row (the totals row, when the table has one, is not source data).
    """

    def __init__(self, plan: _PivotPlan, *, track_all_fields: bool = False):
        # The binary cache types every field, so XLSB tracks the unused ones too.
        self._track_all_fields = track_all_fields
        self.plan = plan
        self.pivot = plan.pivot
        self.table = plan.table
        first_col, first_row, _last_col, last_row = parse_cell_range(plan.table.ref)
        self._first_col = first_col
        self._first_row = first_row
        self._last_row = last_row - 1 if plan.table.totals_row else last_row
        # Field indices below are table-relative; cells are read at first_col + index.
        self._names: list[str] | None = None
        self._row_index = None
        self._col_index = None
        self._filter_indices: list[int] = []
        self._value_indices: list[int] = []
        self._records = 0
        self._groups: dict[tuple, list[_Metric]] = {}
        self._row_totals: dict[object, list[_Metric]] = {}
        self._col_totals: dict[object, list[_Metric]] = {}
        self._grand: list[_Metric] = []
        self._row_seen: dict[object, int] = {}
        self._col_seen: dict[object, int] = {}
        self._filter_seen: dict[int, dict[object, int]] = {}
        self._value_stats: dict[int, _Stats] = {}

    def feed(self, row_idx: int, row) -> None:
        if self._names is None:
            if row_idx == self._first_row:
                cells = [_unwrap_cell(cell) for cell in row]
                self._names = table_column_names(self.table, cells)
                self._resolve_fields()
            return
        if row_idx <= self._first_row or row_idx > self._last_row:
            return
        self._feed_row(row)

    def _resolve_fields(self) -> None:
        names = self._names or []

        def find(field: str) -> int:
            for index, name in enumerate(names):
                if name.casefold() == field.casefold():
                    return index
            raise ValueError(
                f"pivot field {field!r} is not a column of {self.table.display_name!r}; "
                f"available columns: {', '.join(names)}"
            )

        self._row_index = find(self.pivot.rows[0])
        self._col_index = find(self.pivot.columns[0]) if self.pivot.columns else None
        self._filter_indices = [find(field) for field in self.pivot.filters]
        self._value_indices = [find(spec.field) for spec in self.pivot.values]
        stat_fields = set(self._value_indices)
        if self._track_all_fields:
            axes = {self._row_index, self._col_index, *self._filter_indices}
            stat_fields |= set(range(len(names))) - axes
        self._value_stats = {index: _Stats() for index in sorted(stat_fields)}
        self._grand = _new_metrics(len(self.pivot.values))

    def _feed_row(self, row) -> None:
        def cell_value_at(index: int):
            column = self._first_col + index
            return _cell_value(row[column]) if column < len(row) else None

        def cell(index: int):
            return _key(cell_value_at(index))

        row_key = cell(self._row_index)
        col_key = cell(self._col_index) if self._col_index is not None else _NO_COLUMN
        if row_key not in self._row_seen:
            self._row_seen[row_key] = len(self._row_seen)
        if self._col_index is not None and col_key not in self._col_seen:
            self._col_seen[col_key] = len(self._col_seen)
        for field_index in self._filter_indices:
            seen = self._filter_seen.setdefault(field_index, {})
            key = cell(field_index)
            if key not in seen:
                seen[key] = len(seen)

        group = self._groups.setdefault((row_key, col_key), _new_metrics(len(self.pivot.values)))
        row_total = self._row_totals.setdefault(row_key, _new_metrics(len(self.pivot.values)))
        metrics = [group, row_total, self._grand]
        col_total = None
        if self._col_index is not None:
            col_total = self._col_totals.setdefault(col_key, _new_metrics(len(self.pivot.values)))
            metrics.append(col_total)
        for index, field_index in enumerate(self._value_indices):
            value = cell_value_at(field_index)
            for target in metrics:
                target[index].add(value)
        for field_index, stats in self._value_stats.items():
            stats.add(cell_value_at(field_index))
        self._records += 1

    def result(self) -> "PivotRender":
        if self._names is None:
            raise ValueError(
                f"pivot source table {self.table.display_name!r} never reached its "
                "header row; the sheet data ended before the table started"
            )
        row_display = sorted(self._row_seen, key=_sort_key)
        col_display = sorted(self._col_seen, key=_sort_key) if self._col_index is not None else []
        return PivotRender(
            plan=self.plan,
            cells=self._render_cells(row_display, col_display),
            ref=self._render_ref(row_display, col_display),
            row_field=self._row_index,
            col_field=self._col_index,
            value_indices=self._value_indices,
            value_captions=[spec.caption for spec in self.pivot.values],
            filter_indices=self._filter_indices,
            row_items=[self._row_seen[key] for key in row_display],
            col_items=[self._col_seen[key] for key in col_display],
            row_display=row_display,
            col_display=col_display,
            shared_items={
                self._row_index: list(self._row_seen),
                **({
                    self._col_index: list(self._col_seen),
                } if self._col_index is not None else {}),
                **{
                    field: list(seen)
                    for field, seen in self._filter_seen.items()
                },
            },
            filter_names=[self._names[index] for index in self._filter_indices],
            column_names=list(self._names),
            value_stats=self._value_stats,
            records=self._records,
            first_header_row=self._first_header_row(),
            first_data_row=self._first_data_row(),
            end_row=0,
            end_col=0,
        )

    def _first_header_row(self) -> int:
        if self._col_index is None:
            return 1 if len(self.pivot.values) == 1 else 0
        return 1

    def _first_data_row(self) -> int:
        if self._col_index is None:
            return 1
        return 2 if len(self.pivot.values) == 1 else 3

    def _render_cells(self, row_display, col_display) -> list[tuple[int, int, object]]:
        cells: list[tuple[int, int, object]] = []
        start_col = self.plan.start_col
        start_row = self.plan.start_row

        def put(row: int, column: int, value, fmt: str | None = None) -> None:
            if value is None:
                return
            cells.append((row, column, (value, fmt) if fmt is not None else value))

        specs = self.pivot.values
        formats = [spec.number_format for spec in specs]
        count = len(specs)

        # Page fields sit above the pivot, outside its location range.
        for offset, (field_index, name) in enumerate(
            zip(self._filter_indices, [self._names[i] for i in self._filter_indices])
        ):
            page_row = start_row - 1 - len(self._filter_indices) + offset
            if page_row < 0:
                raise ValueError(
                    "pivot filters need one free row each above the pivot location"
                )
            put(page_row, start_col, name)
            put(page_row, start_col + 1, "(All)")

        if self._col_index is None:
            put(start_row, start_col, "Row Labels")
            for index, spec in enumerate(specs):
                put(start_row, start_col + 1 + index, spec.caption)
            for offset, key in enumerate(row_display):
                row = start_row + 1 + offset
                put(row, start_col, _label(key))
                for index, metric in enumerate(self._row_totals[key]):
                    put(row, start_col + 1 + index, metric.result(specs[index].function),
                        formats[index])
            grand_row = start_row + 1 + len(row_display)
            put(grand_row, start_col, "Grand Total")
            for index, metric in enumerate(self._grand):
                put(grand_row, start_col + 1 + index, metric.result(specs[index].function),
                    formats[index])
            return cells

        columns = len(col_display)
        if count == 1:
            put(start_row, start_col, specs[0].caption)
            put(start_row, start_col + 1, "Column Labels")
            put(start_row + 1, start_col, "Row Labels")
            for offset, key in enumerate(col_display):
                put(start_row + 1, start_col + 1 + offset, _label(key))
            put(start_row + 1, start_col + 1 + columns, "Grand Total")
            for offset, key in enumerate(row_display):
                row = start_row + 2 + offset
                put(row, start_col, _label(key))
                for column, col_key in enumerate(col_display):
                    metrics = self._groups.get((key, col_key))
                    if metrics is not None:
                        put(row, start_col + 1 + column,
                            metrics[0].result(specs[0].function), formats[0])
                put(row, start_col + 1 + columns,
                    self._row_totals[key][0].result(specs[0].function), formats[0])
            grand_row = start_row + 2 + len(row_display)
            put(grand_row, start_col, "Grand Total")
            for column, col_key in enumerate(col_display):
                put(grand_row, start_col + 1 + column,
                    self._col_totals[col_key][0].result(specs[0].function), formats[0])
            put(grand_row, start_col + 1 + columns,
                self._grand[0].result(specs[0].function), formats[0])
            return cells

        put(start_row, start_col + 1, "Column Labels")
        for column, col_key in enumerate(col_display):
            put(start_row + 1, start_col + 1 + column * count, _label(col_key))
        total_start = start_col + 1 + columns * count
        for index, spec in enumerate(specs):
            put(start_row + 1, total_start + index, f"Total: {spec.caption}")
        put(start_row + 2, start_col, "Row Labels")
        for column in range(columns):
            for index, spec in enumerate(specs):
                put(start_row + 2, start_col + 1 + column * count + index, spec.caption)
        for offset, key in enumerate(row_display):
            row = start_row + 3 + offset
            put(row, start_col, _label(key))
            for column, col_key in enumerate(col_display):
                metrics = self._groups.get((key, col_key))
                if metrics is None:
                    continue
                for index, metric in enumerate(metrics):
                    put(row, start_col + 1 + column * count + index,
                        metric.result(specs[index].function), formats[index])
            for index, metric in enumerate(self._row_totals[key]):
                put(row, total_start + index, metric.result(specs[index].function),
                    formats[index])
        grand_row = start_row + 3 + len(row_display)
        put(grand_row, start_col, "Grand Total")
        for column, col_key in enumerate(col_display):
            for index, metric in enumerate(self._col_totals[col_key]):
                put(grand_row, start_col + 1 + column * count + index,
                    metric.result(specs[index].function), formats[index])
        for index, metric in enumerate(self._grand):
            put(grand_row, total_start + index, metric.result(specs[index].function),
                formats[index])
        return cells

    def _render_ref(self, row_display, col_display) -> str:
        start_col = self.plan.start_col
        start_row = self.plan.start_row
        count = len(self.pivot.values)
        if self._col_index is None:
            end_row = start_row + 1 + len(row_display)
            end_col = start_col + count
        elif count == 1:
            end_row = start_row + 2 + len(row_display)
            end_col = start_col + 1 + len(col_display)
        else:
            end_row = start_row + 3 + len(row_display)
            end_col = start_col + 1 + len(col_display) * count + count - 1
        return (
            f"{column_label(start_col)}{start_row + 1}:"
            f"{column_label(end_col)}{end_row + 1}"
        )


@dataclass
class PivotRender:
    """The rendered pivot: its cells and everything its XML parts need."""

    plan: _PivotPlan
    cells: list[tuple[int, int, object]]
    ref: str
    row_field: int
    col_field: int | None
    value_indices: list[int]
    value_captions: list[str]
    filter_indices: list[int]
    filter_names: list[str]
    row_items: list[int]
    col_items: list[int]
    row_display: list
    col_display: list
    shared_items: dict[int, list]
    column_names: list[str]
    value_stats: dict[int, _Stats]
    records: int
    first_header_row: int
    first_data_row: int
    end_row: int = 0
    end_col: int = 0
    pivot_id: int = 0
    cache_id: int = 0
    numfmt_ids: list[int | None] = field(default_factory=list)

    @property
    def pivot(self) -> PivotTable:
        return self.plan.pivot

    @property
    def start_row(self) -> int:
        return self.plan.start_row

    @property
    def start_col(self) -> int:
        return self.plan.start_col

    @property
    def max_row(self) -> int:
        return max((row for row, _column, _value in self.cells), default=self.start_row)

    @property
    def max_col(self) -> int:
        return max((column for _row, column, _value in self.cells), default=self.start_col)

    def assign_ids(self, pivot_id: int) -> None:
        self.pivot_id = pivot_id
        self.cache_id = pivot_id


def tee_rows(rows, feeders: list[PivotFeeder]):
    """Yield ``rows`` unchanged while each feeder aggregates them."""
    for row_idx, row in enumerate(rows):
        for feeder in feeders:
            feeder.feed(row_idx, row)
        yield row


def merge_cells(renders: list[PivotRender]) -> dict[tuple[int, int], object]:
    """Merge rendered cells, refusing overlaps (including page-field rows)."""
    merged: dict[tuple[int, int], object] = {}
    for render in renders:
        for row, column, value in render.cells:
            key = (row, column)
            if key in merged:
                raise ValueError(
                    "pivot tables overlap at "
                    f"{column_label(column)}{row + 1}; move one location"
                )
            merged[key] = value
    return merged


def overlay_rows(rows, renders: list[PivotRender]):
    """Yield ``rows`` with the pending pivot cells merged in, row by row."""
    merged = merge_cells(renders)
    by_row: dict[int, dict[int, object]] = {}
    for (row, column), value in merged.items():
        by_row.setdefault(row, {})[column] = value
    last = max(by_row, default=-1)
    emitted = -1
    for row_idx, row in enumerate(rows):
        emitted = row_idx
        cells = by_row.get(row_idx)
        if cells:
            yield _merge_row(row, cells, row_idx)
        else:
            yield row
    for row_idx in range(emitted + 1, last + 1):
        cells = by_row.get(row_idx)
        if cells:
            yield _merge_row([], cells, row_idx)
        else:
            yield []


def _merge_row(row, cells: dict[int, object], row_idx: int):
    width = max(max(cells), len(row) - 1) + 1
    merged = list(row) + [None] * (width - len(row))
    for column, value in cells.items():
        if column < len(row) and row[column] is not None:
            raise ValueError(
                "pivot output would overwrite the sheet cell "
                f"{column_label(column)}{row_idx + 1}"
            )
        merged[column] = value
    return merged


def _is_number(value) -> bool:
    return isinstance(value, (int, float, Decimal)) and not isinstance(value, bool)


class _Stats:
    """What a cache field holds, as the flags Excel reads its type by.

    A numeric field needs ``containsNumber``, ``containsInteger`` and the
    ``minValue``/``maxValue`` range; without them Excel mis-lays out a pivot
    with a numeric column field or several value fields.
    """

    __slots__ = ("blank", "other", "text", "unsupported", "count", "integer", "low", "high")

    def __init__(self):
        self.blank = False  # some cells are empty
        self.other = False  # some cells are text, bool, date or NaN/inf
        self.text = False  # some cells are text
        self.unsupported = False  # some cells are bool, date or NaN/inf
        self.count = 0  # numeric cells
        self.integer = True  # every numeric cell is a whole number
        self.low = None
        self.high = None

    def add(self, value) -> None:
        if value is _BLANK or value is None or value == "":
            self.blank = True
            return
        if not _is_number(value):
            self.other = True
            if isinstance(value, str):
                self.text = True
            else:
                self.unsupported = True
            return
        number = float(value)
        if not math.isfinite(number):
            self.other = True
            self.unsupported = True
            return
        self.count += 1
        if not number.is_integer():
            self.integer = False
        if self.low is None or number < self.low:
            self.low = number
        if self.high is None or number > self.high:
            self.high = number

    def attrs(self) -> str:
        attrs = ' containsBlank="1"' if self.blank else ""
        if self.count and not self.other:
            attrs = ' containsSemiMixedTypes="0" containsString="0"' + attrs + ' containsNumber="1"'
            if self.integer:
                attrs += ' containsInteger="1"'
            attrs += (
                f' minValue="{_number_text(self.low)}"'
                f' maxValue="{_number_text(self.high)}"'
            )
        return attrs


def _shared_items_xml(values: list) -> str:
    """The shared items of an axis field, with the flags its values call for."""
    if not values:
        return "<sharedItems/>"
    stats = _Stats()
    for value in values:
        stats.add(value)
    items = "".join(_shared_item_xml(value) for value in values)
    return f'<sharedItems{stats.attrs()} count="{len(values)}">{items}</sharedItems>'


def pivot_cache_xml(render: PivotRender) -> str:
    """The cache definition part; ``saveData="0"`` means Excel refreshes on open."""
    fields = []
    for index, name in enumerate(render.column_names):
        if index in render.shared_items:
            items = _shared_items_xml(render.shared_items[index])
        elif index in render.value_stats:
            items = f"<sharedItems{render.value_stats[index].attrs()}/>"
        else:
            items = "<sharedItems/>"
        fields.append(f'<cacheField name="{_attr(name)}" numFmtId="0">{items}</cacheField>')
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<pivotCacheDefinition xmlns="{_MAIN_NS}" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        f'saveData="0" refreshOnLoad="1" createdVersion="8" refreshedVersion="8" '
        f'minRefreshableVersion="3" recordCount="{render.records}">'
        f'<cacheSource type="worksheet"><worksheetSource name="{_attr(render.plan.table.display_name)}"/>'
        "</cacheSource>"
        f'<cacheFields count="{len(render.column_names)}">{"".join(fields)}</cacheFields>'
        "</pivotCacheDefinition>"
    )


def pivot_table_xml(render: PivotRender, data_field_format_ids: list[int | None]) -> str:
    """The pivot table definition part."""
    pivot_count = len(render.column_names)
    data_fields = set(render.value_indices)
    row_field = render.row_field
    col_field = render.col_field

    fields = []
    for index in range(pivot_count):
        attrs = ""
        if index == row_field:
            attrs += ' axis="axisRow"'
        if index == col_field:
            attrs += ' axis="axisCol"'
        if index in render.filter_indices:
            attrs += ' axis="axisPage"'
        if index in data_fields:
            attrs += ' dataField="1"'
        attrs += ' showAll="0"'
        if index in render.shared_items:
            display = (
                render.row_display if index == row_field
                else render.col_display if index == col_field
                else sorted(render.shared_items[index], key=_sort_key)
            )
            seen = render.shared_items[index]
            positions = {value: position for position, value in enumerate(seen)}
            items = "".join(f'<item x="{positions[value]}"/>' for value in display)
            attrs_children = f'<items count="{len(display) + 1}">{items}<item t="default"/></items>'
        else:
            attrs_children = ""
        fields.append(f"<pivotField{attrs}>{attrs_children}</pivotField>")

    row_items = "".join(
        "<i><x/></i>" if position == 0 else f'<i><x v="{position}"/></i>'
        for position in range(len(render.row_items))
    )

    col_xml = ""
    count = len(render.value_captions)
    if col_field is None:
        if count == 1:
            col_xml = '<colItems count="1"><i/></colItems>'
        else:
            items = ["<i><x/></i>"]
            items.extend(f'<i i="{index}"><x v="{index}"/></i>' for index in range(1, count))
            col_xml = (
                '<colFields count="1"><field x="-2"/></colFields>'
                f'<colItems count="{count}">{"".join(items)}</colItems>'
            )
    elif count == 1:
        items = []
        for position in range(len(render.col_items)):
            items.append(
                "<i><x/></i>" if position == 0
                else f'<i><x v="{position}"/></i>'
            )
        items.append('<i t="grand"><x/></i>')
        col_xml = (
            f'<colFields count="1"><field x="{col_field}"/></colFields>'
            f'<colItems count="{len(items)}">{"".join(items)}</colItems>'
        )
    else:
        items = []
        for position in range(len(render.col_items)):
            items.append(
                f'<i><x v="{position}"/><x/></i>' if position
                else "<i><x/><x/></i>"
            )
            for index in range(1, count):
                # A repeated column item names the data field, not the column.
                items.append(f'<i r="1" i="{index}"><x v="{index}"/></i>')
        items.append('<i t="grand"><x/></i>')
        for index in range(1, count):
            items.append(f'<i t="grand" i="{index}"><x/></i>')
        col_xml = (
            f'<colFields count="2"><field x="{col_field}"/><field x="-2"/></colFields>'
            f'<colItems count="{len(items)}">{"".join(items)}</colItems>'
        )

    page_xml = ""
    if render.filter_indices:
        pages = "".join(
            f'<pageField fld="{index}" hier="-1"/>' for index in render.filter_indices
        )
        page_xml = f'<pageFields count="{len(render.filter_indices)}">{pages}</pageFields>'

    data_fields_xml = []
    for index, (column, spec) in enumerate(zip(render.value_indices, render.pivot.values)):
        numfmt = ""
        if data_field_format_ids[index] is not None:
            numfmt = f' numFmtId="{data_field_format_ids[index]}"'
        subtotal = "" if spec.function == "sum" else f' subtotal="{_SUBTOTALS[spec.function]}"'
        data_fields_xml.append(
            f'<dataField name="{_attr(spec.caption)}" fld="{column}"{subtotal} '
            f'baseField="0" baseItem="0"{numfmt}/>'
        )

    page_attrs = ""
    if render.filter_indices:
        page_attrs = f' rowPageCount="{len(render.filter_indices)}" colPageCount="1"'
    location = (
        f'<location ref="{render.ref}" firstHeaderRow="{render.first_header_row}" '
        f'firstDataRow="{render.first_data_row}" firstDataCol="1"{page_attrs}/>'
    )
    count = len(render.value_captions)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<pivotTableDefinition xmlns="{_MAIN_NS}" '
        f'name="{_attr(render.plan.name)}" cacheId="{render.cache_id}" '
        'applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" '
        'applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" '
        f'dataCaption="{_attr(render.pivot.data_caption)}" updatedVersion="8" '
        'minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" '
        'createdVersion="8" indent="0" outline="1" outlineData="1" '
        'multipleFieldFilters="0">'
        f"{location}"
        f'<pivotFields count="{pivot_count}">{"".join(fields)}</pivotFields>'
        f'<rowFields count="1"><field x="{row_field}"/></rowFields>'
        f'<rowItems count="{len(render.row_items) + 1}">{row_items}'
        '<i t="grand"><x/></i></rowItems>'
        f"{col_xml}{page_xml}"
        f'<dataFields count="{count}">{"".join(data_fields_xml)}</dataFields>'
        f'<pivotTableStyleInfo name="{_attr(render.pivot.style)}" showRowHeaders="1" '
        'showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/>'
        # Without this extension Excel shows a "Values" caption row above
        # multi-field layouts; Excel's own new pivots set it.
        f"{_PIVOT_EXT_LIST}"
        "</pivotTableDefinition>"
    )


def pivot_table_rels_xml(cache_id: int, extension: str = ".xml") -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" '
        f'Target="../pivotCache/pivotCacheDefinition{cache_id}{extension}"/>'
        "</Relationships>"
    )


def append_rows(rows, feeders: list[PivotFeeder], sheet_name: str, sink: list):
    """Yield ``rows``, then the rows of the pivots that sit on this sheet.

    The feeders have seen every row once the stream ends, so the pivots can be
    rendered then; ``sink`` receives all their renders (also those placed on
    other sheets). Output below the data is the only place the pivot cells can
    go, so a pivot that starts inside the data is refused.
    """
    count = 0
    for row in rows:
        count += 1
        yield row
    renders = [feeder.result() for feeder in feeders]
    sink.extend(renders)
    own = [r for r in renders if r.plan.target_sheet.casefold() == sheet_name.casefold()]
    by_row: dict[int, dict[int, object]] = {}
    for (row, column), value in merge_cells(own).items():
        by_row.setdefault(row, {})[column] = value
    if not by_row:
        return
    if min(by_row) < count:
        raise ValueError(
            "a pivot table on the source sheet must start below the sheet's "
            f"data (first free row {count + 1}); got row {min(by_row) + 1}"
        )
    for row_idx in range(count, max(by_row) + 1):
        cells = by_row.get(row_idx)
        if not cells:
            yield []
            continue
        row = [None] * (max(cells) + 1)
        for column, value in cells.items():
            row[column] = value
        yield row
