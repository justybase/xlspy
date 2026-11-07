"""Excel table descriptions and the column names they imply."""

from __future__ import annotations
from dataclasses import dataclass
from .._table_totals_names import TOTALS_FUNCTION_NAMES
from ..formats import _unwrap_cell
from ..hyperlinks import Hyperlink
from ..refs import column_label, parse_cell_range
from ._attrs import _TABLE_NAME_RE, _is_reference_like


@dataclass(frozen=True, slots=True)
class ExcelTable:
    """An Excel table (ListObject) bound to a rectangular range.

    The range must have a header row whose cells match the column names. Table
    names must be unique in the workbook and may not read as a cell reference
    (``T1``, ``AB12``).

    Examples:
        Turn written rows into a filterable, styled table:

        ```python
        from xlspy import ExcelTable, XlsxWriter

        rows = [["Region", "Sales"], ["North", 120], ["South", 90]]
        sales = ExcelTable("Sales", "A1:B3", style_name="TableStyleMedium2")

        with XlsxWriter("table.xlsx") as workbook:
            workbook.add_sheet("Data")
            workbook.write_sheet(rows, tables=[sales], auto_filter=True)
        ```

        A totals row uses the last row of the range:

        ```python
        ExcelTable("Sales", "A1:B4", totals_row=True,
                   totals_row_functions=(None, "sum"))
        ```

    See Also:
        `PivotTable` to summarise a table, and `Slicer` to filter one.
    """

    display_name: str
    ref: str
    columns: list[str] | tuple[str, ...] | None = None
    header_row: bool = True
    totals_row: bool = False
    style_name: str = "TableStyleMedium9"
    show_first_column: bool = False
    show_last_column: bool = False
    show_row_stripes: bool = True
    show_column_stripes: bool = False
    totals_row_functions: tuple[str | None, ...] | None = None

    def __post_init__(self):
        if not isinstance(self.display_name, str) or _TABLE_NAME_RE.fullmatch(
            self.display_name
        ) is None:
            raise ValueError(
                f"invalid table display name {self.display_name!r}"
            )
        if _is_reference_like(self.display_name):
            raise ValueError(
                f"invalid table display name {self.display_name!r}: "
                "it reads as a cell reference, which Excel refuses"
            )
        first_col, _, last_col, _ = parse_cell_range(self.ref)
        if self.columns is not None:
            if not isinstance(self.columns, (list, tuple)):
                raise TypeError("columns must be a list of strings or None")
            columns = tuple(self.columns)
            for column in columns:
                if not isinstance(column, str):
                    raise TypeError("columns must be a list of strings or None")
            if self.header_row and len(columns) != last_col - first_col + 1:
                raise ValueError(
                    "len(columns) must match the table ref width"
                )
            object.__setattr__(self, "columns", columns)
        for name in (
            "header_row", "totals_row",
            "show_first_column", "show_last_column",
            "show_row_stripes", "show_column_stripes",
        ):
            value = getattr(self, name)
            if not isinstance(value, bool):
                raise TypeError(f"{name} must be bool")
        if self.totals_row:
            _, first_row, _, last_row = parse_cell_range(self.ref)
            if last_row - first_row + 1 <= (1 if self.header_row else 0) + 1:
                raise ValueError(
                    f"table {self.display_name!r}: a totals row needs at least one data row"
                )
        if self.totals_row_functions is not None:
            if not self.totals_row:
                raise ValueError("totals_row_functions needs totals_row=True")
            if not isinstance(self.totals_row_functions, (list, tuple)):
                raise TypeError("totals_row_functions must be a list or tuple")
            functions = tuple(self.totals_row_functions)
            if len(functions) != last_col - first_col + 1:
                raise ValueError(
                    "len(totals_row_functions) must match the table ref width"
                )
            for function in functions:
                if function is not None and function not in TOTALS_FUNCTION_NAMES:
                    raise ValueError(
                        f"invalid totals function {function!r}; expected one of "
                        f"{sorted(TOTALS_FUNCTION_NAMES)} or None"
                    )
            object.__setattr__(self, "totals_row_functions", functions)
        if not isinstance(self.style_name, str):
            raise TypeError("style_name must be a string")


def table_column_names(table: ExcelTable, header_cells) -> list[str]:
    """Column names for ``table``, checked against the cells of its header row.

    Excel opens a table only when each header cell holds text equal (ignoring
    case) to its column's name and the names are unique; Excel 16.0 refuses
    the whole file otherwise. Names come from ``table.columns`` or, when that
    is not given, from the header cells. ``header_cells`` is the sheet row
    holding the header (``None`` if the data never reached it).
    """
    first_col, first_row, last_col, _ = parse_cell_range(table.ref)
    width = last_col - first_col + 1
    if not table.header_row:
        if table.columns is not None:
            return list(table.columns)
        return [f"Column{i + 1}" for i in range(width)]
    row = header_cells or ()
    names = []
    for offset in range(width):
        col = first_col + offset
        ref = f"{column_label(col)}{first_row + 1}"
        cell = row[col] if col < len(row) else None
        if isinstance(cell, tuple) and len(cell) == 2 and isinstance(cell[1], str):
            cell = cell[0]  # a (value, number format) cell
        if isinstance(cell, Hyperlink):
            cell = cell.display_text  # written to the cell as plain text
        if not isinstance(cell, str) or cell == "":
            raise ValueError(
                f"table {table.display_name!r}: header cell {ref} must hold text "
                f"(found {cell!r}); Excel requires a text header for every table column"
            )
        if table.columns is not None:
            name = table.columns[offset]
            if name.casefold() != cell.casefold():
                raise ValueError(
                    f"table {table.display_name!r}: column name {name!r} does not match "
                    f"header cell {ref} ({cell!r}); Excel requires them to be equal"
                )
        else:
            name = cell
        names.append(name)
    seen = set()
    for name in names:
        if name.casefold() in seen:
            raise ValueError(
                f"table {table.display_name!r}: duplicate column name {name!r}; "
                "Excel requires unique table column names"
            )
        seen.add(name.casefold())
    return names


def capture_table_headers(rows, tables, store: dict):
    """Yield ``rows`` unchanged, checking each table's header row as it passes.

    A copy of every header row is kept in ``store`` (row index -> cells) for
    ``table_column_names``. Checking while the row streams past means a bad
    header fails before the rest of the sheet is encoded.
    """
    by_row: dict[int, list] = {}
    for table in tables:
        if table.header_row:
            by_row.setdefault(parse_cell_range(table.ref)[1], []).append(table)
    if not by_row:
        yield from rows
        return
    for row_idx, row in enumerate(rows):
        header_tables = by_row.get(row_idx)
        if header_tables is not None:

            # A (value, format) or (value, Style) header is its value.
            cells = [_unwrap_cell(cell) for cell in row]
            store[row_idx] = cells
            for table in header_tables:
                table_column_names(table, cells)
        yield row


def capture_table_totals(rows, tables, store: dict):
    """Yield ``rows`` unchanged, keeping the totals row of every table that has one.

    The XLSB writer encodes each totals cell as a label, a sum or a custom
    formula, so it needs the cells after the row streams past (``store``: row
    index -> cells).
    """
    by_row: set[int] = {
        parse_cell_range(table.ref)[3] for table in tables if table.totals_row
    }
    if not by_row:
        yield from rows
        return
    for row_idx, row in enumerate(rows):
        if row_idx in by_row:
            store[row_idx] = [_unwrap_cell(cell) for cell in row]
        yield row
