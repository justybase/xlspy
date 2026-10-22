"""Inserting, deleting and moving rows, columns and ranges in the model.

The semantics follow openpyxl: cells move with their value, style, comment
and hyperlink; formula text is left as it is, except that
``move_range(..., translate=True)`` shifts the relative references of the
moved formulas. Merged ranges, tables and other objects stay where they are,
as in openpyxl. Unlike openpyxl, row heights and column widths move with
their rows and columns (``insert_*``/``delete_*``), as they do in Excel.
"""

from __future__ import annotations

import re

from ..formulas import Formula
from ..refs import column_label, parse_cell_range

_MAX_ROW = 1048576
_MAX_COL = 16384
# String literals and quoted sheet names, which are never translated.
_QUOTED = re.compile(r'"(?:[^"]|"")*"|\'(?:[^\']|\'\')*\'')
_CELL_REF = re.compile(
    r"(?<![A-Za-z0-9_.])(\$?)([A-Za-z]{1,3})(\$?)([0-9]+)(?![A-Za-z0-9_(.])"
)


def translate_formula(text: str, rows: int, cols: int) -> str:
    """Shift the relative A1 references of ``text`` by ``rows`` and ``cols``.

    ``$``-anchored parts stay; a reference pushed off the sheet becomes
    ``#REF!``, as in openpyxl's Translator.
    """
    def cell(match: re.Match) -> str:
        col_abs, letters, row_abs, digits = match.groups()
        col = _column_number(letters)
        row = int(digits)
        if col is None or not 1 <= row <= _MAX_ROW:
            return match.group(0)  # not a cell reference (e.g. a name)
        if not col_abs:
            col += cols
        if not row_abs:
            row += rows
        if not (1 <= col <= _MAX_COL and 1 <= row <= _MAX_ROW):
            return "#REF!"
        return f"{col_abs}{column_label(col - 1)}{row_abs}{row}"

    parts, last = [], 0
    for quoted in _QUOTED.finditer(text):
        parts.append(_CELL_REF.sub(cell, text[last:quoted.start()]))
        parts.append(quoted.group(0))
        last = quoted.end()
    parts.append(_CELL_REF.sub(cell, text[last:]))
    return "".join(parts)


def _column_number(letters: str) -> int | None:
    number = 0
    for char in letters.upper():
        number = number * 26 + ord(char) - 64
    return number if number <= _MAX_COL else None


def _translate_array_ref(ref: str, rows: int, cols: int) -> str:
    """Shift the coordinates of a CSE array range with its anchor formula."""
    parse_cell_range(ref)

    def cell(match: re.Match) -> str:
        col = _column_number(match[2])
        row = int(match[4])
        if col is None:
            raise ValueError("array formula range leaves the sheet")
        col += cols
        row += rows
        if not (1 <= col <= _MAX_COL and 1 <= row <= _MAX_ROW):
            raise ValueError("translating the array formula range leaves the sheet")
        return f"{match[1]}{column_label(col - 1)}{match[3]}{row}"

    return _CELL_REF.sub(cell, ref)


def _check_amount(idx, amount, limit: int, what: str) -> None:
    for name, value in (("idx", idx), ("amount", amount)):
        if isinstance(value, bool) or not isinstance(value, int):
            raise TypeError(f"{what} {name} must be an int")
    if not 1 <= idx <= limit:
        raise ValueError(f"{what} idx must be 1..{limit}")
    if amount < 1:
        raise ValueError(f"{what} amount must be at least 1")


def _shifted(number: int, idx: int, amount: int, delete: bool) -> int | None:
    """Where row or column ``number`` goes; None when it is deleted."""
    if number < idx:
        return number
    if not delete:
        return number + amount
    return None if number < idx + amount else number - amount


def shift_rows(sheet, idx: int, amount: int, *, delete: bool) -> None:
    _check_amount(idx, amount, _MAX_ROW, "row")
    if not delete and sheet._cells and max(sheet._cells) >= idx and max(sheet._cells) + amount > _MAX_ROW:
        raise ValueError("the insert pushes cells past Excel's last row")
    # Whole rows move: their cell dictionaries are reused, and moved rows
    # never land on rows that stay.
    cells_by_row, vacated, moved = {}, [], []
    loaded = sheet._loaded
    for row, cells in sheet._cells.items():
        new_row = _shifted(row, idx, amount, delete)
        if new_row == row:
            cells_by_row[row] = cells
            continue
        if loaded:
            vacated.extend((row, col, cell) for col, cell in cells.items())
        if new_row is None:
            continue
        for cell in cells.values():
            cell._row = new_row
        if loaded:
            moved.extend(cells.values())
        cells_by_row[new_row] = cells
    _install(sheet, cells_by_row, vacated, moved, recompute=delete)
    if not delete and sheet._max_row >= idx:
        sheet._max_row += amount
    _shift_dimensions(sheet.row_dimensions, idx, amount, delete, int, int, _MAX_ROW)


def shift_cols(sheet, idx: int, amount: int, *, delete: bool) -> None:
    _check_amount(idx, amount, _MAX_COL, "column")
    if not delete and any(cells and max(cells) >= idx and max(cells) + amount > _MAX_COL
                          for cells in sheet._cells.values()):
        raise ValueError("the insert pushes cells past Excel's last column")
    cells_by_row, vacated, moved = {}, [], []
    loaded = sheet._loaded
    for row, cells in sheet._cells.items():
        if not cells or max(cells) < idx:
            cells_by_row[row] = cells
            continue
        new_cells = {}
        for col, cell in cells.items():
            new_col = _shifted(col, idx, amount, delete)
            if new_col == col:
                new_cells[col] = cell
                continue
            if loaded:
                vacated.append((row, col, cell))
            if new_col is not None:
                cell._column = new_col
                new_cells[new_col] = cell
                if loaded:
                    moved.append(cell)
        cells_by_row[row] = new_cells
    _install(sheet, cells_by_row, vacated, moved, recompute=delete)
    if not delete and sheet._max_column >= idx:
        sheet._max_column += amount
    _shift_dimensions(
        sheet.column_dimensions, idx, amount, delete,
        _column_number, lambda n: column_label(n - 1), _MAX_COL,
    )


def move_range(sheet, cell_range, rows: int = 0, cols: int = 0, translate: bool = False) -> None:
    for name, value in (("rows", rows), ("cols", cols)):
        if isinstance(value, bool) or not isinstance(value, int):
            raise TypeError(f"move_range {name} must be an int")
    c1, r1, c2, r2 = parse_cell_range(str(cell_range))
    first_row, last_row, first_col, last_col = r1 + 1, r2 + 1, c1 + 1, c2 + 1
    if not (1 <= first_row + rows and last_row + rows <= _MAX_ROW
            and 1 <= first_col + cols and last_col + cols <= _MAX_COL):
        raise ValueError(f"moving {cell_range} by ({rows}, {cols}) leaves the sheet")

    def target(row, col):
        if first_row <= row <= last_row and first_col <= col <= last_col:
            return row + rows, col + cols
        return row, col

    translated_formulas = {}
    if translate:
        for row, cells in sheet._cells.items():
            if not first_row <= row <= last_row:
                continue
            for col, cell in cells.items():
                if not first_col <= col <= last_col:
                    continue
                formula = cell._formula
                if isinstance(formula, Formula):
                    array_ref = (
                        _translate_array_ref(formula.array_ref, rows, cols)
                        if formula.array_ref is not None else None
                    )
                    translated_formulas[id(cell)] = Formula(
                        translate_formula(formula.text, rows, cols),
                        formula.value,
                        array_ref,
                    )
                elif isinstance(formula, str):
                    translated_formulas[id(cell)] = translate_formula(formula, rows, cols)
    moved = _relocate(sheet, target)
    for cell in moved:
        if id(cell) in translated_formulas:
            cell._formula = translated_formulas[id(cell)]


def _relocate(sheet, target) -> list:
    """Move every cell to ``target(row, col)`` (None deletes it) and return
    the cells that moved. Moved cells replace cells already at their
    destination, as in openpyxl."""
    staying, moving, vacated = {}, [], []
    for row, cells in sheet._cells.items():
        for col, cell in cells.items():
            destination = target(row, col)
            if destination == (row, col):
                staying.setdefault(row, {})[col] = cell
                continue
            vacated.append((row, col, cell))
            if destination is not None:
                new_row, new_col = destination
                if not (1 <= new_row <= _MAX_ROW and 1 <= new_col <= _MAX_COL):
                    raise ValueError("the move pushes cells past the edge of the sheet")
                moving.append((destination, cell))
    for (row, col), cell in moving:
        cell._row, cell._column = row, col
        staying.setdefault(row, {})[col] = cell
    moved = [cell for _, cell in moving]
    _install(sheet, staying, vacated if sheet._loaded else (), moved, recompute=True)
    return moved


def _install(sheet, cells_by_row, vacated, moved, *, recompute: bool) -> None:
    sheet._cells = cells_by_row
    if sheet._loaded:
        _mark_loaded_changes(sheet, vacated, moved)
    if recompute:
        _recompute_extent(sheet)


def _mark_loaded_changes(sheet, vacated, moved) -> None:
    """A loaded sheet is saved cell by cell: record what changed, and clear
    the cells that were left behind."""
    from ._model import Cell

    for row, col, cell in vacated:
        coordinate = f"{column_label(col - 1)}{row}"
        if cell._comment is not None:
            sheet._dirty_comments.add(coordinate)
        if cell._hyperlink is not None:
            sheet._dirty_hyperlinks.add(coordinate)
        if sheet._cells.get(row, {}).get(col) is None:
            blank = Cell(sheet, row, col)
            sheet._cells.setdefault(row, {})[col] = blank
            blank._mark_dirty()
            sheet._dirty_styles.add(coordinate)
    for cell in moved:
        cell._mark_dirty()
        sheet._dirty_styles.add(cell.coordinate)
        if cell._comment is not None:
            sheet._dirty_comments.add(cell.coordinate)
        if cell._hyperlink is not None:
            sheet._dirty_hyperlinks.add(cell.coordinate)


def _recompute_extent(sheet) -> None:
    max_row = max_col = 0
    for row, cells in sheet._cells.items():
        for col, cell in cells.items():
            if (cell._value is not None or cell._formula is not None or cell._style_id
                    or cell._comment is not None or cell._hyperlink is not None):
                max_row = max(max_row, row)
                max_col = max(max_col, col)
    sheet._max_row, sheet._max_column = max_row, max_col


def _shift_dimensions(holder, idx, amount, delete, to_number, to_key, limit) -> None:
    items = list(dict.items(holder))
    dict.clear(holder)
    for key, dimension in items:
        number = to_number(key)
        if number >= idx:
            if delete:
                if number < idx + amount:
                    continue
                number -= amount
            else:
                number += amount
        if number > limit:
            continue  # pushed off the sheet
        new_key = to_key(number)
        dimension.index = new_key
        dict.__setitem__(holder, new_key, dimension)
