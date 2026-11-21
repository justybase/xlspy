"""Turn indexed rows into the row lists both readers return."""

from __future__ import annotations

from collections.abc import Iterable, Iterator, Sequence
from typing import Any


def rows_in_place(
    indexed_rows: Iterable[tuple[int, Sequence[tuple[int, Any]]]],
) -> Iterator[list[Any]]:
    """Lists from ``(1-based row, [(0-based column, value), ...])``, each value at its
    own position: what a row means, for both formats. Rows absent from the file
    between present rows are emitted as empty rows, so row positions match the sheet.
    The first row with cells fixes the width that later rows are padded to; rows
    before it are empty lists, whether the file has a record for them or not."""
    first_width = None
    next_row = 1
    for row_number, cells in indexed_rows:
        for _ in range(row_number - next_row):
            yield [None] * (first_width or 0)
        next_row = row_number + 1
        row = [None] * max(first_width or 0, max((col for col, _ in cells), default=-1) + 1)
        for col, value in cells:
            row[col] = value
        if first_width is None and row:
            first_width = len(row)
        yield row
