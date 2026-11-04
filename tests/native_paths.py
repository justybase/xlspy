"""Which path an XLSB write takes: the native row encoder, or Python.

``count_paths()`` counts, while a workbook is written, the rows the Rust encoder
took, the rows Python wrote cell by cell, and the per-cell calls into Python
(``scalar_hook``: a value the encoder does not know, ``formula_calls``: a formula
compiled in Python). ``SCENARIOS`` is the matrix that ``test_native_write_coverage.py``
asserts and ``benchmarks/bench_xlsb_paths.py`` prints.
"""

from __future__ import annotations

import datetime as dt
import pathlib
from collections import Counter
from collections.abc import Callable, Iterator
from contextlib import ExitStack, contextmanager
from dataclasses import dataclass, field
from decimal import Decimal
from typing import Any
from unittest import mock

from xlspy import CellRichText, Formula, Hyperlink, Style, XlsbWriter
from xlspy import _xlsb_writer_bin as writer_module

DATE = dt.date(2024, 5, 6)
DATETIME = dt.datetime(2024, 5, 6, 7, 8, 9)
HEADER = Style(bold=True, horizontal="center", fill="#DDEBF7")
MONEY = Style(number_format="#,##0.00")


@contextmanager
def count_paths() -> Iterator[Counter[str]]:
    """Counts native rows (``native``), Python rows (``python_rows``) and per-cell calls
    into Python (``scalar_hook``, ``formula_calls``) of the XLSB writers run inside."""
    counts: Counter[str] = Counter()
    cls = next(v for v in vars(writer_module).values() if isinstance(v, type) and hasattr(v, "_write_python_row"))
    with ExitStack() as stack:
        for name in ("encode_xlsb_row", "encode_xlsb_row_inline"):
            stack.enter_context(mock.patch.object(writer_module, name, _counting(getattr(writer_module, name), counts, "native", True)))
        stack.enter_context(mock.patch.object(cls, "_write_python_row", _counting_method(cls._write_python_row, counts, "python_rows")))
        stack.enter_context(mock.patch.object(writer_module, "native_cell_value", _counting(writer_module.native_cell_value, counts, "scalar_hook")))
        stack.enter_context(mock.patch.object(writer_module, "formula_cell", _counting(writer_module.formula_cell, counts, "formula_calls")))
        yield counts


def _counting(function: Callable[..., Any], counts: Counter[str], key: str, only_results: bool = False) -> Callable[..., Any]:
    def wrapper(*args: Any, **kwargs: Any) -> Any:
        result = function(*args, **kwargs)
        if not only_results or result is not None:
            counts[key] += 1
        return result

    return wrapper


def _counting_method(function: Callable[..., Any], counts: Counter[str], key: str) -> Callable[..., Any]:
    def wrapper(self: Any, *args: Any, **kwargs: Any) -> Any:
        counts[key] += 1
        return function(self, *args, **kwargs)

    return wrapper


@dataclass
class Scenario:
    name: str
    make: Callable[[int], Any]  # rows (an iterable of lists) or, with ``frame``, a DataFrame
    options: dict[str, Any] = field(default_factory=dict)
    shared_strings: bool = True
    frame: bool = False
    # The most rows Python may write (a first row per new style variant is expected).
    max_python_rows: int = 0
    # The most per-cell calls into Python; None: not asserted.
    max_scalar_hook: int | None = 0
    gap: str | None = None  # why the scenario is expected to be slow today


def _rows(n: int, head: list[Any], cell: Callable[[int], list[Any]]) -> Iterator[list[Any]]:
    yield head
    for i in range(n):
        yield cell(i)


def _numpy_rows(n: int) -> Iterator[list[Any]]:
    import numpy as np

    return _rows(n, ["i", "f", "b", "f32"], lambda i: [np.int64(i), np.float64(i) / 3, np.bool_(i % 2), np.float32(i)])


def _frame(n: int, nullable: bool = False) -> Any:
    import numpy as np
    import pandas as pd

    frame = pd.DataFrame({
        "i": np.arange(n, dtype="int64"), "f": np.arange(n) / 3, "b": np.arange(n) % 2 == 0,
        "s": [f"s{k % 50}" for k in range(n)], "d": pd.date_range("2024-01-01", periods=n, freq="min"),
    })
    if nullable:
        frame["i"] = frame["i"].astype("Int64")
        frame.loc[::7, "i"] = pd.NA
        frame.loc[::11, "d"] = pd.NaT
        frame["f"] = frame["f"].where(frame.index % 5 != 0)
    return frame


SCENARIOS: list[Scenario] = [
    Scenario("plain", lambda n: _rows(n, list("abcd"), lambda i: [f"s{i % 50}", i, i * 0.5, i % 2 == 0])),
    Scenario("plain, inline strings", lambda n: _rows(n, list("abcd"), lambda i: [f"s{i % 50}", i, i * 0.5, i % 2 == 0]), shared_strings=False),
    Scenario(
        "two header rows, column formats",
        lambda n: _rows(n, ["Wartosc", "Wartosc", "Ilosc", "Nazwa"], lambda i: [i * 1234.5678, i * 0.3333, i, f"opis {i % 100}"]),
        {"header_rows": 2, "header_style": HEADER, "styles": {"A:B": MONEY, "C:C": Style(number_format="#,##0")}},
    ),
    Scenario("(value, format) cells", lambda n: _rows(n, list("xyz"), lambda i: [(i * 0.5, "0.00"), (i, "#,##0"), (i * 2.0, MONEY)]), max_python_rows=3),
    Scenario("dates", lambda n: _rows(n, ["d", "t"], lambda i: [DATE, DATETIME])),
    Scenario("dates, styled columns", lambda n: _rows(n, ["d", "t"], lambda i: [DATE, DATETIME]), {"styles": {"A:B": Style(horizontal="center")}}, max_python_rows=2),
    Scenario("Decimal", lambda n: _rows(n, ["x"], lambda i: [Decimal("1.25") + i])),
    Scenario("blanks", lambda n: _rows(n, list("abc"), lambda i: [None, i, None] if i % 2 else [i, None, "x"])),
    Scenario("dataframe", _frame, frame=True),
    Scenario("dataframe with NA and NaT", lambda n: _frame(n, nullable=True), frame=True),
    Scenario("formulas", lambda n: _rows(n, ["x", "f"], lambda i: [i, Formula(f"A{i + 2}*2")]), max_scalar_hook=0),
    Scenario("hyperlinks", lambda n: _rows(n, ["x", "l"], lambda i: [i, Hyperlink(f"https://example.com/{i % 20}", f"l{i % 20}")]), max_python_rows=2),
    Scenario("numpy scalars", _numpy_rows, max_scalar_hook=None, gap="a numpy int is no int: a Python call per cell"),
    Scenario("rich text", lambda n: _rows(n, ["x", "r"], lambda i: [i, CellRichText(["a", "b"])]), max_python_rows=1 << 30, max_scalar_hook=None, gap="the whole row goes to Python"),
]


def write(scenario: Scenario, rows: int, path: pathlib.Path) -> Counter[str]:
    """Writes the scenario to ``path`` and returns the counts of its paths."""
    data = scenario.make(rows)
    with count_paths() as counts:
        with XlsbWriter(str(path), use_shared_strings=scenario.shared_strings) as writer:
            writer.add_sheet("S")
            if scenario.frame:
                writer.write_dataframe(data, **scenario.options)
            else:
                writer.write_sheet(data, **scenario.options)
    return counts
