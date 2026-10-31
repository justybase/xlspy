"""Type-level checks for ``src/xlspy/_rs_core.pyi``.

Nothing here runs. ``mypy --warn-unused-ignores tests/typing/stub_usage.py``
checks it: each ``# type: ignore[arg-type]`` marks a call the native module
rejects at runtime (it tests for exact list, dict or set values), so if the stub
ever accepts such an argument again, mypy reports the ignore as unused.
"""

from xlspy import _rs_core as core


def _accepted_calls() -> None:
    core.set_error_type(ValueError)
    core.set_invalid_file_type(ValueError)
    core.encode_xlsb_sst(["a", "b"], 2, 2)
    core.read_xlsb_shared_strings(b"")
    core.read_xlsx_shared_strings(b"")
    core.calc_column_widths([["a", 1.5]], 4)
    core.xlsb_row_has_formatted(["a"])
    core.encode_xlsb_row(["a"], {}, [], 0, 0, 0)
    core.encode_xlsb_row(["a"], {}, [], 0, 0, 0, [1, 2], {("s", 1): 3}, b"\x00")
    core.encode_xlsb_row_inline(["a"], 0, 5)
    core.encode_xlsx_row(["a"], {}, [], 0, 0, 0, True, None, None, ' s="1"')

    reader = core.read_xlsb_worksheet(b"", [], [], set(), False)
    if reader is not None:
        for row in reader:
            row_cells: list[object] = row
            _ = row_cells

    xreader = core.read_xlsx_worksheet(_stream(), b"<worksheet/>", [], [False, 1])
    xreader.resume(5)
    first_row, last_row = xreader.bounds()
    _ = (first_row, last_row, xreader.tail())

    loaded = core.xlsb_load_cells(b"", [], [], set(), False, dict, (), None, {})
    max_row, max_column, formulas, headers, groups, start, end = loaded
    _ = (max_row, max_column, formulas, headers, groups, start, end)


def _rejected_calls() -> None:
    # Tuples are not accepted where the native code needs exact lists.
    core.encode_xlsb_sst(("a",), 1, 2)  # type: ignore[arg-type]
    core.encode_xlsb_row(("a",), {}, [], 0, 0, 0)  # type: ignore[arg-type]
    core.encode_xlsx_row(("a",), {}, [], 0, 0, 0, True)  # type: ignore[arg-type]
    # read_xlsb_worksheet requires a set, not a frozenset.
    core.read_xlsb_worksheet(b"", [], [], frozenset())  # type: ignore[arg-type]


def _stream() -> "_Readable":
    return _Readable()


class _Readable:
    def read(self, size: int = -1) -> bytes:
        return b""
