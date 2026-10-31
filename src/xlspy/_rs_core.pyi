"""Type stub for the native accelerator (rust/xlspy_core).

Callers go through ``xlspy._accel``; a None result means the pure-Python
implementation must handle the input.
"""
from collections.abc import Iterator, Sequence
from typing import Any

def set_error_type(type_: type, /) -> None: ...
def set_style_type(type_: type, /) -> None: ...
def encode_xlsb_sst(ss_list: Sequence[str], sst_all: int, sst_unique: int, /) -> bytes: ...
def read_xlsb_shared_strings(data: bytes, /) -> list[str]: ...
def calc_column_widths(rows: Any, max_cols: int, /) -> list[float]: ...
def xlsb_row_has_formatted(row: Any, /) -> bool | None: ...
def encode_xlsb_row(
    row: Sequence[Any],
    ss_dict: dict[str, int],
    ss_list: list[str],
    sst_uniq: int,
    sst_all: int,
    row_idx: int,
    style_ids: Any = None,
    fmt_xf: dict[Any, int] | None = None,
    row_header: bytes | None = None,
) -> tuple[bytes, int, int] | None: ...
def encode_xlsb_row_inline(
    row: Sequence[Any],
    row_idx: int,
    style_ids: Any = None,
    fmt_xf: dict[Any, int] | None = None,
    row_header: bytes | None = None,
) -> bytes | None: ...
def encode_xlsx_row(
    row: Sequence[Any],
    ss_dict: dict[str, int],
    ss_list: list[str],
    sst_uniq: int,
    sst_all: int,
    row_idx: int,
    use_shared_strings: bool,
    style_ids: Any = None,
    fmt_xf: dict[Any, Any] | None = None,
    row_attrs: str | None = None,
) -> tuple[str, int, int] | None: ...

class XlsbReader(Iterator[list[Any]]):
    def __iter__(self) -> XlsbReader: ...
    def __next__(self) -> list[Any]: ...

def read_xlsb_worksheet(
    data: bytes,
    shared_strings: Sequence[str],
    styles_list: Sequence[dict[str, Any]],
    date_num_fmts: set[int],
    date_1904: bool | None = None,
) -> XlsbReader | None: ...
def xlsb_load_cells(
    data: bytes,
    shared_strings: Sequence[Any],
    styles_list: Sequence[dict[str, Any]],
    date_num_fmts: set[int],
    date_1904: bool,
    cell_type: type,
    descriptors: tuple[Any, ...],
    sheet: Any,
    cells_dict: dict[int, dict[int, Any]],
    /,
) -> tuple[int, int, Any, Any, Any, int, int]: ...

class XlsxReader(Iterator[Any]):
    """Items are ``(row_number, cells)``, or ``bytes`` for a row the caller
    parses itself (then calls ``resume``)."""

    def __iter__(self) -> XlsxReader: ...
    def __next__(self) -> Any: ...
    def resume(self, row_number: int) -> None: ...
    def bounds(self) -> tuple[int, int]: ...
    def tail(self) -> bytes: ...

def read_xlsx_worksheet(
    stream: Any,
    head: bytes,
    shared_strings: Sequence[Any],
    is_date_style: Any,
    date_1904: bool | None = None,
    model: tuple[Any, ...] | None = None,
) -> XlsxReader: ...
