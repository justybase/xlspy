"""Type stub for the native accelerator (rust/xlspy_core).

Callers go through ``xlspy._accel``; a None result means the pure-Python
implementation must handle the input.

Argument types are the ones the Rust code really accepts. The readers and
writers test for exact ``list``/``dict``/``set``/``bytes`` values, so the
stub uses those, not the wider ``Sequence``/``Mapping`` protocols.
"""

from collections.abc import Callable, Iterator, Sequence
from typing import Any, final

from _typeshed import ReadableBuffer, SupportsRead

class PanicException(BaseException):
    """A panic in the Rust code (pyo3_runtime.PanicException)."""

__all__ = [
    "__doc__",
    "PanicException",
    "set_error_type",
    "set_style_type",
    "set_formula_type",
    "set_hyperlink_type",
    "set_invalid_file_type",
    "encode_xlsb_sst",
    "read_xlsb_shared_strings",
    "read_xlsx_shared_strings",
    "calc_column_widths",
    "xlsb_row_has_formatted",
    "encode_xlsb_row",
    "encode_xlsb_row_inline",
    "encode_xlsx_row",
    "read_xlsb_worksheet",
    "XlsbReader",
    "xlsb_load_cells",
    "read_xlsx_worksheet",
    "XlsxReader",
    "WidthMeter",
    "general_number_text",
]

def set_error_type(type_: type) -> None: ...
def set_style_type(type_: type) -> None: ...
def set_formula_type(type_: type) -> None: ...
def set_hyperlink_type(type_: type) -> None: ...
def set_invalid_file_type(type_: type) -> None: ...
def encode_xlsb_sst(ss_list: list[str], sst_all: int, sst_unique: int) -> bytes: ...
def read_xlsb_shared_strings(data: ReadableBuffer) -> list[str]: ...
def read_xlsx_shared_strings(data: ReadableBuffer) -> list[str] | None: ...
def calc_column_widths(rows: list[list[Any]], max_cols: int) -> list[float]: ...
def xlsb_row_has_formatted(row: list[Any]) -> bool | None: ...
# cell_hook(cell, column, xf, row): NotImplemented (row fallback), the encoded
# cell as bytes, a (value, xf) pair, or the value to write in the cell's xf.
_CellHook = Callable[[Any, int, int, int], Any]

def encode_xlsb_row(
    row: Sequence[Any],
    ss_dict: dict[str, int],
    ss_list: list[str],
    sst_uniq: int,
    sst_all: int,
    row_idx: int,
    style_ids: int | list[int] | tuple[int, ...] | None = None,
    fmt_xf: dict[Any, int] | None = None,
    row_header: bytes | None = None,
    cell_hook: _CellHook | None = None,
) -> tuple[bytes, int, int] | None: ...
def encode_xlsb_row_inline(
    row: Sequence[Any],
    row_idx: int,
    style_ids: int | list[int] | tuple[int, ...] | None = None,
    fmt_xf: dict[Any, int] | None = None,
    row_header: bytes | None = None,
    cell_hook: _CellHook | None = None,
) -> bytes | None: ...
def encode_xlsx_row(
    row: Sequence[Any],
    ss_dict: dict[str, int],
    ss_list: list[str],
    sst_uniq: int,
    sst_all: int,
    row_idx: int,
    use_shared_strings: bool,
    style_ids: int | list[int] | tuple[int, ...] | None = None,
    fmt_xf: dict[Any, Any] | None = None,
    row_attrs: str | None = None,
    link_xf: int | None = None,
    write_r: bool = True,
    cell_hook: _CellHook | None = None,
) -> tuple[bytes, int, int, list[tuple[int, int, Any]]] | None: ...
@final
class XlsbReader(Iterator[list[Any]]):
    def __iter__(self) -> XlsbReader: ...
    def __next__(self) -> list[Any]: ...

def read_xlsb_worksheet(
    data: ReadableBuffer,
    shared_strings: list[str],
    styles_list: list[dict[str, Any]],
    date_num_fmts: set[int],
    date_1904: bool | None = None,
) -> XlsbReader | None: ...
def xlsb_load_cells(
    data: ReadableBuffer,
    shared_strings: list[Any],
    styles_list: list[dict[str, Any]],
    date_num_fmts: set[int],
    date_1904: bool,
    cell_type: type,
    descriptors: tuple[Any, ...],
    sheet: Any,
    cells_dict: dict[int, dict[int, Any]],
) -> tuple[
    int,
    int,
    list[tuple[int, int, bytes]],
    list[tuple[int, bytes]],
    list[tuple[int, bytes]],
    int,
    int,
]: ...
@final
class XlsxReader(Iterator[Any]):
    """Items are ``(row_number, cells)``, or ``bytes`` for a row the caller
    parses itself (then calls ``resume``)."""

    def __iter__(self) -> XlsxReader: ...
    def __next__(self) -> Any: ...
    def resume(self, row_number: int) -> None: ...
    def next_row(self) -> int | None: ...
    def bounds(self) -> tuple[int, int]: ...
    def tail(self) -> bytes: ...

def read_xlsx_worksheet(
    stream: SupportsRead[bytes],
    head: bytes,
    shared_strings: list[Any],
    is_date_style: list[bool | int],
    date_1904: bool | None = None,
    model: tuple[Any, ...] | None = None,
) -> XlsxReader: ...
@final
class WidthMeter:
    """Column widths in pixels of the rows the native code measures (see
    ``xlspy._widths.ColumnWidthMeter``)."""

    def __new__(
        cls,
        fallback: Callable[[Any, Any, bool], int],
        kind_of: Callable[[Any], tuple[int, tuple[Any, ...] | None]],
        constant_types: list[type],
        overlay_kind: Callable[[Any, Any, bool], tuple[int, tuple[Any, ...] | None]] | None = None,
    ) -> WidthMeter: ...
    def add_kind(
        self, id: int, spec: tuple[dict[str, int], int, bool, bool, Callable[[str], int], tuple[Any, ...] | None, int],
    ) -> None: ...
    def add(
        self, row: list[Any] | tuple[Any, ...], styles: dict[int, Any] | None = None,
        header: bool = False, keep: bool = False,
    ) -> list[int] | None: ...
    def best(self) -> list[int]: ...

def general_number_text(value: object) -> str | None: ...
