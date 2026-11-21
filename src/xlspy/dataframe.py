"""Optional pandas / polars integration for the xlspy writers.

Writing a DataFrame is a thin conversion layer on top of ``write_sheet``: an
existing DataFrame is turned into rows of plain Python values (``str``,
``int``, ``float``, ``bool``, ``Decimal``, ``datetime``, ``date``, ``time``,
``timedelta``, ``None`` and ``(value, format)`` tuples) that the XLSB/XLSX
writers already understand.

Neither pandas nor polars is required by xlspy. This module never imports
either of them at import time -- they are imported lazily only when a
DataFrame-like object is actually passed in. That is why ``import xlspy``
works on a bare installation with no third-party dependencies.
"""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from itertools import chain, islice
import importlib.util
from typing import TYPE_CHECKING, Any, TypeVar, cast

from ._dataframe_rows import dataframe_to_rows as dataframe_to_rows
from .writer import XlsbWriter
from .xlsx_writer import XlsxWriter

if TYPE_CHECKING:
    from ._metadata import SheetProtection
    from ._output import WriteTarget
    from .styles import Style

_Target = TypeVar("_Target", bound="WriteTarget")


def to_xlsx(
    df: Any,
    path: _Target,
    sheet_name: str = "Sheet1",
    *,
    include_index: bool = False,
    na_rep: Any = None,
    formats: dict[str, str] | None = None,
    header_style: Style | None = None,
    styles: dict | None = None,
    merges: list[str] | tuple[str, ...] | None = None,
    conditional_formats: list | tuple | None = None,
    tables: Any = None,
    column_widths: dict | None = None,
    autofit_rows: int | None = 1000,
    autofit_padding: float = 0.0,
    hidden_columns: Any = None,
    row_heights: dict | None = None,
    hidden_rows: Any = None,
    row_groups: Any = None,
    column_groups: Any = None,
    freeze_panes: str | None = None,
    auto_filter: bool | None = None,
    charts: Any = None,
    comments: Any = None,
    threaded_comments: Any = None,
    sheet_protection: SheetProtection | None = None,
    **writer_kwargs: Any,
) -> _Target:
    """Write a DataFrame straight to an ``.xlsx`` file and return ``path``.

    A failed conversion leaves ``path`` as it was: the writer stages the file
    and publishes it only on success.

    Examples:
        ```python
        import pandas as pd

        from xlspy import to_xlsx

        df = pd.DataFrame({"name": ["Alice", "Bob"], "age": [25, 30]})
        to_xlsx(df, "people.xlsx")
        ```

        All sheet options of ``write_dataframe`` are accepted as keywords:

        ```python
        to_xlsx(df, "people.xlsx", freeze_panes="A2", include_index=True)
        ```

    See Also:
        `to_xlsb` for the binary format, and `XlsxWriter.write_dataframe` for
        writing into a workbook that already has sheets and objects.
    """
    with XlsxWriter(path, **writer_kwargs) as writer:
        writer.add_sheet(sheet_name)
        writer.write_dataframe(
            df,
            include_index=include_index,
            na_rep=na_rep,
            formats=formats,
            header_style=header_style,
            styles=styles,
            merges=merges,
            conditional_formats=conditional_formats,
            tables=tables,
            column_widths=column_widths,
            autofit_rows=autofit_rows,
            autofit_padding=autofit_padding,
            hidden_columns=hidden_columns,
            row_heights=row_heights,
            hidden_rows=hidden_rows,
            row_groups=row_groups,
            column_groups=column_groups,
            freeze_panes=freeze_panes,
            auto_filter=auto_filter,
            charts=charts,
            comments=comments,
            threaded_comments=threaded_comments,
            sheet_protection=sheet_protection,
        )
    return path


def to_xlsb(
    df: Any,
    path: _Target,
    sheet_name: str = "Sheet1",
    *,
    include_index: bool = False,
    na_rep: Any = None,
    formats: dict[str, str] | None = None,
    header_style: Style | None = None,
    styles: dict | None = None,
    merges: list[str] | tuple[str, ...] | None = None,
    column_widths: dict | None = None,
    autofit_rows: int | None = 1000,
    autofit_padding: float = 0.0,
    hidden_columns: Any = None,
    row_heights: dict | None = None,
    hidden_rows: Any = None,
    row_groups: Any = None,
    column_groups: Any = None,
    freeze_panes: str | None = None,
    auto_filter: bool | None = None,
    charts: Any = None,
    tables: Any = None,
    conditional_formats: Any = None,
    comments: Any = None,
    sheet_protection: SheetProtection | None = None,
    **writer_kwargs: Any,
) -> _Target:
    """Write a DataFrame straight to an ``.xlsb`` file and return ``path``.

    A failed conversion leaves ``path`` as it was: the writer stages the file
    and publishes it only on success.

    Examples:
        ```python
        import polars as pl

        from xlspy import to_xlsb

        df = pl.DataFrame({"name": ["Alice", "Bob"], "age": [25, 30]})
        to_xlsb(df, "people.xlsb")
        ```

        The result is a binary XLSB workbook with a single ``Sheet1``. Use
        `XlsbWriter.write_dataframe` to add more sheets, styles or charts.

    See Also:
        `to_xlsx` for maximum compatibility, and `XlsbWriter` for the full
        writer API.
    """
    with XlsbWriter(path, **writer_kwargs) as writer:
        writer.add_sheet(sheet_name)
        writer.write_dataframe(
            df,
            include_index=include_index,
            na_rep=na_rep,
            formats=formats,
            header_style=header_style,
            styles=styles,
            merges=merges,
            column_widths=column_widths,
            autofit_rows=autofit_rows,
            autofit_padding=autofit_padding,
            hidden_columns=hidden_columns,
            row_heights=row_heights,
            hidden_rows=hidden_rows,
            row_groups=row_groups,
            column_groups=column_groups,
            freeze_panes=freeze_panes,
            auto_filter=auto_filter,
            charts=charts,
            tables=tables,
            conditional_formats=conditional_formats,
            comments=comments,
            sheet_protection=sheet_protection,
        )
    return path


def _frame_engine(engine: str) -> str:
    if engine not in ("auto", "pandas", "polars"):
        raise ValueError("engine must be 'auto', 'pandas' or 'polars'")
    if engine == "auto":
        if importlib.util.find_spec("pandas") is not None:
            return "pandas"
        if importlib.util.find_spec("polars") is not None:
            return "polars"
        raise ImportError("read_dataframe needs pandas or polars; install one of them")
    return engine


def _column_name(value: Any, index: int) -> str:
    if value is None or value == "":
        return f"column_{index}"
    return str(value)


def _dedupe_columns(columns: Sequence[str]) -> list[str]:
    seen: dict[str, int] = {}
    result = []
    for name in columns:
        count = seen.get(name, 0)
        seen[name] = count + 1
        result.append(name if count == 0 else f"{name}.{count}")
    return result


def _resolve_usecols(usecols: Sequence[str | int], columns: Sequence[str]) -> list[int]:
    indices = []
    for item in usecols:
        if isinstance(item, int) and not isinstance(item, bool):
            index = item + len(columns) if item < 0 else item
            if not 0 <= index < len(columns):
                raise ValueError(f"usecols index {item} is out of range for {len(columns)} columns")
            indices.append(index)
        elif isinstance(item, str):
            try:
                indices.append(columns.index(item))
            except ValueError:
                raise ValueError(f"usecols column {item!r} not found") from None
        else:
            raise TypeError("usecols entries must be column names or indexes")
    return indices


def _read_sheet(
    reader: Any,
    name: str,
    *,
    header: bool | Sequence[str],
    usecols: Sequence[str | int] | None,
    nrows: int | None,
    dtype: Mapping[str, Any] | None,
    engine: str,
    data_only: bool,
) -> Any:
    rows = reader.get_rows(name, data_only=data_only)
    first = next(rows, None)
    if first is None:
        columns: list[str] = []
        data: list[list[Any]] = []
    elif header is True:
        columns = _dedupe_columns([_column_name(value, index) for index, value in enumerate(first)])
        data = list(islice(rows, nrows)) if nrows is not None else list(rows)
    elif header is False:
        columns = [f"column_{index}" for index in range(len(first))]
        source = chain([first], rows)
        data = list(islice(source, nrows)) if nrows is not None else list(source)
    else:
        columns = [str(value) for value in header]
        source = chain([first], rows)
        data = list(islice(source, nrows)) if nrows is not None else list(source)

    width = max([len(columns), *(len(row) for row in data)], default=len(columns))
    if width > len(columns):
        columns.extend(f"column_{index}" for index in range(len(columns), width))
    data = [list(row) + [None] * (width - len(row)) for row in data]

    if usecols is not None:
        indices = _resolve_usecols(usecols, columns)
        columns = [columns[index] for index in indices]
        data = [[row[index] for index in indices] for row in data]

    if engine == "pandas":
        import pandas as pd

        frame = pd.DataFrame(data, columns=columns)
        if dtype:
            frame = frame.astype(dtype)
        return frame

    import polars as pl

    if data:
        frame = pl.DataFrame(data, schema=columns, orient="row", infer_schema_length=None)
    else:
        frame = pl.DataFrame({column: [] for column in columns})
    if dtype:
        frame = cast(Any, frame).cast(dtype)
    return frame


def read_dataframe(
    source: Any,
    sheet_name: str | int | None = 0,
    *,
    header: bool | Sequence[str] = True,
    usecols: Sequence[str | int] | None = None,
    nrows: int | None = None,
    dtype: Mapping[str, Any] | None = None,
    engine: str = "auto",
    data_only: bool = True,
) -> Any:
    """Read a worksheet straight into a pandas or polars DataFrame.

    ``source`` is a path, a binary file object or an open `ExcelReader`.
    ``sheet_name`` is a name, a zero-based index, or None for every sheet, which
    returns a ``{name: DataFrame}`` mapping. The reader streams rows, so a large
    sheet does not need to be held as a list of lists first.

    Args:
        header: True uses the first row as column names, False generates
            ``column_0``, ``column_1``..., and a sequence supplies the names and
            treats the first row as data.
        usecols: Column names or zero-based indexes to keep.
        nrows: Stop after this many data rows.
        dtype: Passed to ``DataFrame.astype`` (pandas) or ``DataFrame.cast``
            (polars).
        engine: ``"pandas"``, ``"polars"`` or ``"auto"`` (pandas first).
        data_only: Read cached formula results instead of formula text.

    Examples:
        ```python
        from xlspy import read_dataframe

        frame = read_dataframe("report.xlsb", sheet_name="Sales")
        frames = read_dataframe("report.xlsx", sheet_name=None)
        ```
    """
    if not isinstance(sheet_name, (str, int, type(None))) or isinstance(sheet_name, bool):
        raise TypeError("sheet_name must be a name, an index or None")
    if isinstance(sheet_name, int) and sheet_name < 0:
        raise ValueError("sheet_name index must be zero or greater")
    if nrows is not None and (not isinstance(nrows, int) or isinstance(nrows, bool) or nrows < 0):
        raise ValueError("nrows must be a non-negative int or None")
    if not isinstance(header, (bool, Sequence)) or isinstance(header, str):
        raise TypeError("header must be a bool or a sequence of column names")
    if usecols is not None and (not isinstance(usecols, Sequence) or isinstance(usecols, str)):
        raise TypeError("usecols must be a sequence of names or indexes")
    if dtype is not None and not isinstance(dtype, Mapping):
        raise TypeError("dtype must be a mapping of column names to dtypes")
    chosen_engine = _frame_engine(engine)

    from .reader import ExcelReader

    def read_all(reader: ExcelReader) -> Any:
        names = reader.get_sheet_names()
        options: dict[str, Any] = {
            "header": header,
            "usecols": usecols,
            "nrows": nrows,
            "dtype": dtype,
            "engine": chosen_engine,
            "data_only": data_only,
        }
        if sheet_name is None:
            return {name: _read_sheet(reader, name, **options) for name in names}
        if isinstance(sheet_name, int):
            try:
                name = names[sheet_name]
            except IndexError:
                raise IndexError(
                    f"sheet index {sheet_name} is out of range; the workbook has {len(names)} sheets"
                ) from None
        else:
            name = sheet_name
            if name not in names:
                raise KeyError(f"sheet {name!r} not found")
        return _read_sheet(reader, name, **options)

    if isinstance(source, ExcelReader):
        return read_all(source)
    with ExcelReader(source) as reader:
        return read_all(reader)
