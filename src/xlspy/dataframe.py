# -*- coding: utf-8 -*-
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

import datetime
import importlib
import os
from decimal import Decimal
from typing import Any, Iterator


_MISSING = object()
_SECONDS_PER_DAY = 86400.0

_import_cache: dict[str, Any] = {}


def _lazy_import(name: str):
    """Import a module once, returning ``None`` when it is not installed."""
    if name not in _import_cache:
        try:
            _import_cache[name] = importlib.import_module(name)
        except ImportError:
            _import_cache[name] = None
    return _import_cache[name]


def _get_pandas():
    return _lazy_import("pandas")


def _get_numpy():
    return _lazy_import("numpy")


def _is_pandas_frame(obj) -> bool:
    """Duck-typed check that avoids importing pandas."""
    return hasattr(obj, "itertuples") and hasattr(obj, "columns") and hasattr(obj, "dtypes")


def _is_polars_frame(obj) -> bool:
    """Duck-typed check that avoids importing polars."""
    return hasattr(obj, "iter_rows") and hasattr(obj, "columns") and not hasattr(obj, "itertuples")


def _flatten_label(label) -> str:
    """Turn a column label into a string, joining MultiIndex tuples with a space."""
    if isinstance(label, tuple):
        return " ".join(str(part) for part in label)
    return str(label)


def _normalize_cell(value, pd, np):
    """Map a cell value to a type the writers understand.

    Returns ``_MISSING`` for every flavour of missing value (``None``,
    ``NaN``, ``pd.NA``, ``pd.NaT``, ``np.datetime64('NaT')``) so the caller
    can substitute ``na_rep``. ``pd`` and ``np`` are the (possibly ``None``)
    pandas/numpy modules, resolved once per conversion by the caller.
    """
    if value is None:
        return _MISSING

    # pd.NA / pd.NaT are handled by identity before the datetime checks:
    # pd.NaT *is* an instance of datetime.datetime, so it must be caught first.
    if pd is not None and (value is pd.NA or value is pd.NaT):
        return _MISSING

    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        return value
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        return _MISSING if value != value else value  # NaN
    if isinstance(value, datetime.datetime):  # includes pd.Timestamp
        try:
            return value.to_pydatetime()
        except AttributeError:
            return value
    if isinstance(value, datetime.date):
        return value
    if isinstance(value, datetime.timedelta):  # includes pd.Timedelta
        # Excel stores durations as a fraction of a day.
        return value.total_seconds() / _SECONDS_PER_DAY
    if isinstance(value, datetime.time):
        # polars "Time" and object columns may hold plain times.
        seconds = (
            value.hour * 3600
            + value.minute * 60
            + value.second
            + value.microsecond / 1_000_000
        )
        return seconds / _SECONDS_PER_DAY
    if isinstance(value, Decimal):
        return value

    if np is not None and isinstance(value, np.generic):
        return _normalize_numpy(value, np)

    return value


def _normalize_numpy(value, np):
    """Convert a numpy scalar, mapping every NaT/NaN flavour to ``_MISSING``."""
    if isinstance(value, np.datetime64):
        if np.isnat(value):
            return _MISSING
        return value.astype("datetime64[us]").item()
    if isinstance(value, np.timedelta64):
        if np.isnat(value):
            return _MISSING
        delta = value.astype("timedelta64[us]").item()
        # Excel stores durations as a fraction of a day.
        return delta.total_seconds() / _SECONDS_PER_DAY
    if isinstance(value, np.floating):
        # np.float16 / np.float32 are not subclasses of float, so the NaN
        # check above does not cover them.
        return _MISSING if np.isnan(value) else value.item()

    result = value.item()
    return _MISSING if result is None else result


def _pandas_rows(df, include_index: bool):
    headers = [_flatten_label(col) for col in df.columns]

    data_rows = df.itertuples(index=False, name=None)

    if not include_index:
        return headers, data_rows

    # pandas 3.x returns a MultiIndex as a single tuple inside itertuples, so
    # the index is prepended separately to keep levels aligned with the header.
    nlevels = getattr(df.index, "nlevels", 1)
    names = list(getattr(df.index, "names", [None])) or [None]
    index_headers = []
    for position in range(nlevels):
        name = names[position] if position < len(names) else None
        if name is None:
            name = "index" if nlevels == 1 else f"level_{position}"
        index_headers.append(str(name))
    headers = index_headers + headers

    def with_index():
        for index_key, data_row in zip(df.index, data_rows):
            levels = index_key if nlevels > 1 else (index_key,)
            yield tuple(levels) + data_row

    return headers, with_index()


def _polars_rows(df, include_index: bool):
    if include_index:
        raise ValueError(
            "polars DataFrames have no index, so include_index=True is not supported."
        )
    headers = [str(col) for col in df.columns]
    return headers, df.iter_rows()


def _format_lookup(headers, formats):
    if not formats:
        return [None] * len(headers)
    return [formats.get(header) for header in headers]


def _convert_cell(cell, fmt_string, na_rep, pd, np):
    value = _normalize_cell(cell, pd, np)
    if value is _MISSING:
        return na_rep
    if fmt_string is not None:
        return (value, fmt_string)
    return value


def dataframe_to_rows(
    df,
    include_index: bool = False,
    na_rep=None,
    formats: dict[str, str] | None = None,
) -> Iterator[list]:
    """Yield a pandas or polars DataFrame as rows ready for ``write_sheet``.

    Parameters
    ----------
    df:
        A ``pandas.DataFrame`` or ``polars.DataFrame``.
    include_index:
        pandas only -- prepend the index (all levels of a MultiIndex) as
        leading columns. polars has no index and raises ``ValueError``.
    na_rep:
        Value written for missing cells (``None``/``NaN``/``NaT``/``NA``).
        Defaults to ``None``, i.e. an empty cell.
    formats:
        Optional mapping of column name to a format string (see ``xlspy.F``),
        e.g. ``{"date": F.DATE_ISO, "price": F.CURRENCY_PLN}``.

    The result is a generator, so the XLSB writer keeps streaming row by row.
    """
    if _is_pandas_frame(df):
        headers, rows = _pandas_rows(df, include_index)
    elif _is_polars_frame(df):
        headers, rows = _polars_rows(df, include_index)
    else:
        raise TypeError(
            "Expected a pandas or polars DataFrame, got "
            f"{type(df).__module__}.{type(df).__qualname__}. "
            "Install pandas or polars to use write_dataframe()."
        )

    # Optional modules resolved once per conversion instead of once per cell.
    pd = _get_pandas()
    np = _get_numpy()

    format_for = _format_lookup(headers, formats)

    yield headers
    for row in rows:
        yield [_convert_cell(cell, fmt, na_rep, pd, np) for cell, fmt in zip(row, format_for)]


def _remove_partial_file(path) -> None:
    """Delete an output file left behind by a failed conversion."""
    if isinstance(path, (str, os.PathLike)) and os.path.exists(path):
        try:
            os.remove(path)
        except OSError:
            pass


def to_xlsx(
    df,
    path,
    sheet_name: str = "Sheet1",
    *,
    include_index: bool = False,
    na_rep=None,
    formats: dict[str, str] | None = None,
    header_style=None,
    styles: dict | None = None,
    merges: list[str] | tuple[str, ...] | None = None,
    conditional_formats: list | tuple | None = None,
    tables=None,
    column_widths: dict | None = None,
    hidden_columns=None,
    row_heights: dict | None = None,
    hidden_rows=None,
    row_groups=None,
    column_groups=None,
    freeze_panes: str | None = None,
    auto_filter: bool | None = None,
    charts=None,
    **writer_kwargs,
):
    """Write a DataFrame straight to an ``.xlsx`` file and return ``path``.

    If the conversion fails, a partially written file is removed.
    """
    from .xlsx_writer import XlsxWriter

    try:
        with XlsxWriter(path, **writer_kwargs) as writer:
            writer.add_sheet(sheet_name)
            writer.write_dataframe(
                df, include_index=include_index, na_rep=na_rep, formats=formats,
                header_style=header_style, styles=styles,
                merges=merges, conditional_formats=conditional_formats,
                tables=tables, column_widths=column_widths,
                hidden_columns=hidden_columns, row_heights=row_heights,
                hidden_rows=hidden_rows, row_groups=row_groups,
                column_groups=column_groups, freeze_panes=freeze_panes,
                auto_filter=auto_filter, charts=charts,
            )
    except Exception:
        _remove_partial_file(path)
        raise
    return path

def to_xlsb(
    df,
    path,
    sheet_name: str = "Sheet1",
    *,
    include_index: bool = False,
    na_rep=None,
    formats: dict[str, str] | None = None,
    header_style=None,
    styles: dict | None = None,
    merges: list[str] | tuple[str, ...] | None = None,
    column_widths: dict | None = None,
    hidden_columns=None,
    row_heights: dict | None = None,
    hidden_rows=None,
    row_groups=None,
    column_groups=None,
    freeze_panes: str | None = None,
    auto_filter: bool | None = None,
    charts=None,
    tables=None,
    conditional_formats=None,
    **writer_kwargs,
):
    """Write a DataFrame straight to an ``.xlsb`` file and return ``path``.

    If the conversion fails, a partially written file is removed.
    """
    from .writer import XlsbWriter

    try:
        with XlsbWriter(path, **writer_kwargs) as writer:
            writer.add_sheet(sheet_name)
            writer.write_dataframe(
                df, include_index=include_index, na_rep=na_rep, formats=formats,
                header_style=header_style, styles=styles,
                merges=merges, column_widths=column_widths,
                hidden_columns=hidden_columns, row_heights=row_heights,
                hidden_rows=hidden_rows, row_groups=row_groups,
                column_groups=column_groups, freeze_panes=freeze_panes,
                auto_filter=auto_filter, charts=charts, tables=tables,
                conditional_formats=conditional_formats,
            )
    except Exception:
        _remove_partial_file(path)
        raise
    return path
