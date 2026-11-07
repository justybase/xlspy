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

from typing import TYPE_CHECKING, Any, TypeVar

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
            df, include_index=include_index, na_rep=na_rep, formats=formats,
            header_style=header_style, styles=styles,
            merges=merges, conditional_formats=conditional_formats,
            tables=tables, column_widths=column_widths,
            autofit_rows=autofit_rows, autofit_padding=autofit_padding,
            hidden_columns=hidden_columns, row_heights=row_heights,
            hidden_rows=hidden_rows, row_groups=row_groups,
            column_groups=column_groups, freeze_panes=freeze_panes,
            auto_filter=auto_filter, charts=charts,
            comments=comments, sheet_protection=sheet_protection,
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
            df, include_index=include_index, na_rep=na_rep, formats=formats,
            header_style=header_style, styles=styles,
            merges=merges, column_widths=column_widths,
            autofit_rows=autofit_rows, autofit_padding=autofit_padding,
            hidden_columns=hidden_columns, row_heights=row_heights,
            hidden_rows=hidden_rows, row_groups=row_groups,
            column_groups=column_groups, freeze_panes=freeze_panes,
            auto_filter=auto_filter, charts=charts, tables=tables,
            conditional_formats=conditional_formats,
            comments=comments, sheet_protection=sheet_protection,
        )
    return path
