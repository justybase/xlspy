"""Small helpers the tests share: write rows to a workbook, read them back, read a part.

Tests import them as ``from helpers import ...`` (the tests directory is on
sys.path, as for ``conftest``). A helper that builds one file's fixture stays
in that file; only the patterns that repeat across files live here.
"""
from __future__ import annotations

import os
import zipfile
from typing import Any, BinaryIO

from xlspy import ExcelReader, XlsbWriter, XlsxWriter


def writer_for(path: str | os.PathLike[str]) -> type[XlsxWriter] | type[XlsbWriter]:
    """The writer class for a file name: XlsbWriter for .xlsb, else XlsxWriter."""
    return XlsbWriter if os.fspath(path).lower().endswith(".xlsb") else XlsxWriter


def write_workbook(
    target: str | os.PathLike[str] | BinaryIO,
    rows: list[list[Any]],
    *,
    sheet: str = "Sheet1",
    writer: type[XlsxWriter] | type[XlsbWriter] | None = None,
    writer_options: dict[str, Any] | None = None,
    **sheet_options: Any,
) -> Any:
    """Write ``rows`` as one sheet and return ``target``. The writer class comes
    from the file name unless given (a stream needs it); ``writer_options`` go to
    the writer, the other keywords to ``write_sheet``."""
    if writer is None:
        writer = writer_for(target)  # type: ignore[arg-type]
    destination = target if hasattr(target, "write") else os.fspath(target)  # type: ignore[arg-type]
    with writer(destination, **(writer_options or {})) as book:
        book.add_sheet(sheet)
        book.write_sheet(rows, **sheet_options)
    return target


def read_rows(path: str | os.PathLike[str], sheet: str | None = None, **options: Any) -> list[list[Any]]:
    """All rows of ``sheet`` (the first sheet by default) through ExcelReader."""
    with ExcelReader(os.fspath(path)) as reader:
        name = sheet if sheet is not None else reader.get_sheet_names()[0]
        return list(reader.get_rows(name, **options))


def part_text(path: str | os.PathLike[str] | BinaryIO, name: str) -> str:
    """The UTF-8 text of package part ``name``."""
    with zipfile.ZipFile(path) as archive:
        return archive.read(name).decode("utf-8")
