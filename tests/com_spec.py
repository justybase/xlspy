"""Write and verify the cases of ``com_cases`` through Microsoft Excel (COM).

The spec is data only. ``write_case`` uses the xlspy writers, ``write_case_with_excel``
builds the same workbook with Excel, and ``verify_in_excel`` reads a workbook back
the way Excel shows it. Each check is a dict so a failure lists every difference,
not only the first.
"""

from __future__ import annotations

import datetime as dt
import math
from pathlib import Path

from xlspy import XlsbWriter, XlsxWriter

EXCEL_EPOCH = dt.datetime(1899, 12, 30)
XL_SHEET_VISIBLE = -1
XL_SHEET_HIDDEN = 0
XL_FILE_XLSX = 51
XL_FILE_XLSB = 50

# Application.International indexes (1-based in VBA, 0-based in the tuple pywin32
# returns): the separators Excel reports number formats with.
XL_DECIMAL_SEPARATOR = 3
XL_THOUSANDS_SEPARATOR = 4


def expand_rows(sheet: dict) -> list[list]:
    """The sheet's rows, expanding a generator sheet."""
    spec = sheet["rows"]
    if isinstance(spec, dict) and "generator" in spec:
        generator = spec["generator"]
        columns = int(generator["columns"])
        rows = [[f"C{index:02d}" for index in range(1, columns + 1)]]
        rows += [
            [float(row * 1000 + column) for column in range(1, columns + 1)]
            for row in range(1, int(generator["rows"]) + 1)
        ]
        return rows
    return [list(row) for row in spec]


def cell_value(cell):
    """The value a case cell holds (``None`` for an empty cell)."""
    if isinstance(cell, dict):
        if "date" in cell:
            return dt.datetime.strptime(cell["date"], "%Y-%m-%dT%H:%M:%S")
        return cell.get("value")
    return cell


def cell_number_format(cell) -> str | None:
    if isinstance(cell, dict):
        return cell.get("numberFormat")
    return None


def _serial(value) -> float:
    moment = value if isinstance(value, dt.datetime) else dt.datetime.combine(value, dt.time())
    return (moment - EXCEL_EPOCH).total_seconds() / 86400


def _column_letter(column: int) -> str:
    letters = ""
    while column:
        column, remainder = divmod(column - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def _cell_ref(row: int, column: int) -> str:
    return f"{_column_letter(column)}{row}"


def _same(expected, actual) -> bool:
    if isinstance(expected, bool):
        return isinstance(actual, bool) and actual is expected
    if isinstance(expected, (int, float)):
        return isinstance(actual, (int, float)) and not isinstance(actual, bool) and math.isclose(
            float(expected), float(actual), rel_tol=1e-9, abs_tol=1e-12
        )
    return expected == actual


def _number_format(cell, code: str) -> bool:
    """Whether ``cell`` has format ``code``; Excel reports it with local separators."""
    if code == "General":
        return cell.NumberFormat == cell.Worksheet.Range("Z1000").NumberFormat
    international = cell.Application.International
    decimal = international[XL_DECIMAL_SEPARATOR - 1]
    thousands = international[XL_THOUSANDS_SEPARATOR - 1]
    local = "".join({",": thousands, ".": decimal}.get(char, char) for char in code)
    return cell.NumberFormat in (code, local)


def write_case(path: Path, case: dict, fmt: str) -> None:
    """Write ``case`` with xlspy as ``fmt`` (``xlsx`` or ``xlsb``)."""
    writer_class = XlsxWriter if fmt == "xlsx" else XlsbWriter
    with writer_class(str(path)) as writer:
        for sheet in case["sheets"]:
            writer.add_sheet(sheet["name"], hidden=bool(sheet.get("hidden", False)))
            writer.write_sheet(
                _writer_rows(expand_rows(sheet)),
                auto_filter=bool(sheet.get("autoFilter", False)),
            )


def _writer_rows(rows: list[list]) -> list[list]:
    """Rows with the spec cells as the xlspy writers take them: a ``(value,
    number_format)`` cell keeps its format."""
    out = []
    for row in rows:
        writer_row = []
        for cell in row:
            value = cell_value(cell)
            format_code = cell_number_format(cell)
            writer_row.append((value, format_code) if format_code else value)
        out.append(writer_row)
    return out


def write_case_with_excel(excel, path: Path, case: dict, fmt: str) -> None:
    """Write ``case`` with Excel itself, so xlspy can be read against it."""
    book = excel.Workbooks.Add()
    try:
        # Late-bound COM ignores Add's After= keyword: a sheet is inserted before
        # the active one, so the sheets are named in reverse order.
        desired = case["sheets"]
        book.Worksheets(1).Name = desired[-1]["name"]
        for sheet in reversed(desired[:-1]):
            book.Worksheets.Add().Name = sheet["name"]
        visible = next((sheet for sheet in desired if not sheet.get("hidden")), desired[0])
        book.Worksheets(visible["name"]).Activate()

        for sheet in desired:
            worksheet = book.Worksheets(sheet["name"])
            rows = expand_rows(sheet)
            for row_number, row in enumerate(rows, start=1):
                for column_number, cell in enumerate(row, start=1):
                    value = cell_value(cell)
                    if value is None:
                        continue
                    target = worksheet.Cells(row_number, column_number)
                    format_code = cell_number_format(cell)
                    if isinstance(value, dt.date):
                        # pywin32's Time converts through the time zone; a serial
                        # with a date format is what Excel stores for a date cell.
                        target.Value = _serial(value)
                        target.NumberFormat = format_code or (
                            "yyyy-mm-dd hh:mm:ss" if isinstance(value, dt.datetime) else "yyyy-mm-dd"
                        )
                        continue
                    target.Value = value
                    if format_code:
                        target.NumberFormat = format_code
            if sheet.get("hidden"):
                worksheet.Visible = XL_SHEET_HIDDEN
            if sheet.get("autoFilter") and rows:
                width = max(len(row) for row in rows)
                worksheet.Range(
                    worksheet.Cells(1, 1), worksheet.Cells(len(rows), width)
                ).AutoFilter()
        book.SaveAs(str(path), XL_FILE_XLSB if fmt == "xlsb" else XL_FILE_XLSX)
    finally:
        book.Close(False)


def _used_bounds(rows: list[list]) -> tuple[int, int, int, int]:
    last_row = last_column = 1
    for row_number, row in enumerate(rows, start=1):
        for column_number, cell in enumerate(row, start=1):
            if cell_value(cell) is not None:
                last_row = max(last_row, row_number)
                last_column = max(last_column, column_number)
    return 1, 1, last_row, last_column


def _sampled_rows(count: int, sample_every: int | None) -> list[int]:
    if not sample_every:
        return list(range(1, count + 1))
    rows = {1, count, *range(1, count + 1, sample_every)}
    return sorted(row for row in rows if 1 <= row <= count)


def _check(checks, kind, sheet, name, expected, actual, message=""):
    checks.append(
        {
            "kind": kind,
            "sheet": sheet,
            "check": name,
            "expected": expected,
            "actual": actual,
            "pass": _same(expected, actual),
            "message": message,
        }
    )


def verify_in_excel(book, case: dict, fmt: str) -> list[dict]:
    """Read ``book`` through Excel and return one check per expectation."""
    checks: list[dict] = []
    names = [book.Worksheets(index).Name for index in range(1, book.Worksheets.Count + 1)]
    _check(checks, "structure", "-", "sheet names", [s["name"] for s in case["sheets"]], names)

    for sheet in case["sheets"]:
        name = sheet["name"]
        worksheet = book.Worksheets(name)
        expected_visible = XL_SHEET_HIDDEN if sheet.get("hidden") else XL_SHEET_VISIBLE
        _check(checks, "structure", name, "Visible", expected_visible, int(worksheet.Visible))

        rows = expand_rows(sheet)
        first_row, first_column, last_row, last_column = _used_bounds(rows)
        used = worksheet.UsedRange
        _check(
            checks,
            "structure",
            name,
            "UsedRange bounds",
            (first_row, first_column, last_row - first_row + 1, last_column - first_column + 1),
            (int(used.Row), int(used.Column), int(used.Rows.Count), int(used.Columns.Count)),
        )

        if sheet.get("autoFilter") and rows:
            width = max(len(row) for row in rows)
            expected_range = f"$A$1:${_column_letter(width)}${len(rows)}"
            actual_range = worksheet.AutoFilter.Range.Address if worksheet.AutoFilterMode else None
            _check(checks, "structure", name, "AutoFilter range", expected_range, actual_range)

        sampled = set(_sampled_rows(len(rows), sheet.get("sample_every")))
        for row_number, row in enumerate(rows, start=1):
            if row_number not in sampled:
                continue
            for column_number, cell in enumerate(row, start=1):
                value = cell_value(cell)
                if value is None:
                    continue
                reference = _cell_ref(row_number, column_number)
                target = worksheet.Cells(row_number, column_number)
                if isinstance(value, (dt.date, dt.datetime)):
                    _check(
                        checks,
                        "cells",
                        name,
                        f"cell {reference}",
                        _serial(value),
                        target.Value2,
                    )
                else:
                    _check(checks, "cells", name, f"cell {reference}", value, target.Value2)
                format_code = cell_number_format(cell)
                if format_code and isinstance(value, (dt.date, dt.datetime)):
                    # yyyy-mm-dd reads the same in every Excel language.
                    if format_code == "yyyy-mm-dd":
                        _check(
                            checks,
                            "cells",
                            name,
                            f"text {reference}",
                            value.strftime("%Y-%m-%d"),
                            target.Text,
                        )
                elif format_code:
                    _check(
                        checks,
                        "cells",
                        name,
                        f"numberFormat {reference}",
                        True,
                        _number_format(target, format_code),
                    )
    return checks


def format_failures(checks: list[dict]) -> str:
    return "\n".join(
        f"{check['sheet']} {check['check']}: expected {check['expected']!r}, "
        f"got {check['actual']!r} {check['message']}"
        for check in checks
        if not check["pass"]
    )
