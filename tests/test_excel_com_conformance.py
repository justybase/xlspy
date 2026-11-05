"""Excel COM conformance suite driven by the spec in ``com_cases``.

Writer tests: xlspy writes every case in XLSX and XLSB; Microsoft Excel opens each
file and reports sheet names, visibility, AutoFilter, used range, every cell value
and number format.

Reader tests: Excel writes the same cases; xlspy reads them back and must return
the spec values.

Updater test: xlspy writes a template, ``replace_sheet_data`` replaces its body,
and Excel reads the result.

Requires Windows, Microsoft Excel and ``pywin32``; opt in with
``XLSPY_RUN_EXCEL_COM=1``.
"""

from __future__ import annotations

import datetime as dt
import math
import shutil

import pytest

from xlspy import ExcelReader, XlsbUpdater, XlsxUpdater

from com_cases import CASES, UPDATE_RESULT, UPDATE_TEMPLATE
from com_spec import (
    cell_value,
    expand_rows,
    format_failures,
    verify_in_excel,
    write_case,
    write_case_with_excel,
)

pytestmark = pytest.mark.com  # skipped unless XLSPY_RUN_EXCEL_COM=1 (conftest.py)

FORMATS = ("xlsx", "xlsb")
CASE_IDS = [case["id"] for case in CASES]
EXCEL_EPOCH = dt.datetime(1899, 12, 30)


@pytest.fixture(scope="module")
def written(tmp_path_factory, excel):
    """Write every case with xlspy and read it back through Excel, once."""
    base = tmp_path_factory.mktemp("com_conformance")
    results = {}
    for case in CASES:
        for fmt in FORMATS:
            path = base / f"{case['id']}.{fmt}"
            write_case(path, case, fmt)
            book = excel.Workbooks.Open(str(path), 0, True)
            try:
                results[(case["id"], fmt)] = verify_in_excel(book, case, fmt)
            finally:
                book.Close(False)
    return results


@pytest.fixture(scope="module")
def excel_authored(tmp_path_factory, excel):
    """Write every case with Excel itself and return the paths."""
    base = tmp_path_factory.mktemp("com_conformance_excel")
    paths = {}
    for case in CASES:
        if not case.get("excel_authored", True):
            continue
        for fmt in FORMATS:
            path = base / f"{case['id']}.{fmt}"
            write_case_with_excel(excel, path, case, fmt)
            paths[(case["id"], fmt)] = path
    return paths


def _checks(written, case_id, fmt, kind):
    checks = [c for c in written[(case_id, fmt)] if c["kind"] == kind]
    assert checks, f"no {kind} checks were produced"
    failures = [c for c in checks if not c["pass"]]
    assert not failures, format_failures(failures)


@pytest.mark.parametrize("fmt", FORMATS)
@pytest.mark.parametrize("case_id", CASE_IDS)
def test_excel_opens_xlspy_structure(written, case_id, fmt):
    """Sheet names, order, visibility, AutoFilter and used range match the spec."""
    _checks(written, case_id, fmt, "structure")


@pytest.mark.parametrize("fmt", FORMATS)
@pytest.mark.parametrize("case_id", CASE_IDS)
def test_excel_reads_xlspy_cell_values(written, case_id, fmt):
    """Every sampled cell value and number format reads back as specified."""
    checks = [c for c in written[(case_id, fmt)] if c["kind"] == "cells"]
    if not checks:
        pytest.skip("the case holds no cells")
    failures = [c for c in checks if not c["pass"]]
    assert not failures, format_failures(failures)


def _as_datetime(value):
    if isinstance(value, dt.datetime):
        return value
    if isinstance(value, dt.date):
        return dt.datetime.combine(value, dt.time())
    return value


def _same_cell(expected, actual) -> bool:
    expected = _as_datetime(expected)
    if expected is None:
        return actual in (None, "")
    if isinstance(expected, dt.datetime):
        if not isinstance(_as_datetime(actual), dt.datetime):
            return False
        expected_serial = (expected - EXCEL_EPOCH).total_seconds()
        actual_serial = (_as_datetime(actual) - EXCEL_EPOCH).total_seconds()
        return math.isclose(expected_serial, actual_serial, abs_tol=1.0)
    if isinstance(expected, bool):
        return isinstance(actual, bool) and actual is expected
    if isinstance(expected, (int, float)):
        return isinstance(actual, (int, float)) and not isinstance(actual, bool) and math.isclose(
            float(expected), float(actual), rel_tol=1e-9, abs_tol=1e-12
        )
    return expected == actual


@pytest.mark.parametrize("fmt", FORMATS)
@pytest.mark.parametrize("case_id", CASE_IDS)
def test_xlspy_reads_excel_authored_file(excel_authored, case_id, fmt):
    case = next(c for c in CASES if c["id"] == case_id)
    if not case.get("excel_authored", True):
        pytest.skip("Excel authoring this case is covered by the writer side")
    with ExcelReader(excel_authored[(case_id, fmt)]) as reader:
        assert reader.get_sheet_names() == [sheet["name"] for sheet in case["sheets"]]
        for sheet in case["sheets"]:
            expected_rows = expand_rows(sheet)
            if not expected_rows:
                continue
            actual_rows = [list(row) for row in reader.read_all(sheet["name"])]
            for row_number, expected_row in enumerate(expected_rows, start=1):
                if row_number > len(actual_rows):
                    break
                for column_number, expected in enumerate(expected_row, start=1):
                    if expected is None:
                        continue
                    actual = (
                        actual_rows[row_number - 1][column_number - 1]
                        if column_number <= len(actual_rows[row_number - 1])
                        else None
                    )
                    assert _same_cell(cell_value(expected), actual), (
                        f"{sheet['name']} row {row_number} col {column_number}: "
                        f"expected {cell_value(expected)!r}, got {actual!r}"
                    )


@pytest.mark.parametrize("fmt", FORMATS)
def test_xlspy_updater_then_excel_verifies(tmp_path, excel, fmt):
    """A replace_sheet_data edit survives the save and Excel reads the new values."""
    template = tmp_path / f"template.{fmt}"
    write_case(template, UPDATE_TEMPLATE, fmt)
    target = tmp_path / f"updated.{fmt}"
    shutil.copyfile(template, target)

    rows = expand_rows(UPDATE_RESULT["sheets"][0])
    headers, body = rows[0], rows[1:]
    updater_class = XlsxUpdater if fmt == "xlsx" else XlsbUpdater
    updater = updater_class(str(target))
    try:
        updater.replace_sheet_data("Data", body, headers=headers)
        updater.save(str(target))
    finally:
        updater.close()

    book = excel.Workbooks.Open(str(target), 0, True)
    try:
        checks = __import__("com_spec").verify_in_excel(book, UPDATE_RESULT, fmt)
    finally:
        book.Close(False)
    failures = [check for check in checks if not check["pass"]]
    assert not failures, format_failures(failures)
