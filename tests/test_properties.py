"""Property-based smoke tests: writer/reader round trips and date maths.

These complement the example-based suites: ``@given`` generates small grids of
mixed values and checks invariants that must hold for every workbook:

* each writer round-trips its own output (XLSX and XLSB);
* both writers agree on the same grid;
* the XLSB C extension and the pure-Python fallback agree;
* Excel serial numbers and ``date``/``datetime`` values convert back and
  forth (compared against openpyxl for random serials).

The settings are deliberately small (``max_examples=50``) and deterministic
(``derandomize=True``) - this is a smoke layer, not fuzzing.
"""

import datetime as dt
import pytest
from hypothesis import given, settings
from hypothesis import strategies as st

from xlspy import ExcelReader, XlsbWriter, XlsxWriter
from xlspy import reader as xlspy_reader
from xlspy.excel_dates import datetime_to_serial, serial_to_datetime

SMOKE = settings(max_examples=50, deadline=None, derandomize=True)

FIRST_DATE = dt.datetime(1900, 1, 1)
LAST_DATE = dt.datetime(9999, 12, 31)

# Surrogates and control characters are excluded: they are not valid XML text
# and the updaters strip them, so they are out of scope for the smoke layer.
_text = st.text(st.characters(blacklist_categories=("Cs", "Cc")), min_size=1, max_size=12)

_cells = st.one_of(
    st.none(),
    st.booleans(),
    st.integers(min_value=-2**53, max_value=2**53),
    st.floats(allow_nan=False, allow_infinity=False, width=64),
    _text,
    st.dates(min_value=FIRST_DATE.date(), max_value=LAST_DATE.date()),
    st.datetimes(min_value=FIRST_DATE, max_value=LAST_DATE),
)
_non_empty_cells = _cells.filter(lambda value: value is not None)


@st.composite
def _rectangular_grids(draw):
    width = draw(st.integers(min_value=1, max_value=4))
    height = draw(st.integers(min_value=1, max_value=4))
    rows = []
    for _ in range(height):
        if width == 1:
            rows.append([draw(_non_empty_cells)])
        else:
            middle = draw(st.lists(_cells, min_size=width - 2, max_size=width - 2))
            rows.append([draw(_non_empty_cells), *middle, draw(_non_empty_cells)])
    return rows


def _write(path, rows):
    writer_cls = XlsbWriter if str(path).endswith(".xlsb") else XlsxWriter
    with writer_cls(str(path)) as writer:
        writer.add_sheet("Sheet1")
        writer.write_sheet(rows)
    return path


def _read(path, force_python_xlsb=False):
    if force_python_xlsb and xlspy_reader.HAVE_C_EXT:
        original = xlspy_reader.HAVE_C_EXT
        xlspy_reader.HAVE_C_EXT = False
        try:
            with ExcelReader(str(path)) as reader:
                return reader.read_all("Sheet1")
        finally:
            xlspy_reader.HAVE_C_EXT = original
    with ExcelReader(str(path)) as reader:
        return reader.read_all("Sheet1")


def _normalize(value):
    if isinstance(value, dt.datetime):
        return value
    if isinstance(value, dt.date):
        return dt.datetime(value.year, value.month, value.day)
    if isinstance(value, bool) or value is None:
        return value
    if isinstance(value, (int, float)):
        return float(value)
    return value


def _same_cell(actual, expected):
    actual = _normalize(actual)
    expected = _normalize(expected)
    if isinstance(actual, dt.datetime) and isinstance(expected, dt.datetime):
        # Serial arithmetic in doubles loses sub-millisecond precision for
        # dates far in the future; anything larger is a real error.
        return abs((actual - expected).total_seconds()) <= 1e-3
    return actual == expected


def _assert_same_grid(actual, expected, label):
    assert len(actual) == len(expected), f"{label}: {len(actual)} rows != {len(expected)}"
    for row_index, (actual_row, expected_row) in enumerate(zip(actual, expected)):
        assert len(actual_row) == len(expected_row), (
            f"{label} row {row_index}: width {len(actual_row)} != {len(expected_row)}"
        )
        for column_index, (actual_cell, expected_cell) in enumerate(zip(actual_row, expected_row)):
            assert _same_cell(actual_cell, expected_cell), (
                f"{label} row {row_index} column {column_index}: "
                f"{actual_cell!r} != {expected_cell!r}"
            )


@SMOKE
@given(rows=_rectangular_grids())
def test_xlsx_roundtrip(tmp_path_factory, rows):
    path = _write(tmp_path_factory.mktemp("xlsx") / "grid.xlsx", rows)
    _assert_same_grid(_read(path), rows, "xlsx roundtrip")


@SMOKE
@given(rows=_rectangular_grids())
def test_xlsb_roundtrip(tmp_path_factory, rows):
    path = _write(tmp_path_factory.mktemp("xlsb") / "grid.xlsb", rows)
    _assert_same_grid(_read(path), rows, "xlsb roundtrip")
    if xlspy_reader.HAVE_C_EXT:
        _assert_same_grid(
            _read(path, force_python_xlsb=True), rows, "xlsb roundtrip (python path)"
        )


@SMOKE
@given(rows=_rectangular_grids())
def test_xlsx_and_xlsb_writers_agree(tmp_path_factory, rows):
    directory = tmp_path_factory.mktemp("both")
    xlsx_rows = _read(_write(directory / "grid.xlsx", rows))
    xlsb_rows = _read(_write(directory / "grid.xlsb", rows))
    _assert_same_grid(xlsb_rows, xlsx_rows, "xlsx vs xlsb")


@SMOKE
@given(value=st.one_of(st.dates(min_value=FIRST_DATE.date(), max_value=LAST_DATE.date()), st.datetimes(min_value=FIRST_DATE, max_value=LAST_DATE)))
def test_date_serial_roundtrip(value):
    moment = _normalize(value)
    restored = serial_to_datetime(datetime_to_serial(value))
    assert abs((restored - moment).total_seconds()) <= 1e-3


@SMOKE
@given(serial=st.floats(min_value=1.0, max_value=100000.0, allow_nan=False, allow_infinity=False))
def test_serial_to_datetime_matches_openpyxl(serial):
    openpyxl_datetime = pytest.importorskip("openpyxl.utils.datetime")
    reference = openpyxl_datetime.from_excel(serial)
    actual = serial_to_datetime(serial)
    # openpyxl rounds fractions to milliseconds; xlspy retains datetime's
    # available microsecond precision.
    assert abs((actual - reference).total_seconds()) <= 1e-3
    if serial < 100:
        reference = openpyxl_datetime.from_excel(
            serial, epoch=openpyxl_datetime.CALENDAR_MAC_1904
        )
        actual = serial_to_datetime(serial, date_1904=True)
        assert abs((actual - reference).total_seconds()) <= 1e-3
