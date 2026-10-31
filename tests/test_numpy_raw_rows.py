"""numpy scalars in raw rows are written as the Python values they hold (C6).

The DataFrame path has always converted numpy scalars with ``.item()``. A row
given to ``write_sheet`` directly now does the same for integers, bool_ and
float16/float32 (a numpy float64 already is a float). Each such row must
produce the same sheet as the same row of Python values, in XLSX and XLSB,
including the auto column widths, which the sheet XML carries.
"""

import io
import zipfile

import pytest

from xlspy import ExcelReader, XlsbWriter, XlsxWriter
from xlspy._scalars import numpy_value

np = pytest.importorskip("numpy")

NUMPY_ROW = [
    np.int64(7), np.int8(-3), np.uint16(9), np.float32(0.1), np.float16(2.5),
    np.bool_(True), np.bool_(False), np.float64(1.5),
]
PYTHON_ROW = [
    7, -3, 9, float(np.float32(0.1)), 2.5, True, False, 1.5,
]


def _xlsx_sheet(rows):
    buffer = io.BytesIO()
    with XlsxWriter(buffer) as writer:
        writer.add_sheet("T")
        writer.write_sheet(rows)
    with zipfile.ZipFile(io.BytesIO(buffer.getvalue())) as archive:
        return archive.read("xl/worksheets/sheet1.xml")


def _xlsb_sheet(rows):
    buffer = io.BytesIO()
    with XlsbWriter(buffer) as writer:
        writer.add_sheet("T")
        writer.write_sheet(rows)
    with zipfile.ZipFile(io.BytesIO(buffer.getvalue())) as archive:
        return archive.read("xl/worksheets/sheet1.bin")


def test_xlsx_numpy_row_is_written_as_its_python_values():
    assert _xlsx_sheet([NUMPY_ROW]) == _xlsx_sheet([PYTHON_ROW])


def test_xlsb_numpy_row_is_written_as_its_python_values():
    assert _xlsb_sheet([NUMPY_ROW]) == _xlsb_sheet([PYTHON_ROW])


def test_numpy_values_read_back_as_numbers_and_booleans(tmp_path):
    path = tmp_path / "numbers.xlsx"
    with XlsxWriter(str(path)) as writer:
        writer.add_sheet("T")
        writer.write_sheet([NUMPY_ROW])
    with ExcelReader(str(path)) as reader:
        (row,) = list(reader.get_rows("T"))
    assert row[:3] == [7, -3, 9]
    assert row[3] == float(np.float32(0.1))
    assert row[4] == 2.5
    assert row[5] is True and row[6] is False
    assert row[7] == 1.5


def test_numpy_float32_with_a_format_is_the_python_float(tmp_path):
    rows = [[(np.float32(0.5), "0.00"), np.int32(4)]]
    expected = [[(0.5, "0.00"), 4]]
    assert _xlsx_sheet(rows) == _xlsx_sheet(expected)


def test_numpy_non_finite_float32_is_refused_like_a_float_nan():
    # A NaN float is refused by the writer; a float32 NaN now is too, not written as text.
    with pytest.raises(ValueError, match="finite"):
        _xlsx_sheet([[np.float32("nan")]])
    with pytest.raises(ValueError, match="finite"):
        _xlsx_sheet([[float("nan")]])


@pytest.mark.parametrize(
    "value",
    [np.float64(2.0), np.str_("text"), np.datetime64("2024-01-15"), np.array([1, 2]), np.complex64(1j)],
    ids=["float64", "str", "datetime64", "array", "complex64"],
)
def test_numpy_values_outside_the_kinds_are_left_alone(value):
    assert numpy_value(value) is value


def test_numpy_value_converts_only_the_scalar_kinds():
    assert type(numpy_value(np.int64(5))) is int
    assert type(numpy_value(np.uint8(5))) is int
    assert type(numpy_value(np.float32(5.0))) is float
    assert type(numpy_value(np.float16(5.0))) is float
    assert type(numpy_value(np.bool_(True))) is bool
    assert numpy_value(5) == 5 and numpy_value("x") == "x" and numpy_value(None) is None
