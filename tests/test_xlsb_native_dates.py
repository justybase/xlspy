"""XLSB reads through the native extension: date cells and wide rows.

``read_xlsb_worksheet`` converts date serials itself and keeps each row in a
buffer that grows as cells arrive. The date results are compared with the
reference ``serial_to_datetime``; a serial the reference cannot convert (NaN,
infinite, outside years 1..9999) stays a number, as Excel shows such a cell as
####. The wide rows have known values, written in both column orders.
"""

import math
import random
import struct

import pytest

from xlspy._accel import HAVE_NATIVE, read_xlsb_worksheet
from xlspy.biff_utils import build_record
from xlspy.excel_dates import serial_to_datetime

needs_native = pytest.mark.skipif(not HAVE_NATIVE, reason="native extension not built")

STYLES = [{"xfId": 0, "numFmtId": 0}, {"xfId": 0, "numFmtId": 14}]  # xf 1 is a date style
DATE_FMTS = {14}
DATE_XF = 1
EDGES = [
    0.0, 0.5, 1.0, 1.5, 59.0, 59.999999999, 60.0, 61.0, -0.5, -1.0, -1.5,
    25569.0, 2958465.0, 2958466.0, 1e9, -1e9, 1e10, float("nan"), float("inf"), float("-inf"),
]


def _row_header(row):
    return build_record(0x00, struct.pack("<I", row) + bytes(12))


def _real_cell(col, xf, value):
    return build_record(0x05, struct.pack("<IId", col, xf, value))


def _single_column_sheet(values, xf=DATE_XF):
    parts = []
    for row, value in enumerate(values):
        parts.append(_row_header(row))
        parts.append(_real_cell(0, xf, value))
    return b"".join(parts)


def _read(data, date_1904=False):
    rows = read_xlsb_worksheet(data, [], STYLES, DATE_FMTS, date_1904)
    assert rows is not None, "the native reader fell back to Python for a plain layout"
    return list(rows)


def _reference(value, date_1904):
    try:
        return serial_to_datetime(value, date_1904)
    except Exception as error:  # the Python reference decides the exception
        return error


@needs_native
@pytest.mark.parametrize("date_1904", [False, True])
def test_date_cells_match_reference(date_1904):
    rng = random.Random(20261030)
    uniform = [rng.uniform(-70000.0, 2960000.0) for _ in range(2000)]
    # Serials whose fraction is a few microseconds: the rounding decides the result.
    micro = [rng.randint(-5000, 2958000) + rng.randint(0, 86399999999) / 86400000000.0
             for _ in range(2000)]
    values = EDGES + uniform + micro

    valid = [v for v in values if not isinstance(_reference(v, date_1904), Exception)]
    got = [row[0] for row in _read(_single_column_sheet(valid), date_1904)]
    assert got == [serial_to_datetime(v, date_1904) for v in valid]
    assert all(type(item) is type(serial_to_datetime(v, date_1904)) for item, v in zip(got, valid))

    invalid = [v for v in values if isinstance(_reference(v, date_1904), Exception)]
    assert invalid, "the edge list must hold serials outside the datetime range"
    for value, row in zip(invalid, _read(_single_column_sheet(invalid), date_1904)):
        assert_same_number(row[0], value)


def assert_same_number(got, serial):
    if math.isnan(serial):
        assert type(got) is float and math.isnan(got)
        return
    want = int(serial) if serial.is_integer() else serial
    assert got == want and type(got) is type(want)


@needs_native
def test_plain_numbers_in_a_non_date_style_stay_numbers():
    rows = _read(_single_column_sheet([2.5, 3.0, -7.0], xf=0))
    assert rows == [[2.5], [3], [-7]]
    assert [type(row[0]) for row in rows] == [float, int, int]


@needs_native
def test_wide_rows_keep_every_column_in_either_order():
    width = 40
    parts, expected = [], []
    for row in range(6):
        parts.append(_row_header(row))
        columns = range(width) if row % 2 == 0 else reversed(range(width))
        for col in columns:
            parts.append(_real_cell(col, 0, row * 100.0 + col + 0.5))
        expected.append([row * 100.0 + col + 0.5 for col in range(width)])
    assert [list(row) for row in _read(b"".join(parts))] == expected


@pytest.mark.parametrize("backend", ["native", "python"])
def test_out_of_range_date_serials_read_as_numbers_everywhere(tmp_path, monkeypatch, backend):
    # A date-formatted cell whose serial is no datetime is not a damaged file:
    # every read path returns the number, as the XLSX readers do.
    from xlspy import ExcelReader, XlsbWriter, _accel, load_workbook

    if backend == "native" and not HAVE_NATIVE:
        pytest.skip("native extension not built")
    if backend == "python":
        import xlspy.reader as reader_module

        monkeypatch.setattr(reader_module, "_accel_read_xlsb_worksheet", lambda *args, **kwargs: None)
        monkeypatch.setattr(_accel, "xlsb_load_cells", None)
    serials = [1e7, 3e6 + 0.5, -1e10]  # the writer refuses NaN; see above
    path = tmp_path / "dates.xlsb"
    with XlsbWriter(str(path)) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([[(serial, "yyyy-mm-dd")] for serial in serials] + [[(45000.0, "yyyy-mm-dd")]])

    with ExcelReader(str(path)) as reader:
        bulk = [row[0] for row in reader.get_rows("Data")]
        streamed = [row[0] for row in reader.get_rows("Data", streaming=True)]
    sheet = load_workbook(path)["Data"]
    model = [sheet.cell(row=i + 1, column=1).value for i in range(len(serials) + 1)]
    for values in (bulk, streamed, model):
        assert values[-1] == serial_to_datetime(45000.0, False)
        for got, serial in zip(values, serials):
            assert_same_number(got, serial)
