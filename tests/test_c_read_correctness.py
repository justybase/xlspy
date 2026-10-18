"""Focused regressions for C-vs-Python XLSB reader parity.

The generated binary mutations exercise cases that ordinary writer output
does not produce: left-ragged cells, record-order-dependent first-row widths,
and short/overrunning cell records. The Python path is selected in-process by
disabling only the reader's C dispatch, leaving both implementations on the
same input file and in the same test environment.
"""

import struct
import zipfile
import datetime as dt

import pytest

from xlspy import ExcelReader, XlsbWriter
from xlspy import reader as reader_module
from xlspy._accel import HAVE_C_EXT


pytestmark = pytest.mark.skipif(not HAVE_C_EXT, reason="C extension not built")
SHEET_PART = "xl/worksheets/sheet1.bin"


def _vint(value):
    out = bytearray()
    while True:
        byte = value & 0x7F
        value >>= 7
        if value:
            out.append(byte | 0x80)
        else:
            out.append(byte)
            return bytes(out)


def _read_records(data):
    records = []
    pos = 0
    while pos < len(data):
        rec_id, pos = _read_vint(data, pos)
        rec_len, pos = _read_vint(data, pos)
        records.append((rec_id, data[pos:pos + rec_len]))
        pos += rec_len
    return records


def _read_vint(data, pos):
    value = shift = 0
    while pos < len(data):
        byte = data[pos]
        pos += 1
        value |= (byte & 0x7F) << shift
        if not byte & 0x80:
            return value, pos
        shift += 7
    raise AssertionError("truncated BIFF varint in generated workbook")


def _write_records(records):
    out = bytearray()
    for rec_id, body in records:
        out += _vint(rec_id) + _vint(len(body)) + body
    return bytes(out)


def _rewrite_sheet(source, destination, transform):
    with zipfile.ZipFile(source) as zin, zipfile.ZipFile(
        destination, "w", zipfile.ZIP_DEFLATED
    ) as zout:
        for name in zin.namelist():
            payload = zin.read(name)
            if name == SHEET_PART:
                payload = _write_records(transform(_read_records(payload)))
            zout.writestr(name, payload)


def _read(path, *, force_python=False, monkeypatch):
    if force_python:
        with monkeypatch.context() as patch:
            patch.setattr(reader_module, "HAVE_C_EXT", False)
            with ExcelReader(str(path)) as reader:
                return reader.read_all("S")
    with ExcelReader(str(path)) as reader:
        return reader.read_all("S")


def _first_row(path, *, force_python=False, monkeypatch):
    if force_python:
        with monkeypatch.context() as patch:
            patch.setattr(reader_module, "HAVE_C_EXT", False)
            with ExcelReader(str(path)) as reader:
                return next(reader.get_rows("S"))
    with ExcelReader(str(path)) as reader:
        return next(reader.get_rows("S"))


def _same_typed_grid(left, right):
    assert len(left) == len(right)
    for lrow, rrow in zip(left, right):
        assert len(lrow) == len(rrow)
        for lcell, rcell in zip(lrow, rrow):
            assert type(lcell) is type(rcell)
            assert lcell == rcell


def _numeric_workbook(path):
    with XlsbWriter(str(path)) as writer:
        writer.add_sheet("S")
        writer.write_sheet([[10.0, 20.0, 30.0], [40.0, 50.0, 60.0]])


def _set_1904_date_system(path):
    with zipfile.ZipFile(path) as zin:
        members = [(info, zin.read(info.filename)) for info in zin.infolist()]
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for info, payload in members:
            if info.filename == "xl/workbook.bin":
                records = _read_records(payload)
                changed = []
                found = False
                for rec_id, body in records:
                    if rec_id == 0x99 and body:
                        body = bytes([body[0] | 1]) + body[1:]
                        found = True
                    changed.append((rec_id, body))
                assert found, "generated XLSB is missing BrtWbProp"
                payload = _write_records(changed)
            zout.writestr(info, payload)


def test_c_reader_drops_cells_left_of_first_data_column(tmp_path, monkeypatch):
    source = tmp_path / "source.xlsb"
    _numeric_workbook(source)

    # Locate the first row and shift its three numeric cell columns by two.
    # The row parser has a record-order state, so use a simpler explicit
    # transform after finding the first row record.
    def transform(records):
        out = []
        row_number = -1
        for rec_id, body in records:
            if rec_id == 0x00:
                row_number += 1
            elif row_number == 0 and 0x01 <= rec_id <= 0x0B and len(body) >= 8:
                col = struct.unpack_from("<I", body)[0]
                body = struct.pack("<I", col + 2) + body[4:]
            out.append((rec_id, body))
        return out

    mutated = tmp_path / "left-ragged.xlsb"
    _rewrite_sheet(source, mutated, transform)
    c_grid = _read(mutated, monkeypatch=monkeypatch)
    py_grid = _read(mutated, force_python=True, monkeypatch=monkeypatch)
    _same_typed_grid(c_grid, py_grid)
    assert c_grid == [[10, 20, 30], [60, None, None]]


def test_c_reader_matches_python_for_first_row_record_order(tmp_path, monkeypatch):
    source = tmp_path / "source.xlsb"
    _numeric_workbook(source)

    def reverse_first_row_columns(records):
        out = []
        row_number = -1
        cell_cols = iter((2, 0, 1))
        for rec_id, body in records:
            if rec_id == 0x00:
                row_number += 1
            elif row_number == 0 and 0x01 <= rec_id <= 0x0B and len(body) >= 8:
                body = struct.pack("<I", next(cell_cols)) + body[4:]
            out.append((rec_id, body))
        return out

    mutated = tmp_path / "unordered-first-row.xlsb"
    _rewrite_sheet(source, mutated, reverse_first_row_columns)
    c_grid = _read(mutated, monkeypatch=monkeypatch)
    py_grid = _read(mutated, force_python=True, monkeypatch=monkeypatch)
    _same_typed_grid(c_grid, py_grid)
    # The reference computes first/last columns in record order (2 then 1),
    # giving a non-positive width and therefore an empty visible row.
    assert c_grid == [[], []]


def test_c_reader_ignores_short_and_overrunning_records(tmp_path, monkeypatch):
    source = tmp_path / "source.xlsb"
    _numeric_workbook(source)

    def inject_invalid_records(records):
        out = []
        inserted = False
        for rec_id, body in records:
            out.append((rec_id, body))
            if rec_id == 0x00 and not inserted:
                # BiffReader ignores all records with fewer than the common
                # 8-byte column/style header, and ignores string records when
                # the claimed UTF-16 length exceeds the record payload.
                out.append((0x01, b"\0" * 4))
                out.append((0x06, struct.pack("<III", 0, 0, 100) + b"x\0"))
                inserted = True
        return out

    mutated = tmp_path / "short-records.xlsb"
    _rewrite_sheet(source, mutated, inject_invalid_records)
    c_grid = _read(mutated, monkeypatch=monkeypatch)
    py_grid = _read(mutated, force_python=True, monkeypatch=monkeypatch)
    _same_typed_grid(c_grid, py_grid)
    assert c_grid == [[10, 20, 30], [40, 50, 60]]


def test_c_reader_propagates_out_of_range_1904_date_error(tmp_path, monkeypatch):
    path = tmp_path / "far-date.xlsb"
    with XlsbWriter(str(path)) as writer:
        writer.add_sheet("S")
        writer.write_sheet([[(dt.datetime(9999, 12, 31, 23, 59, 59, 999999),
                              "yyyy-mm-dd hh:mm:ss")]])
    _set_1904_date_system(path)

    with pytest.raises(OverflowError) as c_error:
        _read(path, monkeypatch=monkeypatch)
    with pytest.raises(OverflowError) as py_error:
        _read(path, force_python=True, monkeypatch=monkeypatch)
    assert type(c_error.value) is type(py_error.value)


def test_c_reader_delegates_unusually_wide_layout_to_python(tmp_path, monkeypatch):
    source = tmp_path / "source.xlsb"
    _numeric_workbook(source)

    def widen_first_row(records):
        out = []
        row_number = -1
        cell_index = 0
        columns = (0, 1, 2_000_000)
        for rec_id, body in records:
            if rec_id == 0x00:
                row_number += 1
                cell_index = 0
            elif row_number == 0 and 0x01 <= rec_id <= 0x0B and len(body) >= 8:
                body = struct.pack("<I", columns[cell_index]) + body[4:]
                cell_index += 1
            out.append((rec_id, body))
        return out

    mutated = tmp_path / "wide-first-row.xlsb"
    _rewrite_sheet(source, mutated, widen_first_row)

    # C's dense row array intentionally declines this nonstandard layout;
    # the public reader must transparently use the Python reference result.
    c_row = _first_row(mutated, monkeypatch=monkeypatch)
    assert len(c_row) == 2_000_001
    assert (c_row[0], c_row[1], c_row[-1]) == (10, 20, 30)
    del c_row

    py_row = _first_row(mutated, force_python=True, monkeypatch=monkeypatch)
    assert len(py_row) == 2_000_001
    assert (py_row[0], py_row[1], py_row[-1]) == (10, 20, 30)
