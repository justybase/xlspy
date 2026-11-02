"""XLSB row records as the readers see them: row indexes and the end of the stream."""
from __future__ import annotations

import struct
import zipfile

import pytest

from xlspy import ExcelReader, XlsbWriter
from xlspy.biff_utils import iter_records


def _xlsb_with_row_index(tmp_path, new_index):
    """A three-row XLSB whose second BrtRowHdr is rewritten to ``new_index``."""
    source = tmp_path / "rows.xlsb"
    with XlsbWriter(source) as writer:
        writer.add_sheet("S")
        writer.write_sheet([["a"], ["b"], ["c"]])
    part = "xl/worksheets/sheet1.bin"
    with zipfile.ZipFile(source) as archive:
        items = {name: archive.read(name) for name in archive.namelist()}
    blob = bytearray(items[part])
    row_records = [r for r in iter_records(bytes(blob)) if r.record_id == 0x0000]
    struct.pack_into("<I", blob, row_records[1].data_start, new_index)
    items[part] = bytes(blob)
    target = tmp_path / "patched.xlsb"
    with zipfile.ZipFile(target, "w") as archive:
        for name, data in items.items():
            archive.writestr(name, data)
    return target


@pytest.fixture(params=["c", "python"])
def xlsb_reader_impl(request, monkeypatch):
    import xlspy.reader as reader_module

    if request.param == "c" and not reader_module.HAVE_NATIVE:
        pytest.skip("native extension not built")
    if request.param == "python":
        monkeypatch.setattr(reader_module, "HAVE_NATIVE", False)
    return request.param


@pytest.mark.parametrize("bad_index", [1048576, 530579456, 0x7FFFFFFF, 0x80000000, 0xFFFFFFFE])
def test_row_index_past_the_last_row_is_refused(tmp_path, xlsb_reader_impl, bad_index):
    path = _xlsb_with_row_index(tmp_path, bad_index)
    with ExcelReader(str(path)) as reader:
        with pytest.raises(ValueError, match="past the last row"):
            list(reader.get_rows("S"))


@pytest.mark.parametrize("index,expected_rows", [(1, 3), (5, 7)])
def test_row_index_inside_the_sheet_still_reads(tmp_path, xlsb_reader_impl, index, expected_rows):
    path = _xlsb_with_row_index(tmp_path, index)
    with ExcelReader(str(path)) as reader:
        rows = list(reader.get_rows("S"))
    assert len(rows) == expected_rows
    assert rows[0] == ["a"] and rows[-1] == ["c"]


def test_row_index_ffffffff_reads_like_a_short_row_record(tmp_path, xlsb_reader_impl):
    # 0xFFFFFFFF is -1 to the readers, the same as a row record with no index.
    path = _xlsb_with_row_index(tmp_path, 0xFFFFFFFF)
    with ExcelReader(str(path)) as reader:
        rows = list(reader.get_rows("S"))
    assert rows[0] == ["a"]


def _truncate_sheet_after_last_cell(source, target):
    """Copy an XLSB whose first sheet stream ends right after its last cell record."""

    def read_varint(data, i):
        value = shift = 0
        while True:
            byte = data[i]
            i += 1
            value |= (byte & 0x7F) << shift
            shift += 7
            if not byte & 0x80:
                return value, i

    with zipfile.ZipFile(source) as zin:
        names = zin.namelist()
        members = {name: zin.read(name) for name in names}
    sheet = members["xl/worksheets/sheet1.bin"]
    end_of_last_cell = None
    i = 0
    while i < len(sheet):
        record_id, i = read_varint(sheet, i)
        size, i = read_varint(sheet, i)
        if record_id in range(0x01, 0x0C):
            end_of_last_cell = i + size
        i += size
    members["xl/worksheets/sheet1.bin"] = sheet[:end_of_last_cell]
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in names:
            zout.writestr(name, members[name])


def test_pure_python_xlsb_reader_stops_at_eof_after_a_cell(tmp_path, monkeypatch):
    """Ignoring read_worksheet's result at EOF repeated the last cell forever."""
    from xlspy import biff_reader
    from xlspy import reader as reader_module
    from xlspy.reader import ExcelReader

    base = tmp_path / "base.xlsb"
    with XlsbWriter(base) as wb:
        wb.add_sheet("S")
        wb.write_sheet([[1, 2], [3, 4]])
    truncated = tmp_path / "eof.xlsb"
    _truncate_sheet_after_last_cell(base, truncated)

    # The native reader does not reach the Python loop; force the fallback.
    monkeypatch.setattr(reader_module, "HAVE_NATIVE", False)
    original = biff_reader.BiffReader.read_worksheet
    calls = []

    def bounded(self):
        calls.append(None)
        if len(calls) > 10_000:
            raise AssertionError("read_worksheet kept returning the last cell past the end of the stream")
        return original(self)

    monkeypatch.setattr(biff_reader.BiffReader, "read_worksheet", bounded)
    with ExcelReader(truncated) as reader:
        assert list(reader.get_rows("S")) == [[1, 2], [3, 4]]
