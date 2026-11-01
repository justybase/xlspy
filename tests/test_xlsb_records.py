"""BIFF12 records of a written worksheet part: the column-info container ids and the dimension.

Excel tolerates both defects these tests pin down; other readers do not. LibreOffice drops a sheet whose
column-info container carries the wrong ids, and readers that size rows from BrtWsDim (pyxlsb, for one)
fail on an all-zero dimension.
"""
import struct
import zipfile

from xlspy import Formula, Style, XlsbWriter
from xlspy.biff_utils import iter_records

BRT_WS_DIM = 0x0094
BRT_COL_INFO = 0x003C
BRT_END_WS_VIEWS = 0x0086
BRT_BEGIN_SHEET_DATA = 0x0091
BRT_BEGIN_COL_INFOS = 0x0186
BRT_END_COL_INFOS = 0x0187


def _sheet_records(path):
    with zipfile.ZipFile(path) as archive:
        data = archive.read("xl/worksheets/sheet1.bin")
    return data, list(iter_records(data))


def _write(path, rows, **options):
    with XlsbWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(rows, **options)


def _dimension(path):
    data, records = _sheet_records(path)
    (record,) = [record for record in records if record.record_id == BRT_WS_DIM]
    return struct.unpack_from("<IIII", data, record.data_start)


def test_column_info_container_uses_the_biff12_ids(tmp_path):
    path = tmp_path / "cols.xlsb"
    _write(path, [["a", "b", "c"], [1, 2, 3]], hidden_columns=["B"], column_widths={"A": 20})
    _, records = _sheet_records(path)
    ids = [record.record_id for record in records]
    assert ids.count(BRT_BEGIN_COL_INFOS) == 1
    assert ids.count(BRT_END_COL_INFOS) == 1
    begin, end = ids.index(BRT_BEGIN_COL_INFOS), ids.index(BRT_END_COL_INFOS)
    assert begin < end < ids.index(BRT_BEGIN_SHEET_DATA)
    assert all(record_id == BRT_COL_INFO for record_id in ids[begin + 1:end])
    # 0x86 is BrtEndWsViews: once, closing the view block, before the columns.
    assert ids.count(BRT_END_WS_VIEWS) == 1
    assert ids.index(BRT_END_WS_VIEWS) < begin


def test_dimension_offset_in_the_sheet_template():
    assert XlsbWriter._sheet1Bytes[29:32] == b"\x94\x01\x10"


def test_dimension_record_covers_the_used_range(tmp_path):
    path = tmp_path / "dim.xlsb"
    rows = [
        [None, None, None],
        [None, "Particulars", "Q1", None],
        [None, "Sales", 10.5],
        [None, None],
        [None, "Total", Formula("=C3", value=10.5), None, None],
    ]
    _write(path, rows)
    assert _dimension(path) == (1, 4, 1, 2)


def test_dimension_counts_styled_blank_cells(tmp_path):
    path = tmp_path / "styled.xlsb"
    _write(path, [["a", None, None]], styles={"A1:C1": Style(bold=True)})
    assert _dimension(path) == (0, 0, 0, 2)


def test_empty_sheet_keeps_the_zero_dimension(tmp_path):
    path = tmp_path / "empty.xlsb"
    _write(path, [])
    assert _dimension(path) == (0, 0, 0, 0)


def test_dimension_survives_a_sheet_larger_than_the_write_buffer(tmp_path):
    path = tmp_path / "big.xlsb"
    _write(path, ([f"row {i}", i, i * 1.5] for i in range(20000)))
    assert _dimension(path) == (0, 19999, 0, 2)
