"""Worksheet records of XlsbWriter output that readers other than Excel depend on (PR #1).

* The column-info container must use ``BrtBeginColInfos`` / ``BrtEndColInfos`` (``0x186`` / ``0x187``).
  xlspy before 0.11 wrote ``0x86`` / ``0x87`` (``BrtEndWsViews`` / ``BrtBeginBookViews``); Excel
  tolerated it, LibreOffice dropped every cell of the sheet (see test_libreoffice_xlsb.py).
* ``BrtWsDim`` must hold the used range. xlspy before 0.11 wrote ``(0, 0, 0, 0)`` whatever the data;
  Excel ignores the record on load, but readers that size rows from it (pyxlsb) failed past
  column A. ``DIMENSION_CASES`` holds the dimension Excel writes when it resaves the same data;
  test_excel_com_xlsb_dimension.py checks that against Excel itself.
"""
import io
import random
import struct
import zipfile

import pytest

from xlspy import (
    Formula, Group, Style, XlsbUpdater, XlsbWriter, Workbook, load_workbook,
)
from xlspy.biff_utils import iter_records

BRT_ROW_HDR = 0x0000
BRT_COL_INFO = 0x003C
BRT_END_WS_VIEWS = 0x0086
BRT_BEGIN_SHEET_DATA = 0x0091
BRT_WS_DIM = 0x0094
BRT_BEGIN_COL_INFOS = 0x0186
BRT_END_COL_INFOS = 0x0187
# Cell records: BrtCellBlank .. BrtFmlaError, then BrtCellRString.
CELL_RECORDS = frozenset(range(0x01, 0x0C)) | {0x3E}
# BrtRowHdr flags byte (payload offset 11): outline level, collapsed, hidden, custom height.
ROW_HDR_FLAGS = 0x3F


# name -> (rows, write_sheet options, BrtWsDim Excel writes: rwFirst, rwLast, colFirst, colLast).
DIMENSION_CASES = {
    "sparse": (
        [
            [None, None, None],
            [None, "Particulars", "Q1", None],
            [None, "Sales", 10.5],
            [None, None],
            [None, "Total", Formula("=C3", value=10.5), None, None],
        ],
        {},
        (1, 4, 1, 2),
    ),
    "formula": ([[1, 2, Formula("A1*2")], [4]], {}, (0, 1, 0, 2)),
    "styled_blanks": ([["a", None, None]], {"styles": {"A1:C1": Style(bold=True)}}, (0, 0, 0, 2)),
    "header_style": ([["a", "b", None]], {"header_style": Style(italic=True)}, (0, 0, 0, 2)),
    "first_column_moves_left": ([[None, None, "c"], ["a"]], {}, (0, 1, 0, 2)),
    "trailing_empty_rows": ([[None, 1], [], [None], []], {}, (0, 0, 1, 1)),
    "height_on_empty_row": ([[None, "x"], [], [], []], {"row_heights": {3: 30}}, (0, 3, 1, 1)),
    "hidden_empty_row": ([[None, "x"], [], []], {"hidden_rows": [2]}, (0, 2, 1, 1)),
    "outline_past_data": ([[None, "x"]], {"row_groups": [Group(3, 4)]}, (0, 4, 1, 1)),
    "array_past_data": ([[None, Formula("=1", value=1, array_ref="B1:C3")]], {}, (0, 2, 1, 2)),
    "only_empty_rows": ([[], [None, None]], {}, (0, 0, 0, 0)),
    "empty": ([], {}, (0, 0, 0, 0)),
}


def write_case(path, case):
    rows, options, _ = DIMENSION_CASES[case]
    _write(path, rows, **options)


def _write(path, rows, **options):
    with XlsbWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(rows, **options)


def _sheet_part(path, part="xl/worksheets/sheet1.bin"):
    with zipfile.ZipFile(path) as archive:
        return archive.read(part)


def sheet_dimension(path, part="xl/worksheets/sheet1.bin"):
    data = _sheet_part(path, part)
    (record,) = [record for record in iter_records(data) if record.record_id == BRT_WS_DIM]
    return struct.unpack_from("<IIII", data, record.data_start)


def used_range_of_records(data):
    """The used range by Excel's rule, from the records themselves: rows with a cell
    or a row header that is not the default, columns of every cell record."""
    rows, cols, row = [], [], None
    for record in iter_records(data):
        if record.record_id == BRT_ROW_HDR:
            row, ixfe = struct.unpack_from("<II", data, record.data_start)
            flags = data[record.data_start + 11]
            height = struct.unpack_from("<H", data, record.data_start + 8)[0]
            if ixfe or flags & ROW_HDR_FLAGS or height != 300:
                rows.append(row)
        elif record.record_id in CELL_RECORDS:
            rows.append(row)
            cols.append(struct.unpack_from("<I", data, record.data_start)[0])
    if not rows:
        return (0, 0, 0, 0)
    return (min(rows), max(rows), min(cols, default=0), max(cols, default=0))


def assert_dimension_matches_records(path, part="xl/worksheets/sheet1.bin"):
    assert sheet_dimension(path, part) == used_range_of_records(_sheet_part(path, part))


# Column-info container ------------------------------------------------------------------------

def test_column_info_container_uses_the_biff12_ids(tmp_path):
    path = tmp_path / "cols.xlsb"
    _write(path, [["a", "b", "c"], [1, 2, 3]], hidden_columns=["B"], column_widths={"A": 20})
    ids = [record.record_id for record in iter_records(_sheet_part(path))]
    begin, end = ids.index(BRT_BEGIN_COL_INFOS), ids.index(BRT_END_COL_INFOS)
    assert ids.count(BRT_BEGIN_COL_INFOS) == ids.count(BRT_END_COL_INFOS) == 1
    assert begin < end < ids.index(BRT_BEGIN_SHEET_DATA)
    assert all(record_id == BRT_COL_INFO for record_id in ids[begin + 1:end])
    # 0x86 is BrtEndWsViews: once, closing the view block, before the columns.
    assert ids.count(BRT_END_WS_VIEWS) == 1 and ids.index(BRT_END_WS_VIEWS) < begin


def test_column_info_ids_after_a_model_save(tmp_path):
    source, saved = tmp_path / "source.xlsb", tmp_path / "saved.xlsb"
    _write(source, [["a", "b"], [1, 2]])
    workbook = load_workbook(source)
    try:
        workbook["Data"].column_dimensions["B"].hidden = True
        workbook["Data"].column_dimensions["A"].width = 30
        workbook.save(saved)
    finally:
        workbook.close()
    ids = [record.record_id for record in iter_records(_sheet_part(saved))]
    assert ids.count(BRT_BEGIN_COL_INFOS) == ids.count(BRT_END_COL_INFOS) == 1
    assert ids.count(BRT_END_WS_VIEWS) == 1


# BrtWsDim ------------------------------------------------------------------------------------

@pytest.mark.parametrize("case", sorted(DIMENSION_CASES))
def test_dimension_holds_the_used_range(tmp_path, case):
    path = tmp_path / f"{case}.xlsb"
    write_case(path, case)
    assert sheet_dimension(path) == DIMENSION_CASES[case][2]
    assert_dimension_matches_records(path)


def test_dimension_of_a_sheet_larger_than_the_row_buffer(tmp_path):
    # Past 64 KB the body is spooled and the head is written after the rows.
    path = tmp_path / "big.xlsb"
    _write(path, ([None, f"row {i}", i, i * 1.5] for i in range(30000)))
    assert sheet_dimension(path) == (0, 29999, 1, 3)
    with zipfile.ZipFile(path) as archive:
        assert archive.testzip() is None
    workbook = load_workbook(path, read_only=True)
    try:
        sheet = workbook["Data"]
        assert sheet.dimensions == "B1:D30000"
        rows = list(sheet.values)
        # As in openpyxl, read-only rows start at column A.
        assert len(rows) == 30000
        assert rows[0] == (None, "row 0", 0, 0) and rows[-1] == (None, "row 29999", 29999, 44998.5)
    finally:
        workbook.close()


@pytest.mark.parametrize("compression_level", [0, 1, 6, 9])
def test_spooled_sheet_part_is_valid_at_every_compression_level(tmp_path, compression_level):
    path = tmp_path / "level.xlsb"
    with XlsbWriter(path, compression_level=compression_level) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([f"value {i}", i] for i in range(20000))
    with zipfile.ZipFile(path) as archive:
        assert archive.testzip() is None
    assert sheet_dimension(path) == (0, 19999, 0, 1)


def test_dimension_of_a_one_pass_stream_into_a_file_object():
    buffer = io.BytesIO()
    with XlsbWriter(buffer) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(iter([[None, None, i] for i in range(15000)]))
    buffer.seek(0)
    assert sheet_dimension(buffer) == (0, 14999, 2, 2)


def test_each_sheet_gets_its_own_dimension(tmp_path):
    path = tmp_path / "sheets.xlsb"
    with XlsbWriter(path) as writer:
        writer.add_sheet("One")
        writer.write_sheet([["a", "b"]])
        writer.add_sheet("Two")
        writer.write_sheet([[], [None, None, None, "d"]])
        writer.add_sheet("Never written")
    assert sheet_dimension(path, "xl/worksheets/sheet1.bin") == (0, 0, 0, 1)
    assert sheet_dimension(path, "xl/worksheets/sheet2.bin") == (1, 1, 3, 3)
    assert sheet_dimension(path, "xl/worksheets/sheet3.bin") == (0, 0, 0, 0)


def test_failed_sheet_leaves_no_partial_part(tmp_path):
    def rows():
        for i in range(20000):
            yield [i]
        raise RuntimeError("source failed")

    path = tmp_path / "failed.xlsb"
    with pytest.raises(RuntimeError, match="source failed"):
        with XlsbWriter(path) as writer:
            writer.add_sheet("Data")
            writer.write_sheet(rows())
    assert not path.exists()


def test_dimension_matches_records_for_random_sheets(tmp_path):
    """Random sparse rows with styles, heights, hidden rows and groups: the
    dimension is always the bounding box Excel's rule gives for the records."""
    generator = random.Random(20261102)
    values = [None, None, None, 1, 2.5, "s", True, Formula("1+1", value=2), (None, "0.00"), ""]
    for case in range(60):
        rows = [
            [generator.choice(values) for _ in range(generator.randint(0, 7))]
            for _ in range(generator.randint(0, 12))
        ]
        options = {}
        if generator.random() < 0.3:
            options["styles"] = {"B2:D4": Style(bold=True)}
        if generator.random() < 0.3:
            options["row_heights"] = {generator.randint(0, 14): 25}
        if generator.random() < 0.3:
            options["hidden_rows"] = [generator.randint(0, 14)]
        if generator.random() < 0.2:
            first = generator.randint(1, 15)
            options["row_groups"] = [Group(first, first + generator.randint(0, 3))]
        path = tmp_path / f"random{case}.xlsb"
        _write(path, rows, **options)
        expected = used_range_of_records(_sheet_part(path))
        assert sheet_dimension(path) == expected, (rows, options)


# Other write paths -------------------------------------------------------------------------

@pytest.fixture
def written(tmp_path):
    path = tmp_path / "written.xlsb"
    _write(path, [[None, None], [None, "b", 2], [None, None, 3]])
    assert sheet_dimension(path) == (1, 2, 1, 2)
    return path


@pytest.mark.parametrize("updates", [{"E7": 9}, {"A1": 1}, {"B2": 5}])
def test_updater_keeps_the_dimension_exact(written, tmp_path, updates):
    path = tmp_path / "updated.xlsb"
    updater = XlsbUpdater(written)
    try:
        updater.update_cells("Data", updates)
        updater.save(path)
    finally:
        updater.close()
    assert_dimension_matches_records(path)


def test_model_save_keeps_the_dimension_exact(written, tmp_path):
    path = tmp_path / "model.xlsb"
    workbook = load_workbook(written)
    try:
        workbook["Data"]["A1"] = "x"
        workbook["Data"]["F9"] = 1
        workbook.save(path)
    finally:
        workbook.close()
    assert sheet_dimension(path) == (0, 8, 0, 5)
    assert_dimension_matches_records(path)


def test_new_sheet_saved_by_the_model_does_not_start_at_a1(written, tmp_path):
    path = tmp_path / "new_sheet.xlsb"
    workbook = load_workbook(written)
    try:
        workbook.create_sheet("New")["D5"] = 1
        workbook.save(path)
    finally:
        workbook.close()
    assert sheet_dimension(path, "xl/worksheets/sheet2.bin") == (4, 4, 3, 3)
    assert_dimension_matches_records(path, "xl/worksheets/sheet2.bin")


def test_new_workbook_saved_as_xlsb(tmp_path):
    path = tmp_path / "new.xlsb"
    workbook = Workbook()
    workbook.active["C4"] = 1
    workbook.active["B2"] = "x"
    workbook.save(path)
    workbook.close()
    assert sheet_dimension(path) == (1, 3, 1, 2)
    assert_dimension_matches_records(path)


# Reading --------------------------------------------------------------------------------------

def test_read_only_load_sizes_the_sheet_from_the_dimension(tmp_path):
    path = tmp_path / "sparse.xlsb"
    write_case(path, "sparse")
    workbook = load_workbook(path, read_only=True)
    try:
        sheet = workbook["Data"]
        assert sheet.dimensions == "B2:C5"
        assert (sheet.min_row, sheet.max_row, sheet.min_column, sheet.max_column) == (2, 5, 2, 3)
    finally:
        workbook.close()


def test_read_only_load_of_a_legacy_all_zero_dimension(tmp_path):
    """Files from xlspy before 0.11 carry (0, 0, 0, 0): they still read, unsized."""
    path, legacy = tmp_path / "current.xlsb", tmp_path / "legacy.xlsb"
    _write(path, [[1, 2, Formula("A1*2")], [4]])
    with zipfile.ZipFile(path) as source, zipfile.ZipFile(legacy, "w", zipfile.ZIP_DEFLATED) as target:
        for info in source.infolist():
            data = source.read(info.filename)
            if info.filename == "xl/worksheets/sheet1.bin":
                (record,) = [r for r in iter_records(data) if r.record_id == BRT_WS_DIM]
                data = data[:record.data_start] + bytes(16) + data[record.data_end:]
            target.writestr(info, data)
    workbook = load_workbook(legacy, read_only=True)
    try:
        assert workbook["Data"].max_row is None
        assert list(workbook["Data"].values) == [(1, 2, "=A1*2"), (4,)]
    finally:
        workbook.close()


def test_pyxlsb_reads_cells_past_column_a(tmp_path):
    pyxlsb = pytest.importorskip("pyxlsb")
    path = tmp_path / "pyxlsb.xlsb"
    _write(path, [["a", 1, 2], [None, "x", 3.5]], hidden_columns=["B"])
    with pyxlsb.open_workbook(str(path)) as workbook, workbook.get_sheet("Data") as sheet:
        rows = [[cell.v for cell in row] for row in sheet.rows()]
    assert rows == [["a", 1, 2], [None, "x", 3.5]]
