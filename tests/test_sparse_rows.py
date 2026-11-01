"""Row records for dense data stay small; gaps and empty rows keep their place."""

import os
import re
import zipfile

import pytest

from xlspy import ExcelReader, Workbook, XlsbWriter, XlsxWriter, load_workbook
from xlspy.workbook import _reader

ROW_TAG = re.compile(rb"<row\b[^>]*>")


def _sheet(path, part="xl/worksheets/sheet1.xml"):
    with zipfile.ZipFile(path) as archive:
        return archive.read(part)


def test_dense_rows_carry_no_r_attribute(tmp_path):
    path = tmp_path / "dense.xlsx"
    with XlsxWriter(path) as writer:
        writer.write_sheet([[1], [2], [3]])
    assert ROW_TAG.findall(_sheet(path)) == [b"<row>", b"<row>", b"<row>"]
    with ExcelReader(str(path)) as reader:
        assert reader.read_all("Sheet1") == [[1], [2], [3]]


def test_gap_gets_an_r_attribute_and_keeps_its_position(tmp_path):
    path = tmp_path / "gap.xlsx"
    with XlsxWriter(path) as writer:
        writer.write_sheet([[1], [], [], [4]])
    assert ROW_TAG.findall(_sheet(path)) == [b"<row>", b'<row r="4">']
    with ExcelReader(str(path)) as reader:
        assert reader.read_all("Sheet1") == [[1], [None], [None], [4]]


def test_empty_rows_at_the_end_are_kept(tmp_path):
    """An empty frame is still one row, so the shape survives the round trip."""
    for writer_cls, ext in ((XlsxWriter, "xlsx"), (XlsbWriter, "xlsb")):
        path = tmp_path / f"trailing.{ext}"
        with writer_cls(path) as writer:
            writer.write_sheet([[1], [], []])
        with ExcelReader(str(path)) as reader:
            assert reader.read_all("Sheet1") == [[1], [None], [None]], ext


@pytest.mark.parametrize("writer_cls,ext", [(XlsxWriter, "xlsx"), (XlsbWriter, "xlsb")])
def test_far_cell_file_stays_small(tmp_path, writer_cls, ext):
    """Two cells a million rows apart used to write every row between them."""
    path = tmp_path / f"far.{ext}"

    def rows():
        yield [1]
        for _ in range(1048574):
            yield []
        yield [2]

    with writer_cls(path) as writer:
        writer.write_sheet(rows())
    assert os.path.getsize(path) < 20_000
    with ExcelReader(str(path)) as reader:
        present = [(i, row) for i, row in enumerate(reader.get_rows("Sheet1"), 1)
                   if any(value is not None for value in row)]
    assert [(i, row[0]) for i, row in present] == [(1, 1), (1048576, 2)]


def test_row_height_on_an_empty_row_is_saved(tmp_path):
    workbook = Workbook()
    sheet = workbook.active
    sheet["A1"] = "top"
    sheet["A2"] = "bottom"
    sheet.row_dimensions[5].height = 30
    path = tmp_path / "height.xlsx"
    workbook.save(path)
    workbook.close()
    loaded = load_workbook(path).active
    assert loaded.row_dimensions[5].height == 30
    assert loaded["A2"].value == "bottom"


@pytest.mark.parametrize("fast", [True, False])
def test_height_after_dense_rows_is_on_the_right_row(tmp_path, monkeypatch, fast):
    """Rows without r are numbered in order, also by the fast loader."""
    workbook = Workbook()
    sheet = workbook.active
    for value in ("a", "b", "c"):
        sheet.append([value])
    sheet.row_dimensions[2].height = 30
    path = tmp_path / "dense_height.xlsx"
    workbook.save(path)
    workbook.close()
    monkeypatch.setattr(_reader, "_FAST_SHEET_DATA", fast)
    loaded = load_workbook(path).active
    assert loaded.row_dimensions[2].height == 30
    assert loaded.row_dimensions[1].height is None
    assert loaded["B1"].value is None
    assert [loaded.cell(row=r, column=1).value for r in (1, 2, 3)] == ["a", "b", "c"]
