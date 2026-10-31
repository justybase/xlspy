"""Sheet view switches: row and column headers, zero values and right-to-left."""

import re
import zipfile
from typing import Any

import pytest

from xlspy import XlsbWriter, XlsxWriter
from xlspy.biff_utils import iter_records
from xlspy.workbook import load_workbook

ROWS = [["a", 0], ["b", 1]]
OFF = {"show_row_col_headers": False, "show_zero_values": False, "right_to_left": True}


def _write(writer_class, path, **options):
    with writer_class(path) as writer:
        writer.add_sheet("View")
        writer.write_sheet(ROWS, **options)
    return path


def _sheet_view(path):
    with zipfile.ZipFile(path) as archive:
        xml = archive.read("xl/worksheets/sheet1.xml").decode()
    return re.search(r"<sheetView\b[^>]*>", xml).group(0)


def _view_flags(path):
    with zipfile.ZipFile(path) as archive:
        data = archive.read("xl/worksheets/sheet1.bin")
    record = next(record for record in iter_records(data) if record.record_id == 0x89)
    return data[record.data_start]


def test_xlsx_defaults_add_no_attributes(tmp_path):
    view = _sheet_view(_write(XlsxWriter, tmp_path / "v.xlsx"))
    assert "showRowColHeaders" not in view and "showZeros" not in view and "rightToLeft" not in view


def test_xlsx_writes_the_attributes_in_schema_order(tmp_path):
    view = _sheet_view(_write(XlsxWriter, tmp_path / "v.xlsx", show_grid_lines=False, zoom=90, **OFF))
    assert view.startswith(
        '<sheetView showGridLines="0" showRowColHeaders="0" showZeros="0" rightToLeft="1" tabSelected="1"'
    )
    assert 'zoomScale="90"' in view


def test_xlsb_sets_the_view_flag_bits(tmp_path):
    assert _view_flags(_write(XlsbWriter, tmp_path / "d.xlsb")) == 0xDC
    flags = _view_flags(_write(XlsbWriter, tmp_path / "v.xlsb", **OFF))
    assert flags == (0xDC & ~0x08 & ~0x10) | 0x20


@pytest.mark.parametrize("name", sorted(OFF))
def test_options_must_be_bools(tmp_path, name):
    options: dict[str, Any] = {name: 0}
    with XlsxWriter(tmp_path / "v.xlsx") as writer:
        with pytest.raises(TypeError, match=name):
            writer.write_sheet(ROWS, **options)


@pytest.mark.parametrize(("writer_class", "suffix"), [(XlsxWriter, "xlsx"), (XlsbWriter, "xlsb")])
def test_loaded_model_reads_and_keeps_the_view(tmp_path, writer_class, suffix):
    source = _write(writer_class, tmp_path / f"v.{suffix}", **OFF)
    workbook = load_workbook(source)
    try:
        view = workbook.worksheets[0].sheet_view
        assert (view.showRowColHeaders, view.showZeros, view.rightToLeft) == (False, False, True)
    finally:
        workbook.close()


@pytest.mark.com
@pytest.mark.parametrize(("writer_class", "suffix"), [(XlsxWriter, "xlsx"), (XlsbWriter, "xlsb")])
def test_excel_shows_the_view(tmp_path, writer_class, suffix):
    win32 = pytest.importorskip("win32com.client")
    path = _write(writer_class, tmp_path / f"com.{suffix}", **OFF)
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    try:
        book = excel.Workbooks.Open(str(path))
        window = book.Windows(1)
        assert not window.DisplayHeadings
        assert not window.DisplayZeros
        assert window.DisplayRightToLeft
        book.Close(SaveChanges=False)
    finally:
        excel.Quit()
