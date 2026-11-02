"""Writer lifecycle: one write per sheet, sheets never written, recovery after a refused write."""
from __future__ import annotations

import zipfile

import pytest

from xlspy import ExcelTable, XlsbUpdater, XlsbWriter, XlsxUpdater, XlsxWriter, load_workbook
from xlspy.exceptions import UnsupportedFeatureError


def parts(path):
    with zipfile.ZipFile(path) as zf:
        return {name: zf.read(name) for name in zf.namelist()}


def test_xlsb_writes_a_part_for_a_sheet_never_written(tmp_path):
    path = tmp_path / "empty.xlsb"
    with XlsbWriter(path) as wb:
        wb.add_sheet("First")
        wb.write_sheet([["a"]])
        wb.add_sheet("Never written")
    package = parts(path)
    assert "xl/worksheets/sheet2.bin" in package
    assert load_workbook(path).sheetnames == ["First", "Never written"]


@pytest.mark.parametrize("writer", [XlsxWriter, XlsbWriter])
def test_second_write_sheet_on_one_sheet_is_rejected(tmp_path, writer):
    path = tmp_path / "twice"
    with pytest.raises(ValueError, match="already has its data"):
        with writer(path) as wb:
            wb.add_sheet("S")
            wb.write_sheet([["a"]])
            wb.write_sheet([["b"]])


FORMATS = pytest.mark.parametrize("suffix", ["xlsx", "xlsb"])


TABLE_ROWS = [["Name", "Qty"], ["a", 1], ["b", 2]]


def _table_base(tmp_path, suffix, rows=TABLE_ROWS, ref="A1:B3"):
    base = tmp_path / f"base.{suffix}"
    writer = XlsxWriter if suffix == "xlsx" else XlsbWriter
    with writer(base) as wb:
        wb.add_sheet("S")
        wb.write_sheet(rows, tables=[ExcelTable("Tbl", ref)])
    return base


def _updater(path, suffix):
    return XlsxUpdater(path) if suffix == "xlsx" else XlsbUpdater(path)


@FORMATS
def test_mismatch_after_refused_write_leaves_later_valid_write_working(tmp_path, suffix):
    base = _table_base(tmp_path, suffix)
    updater = _updater(base, suffix)
    with pytest.raises(UnsupportedFeatureError):
        updater.replace_sheet_data("S", [["Other", "Qty"], ["c", 3]])
    updater.replace_sheet_data("S", [["Name", "Qty"], ["c", 3]])
    out = tmp_path / f"out.{suffix}"
    updater.save(out)
    assert load_workbook(out)["S"].cell(row=2, column=1).value == "c"
