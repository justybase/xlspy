"""Automatic column widths: displayed text, per-character widths, writers."""

import datetime
import re
import struct
import zipfile

import pytest

from xlspy import Alignment, ExcelTable, Font, Style, XlsbWriter, XlsxWriter
from xlspy._widths import (
    DEFAULT_WIDTH, FILTER_BUTTON, cell_text, cell_width, estimate_column_widths, file_width,
    text_width,
)
from xlspy.biff_utils import iter_records


@pytest.mark.parametrize("value,number_format,expected", [
    (0.1 + 0.2, None, "0.3"),
    (1 / 3, None, "0.333333333"),
    (1234567890123, None, "1.23457E+12"),
    (2.5e15, None, "2.5E+15"),
    (1e-7, None, "0.0000001"),
    (-98765.4321, None, "-98765.4321"),
    (42, None, "42"),
    (1234567.891, "#,##0.00", "1,234,567.89"),
    (1234567.891, "#,##0", "1,234,568"),
    (0.4567, "0%", "46%"),
    (0.4567, "0.00%", "45.67%"),
    (12345, "000000000", "000012345"),
    (1234.5678, "0.00E+00", "1.23E+03"),
    (-1234.5, "#,##0.00;[Red]-#,##0.00", "-1,234.50"),
    (1234.5, "[$€-407] #,##0.00", "€ 1,234.50"),
    (True, None, "PRAWDA"),
    ("text", None, "text"),
    (datetime.date(2024, 9, 30), None, "30.09.2024"),
    (datetime.datetime(2024, 9, 30, 13, 45), None, "30.09.2024 13:45"),
    (datetime.date(2024, 9, 30), "yyyy-mm-dd", "2024-12-30"),
])
def test_cell_text(value, number_format, expected):
    assert cell_text(value, number_format) == expected


def test_long_date_names_are_generous():
    # Month and day names come from the reader's language: take long ones.
    text = cell_text(datetime.date(2024, 9, 30), "dddd, d mmmm yyyy")
    assert text_width(text) >= text_width("poniedziałek, 30 wrzesień 2024")
    assert "zł" in cell_text(1234.5, '#,##0.00 "zł"')


def test_character_widths():
    assert text_width("WWWW") > text_width("abcd") > text_width("iiii")
    assert text_width("中文") == pytest.approx(2 * text_width("00"), rel=0.01)
    assert text_width("0123456789") == pytest.approx(10.2, abs=0.1)


def test_cell_width_options():
    plain = cell_width("Customer name")
    assert cell_width("Customer name", bold=True) > plain
    assert cell_width("Customer name", size=22) == pytest.approx((plain - 1) * 2 + 1)
    two_lines = "short\na much longer second line"
    assert cell_width(two_lines, wrap=True) < cell_width(two_lines)
    assert cell_width(None) == 0.0


def test_estimate_column_widths():
    rows = [["ID", "Customer name", None], [1, "x", None], [2, "y", (1234567.891, "#,##0.00")]]
    widths = estimate_column_widths(rows, 3)
    assert widths[0] == DEFAULT_WIDTH  # never narrower than Excel's default
    assert widths[1] > DEFAULT_WIDTH
    assert widths[2] == pytest.approx(cell_width(1234567.891, "#,##0.00"))
    with_filter = estimate_column_widths(rows, 3, filter_cells=[(0, 0, 2)])
    assert with_filter[1] == pytest.approx(widths[1] + FILTER_BUTTON)
    styled = estimate_column_widths(
        rows, 3, lambda row_idx, length: {1: Style(font=Font(size=20))} if row_idx else {},
    )
    assert styled[1] == widths[1]  # the big font only applies to the short cells
    wrapped = estimate_column_widths(
        [["a fairly long heading\nx"]], 1, lambda r, n: {0: Style(alignment=Alignment(wrap_text=True))},
    )
    assert wrapped[0] < estimate_column_widths([["a fairly long heading\nx"]], 1)[0]


def _xlsx_widths(path):
    with zipfile.ZipFile(path) as archive:
        xml = archive.read("xl/worksheets/sheet1.xml").decode()
    return {int(col) - 1: float(width) for col, width in
            re.findall(r'<col min="(\d+)" max="\d+" width="([0-9.]+)"', xml)}


def _xlsb_widths(path):
    with zipfile.ZipFile(path) as archive:
        data = archive.read("xl/worksheets/sheet1.bin")
    widths = {}
    for record in iter_records(data):
        if record.record_id == 0x003C:
            first, _last, coldx = struct.unpack_from("<III", data, record.data_start)
            widths[first] = coldx / 256
    return widths


ROWS = [
    ["Order number", "Amount", "Note"],
    ["ORDER-2024-000123", (1234567.891, "#,##0.00"), "WWWWWWWWWWWW"],
    ["x", 2.5, "中文字符测试中文字符测试"],
]


@pytest.mark.parametrize("writer_cls,ext,read", [
    (XlsxWriter, "xlsx", _xlsx_widths), (XlsbWriter, "xlsb", _xlsb_widths),
])
def test_writers_use_the_estimate(tmp_path, writer_cls, ext, read):
    path = tmp_path / f"widths.{ext}"
    with writer_cls(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(ROWS)
    widths = read(path)
    expected = estimate_column_widths(ROWS, 3, filter_cells=[(0, 0, 2)])
    for col in range(3):
        # The file stores the width with Excel's cell padding.
        assert widths[col] == pytest.approx(file_width(expected[col]), abs=0.01)
    assert widths[2] > 20  # the CJK text needs about two digits per character
    assert file_width(DEFAULT_WIDTH) == pytest.approx(9.140625, abs=0.005)  # Excel's own default


def test_xlsb_counts_columns_after_empty_cells(tmp_path):
    path = tmp_path / "gaps.xlsb"
    with XlsbWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([[None, None, "a long text in column C"]], auto_filter=False)
    widths = _xlsb_widths(path)
    assert widths[2] > 15


def test_table_header_gets_room_for_its_button(tmp_path):
    rows = [["Customer name", "Value"], ["a", 1], ["b", 2]]
    path = tmp_path / "table.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(rows, tables=[ExcelTable(ref="A1:B3", display_name="Orders")])
    with_button = estimate_column_widths(rows, 2, filter_cells=[(0, 0, 1)])
    assert _xlsx_widths(path)[0] == pytest.approx(file_width(with_button[0]), abs=0.01)
    assert with_button[0] > estimate_column_widths(rows, 2)[0]
