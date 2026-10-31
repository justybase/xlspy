"""Phase 1 compatibility tests: utils, number formats, formula, model surface."""

from __future__ import annotations

import datetime as dt

import pytest

openpyxl = pytest.importorskip("openpyxl")

import openpyxl.utils.cell as ocell
import openpyxl.utils.datetime as odt
import openpyxl.utils.escape as oescape
import openpyxl.utils.units as ounits
import openpyxl.styles.numbers as onumbers
from openpyxl.formula.tokenizer import Tokenizer as OpenpyxlTokenizer
from openpyxl.formula.translate import Translator as OpenpyxlTranslator

import xlspy
from xlspy.formula import Tokenizer, Translator
from xlspy.styles import numbers
from xlspy.utils import cell, datetime as xdt, escape, units


@pytest.mark.parametrize("index", [1, 2, 26, 27, 52, 702, 703, 16384, 18278])
def test_get_column_letter(index):
    assert cell.get_column_letter(index) == ocell.get_column_letter(index)


@pytest.mark.parametrize("value", ["A", "Z", "AA", "AZ", "ZZ", "AAA", "XFD", "ZZZ"])
def test_column_index_from_string(value):
    assert cell.column_index_from_string(value) == ocell.column_index_from_string(value)


@pytest.mark.parametrize("value", ["A1", "b12", "$C$3", "XFD1048576"])
def test_coordinate_from_string(value):
    assert cell.coordinate_from_string(value) == ocell.coordinate_from_string(value)


@pytest.mark.parametrize("value", ["A1", "b12", "XFD1048576"])
def test_coordinate_to_tuple(value):
    assert cell.coordinate_to_tuple(value) == ocell.coordinate_to_tuple(value)


@pytest.mark.parametrize(
    "value",
    ["A1", "A1:B2", "$A$1:$B$2", "A:A", "1:2", "B2:D9"],
)
def test_range_boundaries(value):
    assert cell.range_boundaries(value) == ocell.range_boundaries(value)


@pytest.mark.parametrize("value", ["Sheet1!A1", "Sheet1!A1:B2", "'My Sheet'!A1:B2"])
def test_range_to_tuple(value):
    assert cell.range_to_tuple(value) == ocell.range_to_tuple(value)


@pytest.mark.parametrize("value", ["A1:B2", "B2:D4"])
def test_rows_and_cols_from_range(value):
    assert list(cell.rows_from_range(value)) == list(ocell.rows_from_range(value))
    assert list(cell.cols_from_range(value)) == list(ocell.cols_from_range(value))


def test_column_interval_and_absolute_coordinate():
    assert cell.get_column_interval("A", "D") == ocell.get_column_interval("A", "D")
    assert cell.get_column_interval(1, 4) == ocell.get_column_interval(1, 4)
    for value in ("A1", "B2:D4", "$A$1"):
        assert cell.absolute_coordinate(value) == ocell.absolute_coordinate(value)
    assert cell.quote_sheetname("My Sheet") == ocell.quote_sheetname("My Sheet")


@pytest.mark.parametrize("serial", [0, 0.5, 1, 59, 60, 61, 40000, 45000.25])
def test_from_excel(serial):
    assert xdt.from_excel(serial) == odt.from_excel(serial)


@pytest.mark.parametrize(
    "value",
    [
        dt.date(1900, 1, 1),
        dt.date(1900, 2, 28),
        dt.date(1900, 3, 1),
        dt.date(2024, 1, 2),
        dt.datetime(2024, 1, 2, 3, 4, 5),
        dt.time(12, 30, 15),
        dt.timedelta(days=1, hours=2, minutes=3),
    ],
)
def test_to_excel(value):
    assert xdt.to_excel(value) == odt.to_excel(value)


def test_time_helpers():
    delta = dt.timedelta(hours=12, minutes=30, seconds=15)
    moment = dt.time(12, 30, 15)
    assert xdt.days_to_time(delta) == odt.days_to_time(delta)
    assert xdt.time_to_days(moment) == odt.time_to_days(moment)
    assert xdt.timedelta_to_days(delta) == odt.timedelta_to_days(delta)


@pytest.mark.parametrize(
    "value",
    ["2024-01-02", "2024-01-02T03:04:05", "2024-01-02T03:04:05Z", "03:04:05"],
)
def test_iso8601(value):
    assert xdt.from_ISO8601(value) == odt.from_ISO8601(value)
    parsed = xdt.from_ISO8601(value)
    assert xdt.to_ISO8601(parsed) == odt.to_ISO8601(parsed)


def test_iso8601_duration():
    assert xdt.from_ISO8601("PT2H") == odt.from_ISO8601("PT2H")


@pytest.mark.parametrize("value", ["abc", "a\x01b", "\x1f", "plain"])
def test_escape(value):
    assert escape.escape(value) == oescape.escape(value)
    assert escape.unescape(escape.escape(value)) == oescape.unescape(oescape.escape(value))


def test_units():
    for name in (
        "EMU_to_cm", "cm_to_EMU", "EMU_to_inch", "inch_to_EMU",
        "EMU_to_pixels", "pixels_to_EMU", "pixels_to_points",
        "points_to_pixels", "dxa_to_cm", "cm_to_dxa", "dxa_to_inch",
        "inch_to_dxa",
    ):
        assert getattr(units, name)(1000) == getattr(ounits, name)(1000)
    assert units.short_color("FF00FF00") == ounits.short_color("FF00FF00")


def test_exception_hierarchy():
    for name in (
        "CellCoordinatesException", "IllegalCharacterError", "InvalidFileException",
        "NamedRangeException", "ReadOnlyWorkbookException", "SheetTitleException",
        "WorkbookAlreadySaved",
    ):
        assert issubclass(getattr(xlspy.utils.exceptions, name), Exception)


@pytest.mark.parametrize(
    "fmt",
    [
        "General", "0", "0.00", "#,##0.00", "0%", "yyyy-mm-dd", "mm-dd-yy",
        "h:mm:ss", "[h]:mm:ss", "[hh]:mm:ss", '"yyyy"0.00', "0.00E+00", "@",
    ],
)
def test_number_format_helpers(fmt):
    assert numbers.is_date_format(fmt) == onumbers.is_date_format(fmt)
    assert numbers.is_datetime(fmt) == onumbers.is_datetime(fmt)
    assert numbers.is_timedelta_format(fmt) == onumbers.is_timedelta_format(fmt)
    assert numbers.is_builtin(fmt) == onumbers.is_builtin(fmt)
    assert numbers.builtin_format_id(fmt) == onumbers.builtin_format_id(fmt)
    if onumbers.builtin_format_id(fmt) is not None:
        index = onumbers.builtin_format_id(fmt)
        assert numbers.builtin_format_code(index) == onumbers.builtin_format_code(index)


FORMULAS = [
    "=SUM(A1:A3)",
    "=A1+B2*3",
    '=IF(A1>0,"yes","no")',
    "='My Sheet'!A1+$B$2",
    "=1.5E+2",
    "=RATE+1",
    "=PRICE-2*A1E-3",
    "=#DIV/0!",
    "={1,2;3,4}",
    "=A1%",
    "=TRUE",
    '=A1&"x"',
    "=SUM(A:A)",
]


@pytest.mark.parametrize("formula", FORMULAS)
def test_formula_tokenizer(formula):
    expected = [(t.value, t.type, t.subtype) for t in OpenpyxlTokenizer(formula).items]
    actual = [(t.value, t.type, t.subtype) for t in Tokenizer(formula).items]
    assert actual == expected


@pytest.mark.parametrize(
    "formula,origin,dest",
    [
        ("=SUM(A1:A3)", "B2", "C3"),
        ("=$A$1+B1", "B2", "C3"),
        ("=SUM(A:A)", "B2", "C3"),
        ("=SUM(1:1)", "B2", "C3"),
        ("=A1+Sheet2!B2", "A1", "B2"),
    ],
)
def test_formula_translator(formula, origin, dest):
    assert (
        Translator(formula, origin).translate_formula(dest)
        == OpenpyxlTranslator(formula, origin).translate_formula(dest)
    )


def test_write_only_cell_and_append_dict():
    wb = xlspy.Workbook(write_only=True)
    ws = wb.create_sheet()
    cell = xlspy.WriteOnlyCell(ws, value=1)
    cell.font = xlspy.Font(bold=True)
    ws.append([cell, 2])
    ws.append({"A": 3, "C": 4})
    assert ws._rows[0] == [1, 2]
    assert ws._rows[1] == [3, None, 4]


def test_append_dict_on_worksheet():
    wb = xlspy.Workbook()
    ws = wb.active
    ws.append({"A": 1, "C": 3})
    assert ws["A1"].value == 1
    assert ws["C1"].value == 3


def test_copy_worksheet_copies_hyperlinks():
    wb = xlspy.Workbook()
    ws = wb.active
    ws["A1"] = "link"
    ws["A1"].hyperlink = xlspy.Hyperlink("https://example.com")
    copy = wb.copy_worksheet(ws)
    assert copy["A1"].hyperlink is not None
    assert copy["A1"].hyperlink.target == "https://example.com"


def test_cell_range_operations():
    first = xlspy.CellRange("A1:B2")
    second = xlspy.CellRange("B2:C3")
    assert first.bounds == (1, 1, 2, 2)
    assert first.coord == "A1:B2"
    assert first.union(second).bounds == (1, 1, 3, 3)
    assert first.intersection(second).bounds == (2, 2, 2, 2)
    assert first.shift(1, 1).bounds == (2, 2, 3, 3)
    assert first.issubset(xlspy.CellRange("A1:C3"))
    assert first.isdisjoint(xlspy.CellRange("D1:D2"))


def test_workbook_metadata_defaults():
    wb = xlspy.Workbook()
    assert wb.read_only is False
    assert wb.template is False
    assert wb.epoch == dt.datetime(1899, 12, 30)
    assert wb.vba_archive is None
    assert wb.security.lockStructure is None
    assert wb.calculation.fullCalcOnLoad is True
    assert wb.views[0].activeTab == 0
    assert wb.custom_doc_props == []


def test_defined_names_add_remove():
    wb = xlspy.Workbook()
    name = xlspy.DefinedName("MyName", "'Sheet'!$A$1")
    wb.defined_names.add(name)
    assert wb.defined_names["MyName"] is name
    assert "MyName" in wb.defined_names
    wb.defined_names.remove("MyName")
    assert "MyName" not in wb.defined_names


def test_read_only_cell_has_number_format(tmp_path):
    path = tmp_path / "read_only.xlsx"
    wb = xlspy.Workbook()
    wb.active["A1"] = 1
    wb.save(path)
    loaded = xlspy.load_workbook(path, read_only=True)
    try:
        assert loaded.active["A1"].number_format == "General"
    finally:
        loaded.close()


def test_copy_worksheet_copies_are_independent():
    from xlspy import Comment, Workbook

    wb = Workbook()
    ws = wb.active
    ws["A1"] = 1
    ws["A1"].comment = Comment("note", "me")
    ws.column_dimensions["A"].width = 20
    ws.row_dimensions[2].height = 30
    ws.page_setup.orientation = "landscape"
    ws.sheet_view.zoomScale = 150
    ws.auto_filter.ref = "A1:B5"
    ws.auto_filter.add_filter_column(0, ["x"])
    copied = wb.copy_worksheet(ws)
    copied.column_dimensions["A"].width = 5
    copied.row_dimensions[2].height = 10
    copied.page_setup.orientation = "portrait"
    copied.sheet_view.zoomScale = 50
    copied["A1"].comment = Comment("changed", "me")
    copied.auto_filter.filterColumn.clear()
    assert ws.column_dimensions["A"].width == 20
    assert ws.row_dimensions[2].height == 30
    assert ws.page_setup.orientation == "landscape"
    assert ws.sheet_view.zoomScale == 150
    assert ws["A1"].comment.text == "note"
    assert len(ws.auto_filter.filterColumn) == 1
    assert copied.column_dimensions["A"].width == 5 and copied.auto_filter.ref == "A1:B5"
