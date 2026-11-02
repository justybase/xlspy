"""Regression tests for the audit fixes: charts, to_xlsx, loaded-file round trips,
new-sheet metadata, table saving, updater text rules and named styles."""

import math
import zipfile
from decimal import Decimal

import pytest

from xlspy import (
    Chart,
    ChartSeries,
    IgnoredError,
    NamedStyle,
    ReadOnlyCell,
    Table,
    Workbook,
    XlsbWriter,
    XlsxWriter,
    load_workbook,
)
from xlspy.dataframe import to_xlsx
from xlspy.workbook._model import Break, RowBreak
from xlspy.xlsb_updater import XlsbUpdater
from xlspy.xlsx_updater import XlsxUpdater


def _part(path, name):
    with zipfile.ZipFile(path) as archive:
        return archive.read(name).decode("utf-8")


def _first_cell(path):
    workbook = load_workbook(path)
    try:
        return workbook.active["A1"].value
    finally:
        workbook.close()


def _save_plain(path):
    workbook = Workbook()
    workbook.active["A1"] = "x"
    workbook.save(path)
    workbook.close()
    return path


# --- charts -----------------------------------------------------------------

def test_radar_chart_is_written_as_radar_not_pie(tmp_path):
    path = tmp_path / "radar.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    for row in range(1, 6):
        sheet.cell(row=row, column=2, value=row)
    sheet.add_chart(Chart("radar", [ChartSeries("B1:B5")]), "E2")
    workbook.save(path)
    workbook.close()

    chart = _part(path, "xl/charts/chart1.xml")
    assert "<radarChart>" in chart
    assert "<radarStyle" in chart
    assert "pieChart" not in chart


@pytest.mark.parametrize("chart_type", [
    "area3d", "bar3d", "bubble", "line3d", "pie3d",
    "projected_pie", "stock", "surface", "surface3d",
])
def test_unsupported_chart_types_raise_instead_of_becoming_pie(chart_type):
    with pytest.raises(ValueError):
        Chart(chart_type, [ChartSeries("B1:B5")])


# --- to_xlsx ----------------------------------------------------------------

def test_to_xlsx_failure_keeps_existing_file(tmp_path):
    path = tmp_path / "keep.xlsx"
    _save_plain(path)
    before = path.read_bytes()

    with pytest.raises(Exception):
        to_xlsx(object(), path)

    assert path.read_bytes() == before


# --- loaded-file round trips ------------------------------------------------

def test_custom_autofilter_survives_column_width_edit(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")
    from openpyxl.worksheet.filters import CustomFilter, CustomFilters, FilterColumn

    source = tmp_path / "filters.xlsx"
    book = openpyxl.Workbook()
    sheet = book.active
    sheet.append(["n"])
    for value in range(1, 6):
        sheet.append([value])
    sheet.auto_filter.ref = "A1:A6"
    sheet.auto_filter.filterColumn.append(FilterColumn(
        colId=0,
        customFilters=CustomFilters(customFilter=[CustomFilter(operator="greaterThan", val="2")]),
    ))
    book.save(source)

    out = tmp_path / "out.xlsx"
    workbook = load_workbook(source)
    try:
        workbook.active.column_dimensions["A"].width = 20
        workbook.save(out)
    finally:
        workbook.close()

    xml = _part(out, "xl/worksheets/sheet1.xml")
    assert "customFilter" in xml
    assert 'operator="greaterThan"' in xml


def test_style_only_column_survives_column_width_edit(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")
    from openpyxl.styles import Font

    source = tmp_path / "col_style.xlsx"
    book = openpyxl.Workbook()
    book.active["A1"] = "a"
    book.active["B1"] = "b"
    book.active.column_dimensions["B"].font = Font(bold=True)
    book.save(source)

    out = tmp_path / "out.xlsx"
    workbook = load_workbook(source)
    try:
        workbook.active.column_dimensions["A"].width = 25
        workbook.save(out)
    finally:
        workbook.close()

    xml = _part(out, "xl/worksheets/sheet1.xml")
    cols = xml[xml.index("<cols>"):xml.index("</cols>")]
    assert 'min="2" max="2"' in cols
    assert ' style="' in cols


def test_new_sheet_in_loaded_workbook_keeps_protection_breaks_and_ignored_errors(tmp_path):
    source = _save_plain(tmp_path / "base.xlsx")
    out = tmp_path / "out.xlsx"

    workbook = load_workbook(source)
    try:
        sheet = workbook.create_sheet("New")
        sheet["A1"] = 1
        sheet.protection.sheet = True
        sheet.ignored_errors.append(IgnoredError("A1", numberStoredAsText=True))
        sheet.row_breaks = RowBreak([Break(id=3)])
        sheet.sheet_format.defaultRowHeight = 20
        workbook.save(out)
    finally:
        workbook.close()

    reloaded = load_workbook(out)
    try:
        sheet = reloaded["New"]
        assert sheet.protection.sheet is True
        assert [item.sqref for item in sheet.ignored_errors] == ["A1"]
        assert [item.id for item in sheet.row_breaks.brk] == [3]
        assert sheet.sheet_format.defaultRowHeight == 20
    finally:
        reloaded.close()


# --- Workbook() with openpyxl-style tables ----------------------------------

def test_new_workbook_saves_table_objects(tmp_path):
    path = tmp_path / "table.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    sheet["A1"] = "Item"
    sheet["B1"] = "Qty"
    sheet["A2"] = "x"
    sheet["B2"] = 1
    sheet.add_table(Table(displayName="Sales1", ref="A1:B2"))
    workbook.save(path)
    workbook.close()

    reloaded = load_workbook(path)
    try:
        assert "Sales1" in [table.displayName for table in reloaded["Data"].tables]
    finally:
        reloaded.close()


# --- updater text rules -----------------------------------------------------

def test_xlsx_updater_escapes_control_characters_and_round_trips(tmp_path):
    path = _save_plain(tmp_path / "ctrl.xlsx")
    updater = XlsxUpdater(path)
    try:
        updater.replace_sheet_data("Sheet", [["a\x01b"]])
        path.write_bytes(updater.to_bytes())
    finally:
        updater.close()
    with zipfile.ZipFile(path) as archive:
        package = "".join(archive.read(name).decode("utf-8") for name in archive.namelist() if name.endswith(".xml"))
    assert "\x01" not in package
    assert "_x0001_" in package
    assert _first_cell(path) == "a\x01b"


def test_xlsx_updater_rejects_text_over_excel_limit(tmp_path):
    path = _save_plain(tmp_path / "long.xlsx")
    updater = XlsxUpdater(path)
    try:
        with pytest.raises(ValueError):
            updater.replace_sheet_data("Sheet", [["x" * 40000]])
    finally:
        updater.close()


@pytest.mark.parametrize("value", [math.nan, math.inf, -math.inf])
def test_xlsx_updater_rejects_non_finite_numbers(tmp_path, value):
    path = _save_plain(tmp_path / "nan.xlsx")
    updater = XlsxUpdater(path)
    try:
        with pytest.raises(ValueError):
            updater.replace_sheet_data("Sheet", [[value]])
    finally:
        updater.close()


def test_xlsb_updater_keeps_control_characters_and_rejects_nan(tmp_path):
    path = tmp_path / "ctrl.xlsb"
    workbook = Workbook()
    workbook.active["A1"] = "x"
    workbook.save(path)
    workbook.close()

    updater = XlsbUpdater(path)
    try:
        with pytest.raises(ValueError):
            updater.replace_sheet_data("Sheet", [[math.nan]])
        updater.replace_sheet_data("Sheet", [["a\x01b"]])
        path.write_bytes(updater.to_bytes())
    finally:
        updater.close()
    assert _first_cell(path) == "a\x01b"


# --- named styles -----------------------------------------------------------

def test_named_style_on_new_workbook_raises_instead_of_being_dropped(tmp_path):
    workbook = Workbook()
    workbook.active["A1"] = 1
    workbook.add_named_style(NamedStyle(name="Mine"))
    try:
        with pytest.raises(NotImplementedError, match="Mine"):
            workbook.save(tmp_path / "new.xlsx")
    finally:
        workbook.close()


def test_named_style_on_loaded_workbook_is_written(tmp_path):
    source = _save_plain(tmp_path / "src.xlsx")
    out = tmp_path / "out.xlsx"
    workbook = load_workbook(source)
    try:
        workbook.add_named_style(NamedStyle(name="Mine"))
        workbook.save(out)
    finally:
        workbook.close()
    assert "Mine" in _part(out, "xl/styles.xml")


# --- writers reject non-finite numbers --------------------------------------

@pytest.mark.parametrize("use_shared_strings", [True, False])
@pytest.mark.parametrize("value", [math.nan, math.inf, -math.inf, Decimal("NaN"), Decimal("Infinity")])
def test_xlsx_writer_rejects_non_finite_numbers(tmp_path, value, use_shared_strings):
    with pytest.raises(ValueError, match="finite"):
        with XlsxWriter(tmp_path / "bad.xlsx", use_shared_strings=use_shared_strings) as writer:
            writer.add_sheet("S")
            writer.write_sheet([[1, value, "x"]])


@pytest.mark.parametrize("use_shared_strings", [True, False])
@pytest.mark.parametrize("value", [math.nan, math.inf, -math.inf, Decimal("NaN")])
def test_xlsb_writer_rejects_non_finite_numbers(tmp_path, value, use_shared_strings):
    with pytest.raises(ValueError, match="finite"):
        with XlsbWriter(tmp_path / "bad.xlsb", use_shared_strings=use_shared_strings) as writer:
            writer.add_sheet("S")
            writer.write_sheet([[1, value, "x"]])


def test_writers_still_write_finite_numbers(tmp_path):
    path = tmp_path / "finite.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("S")
        writer.write_sheet([[1.5, -2.0, Decimal("3.25")]])
    assert [cell.value for cell in load_workbook(path).active[1]] == [1.5, -2, 3.25]


# --- exports ----------------------------------------------------------------

def test_read_only_cell_is_exported():
    assert ReadOnlyCell.__name__ == "ReadOnlyCell"
