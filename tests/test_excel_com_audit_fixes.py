"""Opt-in Excel COM checks for the audit fixes: what Excel itself reads back from
the files xlspy writes (chart type, named style, autofilter criteria, column
style, sheet protection, page breaks, default row height, tables and control
characters)."""

import gc
import zipfile

import pytest

from xlspy import Chart, ChartSeries, NamedStyle, Table, Workbook, load_workbook
from xlspy.workbook._model import Break, RowBreak

pytestmark = pytest.mark.com  # skipped unless XLSPY_RUN_EXCEL_COM=1 (conftest.py)

XL_RADAR = {-4151, 81, 82}  # xlRadar, xlRadarMarkers, xlRadarFilled


@pytest.fixture(scope="module")
def excel():
    client = pytest.importorskip("win32com.client")
    app = client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    app.EnableEvents = False
    yield app
    gc.collect()
    app.Quit()
    del app
    gc.collect()


@pytest.fixture
def open_workbook(excel):
    opened = []

    def open_file(path):
        workbook = excel.Workbooks.Open(str(path), 0, True)
        opened.append(workbook)
        return workbook

    yield open_file
    while opened:
        opened.pop().Close(False)


def test_radar_chart_is_a_radar_chart_in_excel(open_workbook, tmp_path):
    path = tmp_path / "radar.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    for row in range(1, 6):
        sheet.cell(row=row, column=2, value=row)
    sheet.add_chart(Chart("radar", [ChartSeries("B1:B5")]), "E2")
    workbook.save(path)
    workbook.close()

    opened = open_workbook(path)
    chart = opened.Worksheets(1).ChartObjects(1).Chart
    assert int(chart.ChartType) in XL_RADAR


def test_named_style_is_defined_in_excel(open_workbook, tmp_path):
    openpyxl = pytest.importorskip("openpyxl")
    source = tmp_path / "source.xlsx"
    book = openpyxl.Workbook()
    book.active["A1"] = 1
    book.save(source)

    out = tmp_path / "styled.xlsx"
    workbook = load_workbook(source)
    try:
        workbook.add_named_style(NamedStyle(name="Mine"))
        workbook.active["A1"].style = "Mine"
        workbook.save(out)
    finally:
        workbook.close()

    opened = open_workbook(out)
    assert opened.Styles("Mine").Name == "Mine"


def test_custom_autofilter_criterion_is_kept_in_excel(open_workbook, tmp_path):
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

    out = tmp_path / "filters_out.xlsx"
    workbook = load_workbook(source)
    try:
        workbook.active.column_dimensions["A"].width = 20
        workbook.save(out)
    finally:
        workbook.close()

    opened = open_workbook(out)
    criteria = opened.Worksheets(1).AutoFilter.Filters(1)
    assert bool(criteria.On)
    assert criteria.Criteria1 == ">2"


def test_style_only_column_keeps_its_font_in_excel(open_workbook, tmp_path):
    openpyxl = pytest.importorskip("openpyxl")
    from openpyxl.styles import Font

    source = tmp_path / "col_style.xlsx"
    book = openpyxl.Workbook()
    book.active["A1"] = "a"
    book.active["B1"] = "b"
    book.active.column_dimensions["B"].font = Font(bold=True)
    book.save(source)

    out = tmp_path / "col_style_out.xlsx"
    workbook = load_workbook(source)
    try:
        workbook.active.column_dimensions["A"].width = 25
        workbook.save(out)
    finally:
        workbook.close()

    # An empty cell takes its font from the column; Columns(2).Font reports None.
    opened = open_workbook(out)
    assert bool(opened.Worksheets(1).Range("B2").Font.Bold)


def test_new_sheet_metadata_is_applied_in_excel(open_workbook, tmp_path):
    source = tmp_path / "base.xlsx"
    base = Workbook()
    base.active["A1"] = "x"
    base.save(source)
    base.close()

    out = tmp_path / "new_sheet.xlsx"
    workbook = load_workbook(source)
    try:
        sheet = workbook.create_sheet("New")
        sheet["A1"] = 1
        sheet.protection.sheet = True
        sheet.row_breaks = RowBreak([Break(id=3)])
        sheet.sheet_format.defaultRowHeight = 20
        sheet.sheet_format.customHeight = True
        workbook.save(out)
    finally:
        workbook.close()

    # Page breaks are not read here: Excel computes them only with a printer
    # (see XLSPY_RUN_EXCEL_PRINT), so the break is checked in the sheet XML.
    with zipfile.ZipFile(out) as archive:
        sheet_xml = next(
            archive.read(name).decode("utf-8")
            for name in archive.namelist()
            if name.startswith("xl/worksheets/") and b"rowBreaks" in archive.read(name)
        )
    assert '<brk id="3"' in sheet_xml

    opened = open_workbook(out)
    new_sheet = opened.Worksheets("New")
    assert bool(new_sheet.ProtectContents)
    assert abs(float(new_sheet.StandardHeight) - 20.0) < 0.5


def test_table_added_to_new_workbook_is_listed_in_excel(open_workbook, tmp_path):
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

    opened = open_workbook(path)
    assert opened.Worksheets("Data").ListObjects("Sales1").Name == "Sales1"


def test_control_characters_round_trip_through_excel(open_workbook, tmp_path):
    from xlspy.xlsx_updater import XlsxUpdater

    path = tmp_path / "ctrl.xlsx"
    workbook = Workbook()
    workbook.active["A1"] = "x"
    workbook.save(path)
    workbook.close()

    updater = XlsxUpdater(path)
    try:
        updater.replace_sheet_data("Sheet", [["a\x01b"]])
        path.write_bytes(updater.to_bytes())
    finally:
        updater.close()

    opened = open_workbook(path)
    assert opened.Worksheets(1).Range("A1").Value2 == "a\x01b"


XL_EXCEL12 = 50  # xlExcel12, the .xlsb file format


@pytest.fixture
def custom_filter_xlsb(excel, tmp_path):
    path = tmp_path / "custom.xlsb"
    workbook = excel.Workbooks.Add()
    try:
        sheet = workbook.Worksheets(1)
        sheet.Range("A1").Value2 = "n"
        for row in range(2, 7):
            sheet.Cells(row, 1).Value2 = row - 1
        sheet.Range("A1:A6").AutoFilter(1, ">2")
        workbook.SaveAs(str(path), XL_EXCEL12)
    finally:
        workbook.Close(False)
    return path


def test_custom_xlsb_filter_is_kept_when_unchanged(open_workbook, custom_filter_xlsb, tmp_path):
    out = tmp_path / "unchanged.xlsb"
    workbook = load_workbook(custom_filter_xlsb)
    try:
        workbook.save(out)
    finally:
        workbook.close()

    criteria = open_workbook(out).Worksheets(1).AutoFilter.Filters(1)
    assert bool(criteria.On)
    assert criteria.Criteria1 == ">2"


def test_custom_xlsb_filter_change_raises_instead_of_dropping_criteria(custom_filter_xlsb, tmp_path):
    workbook = load_workbook(custom_filter_xlsb)
    try:
        workbook.active.auto_filter.filterColumn[0].blank = True
        with pytest.raises(NotImplementedError, match="custom auto-filter"):
            workbook.save(tmp_path / "changed.xlsb")
    finally:
        workbook.close()


def test_custom_xlsb_filter_can_be_removed(open_workbook, custom_filter_xlsb, tmp_path):
    out = tmp_path / "removed.xlsb"
    workbook = load_workbook(custom_filter_xlsb)
    try:
        workbook.active.auto_filter = None
        workbook.save(out)
    finally:
        workbook.close()

    assert not bool(open_workbook(out).Worksheets(1).AutoFilterMode)
