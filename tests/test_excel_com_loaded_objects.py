"""Objects added to a loaded workbook, checked in Excel.

Excel opens files with alerts off, so a file it had to repair loses parts;
the counts below would show that. Print setup is checked through the
Print_Area name: reading PageSetup needs a reachable printer.
"""

import gc
import os

import pytest

from xlspy import load_workbook
from test_excel_com_metadata import excel, open_workbook  # noqa: F401
from test_model_loaded_objects import ROWS, _add_everything, _source

pytestmark = [pytest.mark.com, pytest.mark.skipif(os.environ.get("XLSPY_RUN_EXCEL_COM") != "1",
                                                   reason="Set XLSPY_RUN_EXCEL_COM=1")]

XL_VALIDATE_WHOLE_NUMBER = 1
XL_CELL_VALUE = 1
MSO_PICTURE = 13
XL_CHART_TYPE_COLUMN = 51  # xlColumnClustered
XL_CHART_TYPE_LINE = 4


def _check_added_objects(book, *, existing_charts: int, existing_tables: set[str]):
    sheet = book.Worksheets("Data")
    charts = [sheet.ChartObjects(i).Chart for i in range(1, sheet.ChartObjects().Count + 1)]
    assert len(charts) == existing_charts + 1
    assert "New" in [chart.ChartTitle.Text for chart in charts if chart.HasTitle]
    pictures = [shape for shape in sheet.Shapes if shape.Type == MSO_PICTURE]
    assert len(pictures) == 1 and pictures[0].TopLeftCell.Address == "$D$30"
    tables = {sheet.ListObjects(i).Name for i in range(1, sheet.ListObjects.Count + 1)}
    assert tables == existing_tables | {"NewTable"}
    assert sheet.ListObjects("NewTable").Range.Address == "$D$1:$D$4"
    validation = sheet.Range("B3").Validation
    assert validation.Type == XL_VALIDATE_WHOLE_NUMBER and validation.Formula2 == "100"
    rules = sheet.Range("B2:B4").FormatConditions
    # Only cell-value rules have a font (a data bar does not).
    bold = [rules(i) for i in range(1, rules.Count + 1)
            if rules(i).Type == XL_CELL_VALUE and rules(i).Font.Bold is True]
    assert len(bold) == 1 and bold[0].Formula1 == "=15"
    # The value 30 is above 15: the new rule paints it, with its number format.
    # DisplayFormat.NumberFormat stays "General" in Excel, so read the text
    # (the decimal separator follows Excel's locale).
    book.Application.Calculate()
    assert sheet.Range("B3").DisplayFormat.Font.Bold is True
    assert sheet.Range("B3").Text.replace(",", ".") == "3000.0%"
    assert sheet.Names("Print_Area").RefersTo == "=Data!$A$1:$D$4"
    assert sheet.Names("Print_Titles").RefersTo == "=Data!$1:$1"


def test_objects_added_to_a_loaded_xlspy_file_in_excel(tmp_path, open_workbook):
    workbook = load_workbook(_source(tmp_path / "source.xlsx"))
    try:
        _add_everything(workbook["Data"])
        workbook.save(tmp_path / "out.xlsx")
    finally:
        workbook.close()
    book = open_workbook(tmp_path / "out.xlsx")
    _check_added_objects(book, existing_charts=1, existing_tables={"OldTable"})
    # The file's own rule still applies (10 is below 15).
    assert book.Worksheets("Data").Range("B2").DisplayFormat.Font.Italic is True


def _excel_source(excel, path):
    """A workbook saved by Excel: its own chart, table, rule and validation,
    with Excel's prefixes (xdr:, x14:) and part layout."""
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Name = "Data"
        for row, values in enumerate(ROWS, 1):
            for column, value in enumerate(values, 1):
                sheet.Cells(row, column).Value = value
        chart = sheet.Shapes.AddChart2(-1, XL_CHART_TYPE_LINE, 400, 10, 300, 200).Chart
        chart.SetSourceData(sheet.Range("A1:B4"))
        sheet.ListObjects.Add(1, sheet.Range("A1:B4"), None, 1).Name = "OldTable"
        sheet.Range("B2:B4").FormatConditions.AddDatabar()
        sheet.Range("A2:A4").Validation.Add(3, 1, 1, "North,South,East")
        book.Worksheets.Add(After=sheet).Name = "Other"
        book.SaveAs(str(path))
    finally:
        book.Close(False)
        del book
        gc.collect()
    return path


def test_objects_added_to_a_loaded_excel_file_in_excel(tmp_path, excel, open_workbook):
    source = _excel_source(excel, tmp_path / "excel.xlsx")
    workbook = load_workbook(source)
    try:
        _add_everything(workbook["Data"])
        workbook.save(tmp_path / "out.xlsx")
    finally:
        workbook.close()
    book = open_workbook(tmp_path / "out.xlsx")
    _check_added_objects(book, existing_charts=1, existing_tables={"OldTable"})
    sheet = book.Worksheets("Data")
    # Excel's own objects are kept: the data bar, the list validation and the chart.
    kinds = [sheet.Range("B2:B4").FormatConditions(i).Type
             for i in range(1, sheet.Range("B2:B4").FormatConditions.Count + 1)]
    assert 4 in kinds  # xlDatabar
    assert sheet.Range("A2").Validation.Formula1 == "North,South,East"
    assert sheet.ChartObjects(1).Chart.ChartType == XL_CHART_TYPE_LINE
