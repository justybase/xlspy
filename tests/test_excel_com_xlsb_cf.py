"""Excel reads conditional formats written into XLSB files, new and loaded."""


from dataclasses import replace
from pathlib import Path

import pytest

from xlspy import Workbook, load_workbook
from xlspy.styles import (
    CellIsRule, ColorScaleRule, DataBarRule, DifferentialStyle, Font, FormatObject, IconSetRule,
    PatternFill,
)

pytestmark = pytest.mark.com  # skipped unless XLSPY_RUN_EXCEL_COM=1 (conftest.py)

XL_XLSB = 50
XL_CELL_VALUE, XL_COLOR_SCALE, XL_DATABAR, XL_ICON_SETS = 1, 3, 4, 6
XL_CONDITION_VALUE_FORMULA = 4
XL_3_ARROWS, XL_5_RATING = 1, 16

X14_DATABARS = Path(__file__).parent / "data" / "x14_databars.xlsb"


def _rules():
    red = DifferentialStyle(font=Font(color="FF9C0006", bold=True),
                            fill=PatternFill("solid", start_color="FFFFC7CE"))
    return [
        ("A1:A10", CellIsRule("A1", "greaterThan", ["5"], dxf=red)),
        ("B1:B10", ColorScaleRule("B1", FormatObject("num", 2), FormatObject("formula", "$D$1"),
                                  FormatObject("percentile", 90), ["FFF8696B", "FFFFEB84", "FF63BE7B"])),
        ("C1:C10", DataBarRule("C1", FormatObject("min"), FormatObject("formula", "$A$1+$A$3*2"),
                               "FF638EC6", show_value=False)),
        ("D1:D10", IconSetRule("D1", "3Arrows", [0, 33, 67], reverse=True, show_value=False)),
        ("A12:A14 D12:D14", IconSetRule("A12", "5Rating", ["$A$1", "$A$2", "$A$3", "$A$4", "$A$5"],
                                        cfvo_type="formula")),
    ]


def _check(book):
    sheet = book.Worksheets(1)
    found = {}
    for index in range(1, sheet.Cells.FormatConditions.Count + 1):
        condition = sheet.Cells.FormatConditions(index)
        found[condition.AppliesTo.Address] = condition
    cell_is = found["$A$1:$A$10"]
    assert cell_is.Type == XL_CELL_VALUE
    assert cell_is.Formula1 == "=5"
    assert cell_is.Font.Bold
    assert cell_is.Font.Color == 0x06009C  # BGR
    scale = found["$B$1:$B$10"]
    assert scale.Type == XL_COLOR_SCALE
    assert scale.ColorScaleCriteria(2).Type == XL_CONDITION_VALUE_FORMULA
    assert scale.ColorScaleCriteria(2).Value == "=$D$1"
    bar = found["$C$1:$C$10"]
    assert bar.Type == XL_DATABAR
    assert bar.ShowValue is False
    assert bar.MaxPoint.Type == XL_CONDITION_VALUE_FORMULA
    assert bar.MaxPoint.Value == "=$A$1+$A$3*2"
    arrows = found["$D$1:$D$10"]
    assert arrows.Type == XL_ICON_SETS
    assert arrows.IconSet.ID == XL_3_ARROWS
    assert arrows.ReverseOrder and arrows.ShowIconOnly
    rating = found["$A$12:$A$14,$D$12:$D$14"]
    assert rating.IconSet.ID == XL_5_RATING
    assert rating.IconCriteria(2).Type == XL_CONDITION_VALUE_FORMULA
    assert rating.IconCriteria(2).Value == "=$A$2"


def _fill(sheet):
    for row in range(1, 15):
        sheet.append([row, row * 2, row * 3, row * 4])


def test_new_workbook_conditional_formats(tmp_path, open_workbook):
    workbook = Workbook()
    _fill(workbook.active)
    for sqref, rule in _rules():
        workbook.active.conditional_formatting.add(sqref, rule)
    workbook.save(tmp_path / "new.xlsb")
    workbook.close()
    _check(open_workbook(tmp_path / "new.xlsb"))


def test_conditional_formats_added_to_an_excel_file(tmp_path, excel, open_workbook):
    book = excel.Workbooks.Add()
    sheet = book.Worksheets(1)
    for row in range(1, 15):
        for column in range(1, 5):
            sheet.Cells(row, column).Value = row * column
    # Excel's own data bar carries an Excel 2010 extension.
    sheet.Range("F1:F5").FormatConditions.AddDatabar()
    book.SaveAs(str(tmp_path / "native.xlsb"), XL_XLSB)
    book.Close(False)

    workbook = load_workbook(tmp_path / "native.xlsb")
    try:
        assert [type(rule).__name__ for rule in workbook.active.conditional_formatting.all_rules()] == [
            "DataBarRule",
        ]
        for sqref, rule in _rules():
            workbook.active.conditional_formatting.add(sqref, rule)
        workbook.save(tmp_path / "edited.xlsb")
    finally:
        workbook.close()
    book = open_workbook(tmp_path / "edited.xlsb")
    _check(book)
    assert book.Worksheets(1).Range("F1:F5").FormatConditions(1).Type == XL_DATABAR


def test_priority_edit_keeps_x14_data_bar_settings_in_excel(tmp_path, open_workbook):
    workbook = load_workbook(X14_DATABARS)
    try:
        workbook.active.conditional_formatting[0].rules[0] = replace(
            workbook.active.conditional_formatting[0].rules[0], priority=7,
        )
        target = tmp_path / "priority.xlsb"
        workbook.save(target)
    finally:
        workbook.close()

    sheet = open_workbook(target).Worksheets(1)
    first = sheet.Range("A1").FormatConditions(1)
    second = sheet.Range("C1").FormatConditions(1)
    assert first.AxisColor.Color == 65280
    assert second.AxisPosition == 2
    assert first.Priority == 7
