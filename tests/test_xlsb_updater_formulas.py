"""Formulas written by ``XlsbUpdater.replace_sheet_data``."""

import shutil

import pytest

from xlspy import ExcelReader, Formula, XlsbUpdater, XlsbWriter

HEADERS = ["Name", "Qty", "Price"]
BODY = [
    ["alpha", 10, 1.5],
    ["beta", 20, 2.5],
    ["gamma", 30, 3.5],
    ["total", Formula("SUM(B2:B4)", 60), Formula("B5*2", 120)],
]


def _template(path):
    with XlsbWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([HEADERS, ["old", 1, 1.0]])


def _replace(updater, stream):
    replace = updater.replace_sheet_data_stream if stream else updater.replace_sheet_data
    replace("Data", [list(row) for row in BODY], headers=HEADERS)


@pytest.mark.parametrize("stream", [False, True])
def test_replace_sheet_data_writes_formulas(tmp_path, stream):
    source = tmp_path / "source.xlsb"
    _template(source)
    target = tmp_path / "target.xlsb"
    shutil.copyfile(source, target)

    updater = XlsbUpdater(target)
    try:
        _replace(updater, stream)
        updater.save(target)
    finally:
        updater.close()

    with ExcelReader(target) as reader:
        assert reader.read_all("Data")[4] == ["total", 60, 120]
    with ExcelReader(target) as reader:
        assert reader.read_all("Data", data_only=False)[4] == [
            "total", "=SUM(B2:B4)", "=B5*2"
        ]


def test_replace_sheet_data_keeps_formulas_in_later_rows(tmp_path):
    source = tmp_path / "source.xlsb"
    _template(source)
    target = tmp_path / "target.xlsb"
    shutil.copyfile(source, target)
    body = [["alpha", Formula("1+1", 2)], ["beta", Formula("A2*2", 4)]]

    updater = XlsbUpdater(target)
    try:
        updater.replace_sheet_data("Data", body, headers=["Name", "Value"])
        updater.save(target)
    finally:
        updater.close()

    with ExcelReader(target) as reader:
        assert reader.read_all("Data", data_only=False) == [
            ["Name", "Value"], ["alpha", "=1+1"], ["beta", "=A2*2"],
        ]


@pytest.mark.parametrize("stream", [False, True])
def test_replace_sheet_data_writes_array_formulas(tmp_path, stream):
    source = tmp_path / "source.xlsb"
    _template(source)
    target = tmp_path / "target.xlsb"
    shutil.copyfile(source, target)
    updater = XlsbUpdater(target)
    try:
        replace = updater.replace_sheet_data_stream if stream else updater.replace_sheet_data
        replace(
            "Data",
            [["x", Formula("1+1", array_ref="B2:B3")]],
            headers=["Name", "Value"],
        )
        updater.save(target)
    finally:
        updater.close()

    with ExcelReader(target) as reader:
        rows = reader.read_all("Data", data_only=False)
    assert rows[0] == ["Name", "Value"]
    assert rows[1] == ["x", "=1+1"]
    assert rows[2] == [None, 0]


@pytest.mark.com
def test_excel_calculates_the_replaced_formulas(tmp_path, open_workbook):
    """Excel recalculates the formulas on load, not the cached values xlspy wrote."""
    source = tmp_path / "source.xlsb"
    _template(source)
    target = tmp_path / "target.xlsb"
    shutil.copyfile(source, target)
    body = [list(row) for row in BODY]
    body[3][1] = Formula("SUM(B2:B4)", 1)
    body[3][2] = Formula("B5*2", 2)

    updater = XlsbUpdater(target)
    try:
        updater.replace_sheet_data("Data", body, headers=HEADERS)
        updater.save(target)
    finally:
        updater.close()

    workbook = open_workbook(target)
    workbook.Application.Calculate()
    sheet = workbook.Worksheets("Data")
    assert sheet.Range("B5").Formula == "=SUM(B2:B4)"
    assert sheet.Range("B5").Value == 60
    assert sheet.Range("C5").Value == 120