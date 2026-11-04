"""The table XML of an XLSX table with a totals row, as Excel reads it."""

import xml.etree.ElementTree as ET
import zipfile

import pytest

from xlspy import ExcelTable, Formula, Workbook

NAMESPACE = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def _make(path, value):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    for row, values in enumerate([["name", "qty"], ["a", 2], ["b", 3], ["Total", value]], start=1):
        for column, cell in enumerate(values, start=1):
            sheet.cell(row=row, column=column, value=cell)
    sheet.add_table(ExcelTable("Tbl", "A1:B4", totals_row=True))
    workbook.save(path)
    workbook.close()


def _table(path):
    with zipfile.ZipFile(path) as archive:
        return ET.fromstring(archive.read("xl/tables/table1.xml"))


@pytest.mark.parametrize(
    ("value", "function", "label", "formula"),
    [
        (5, "sum", None, None),
        ("5", None, "5", None),
        (None, None, None, None),
        (Formula("SUM(B2:B3)", 5), "custom", None, "SUM(B2:B3)"),
    ],
)
def test_totals_attributes_follow_the_cells(tmp_path, value, function, label, formula):
    path = tmp_path / "totals.xlsx"
    _make(path, value)
    table = _table(path)

    assert table.get("totalsRowCount") == "1"
    assert table.get("totalsRowShown") == "1"
    assert table.find("m:autoFilter", NAMESPACE).get("ref") == "A1:B3"

    columns = table.find("m:tableColumns", NAMESPACE).findall("m:tableColumn", NAMESPACE)
    assert columns[0].get("totalsRowLabel") == "Total"
    assert columns[1].get("totalsRowFunction") == function
    assert columns[1].get("totalsRowLabel") == label
    formula_node = columns[1].find("m:totalsRowFormula", NAMESPACE)
    assert (formula_node.text if formula_node is not None else None) == formula


def test_explicit_totals_function_is_written(tmp_path):
    path = tmp_path / "totals.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    for row, values in enumerate([["name", "qty"], ["a", 2], ["b", 3], ["Total", 5]], start=1):
        for column, cell in enumerate(values, start=1):
            sheet.cell(row=row, column=column, value=cell)
    sheet.add_table(ExcelTable(
        "Tbl", "A1:B4", totals_row=True, totals_row_functions=("label", "average"),
    ))
    workbook.save(path)
    workbook.close()

    columns = _table(path).find("m:tableColumns", NAMESPACE).findall("m:tableColumn", NAMESPACE)
    assert columns[0].get("totalsRowLabel") == "Total"
    assert columns[1].get("totalsRowFunction") == "average"
