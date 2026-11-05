"""Round-trip checks for interacting table, formula and worksheet features."""

import zipfile
from xml.etree import ElementTree as ET

import pytest

from xlspy import ExcelTable, Formula, IgnoredError, Workbook, load_workbook


def _create_workbook(path):
    workbook = Workbook()
    data = workbook.active
    data.title = "Data"
    for row in (
        ["name", "qty"],
        ["a", 2],
        ["b", 3],
        ["Total", Formula("SUM(B2:B3)", 5)],
    ):
        data.append(row)
    data.add_table(ExcelTable("Sales", "A1:B4", totals_row=True))
    data.ignored_errors.append(IgnoredError("D1", formula=True))
    data.protection.sheet = True

    calc = workbook.create_sheet("Calc")
    calc["A1"] = Formula("SEQUENCE(2,2)", spill_ref="A1:B2")
    calc["D1"] = Formula("LET(x,1,x+1)", 2)
    workbook.save(path)
    workbook.close()


def _edit_loaded_workbook(source, target):
    workbook = load_workbook(source, data_only=False)
    try:
        data = workbook["Data"]
        table = data.tables["Sales"]
        table.displayName = table.name = "Revenue"
        table.tableColumns[1].name = "units"
        data["B1"] = "units"
        data["D1"] = "=SUM(Revenue[units])"

        calc = workbook["Calc"]
        calc["E1"] = 42
        workbook.save(target)
    finally:
        workbook.close()


def _parts(path):
    with zipfile.ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


@pytest.mark.parametrize("suffix", ["xlsx", "xlsb"])
def test_loaded_save_keeps_interacting_features_and_is_idempotent(tmp_path, suffix):
    source = tmp_path / f"source.{suffix}"
    target = tmp_path / f"edited.{suffix}"
    second_save = tmp_path / f"second.{suffix}"
    _create_workbook(source)
    _edit_loaded_workbook(source, target)

    content_types = ET.fromstring(_parts(target)["[Content_Types].xml"])
    overrides = [
        node.get("PartName")
        for node in content_types.findall(
            "{http://schemas.openxmlformats.org/package/2006/content-types}Override"
        )
    ]
    assert len(overrides) == len(set(overrides))

    workbook = load_workbook(target, data_only=False)
    try:
        data = workbook["Data"]
        assert data.tables["Revenue"].tableColumns[1].name == "units"
        assert data.tables["Revenue"].totalsRowCount == 1
        assert data["D1"].value == "=SUM(Revenue[units])"
        assert data.protection.sheet
        assert [(item.sqref, item.formula) for item in data.ignored_errors] == [
            ("D1", True),
        ]

        calc = workbook["Calc"]
        assert calc["A1"].value == "=_xlfn.SEQUENCE(2,2)"
        assert calc["D1"].value == "=_xlfn.LET(_xlpm.x,1,_xlpm.x+1)"
        assert calc["E1"].value == 42
        workbook.save(second_save)
    finally:
        workbook.close()

    assert _parts(second_save) == _parts(target)


@pytest.mark.com
@pytest.mark.parametrize("suffix", ["xlsx", "xlsb"])
def test_excel_opens_the_mixed_round_trip(tmp_path, open_workbook, suffix):
    source = tmp_path / f"source.{suffix}"
    target = tmp_path / f"edited.{suffix}"
    _create_workbook(source)
    _edit_loaded_workbook(source, target)

    data = open_workbook(target).Worksheets("Data")
    assert data.ListObjects("Revenue").ShowTotals
    assert data.Range("D1").Value2 == 5
    assert data.ProtectContents

    calc = data.Parent.Worksheets("Calc")
    calc.Application.Calculate()
    assert calc.Range("A1").HasSpill
    assert calc.Range("A1").SpillingToRange.Address == "$A$1:$B$2"
    assert calc.Range("D1").Value2 == 2
