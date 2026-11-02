"""Native Excel checks for edits saved into a loaded XLSB workbook."""


import pytest

from xlspy import Border, Font, PatternFill, Side, Workbook, load_workbook

pytestmark = pytest.mark.com  # skipped unless XLSPY_RUN_EXCEL_COM=1 (conftest.py)


def test_loaded_xlsb_style_edits_in_excel(tmp_path, open_workbook):
    created = Workbook()
    for number in range(1, 6):
        created.active.append([f"name{number}", number, number * 1.5])
    source = tmp_path / "source.xlsb"
    created.save(source)
    created.close()

    workbook = load_workbook(source)
    try:
        sheet = workbook.active
        sheet["A2"].font = Font(bold=True, color="FFFF0000")
        sheet["B2"].fill = PatternFill(fill_type="solid", start_color="FF00FF00")
        sheet["C3"].number_format = "0.000"
        sheet["E5"].border = Border(top=Side(style="thick"))
        workbook.save(tmp_path / "edited.xlsb")
    finally:
        workbook.close()

    book = open_workbook(tmp_path / "edited.xlsb")
    sheet = book.Worksheets(1)
    assert sheet.Range("A2").Font.Bold and sheet.Range("A2").Font.Color == 255
    assert sheet.Range("A2").Value == "name2"
    assert sheet.Range("B2").Interior.Color == 0x00FF00
    assert sheet.Range("C3").NumberFormat == "0.000" and sheet.Range("C3").Value == 4.5
    assert sheet.Range("E5").Borders(8).Weight == 4  # xlEdgeTop, xlThick


def test_loaded_xlsb_merges_and_links_in_excel(tmp_path, open_workbook):
    from xlspy import Hyperlink

    created = Workbook()
    for number in range(1, 8):
        created.active.append([f"r{number}", number, number * 2])
    created.active.merge_cells("A1:B1")
    created.active["C7"] = Hyperlink("https://a.example/x", text="A link")
    source = tmp_path / "source.xlsb"
    created.save(source)
    created.close()

    workbook = load_workbook(source)
    try:
        sheet = workbook.active
        sheet.unmerge_cells("A1:B1")
        sheet.merge_cells("A3:C4")
        sheet["A5"].hyperlink = Hyperlink("https://b.example/y")
        sheet["C6"].hyperlink = Hyperlink("#Sheet!A2")
        sheet["C7"].hyperlink = None
        workbook.save(tmp_path / "edited.xlsb")
    finally:
        workbook.close()

    book = open_workbook(tmp_path / "edited.xlsb")
    sheet = book.Worksheets(1)
    assert sheet.Range("A3").MergeArea.Address == "$A$3:$C$4"
    assert not sheet.Range("A1").MergeCells
    links = {link.Range.Address: (link.Address, link.SubAddress) for link in sheet.Hyperlinks}
    assert links == {"$A$5": ("https://b.example/y", ""), "$C$6": ("", "Sheet!A2")}


def test_loaded_xlsb_validations_in_excel(tmp_path, open_workbook):
    from xlspy import DataValidation

    created = Workbook()
    created.active.add_data_validation(
        DataValidation(type="list", sqref="A1:A5", formula1='"New,Active"'))
    created.active.add_data_validation(
        DataValidation(type="whole", sqref="B1:B7", formula1="1", formula2="10"))
    source = tmp_path / "source.xlsb"
    created.save(source)
    created.close()

    workbook = load_workbook(source)
    try:
        rules = workbook.active.data_validations
        rules.remove(list(rules)[0])
        workbook.active.add_data_validation(
            DataValidation(type="custom", sqref="E1:E7", formula1="LEN(E1)<5"))
        workbook.save(tmp_path / "edited.xlsb")
    finally:
        workbook.close()

    book = open_workbook(tmp_path / "edited.xlsb")
    sheet = book.Worksheets(1)
    assert sheet.Range("B2").Validation.Type == 1  # xlValidateWholeNumber
    assert sheet.Range("B2").Validation.Formula2 == "10"
    assert sheet.Range("E3").Validation.Type == 7  # xlValidateCustom
    with pytest.raises(Exception):
        sheet.Range("A1").Validation.Type  # the removed rule
