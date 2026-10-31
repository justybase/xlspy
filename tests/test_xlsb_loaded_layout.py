"""Layout, sheet-list and property edits saved into a loaded workbook.

Every edit runs against the same workbook saved as XLSX and as XLSB, so the
two formats are checked for the same behaviour.
"""

import datetime

import pytest

from xlspy import Comment, DataValidation, Hyperlink, Workbook, load_workbook
from xlspy._metadata import password_hash
from xlspy.workbook import CustomProperty


@pytest.fixture(params=["xlsx", "xlsb"])
def source(request, tmp_path):
    workbook = Workbook()
    first = workbook.active
    first.title = "Data"
    for number in range(1, 11):
        first.append([f"name{number}", number, number * 1.5])
    second = workbook.create_sheet("Other")
    second["A1"] = "=Data!B2*2"
    second["B1"] = "plain"
    path = tmp_path / f"source.{request.param}"
    workbook.save(path)
    workbook.close()
    return path


def round_trip(path, tmp_path, edit):
    out = tmp_path / f"out{path.suffix}"
    workbook = load_workbook(path)
    try:
        edit(workbook)
        workbook.save(out)
    finally:
        workbook.close()
    return load_workbook(out)


def _set_dimensions(workbook):
    sheet = workbook["Data"]
    sheet.column_dimensions["A"].width = 30
    sheet.column_dimensions["C"].hidden = True
    sheet.row_dimensions[2].height = 33
    sheet.row_dimensions[3].hidden = True
    sheet.row_dimensions[20].height = 40  # a row without cells


def _check_dimensions(workbook):
    sheet = workbook["Data"]
    assert sheet.column_dimensions["A"].width == pytest.approx(30, abs=0.01)
    assert sheet.column_dimensions["C"].hidden
    assert sheet.row_dimensions[2].height == pytest.approx(33)
    assert sheet.row_dimensions[3].hidden
    assert sheet.row_dimensions[20].height == pytest.approx(40)
    assert sheet["A2"].value == "name2"


def _set_view(workbook):
    sheet = workbook["Data"]
    sheet.freeze_panes = "B2"
    sheet.sheet_view.zoomScale = 150
    sheet.sheet_view.showGridLines = False
    sheet.sheet_properties.tabColor = "FF00AA00"


def _check_view(workbook):
    sheet = workbook["Data"]
    assert sheet.freeze_panes == "B2"
    assert sheet.sheet_view.zoomScale == 150
    assert sheet.sheet_view.showGridLines is False
    assert sheet.sheet_properties.tabColor.rgb.endswith("00AA00")


def _set_print(workbook):
    sheet = workbook["Data"]
    sheet.page_setup.orientation = "landscape"
    sheet.page_setup.paperSize = 9
    sheet.page_setup.scale = 80
    sheet.print_options.gridLines = True
    sheet.print_options.horizontalCentered = True
    sheet.page_margins.left = 1.25
    sheet.oddHeader.center.text = "Report &P"
    sheet.oddFooter.right.text = "&D"
    sheet.print_area = "A1:C5"
    sheet.print_title_rows = "1:1"


def _check_print(workbook):
    sheet = workbook["Data"]
    assert sheet.page_setup.orientation == "landscape"
    assert sheet.page_setup.paperSize == 9
    assert sheet.page_setup.scale == 80
    assert sheet.print_options.gridLines
    assert sheet.print_options.horizontalCentered
    assert sheet.page_margins.left == pytest.approx(1.25)
    assert sheet.oddHeader.center.text == "Report &P"
    assert sheet.oddFooter.right.text == "&D"
    assert sheet.print_area == "A1:C5"
    assert sheet.print_title_rows == "1:1"


def _set_protection(workbook):
    protection = workbook["Data"].protection
    protection.sheet = True
    protection.password = "secret"
    protection.formatColumns = False
    protection.selectLockedCells = True


def _check_protection(workbook):
    protection = workbook["Data"].protection
    assert protection.sheet
    assert protection.password == password_hash("secret")
    assert protection.formatColumns is False
    assert protection.formatCells is True
    assert protection.selectLockedCells is True
    assert protection.selectUnlockedCells is False


def _set_filter(workbook):
    sheet = workbook["Data"]
    sheet.auto_filter.ref = "A1:C10"
    sheet.auto_filter.add_filter_column(0, ["name1", "name2"])


def _check_filter(workbook):
    sheet = workbook["Data"]
    assert sheet.auto_filter.ref == "A1:C10"
    assert [(column.colId, column.vals) for column in sheet.auto_filter.filterColumn] == [
        (0, ["name1", "name2"]),
    ]


def _set_sheets(workbook):
    workbook["Data"].title = "Renamed"
    new = workbook.create_sheet("New", 0)
    new["A1"] = "fresh"
    new["B2"] = 42
    new["C3"] = datetime.date(2024, 1, 2)
    new.merge_cells("D1:E2")
    new["A1"].hyperlink = Hyperlink("https://example.com")
    new["B2"].comment = Comment("note", "me")
    new.data_validations.append(DataValidation(type="list", formula1='"a,b"', sqref="F1:F5"))
    new.column_dimensions["A"].width = 20
    workbook.move_sheet("Other", offset=-1)
    workbook["Other"].sheet_state = "hidden"
    workbook.active = 0


def _check_sheets(workbook):
    assert workbook.sheetnames == ["New", "Other", "Renamed"]
    assert workbook["Other"].sheet_state == "hidden"
    assert workbook.active.title == "New"
    new = workbook["New"]
    assert new["A1"].value == "fresh"
    assert new["B2"].value == 42
    assert new["C3"].value.date() == datetime.date(2024, 1, 2)
    assert "D1:E2" in [str(item) for item in new.merged_cells.ranges]
    assert new["A1"].hyperlink.target == "https://example.com"
    assert new["B2"].comment.text == "note"
    assert [rule.sqref for rule in new.data_validations] == ["F1:F5"]
    assert new.column_dimensions["A"].width == pytest.approx(20, abs=0.01)
    # The formula still points at the renamed sheet.
    assert workbook["Other"]["A1"].value.replace("'", "") == "=Renamed!B2*2"


def _set_removal(workbook):
    workbook.remove(workbook["Data"])


def _check_removal(workbook):
    assert workbook.sheetnames == ["Other"]
    assert workbook["Other"]["B1"].value == "plain"
    # XLSB formulas point at sheets by index: Excel's #REF! replaces the
    # removed one. XLSX keeps the formula text, which Excel resolves on load.
    assert workbook["Other"]["A1"].value in ("=#REF!B2*2", "=Data!B2*2")


def _set_properties(workbook):
    workbook.properties.title = "Quarterly"
    workbook.properties.creator = "Analyst"
    workbook.custom_doc_props.append(CustomProperty("Project", "Apollo"))


def _check_properties(workbook):
    assert workbook.properties.title == "Quarterly"
    assert workbook.properties.creator == "Analyst"
    assert [(item.name, item.value) for item in workbook.custom_doc_props] == [("Project", "Apollo")]


CASES = {
    "dimensions": (_set_dimensions, _check_dimensions),
    "view": (_set_view, _check_view),
    "print": (_set_print, _check_print),
    "protection": (_set_protection, _check_protection),
    "filter": (_set_filter, _check_filter),
    "sheets": (_set_sheets, _check_sheets),
    "removal": (_set_removal, _check_removal),
    "properties": (_set_properties, _check_properties),
}


@pytest.mark.parametrize("case", sorted(CASES))
def test_edit_is_saved(source, tmp_path, case):
    edit, check = CASES[case]
    reloaded = round_trip(source, tmp_path, edit)
    try:
        check(reloaded)
    finally:
        reloaded.close()


def test_unedited_layout_keeps_the_sheet_part(tmp_path):
    workbook = Workbook()
    workbook.active["A1"] = 1
    workbook.active.column_dimensions["B"].width = 12
    path = tmp_path / "plain.xlsb"
    workbook.save(path)
    workbook.close()

    import zipfile

    loaded = load_workbook(path)
    try:
        loaded.save(tmp_path / "same.xlsb")
    finally:
        loaded.close()
    with zipfile.ZipFile(path) as before, zipfile.ZipFile(tmp_path / "same.xlsb") as after:
        assert before.read("xl/worksheets/sheet1.bin") == after.read("xl/worksheets/sheet1.bin")


def test_repeated_saves_keep_edits(tmp_path):
    workbook = Workbook()
    workbook.active["A1"] = 1
    path = tmp_path / "twice.xlsb"
    workbook.save(path)
    workbook.close()

    loaded = load_workbook(path)
    try:
        loaded.active.column_dimensions["A"].width = 25
        loaded.save(tmp_path / "first.xlsb")
        loaded.active.freeze_panes = "A2"
        loaded.create_sheet("Second")["A1"] = "x"
        loaded.save(tmp_path / "second.xlsb")
    finally:
        loaded.close()
    reloaded = load_workbook(tmp_path / "second.xlsb")
    try:
        assert reloaded.sheetnames == ["Sheet", "Second"]
        assert reloaded.active.column_dimensions["A"].width == pytest.approx(25, abs=0.01)
        assert reloaded["Sheet"].freeze_panes == "A2"
        assert reloaded["Second"]["A1"].value == "x"
    finally:
        reloaded.close()


def test_column_records_use_spec_ids_and_legacy_wrappers_are_dropped(tmp_path):
    import zipfile

    from xlspy.biff_utils import build_record, iter_records

    workbook = Workbook()
    workbook.active["A1"] = 1
    workbook.active.column_dimensions["B"].width = 12
    path = tmp_path / "written.xlsb"
    workbook.save(path)
    workbook.close()

    def ids(file):
        with zipfile.ZipFile(file) as archive:
            data = archive.read("xl/worksheets/sheet1.bin")
        return data, [record.record_id for record in iter_records(data)]

    data, written = ids(path)
    assert 0x186 in written and 0x187 in written and 0x87 not in written

    # Files from xlspy < 0.11 used 0x86/0x87, and 0x86 is also BrtEndWsViews.
    legacy = b"".join(
        build_record({0x186: 0x86, 0x187: 0x87}.get(record.record_id, record.record_id),
                     data[record.data_start:record.data_end])
        for record in iter_records(data)
    )
    legacy_path = tmp_path / "legacy.xlsb"
    with zipfile.ZipFile(path) as source, zipfile.ZipFile(legacy_path, "w") as target:
        for item in source.infolist():
            content = legacy if item.filename == "xl/worksheets/sheet1.bin" else source.read(item)
            target.writestr(item, content)
    loaded = load_workbook(legacy_path)
    try:
        loaded.active.column_dimensions["C"].width = 20
        loaded.save(tmp_path / "fixed.xlsb")
    finally:
        loaded.close()
    _, fixed = ids(tmp_path / "fixed.xlsb")
    assert fixed.count(0x86) == 1 and 0x87 not in fixed
    assert fixed.count(0x186) == 1
