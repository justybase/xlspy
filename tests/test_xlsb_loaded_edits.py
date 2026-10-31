"""Edits saved into a loaded XLSB workbook: persisted, or refused, never dropped."""

import pytest

from xlspy import (
    AboveAverageRule, BlanksRule, CellIsRule, DataValidation, DifferentialStyle, DuplicateRule,
    ExcelTable, Font, FormulaRule, Hyperlink, PatternFill, TextRule, TopBottomRule, Workbook,
    load_workbook,
)


@pytest.fixture
def source(tmp_path):
    workbook = Workbook()
    sheet = workbook.active
    for number in range(1, 6):
        sheet.append([f"name{number}", number, number * 1.5])
    path = tmp_path / "source.xlsb"
    workbook.save(path)
    workbook.close()
    return path


def save_after(path, tmp_path, edit):
    workbook = load_workbook(path)
    try:
        edit(workbook.active)
        workbook.save(tmp_path / "out.xlsb")
    finally:
        workbook.close()


def test_reading_dimensions_is_not_an_edit(source, tmp_path):
    def read_only(sheet):
        assert sheet.row_dimensions[30].height is None
        assert sheet.column_dimensions["Z"].width is None
        sheet["A1"] = "edited"

    save_after(source, tmp_path, read_only)
    reloaded = load_workbook(tmp_path / "out.xlsb")
    try:
        assert reloaded.active["A1"].value == "edited"
    finally:
        reloaded.close()


def styled_workbook():
    import datetime

    from xlspy import Alignment, Border, Font, PatternFill, Protection, Side

    workbook = Workbook()
    sheet = workbook.active
    sheet["A1"] = "bold"
    sheet["A1"].font = Font(bold=True, italic=True, color="FFFF0000", size=14, name="Arial",
                            underline="single")
    sheet["B1"] = 1.5
    sheet["B1"].number_format = "0.00%"
    sheet["C1"] = "fill"
    sheet["C1"].fill = PatternFill(fill_type="solid", start_color="FF00FF00")
    sheet["D1"] = "border"
    sheet["D1"].border = Border(left=Side(style="thin", color="FF0000FF"),
                                bottom=Side(style="double"))
    sheet["E1"] = "align"
    sheet["E1"].alignment = Alignment(horizontal="center", vertical="top", wrap_text=True,
                                      indent=2, text_rotation=45)
    sheet["F1"] = datetime.date(2024, 5, 1)
    sheet["G1"] = "protected"
    sheet["G1"].protection = Protection(locked=False, hidden=True)
    sheet["H1"].fill = PatternFill(fill_type="solid", start_color="FFFFFF00")
    return workbook


def style_view(cell):
    font, fill, border = cell.font, cell.fill, cell.border
    alignment, protection = cell.alignment, cell.protection
    return (
        cell.value, cell.number_format,
        (font.bold, font.italic, font.size, font.name, font.underline,
         font.color.rgb if font.color else None),
        (fill.fill_type, fill.start_color.rgb if fill.start_color else None) if fill else None,
        tuple(
            (side.style, side.color.rgb if side.color else None) if side else None
            for side in ((border.left, border.right, border.top, border.bottom) if border else ())
        ),
        (alignment.horizontal, alignment.vertical, alignment.wrap_text, alignment.indent,
         alignment.text_rotation) if alignment else None,
        (protection.locked, protection.hidden) if protection else None,
    )


def test_loaded_xlsb_styles_match_xlsx(tmp_path):
    workbook = styled_workbook()
    for extension in ("xlsx", "xlsb"):
        workbook.save(tmp_path / f"styled.{extension}")
    workbook.close()
    views = {}
    for extension in ("xlsx", "xlsb"):
        loaded = load_workbook(tmp_path / f"styled.{extension}")
        try:
            views[extension] = [style_view(cell) for row in loaded.active.iter_rows() for cell in row]
        finally:
            loaded.close()
    assert views["xlsb"] == views["xlsx"]


def test_style_edits_are_saved_into_loaded_xlsb(source, tmp_path):
    import datetime
    import zipfile

    from xlspy import Border, Font, PatternFill, Side

    def edit(sheet):
        sheet["A2"].font = Font(bold=True, color="FFFF0000")
        sheet["B2"].fill = PatternFill(fill_type="solid", start_color="FF00FF00")
        sheet["C3"].number_format = "0.000"
        sheet["E5"].border = Border(top=Side(style="thick"))  # a styled blank
        sheet["D1"] = datetime.date(2024, 1, 2)
        sheet["B4"] = 99

    workbook = load_workbook(source)
    try:
        edit(workbook.active)
        workbook.save(tmp_path / "first.xlsb")
        workbook.save(tmp_path / "second.xlsb")
    finally:
        workbook.close()
    for name in ("first.xlsb", "second.xlsb"):
        loaded = load_workbook(tmp_path / name)
        try:
            sheet = loaded.active
            assert sheet["A2"].font.bold and sheet["A2"].font.color.rgb == "FFFF0000"
            assert sheet["A2"].value == "name2"
            assert sheet["B2"].fill.start_color.rgb == "FF00FF00" and sheet["B2"].value == 2
            assert sheet["C3"].number_format == "0.000" and sheet["C3"].value == 4.5
            assert sheet["E5"].border.top.style == "thick" and sheet["E5"].value is None
            assert sheet["D1"].value == datetime.datetime(2024, 1, 2)
            assert sheet["D1"].number_format != "General"
            assert sheet["B4"].value == 99
        finally:
            loaded.close()
    styles = [
        zipfile.ZipFile(tmp_path / name).read("xl/styles.bin")
        for name in ("first.xlsb", "second.xlsb")
    ]
    assert styles[0] == styles[1]


def test_unchanged_loaded_xlsb_keeps_its_style_sheet(source, tmp_path):
    import zipfile

    workbook = load_workbook(source)
    try:
        workbook.active["A1"] = "value only"
        workbook.save(tmp_path / "out.xlsb")
    finally:
        workbook.close()
    with zipfile.ZipFile(source) as before, zipfile.ZipFile(tmp_path / "out.xlsb") as after:
        assert before.read("xl/styles.bin") == after.read("xl/styles.bin")


def test_xlsb_only_styles_are_refused(source, tmp_path):
    from xlspy import Alignment

    def edit(sheet):
        sheet["A1"].alignment = Alignment(relative_indent=1)

    with pytest.raises(NotImplementedError, match="relative_indent"):
        save_after(source, tmp_path, edit)


@pytest.fixture
def linked(tmp_path):
    workbook = Workbook()
    sheet = workbook.active
    for number in range(1, 8):
        sheet.append([f"r{number}", number, number * 2])
    sheet.merge_cells("A1:B1")
    sheet["C7"] = Hyperlink("https://a.example/x", text="A link")
    sheet["C6"] = Hyperlink("#Sheet!A2", text="to A2")
    path = tmp_path / "linked.xlsb"
    workbook.save(path)
    workbook.close()
    return path


def test_loaded_xlsb_hyperlinks_keep_cell_values(linked):
    workbook = load_workbook(linked)
    try:
        sheet = workbook.active
        assert sheet["C7"].value == "A link"
        assert sheet["C7"].hyperlink.target == "https://a.example/x"
        assert sheet["C6"].value == "to A2"
        assert sheet["C6"].hyperlink.target == "#Sheet!A2"
    finally:
        workbook.close()


def test_merges_and_hyperlinks_are_saved_into_loaded_xlsb(linked, tmp_path):
    import zipfile

    workbook = load_workbook(linked)
    try:
        sheet = workbook.active
        sheet.unmerge_cells("A1:B1")
        sheet.merge_cells("A3:C4")
        sheet["A5"].hyperlink = Hyperlink("https://b.example/y")
        sheet["C7"].hyperlink = None
        workbook.save(tmp_path / "out.xlsb")
    finally:
        workbook.close()
    loaded = load_workbook(tmp_path / "out.xlsb")
    try:
        sheet = loaded.active
        assert sorted(map(str, sheet.merged_cells.ranges)) == ["A3:C4"]
        assert sheet["A5"].value == "r5"
        assert sheet["A5"].hyperlink.target == "https://b.example/y"
        assert sheet["C7"].hyperlink is None and sheet["C7"].value == "A link"
        assert sheet["C6"].hyperlink.target == "#Sheet!A2"
    finally:
        loaded.close()
    rels = zipfile.ZipFile(tmp_path / "out.xlsb").read("xl/worksheets/_rels/sheet1.bin.rels").decode()
    assert "b.example" in rels and "a.example" not in rels


def test_removing_all_merges_and_links(linked, tmp_path):
    workbook = load_workbook(linked)
    try:
        sheet = workbook.active
        sheet.unmerge_cells("A1:B1")
        sheet["C6"].hyperlink = None
        sheet["C7"].hyperlink = None
        workbook.save(tmp_path / "out.xlsb")
    finally:
        workbook.close()
    loaded = load_workbook(tmp_path / "out.xlsb")
    try:
        sheet = loaded.active
        assert not sheet.merged_cells.ranges
        assert sheet["C6"].hyperlink is None and sheet["C7"].hyperlink is None
    finally:
        loaded.close()


def validation_rules():
    return [
        DataValidation(type="list", sqref="A1:A5", formula1='"New,Active,Closed"',
                       prompt="Pick", show_input_message=True),
        DataValidation(type="whole", sqref="B1:B7 D1", formula1="1", formula2="10",
                       error="1..10", error_title="Bad", error_style="warning"),
        DataValidation(type="decimal", sqref="C1:C7", operator="greater_than", formula1="0.5"),
        DataValidation(type="custom", sqref="E1:E7", formula1="LEN(E1)<5"),
        DataValidation(type="list", sqref="F1:F3", formula1="$A$1:$A$3", show_dropdown=False),
    ]


def test_loaded_xlsb_validations_match_xlsx(tmp_path):
    workbook = Workbook()
    for rule in validation_rules():
        workbook.active.add_data_validation(rule)
    for extension in ("xlsx", "xlsb"):
        workbook.save(tmp_path / f"rules.{extension}")
    workbook.close()
    found = {}
    for extension in ("xlsx", "xlsb"):
        loaded = load_workbook(tmp_path / f"rules.{extension}")
        try:
            found[extension] = [repr(rule) for rule in loaded.active.data_validations]
        finally:
            loaded.close()
    assert found["xlsb"] == found["xlsx"]
    assert len(found["xlsb"]) == 5


def test_validation_edits_are_saved_into_loaded_xlsb(tmp_path):
    workbook = Workbook()
    for rule in validation_rules()[:3]:
        workbook.active.add_data_validation(rule)
    workbook.save(tmp_path / "rules.xlsb")
    workbook.close()

    loaded = load_workbook(tmp_path / "rules.xlsb")
    try:
        rules = loaded.active.data_validations
        rules.remove(list(rules)[0])
        loaded.active.add_data_validation(validation_rules()[3])
        loaded.save(tmp_path / "edited.xlsb")
    finally:
        loaded.close()
    edited = load_workbook(tmp_path / "edited.xlsb")
    try:
        assert [(rule.type, rule.sqref) for rule in edited.active.data_validations] == [
            ("whole", "B1:B7 D1"), ("decimal", "C1:C7"), ("custom", "E1:E7"),
        ]
        assert list(edited.active.data_validations)[2].formula1 == "LEN(E1)<5"
    finally:
        edited.close()


PNG = bytes.fromhex(
    "89504e470d0a1a0a0000000d4948445200000001000000010802000000907753de"
    "0000000c4944415408d763f8cfc000000301010018dd8db40000000049454e44ae426082"
)


def metadata_workbook():
    from xlspy import Chart, ChartSeries, Image

    workbook = Workbook()
    sheet = workbook.active
    for number in range(1, 12):
        sheet.append([f"r{number}", number, number * 2])
    sheet.column_dimensions["B"].width = 25
    sheet.column_dimensions["C"].hidden = True
    sheet.row_dimensions[3].height = 33
    sheet.row_dimensions[4].hidden = True
    sheet.freeze_panes = "B3"
    sheet.sheet_view.zoomScale = 150
    sheet.sheet_view.showGridLines = False
    sheet.sheet_properties.tabColor = "FF00AA00"
    sheet.page_setup.orientation = "landscape"
    sheet.page_setup.paperSize = 9
    sheet.page_margins.left = 1.25
    sheet.print_options.horizontalCentered = True
    sheet.oddHeader.center.text = "Report"
    sheet.add_chart(Chart("bar", [ChartSeries("B1:B5", title="B")], title="Values"), "E2")
    sheet.add_image(Image(PNG, description="dot"), "H2")
    workbook.properties.title = "Meta test"
    return workbook


def metadata_view(workbook):
    sheet = workbook.active
    return (
        sheet.column_dimensions["B"].width, sheet.column_dimensions["C"].hidden,
        sheet.row_dimensions[3].height, sheet.row_dimensions[4].hidden,
        sheet.freeze_panes, sheet.sheet_view.zoomScale, sheet.sheet_view.showGridLines,
        sheet.sheet_properties.tabColor.rgb,
        sheet.page_setup.orientation, sheet.page_setup.paperSize, sheet.page_margins.left,
        sheet.print_options.horizontalCentered, sheet.oddHeader.center.text,
        [(chart.chart_type, chart.title) for chart in sheet.charts],
        [(image.description, image.anchor) for image in sheet.images],
        workbook.properties.title,
    )


def test_loaded_xlsb_sheet_metadata_matches_xlsx(tmp_path):
    workbook = metadata_workbook()
    for extension in ("xlsx", "xlsb"):
        workbook.save(tmp_path / f"meta.{extension}")
    workbook.close()
    views = {}
    for extension in ("xlsx", "xlsb"):
        loaded = load_workbook(tmp_path / f"meta.{extension}")
        try:
            views[extension] = metadata_view(loaded)
        finally:
            loaded.close()
    assert views["xlsb"] == views["xlsx"]
    assert views["xlsb"][13] == [("bar", "Values")] and len(views["xlsb"][14]) == 1


def test_loaded_xlsb_keeps_read_only_objects_on_save(tmp_path):
    workbook = metadata_workbook()
    workbook.save(tmp_path / "meta.xlsb")
    workbook.close()
    loaded = load_workbook(tmp_path / "meta.xlsb")
    try:
        loaded.active["A1"] = "edited"
        loaded.save(tmp_path / "saved.xlsb")
        before = metadata_view(loaded)
    finally:
        loaded.close()
    again = load_workbook(tmp_path / "saved.xlsb")
    try:
        assert metadata_view(again) == before
        assert again.active["A1"].value == "edited"
    finally:
        again.close()


def formatted_workbook():
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["Name", "Value", "Other"])
    for number in range(1, 10):
        sheet.append([f"n{number}", number, number * 2])
    red = DifferentialStyle(
        font=Font(bold=True, color="FFFF0000"),
        fill=PatternFill(fill_type="solid", start_color="FFFFFF00"),
    )
    rules = sheet.conditional_formatting
    rules.add("B2:B10", CellIsRule("B2:B10", "between", ["2", "$C$1"], red))
    rules.add("C2:C10", FormulaRule("C2:C10", ["C2>10"], red))
    rules.add("B2:B10", TopBottomRule("B2:B10", rank=3, dxf=red))
    rules.add("C2:C10", AboveAverageRule("C2:C10", dxf=red))
    rules.add("A2:A10", DuplicateRule("A2:A10", dxf=red))
    rules.add("A2:A10", TextRule("A2:A10", "contains", "n1", red))
    rules.add("A2:A10", BlanksRule("A2:A10", "blanks", dxf=red))
    sheet.add_table(ExcelTable(display_name="Data", ref="A1:C10"))
    return workbook


def formatting_view(sheet):
    rules = []
    for rule in sheet.conditional_formatting.all_rules():
        font = rule.dxf.font if rule.dxf else None
        fill = rule.dxf.fill if rule.dxf else None
        rules.append((
            type(rule).__name__, rule.sqref, getattr(rule, "operator", None),
            getattr(rule, "formula", None), font and font.bold, font and font.color.rgb,
            fill and fill.start_color.rgb,
        ))
    tables = [
        (table.displayName, table.ref, [column.name for column in table.tableColumns],
         table.headerRowCount, table.tableStyleInfo and table.tableStyleInfo.name)
        for table in sheet.tables
    ]
    return rules, tables


@pytest.fixture
def formatted(tmp_path):
    workbook = formatted_workbook()
    for suffix in ("xlsx", "xlsb"):
        workbook.save(tmp_path / f"cf.{suffix}")
    workbook.close()
    return tmp_path / "cf.xlsb"


def test_loaded_xlsb_formatting_and_tables_match_xlsx(formatted, tmp_path):
    views = {}
    for suffix in ("xlsx", "xlsb"):
        loaded = load_workbook(tmp_path / f"cf.{suffix}")
        try:
            views[suffix] = formatting_view(loaded.active)
        finally:
            loaded.close()
    assert views["xlsb"] == views["xlsx"]
    rules, tables = views["xlsb"]
    assert len(rules) == 7 and rules[0][3] == ("2", "$C$1")
    assert tables == [("Data", "A1:C10", ["Name", "Value", "Other"], 1, "TableStyleMedium9")]


def test_loaded_xlsb_saves_with_formatting_and_tables(formatted, tmp_path):
    save_after(formatted, tmp_path, lambda sheet: sheet.cell(2, 2, 42))
    loaded = load_workbook(tmp_path / "out.xlsb")
    try:
        assert loaded.active["B2"].value == 42
        rules, tables = formatting_view(loaded.active)
        assert len(rules) == 7 and len(tables) == 1
    finally:
        loaded.close()


def test_table_edits_are_refused(formatted, tmp_path):
    with pytest.raises(NotImplementedError, match="table"):
        save_after(formatted, tmp_path, lambda s: s.tables.pop(0))


def test_conditional_formats_are_added_next_to_the_files_own(formatted, tmp_path):
    before = load_workbook(formatted)
    try:
        kept = [repr(rule) for rule in before.active.conditional_formatting.all_rules()]
    finally:
        before.close()
    save_after(formatted, tmp_path, lambda s: s.conditional_formatting.add("A1", FormulaRule("A1", ["TRUE"])))
    reloaded = load_workbook(tmp_path / "out.xlsb")
    try:
        rules = reloaded.active.conditional_formatting.all_rules()
        assert [repr(rule) for rule in rules[:len(kept)]] == kept
        assert isinstance(rules[-1], FormulaRule) and list(rules[-1].formula) == ["TRUE"]
    finally:
        reloaded.close()
