import re
import zipfile

import pytest

from xlspy import (
    Alignment,
    Border,
    CellIsRule,
    Color,
    ColorScaleRule,
    DataBarRule,
    DifferentialStyle,
    ExcelTable,
    Font,
    FormatObject,
    FormulaRule,
    GradientFill,
    IconSetRule,
    PatternFill,
    Protection,
    Side,
    Stop,
    Style,
    XlsxWriter,
)


def test_rich_style_round_trip_through_openpyxl(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")
    path = tmp_path / "rich.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(
            [["h1", "h2", "h3"], ["a", 1, 2], ["b", 3, 4]],
            header_style=Style(
                font=Font(
                    name="Calibri", size=14, bold=True, color="FF0000",
                ),
                pattern=PatternFill("solid", start_color="FFFF00"),
                border=Border(
                    left=Side("thin", "000000"),
                    bottom=Side("double", "FF0000"),
                ),
                alignment=Alignment(
                    horizontal="center", vertical="center",
                    wrap_text=True, text_rotation=45,
                ),
                number_format="#,##0",
            ),
            merges=["B2:D4"],
            column_widths={"B": 24},
            freeze_panes="B2",
        )

    sheet = openpyxl.load_workbook(path)["Data"]
    font = sheet["A1"].font
    assert font.bold is True
    assert font.size == 14
    assert font.color.rgb == "FFFF0000"
    assert sheet["A1"].fill.start_color.rgb == "FFFFFF00"
    assert sheet["A1"].fill.patternType == "solid"
    assert sheet["A1"].border.left.style == "thin"
    assert sheet["A1"].border.bottom.style == "double"
    alignment = sheet["A1"].alignment
    assert alignment.horizontal == "center"
    assert alignment.vertical == "center"
    assert alignment.wrap_text is True
    assert alignment.textRotation == 45
    assert sheet["A1"].number_format == "#,##0"
    assert {str(r) for r in sheet.merged_cells.ranges} == {"B2:D4"}
    assert sheet.column_dimensions["B"].width == 24
    assert sheet.freeze_panes == "B2"

    with zipfile.ZipFile(path) as archive:
        styles_xml = archive.read("xl/styles.xml").decode()
    assert re.search(
        r"numFmts.*fonts.*fills.*borders.*cellStyleXfs.*cellXfs"
        r".*cellStyles.*dxfs.*tableStyles",
        styles_xml,
        re.S,
    )


def test_polymorphic_fill_and_precedence(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")
    path = tmp_path / "poly.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(
            [["v"], ["w"], ["x"]],
            styles={
                # fill= is PatternFill, no separate pattern: fills yellow.
                "A1": Style(fill=PatternFill("solid", start_color="FFFF00")),
                # pattern wins over fill str.
                "A2": Style(
                    fill="FF0000",
                    pattern=PatternFill("solid", start_color="00FF00"),
                ),
                # structured font wins per property; wrap falls back to flat.
                "A3": Style(
                    font=Font(bold=True),
                    font_color="0000FF",
                    italic=True,
                    alignment=Alignment(horizontal="right"),
                    horizontal="left",
                    wrap=True,
                ),
            },
        )

    sheet = openpyxl.load_workbook(path)["Data"]
    assert sheet["A1"].fill.start_color.rgb == "FFFFFF00"
    assert sheet["A2"].fill.start_color.rgb == "FF00FF00"
    assert sheet["A3"].font.bold is True
    assert sheet["A3"].font.italic is True
    assert sheet["A3"].font.color.rgb == "FF0000FF"
    assert sheet["A3"].alignment.horizontal == "right"
    assert sheet["A3"].alignment.wrap_text is True


def test_cf_and_table_round_trip(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")
    path = tmp_path / "cf-table.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(
            [["a", "b", "c"], [1, 2, 3], [4, 5, 6], [7, 8, 9], [10, 11, 12]],
            conditional_formats=[
                CellIsRule(
                    "A1:A10", "lessThan", ["5"],
                    dxf=DifferentialStyle(
                        fill=PatternFill("solid", "FFC7CE"),
                    ),
                ),
                ColorScaleRule(
                    "B1:B10", FormatObject("min"), None,
                    FormatObject("max"), ["FF0000", "00FF00"],
                ),
                DataBarRule(
                    "C1:C10", FormatObject("min"), FormatObject("max"),
                    "638EC6",
                ),
            ],
            tables=[ExcelTable("Tbl", "A1:C5", columns=["a", "b", "c"])],
        )

    sheet = openpyxl.load_workbook(path)["Data"]
    total = sum(
        len(sheet.conditional_formatting[range_string])
        for range_string in sheet.conditional_formatting
    )
    assert total == 3
    assert sheet.tables["Tbl"].ref == "A1:C5"

    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
        assert "xl/tables/table1.xml" in names
        assert "xl/worksheets/_rels/sheet1.xml.rels" in names
        styles_xml = archive.read("xl/styles.xml").decode()
        assert '<dxfs count="1">' in styles_xml


def test_dimensions_hidden_and_row_heights(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")
    path = tmp_path / "dims.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(
            [["a", "b"], [1, 2]],
            column_widths={"B": 30, 0: 12},
            hidden_columns=["B"],
            row_heights={1: 30},
            hidden_rows=[0],
            freeze_panes="",
            auto_filter=False,
        )

    sheet = openpyxl.load_workbook(path)["Data"]
    assert sheet.column_dimensions["B"].width == 30
    assert sheet.column_dimensions["B"].hidden is True
    assert sheet.column_dimensions["A"].width == 12
    assert sheet.row_dimensions[2].height == 30
    assert sheet.row_dimensions[1].hidden is True
    assert sheet.freeze_panes is None
    with zipfile.ZipFile(path) as archive:
        sheet_xml = archive.read("xl/worksheets/sheet1.xml").decode()
    assert "autoFilter" not in sheet_xml


def test_plain_write_stays_lean(tmp_path):
    path = tmp_path / "plain.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([["h1", "h2"]] + [[1, "a"]] * 2000)
        assert writer._style_xf_map == {}
        assert writer._font_registry == {}
        assert writer._fill_registry == {}
        assert writer._border_registry == {}
        assert writer._dxf_list == []

    with zipfile.ZipFile(path) as archive:
        styles_xml = archive.read("xl/styles.xml").decode()
        sheet_xml = archive.read("xl/worksheets/sheet1.xml").decode()
    assert "<borders count=\"1\">" in styles_xml
    assert '<dxfs count="0"/>' in styles_xml
    assert "applyBorder" not in styles_xml
    assert "applyProtection" not in styles_xml
    assert "mergeCell" not in sheet_xml
    assert "conditionalFormatting" not in sheet_xml
    assert "tableParts" not in sheet_xml


def test_model_validation_rejects_bad_input():
    with pytest.raises(ValueError):
        Color(rgb="FF0000", indexed=1)
    with pytest.raises(ValueError):
        Color(rgb="not-a-color")
    with pytest.raises(ValueError):
        Color(indexed=99)
    with pytest.raises(ValueError):
        Font(size=0)
    with pytest.raises(ValueError):
        Font(underline="wavy")
    with pytest.raises(ValueError):
        PatternFill("gradient")
    with pytest.raises(ValueError):
        Side("wavy")
    with pytest.raises(ValueError):
        Border(diagonal_direction="sideways")
    with pytest.raises(ValueError):
        Alignment(indent=300)
    with pytest.raises(ValueError):
        Alignment(text_rotation=200)
    with pytest.raises(ValueError):
        CellIsRule("A1:A5 C1:C5", "lessThan", ["5"])
    with pytest.raises(ValueError):
        CellIsRule("A1:A5", "outside", ["5"])
    with pytest.raises(TypeError):
        CellIsRule("A1:A5", "lessThan", ["5"], dxf="red")
    with pytest.raises(ValueError):
        ColorScaleRule(
            "A1:A5", FormatObject("min"), None, FormatObject("max"),
            ["FF0000"],
        )
    with pytest.raises(ValueError):
        ExcelTable("1bad", "A1:B2")
    with pytest.raises(ValueError):
        ExcelTable("Tbl", "A1:C3", columns=["only-one"])
    with pytest.raises(TypeError):
        Style(fill={"pattern": "solid"})
    namespace = {"GradientFill": type("GradientFill", (), {}), "Style": Style}
    with pytest.raises(TypeError):
        # Foreign (for example openpyxl) fill objects are rejected: the model
        # is structurally similar but not interchangeable.
        eval("Style(fill=GradientFill())", namespace)
    # Foreign conditional-formatting classes named like ours are rejected too.
    with pytest.raises(TypeError):
        XlsxWriter.__new__(XlsxWriter)._normalize_cf_rules(
            [type("IconSetRule", (), {"sqref": "A1"})()]
        )
    with pytest.raises(ValueError):
        Color(auto=False)
    with pytest.raises(ValueError):
        Color(rgb="FF0000", auto=True)
    with pytest.raises(ValueError):
        Color(auto=True, tint=0.5)
    with pytest.raises(ValueError):
        Font(charset=300)
    with pytest.raises(TypeError):
        Font(charset=True)
    with pytest.raises(ValueError):
        Font(family=20)
    with pytest.raises(ValueError):
        Font(scheme="small")
    with pytest.raises(TypeError):
        Font(outline="yes")
    with pytest.raises(TypeError):
        Border(diagonal_up="yes")
    with pytest.raises(TypeError):
        Border(start="thin")
    with pytest.raises(ValueError):
        Alignment(reading_order=3)
    with pytest.raises(ValueError):
        Alignment(relative_indent=-1)
    with pytest.raises(TypeError):
        Alignment(justify_last_line="yes")
    with pytest.raises(ValueError):
        Stop("FF0000", 2)
    with pytest.raises(TypeError):
        Stop("FF0000", "middle")
    with pytest.raises(ValueError):
        GradientFill("radial", stop=["FF0000", "00FF00"])
    with pytest.raises(ValueError):
        GradientFill()
    with pytest.raises(ValueError):
        GradientFill("linear", degree=400, stop=["FF0000"])
    with pytest.raises(ValueError):
        GradientFill("path", degree=45, stop=["FF0000"])
    with pytest.raises(TypeError):
        GradientFill("linear", stop="FF0000")
    with pytest.raises(TypeError):
        DifferentialStyle(number_format=123)
    with pytest.raises(TypeError):
        DifferentialStyle(alignment="center")
    with pytest.raises(TypeError):
        DifferentialStyle(protection="locked")
    with pytest.raises(ValueError):
        IconSetRule("A1:A5", "3Stars", [0, 50])
    with pytest.raises(ValueError):
        IconSetRule("A1:A5", "5Arrows", [])
    with pytest.raises(ValueError):
        IconSetRule("A1:A5", "5Arrows", [0], type="bogus")
    with pytest.raises(ValueError):
        IconSetRule("A1:A5", "5Arrows", [0, "=A1"], type="percent")
    with pytest.raises(TypeError):
        IconSetRule("A1:A5", "5Arrows", [0], show_value="yes")
    with pytest.raises(ValueError):
        IconSetRule("A1:A5", "5Arrows", [0], priority=0)


def test_structured_colors_and_protection(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")
    path = tmp_path / "colors.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(
            [["v"]],
            styles={
                "A1": Style(
                    font=Font(color=Color(theme=1), name="Arial", size=9),
                    pattern=PatternFill(
                        "solid",
                        start_color=Color(indexed=3),
                        end_color=Color(indexed=3),
                    ),
                    protection=Protection(locked=False, hidden=True),
                    alignment=Alignment(
                        shrink_to_fit=True, indent=2, text_rotation=255,
                    ),
                ),
            },
        )
    sheet = openpyxl.load_workbook(path)["Data"]
    assert sheet["A1"].font.name == "Arial"
    assert sheet["A1"].protection.locked is False
    assert sheet["A1"].alignment.shrinkToFit is True
    assert sheet["A1"].alignment.indent == 2
    assert sheet["A1"].alignment.textRotation == 255


def test_named_font_does_not_force_theme_scheme(tmp_path):
    path = tmp_path / "font-scheme.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(
            [["v"]],
            styles={"A1": Style(font=Font(name="Arial", size=12))},
        )
    styles = zipfile.ZipFile(path).read("xl/styles.xml").decode()
    font_elements = re.findall(r"<font>.*?</font>", styles)
    arial = next(font for font in font_elements if "Arial" in font)
    # scheme="minor" next to an explicit name would make Excel substitute
    # the theme font (Calibri) for Arial.
    assert "<scheme" not in arial
    assert "<name val=\"Arial\"/>" in arial
    # The built-in default font keeps its scheme.
    assert "<scheme val=\"minor\"/>" in font_elements[0]


def test_font_extended_attributes_round_trip(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")
    path = tmp_path / "font-extras.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(
            [["v"]],
            styles={
                "A1": Style(
                    font=Font(
                        name="Arial",
                        size=11,
                        bold=True,
                        color=Color(auto=True),
                        charset=0,
                        family=2,
                        scheme="major",
                        outline=True,
                        shadow=True,
                        condense=True,
                        extend=True,
                    ),
                ),
            },
        )
    font = openpyxl.load_workbook(path)["Data"]["A1"].font
    assert font.name == "Arial"
    assert font.charset == 0
    assert font.family == 2
    assert font.scheme == "major"
    assert font.outline is True
    assert font.shadow is True
    assert font.condense is True
    assert font.extend is True
    assert font.color.type == "auto"


def test_border_extended_attributes_round_trip(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")
    path = tmp_path / "border-extras.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(
            [["v"]],
            styles={
                "A1": Style(
                    border=Border(
                        start=Side("thin", "000000"),
                        end=Side("thick", "FF0000"),
                        vertical=Side("dotted"),
                        horizontal=Side("dashed"),
                        diagonal=Side("hair", "00FF00"),
                        diagonal_up=True,
                        outline=False,
                    ),
                ),
            },
        )
    border = openpyxl.load_workbook(path)["Data"]["A1"].border
    assert border.start.style == "thin"
    assert border.end.style == "thick"
    assert border.vertical.style == "dotted"
    assert border.horizontal.style == "dashed"
    assert border.diagonal.style == "hair"
    assert border.diagonalUp is True
    assert border.outline is False


def test_alignment_extended_attributes_round_trip(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")
    path = tmp_path / "alignment-extras.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(
            [["v"]],
            styles={
                "A1": Style(
                    alignment=Alignment(
                        horizontal="justify",
                        vertical="top",
                        indent=1,
                        relative_indent=2,
                        justify_last_line=True,
                        reading_order=2,
                    ),
                ),
            },
        )
    alignment = openpyxl.load_workbook(path)["Data"]["A1"].alignment
    assert alignment.horizontal == "justify"
    assert alignment.indent == 1
    assert alignment.relativeIndent == 2
    assert alignment.justifyLastLine is True
    assert alignment.readingOrder == 2


def test_gradient_fill_round_trip(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")
    path = tmp_path / "gradient.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(
            [["a"], ["b"]],
            styles={
                "A1": Style(
                    fill=GradientFill(
                        "linear",
                        degree=90,
                        stop=[Stop("FF0000", 0), Stop("00FF00", 1)],
                    ),
                ),
                "A2": Style(
                    fill=GradientFill("path", top=0.2, stop=["000000", "FFFFFF"]),
                ),
            },
        )
    sheet = openpyxl.load_workbook(path)["Data"]
    linear = sheet["A1"].fill
    assert linear.type == "linear"
    assert linear.degree == 90
    assert [(s.position, s.color.rgb) for s in linear.stop] == [
        (0, "FFFF0000"),
        (1, "FF00FF00"),
    ]
    path_fill = sheet["A2"].fill
    assert path_fill.type == "path"
    assert path_fill.top == 0.2
    # Plain colors are spread evenly between position 0 and 1.
    assert [s.position for s in path_fill.stop] == [0, 1]
    assert [s.color.rgb for s in path_fill.stop] == [
        "FF000000",
        "FFFFFFFF",
    ]


def test_differential_style_full_round_trip(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")
    path = tmp_path / "dxf-full.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(
            [["v"]],
            conditional_formats=[
                CellIsRule(
                    "A1:A5",
                    "lessThan",
                    ["5"],
                    dxf=DifferentialStyle(
                        font=Font(bold=True, color="FF0000"),
                        number_format="0.00%",
                        fill=PatternFill("solid", start_color="FFFF00"),
                        alignment=Alignment(horizontal="center"),
                        border=Border(left=Side("thin", "0000FF")),
                        protection=Protection(locked=False),
                    ),
                ),
            ],
        )
    sheet = openpyxl.load_workbook(path)["Data"]
    rules = [
        rule
        for range_string in sheet.conditional_formatting
        for rule in sheet.conditional_formatting[range_string]
    ]
    assert len(rules) == 1
    dxf = rules[0].dxf
    assert dxf.font.bold is True
    assert dxf.numFmt.formatCode == "0.00%"
    assert dxf.fill.patternType == "solid"
    assert dxf.fill.fgColor.rgb == "FFFFFF00"
    assert dxf.alignment.horizontal == "center"
    assert dxf.border.left.style == "thin"
    assert dxf.protection.locked is False


def test_dxf_components_are_deduplicated_and_schema_ordered(tmp_path):
    path = tmp_path / "dxf-dedup.xlsx"
    dxf = DifferentialStyle(
        font=Font(bold=True),
        number_format="0.00",
        fill=PatternFill("solid", start_color="FF0000"),
        alignment=Alignment(horizontal="center"),
        border=Border(left=Side("thin")),
        protection=Protection(locked=False),
    )
    with XlsxWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(
            [["v"]],
            conditional_formats=[
                CellIsRule("A1:A5", "lessThan", ["5"], dxf=dxf),
                FormulaRule("B1:B5", "A1>0", dxf=dxf),
            ],
        )
    styles = zipfile.ZipFile(path).read("xl/styles.xml").decode()
    # The same dxf used by two rules must be interned once.
    assert '<dxfs count="1">' in styles
    inner = re.search(r"<dxf>(.*?)</dxf>", styles).group(1)
    tags = ["<font>", "<numFmt", "<fill>", "<alignment", "<border", "<protection"]
    positions = [inner.index(tag) for tag in tags]
    assert positions == sorted(positions)


def test_dxf_font_emits_only_explicit_properties(tmp_path):
    path = tmp_path / "dxf-font.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(
            [["v"]],
            conditional_formats=[
                CellIsRule(
                    "A1:A5",
                    "lessThan",
                    ["5"],
                    dxf=DifferentialStyle(font=Font(bold=True)),
                ),
            ],
        )
    styles = zipfile.ZipFile(path).read("xl/styles.xml").decode()
    inner = re.search(r"<dxf>(.*?)</dxf>", styles).group(1)
    assert "<b/>" in inner
    # A dxf font must not force Calibri/theme-1 defaults onto the cell.
    assert "<name" not in inner
    assert "<sz" not in inner
    assert "<family" not in inner
    assert "<scheme" not in inner
    assert "theme=" not in inner


def test_icon_set_rule_round_trip(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")
    path = tmp_path / "icon-set.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(
            [[i] for i in range(10)],
            conditional_formats=[
                IconSetRule(
                    "A1:A10",
                    "5Arrows",
                    [0, 20, 40, 60, 80],
                    type="percent",
                    show_value=True,
                    reverse=True,
                    priority=1,
                ),
                IconSetRule(
                    "B1:B10",
                    "3TrafficLights1",
                    ["=A1", "=A2", "=A3"],
                    type="formula",
                    percent=False,
                    priority=2,
                ),
            ],
        )
    sheet = openpyxl.load_workbook(path)["Data"]
    rules = {
        rule.iconSet.iconSet: rule
        for range_string in sheet.conditional_formatting
        for rule in sheet.conditional_formatting[range_string]
        if rule.type == "iconSet"
    }
    arrows = rules["5Arrows"]
    assert arrows.priority == 1
    assert arrows.iconSet.showValue is True
    assert arrows.iconSet.reverse is True
    assert [(cfvo.type, cfvo.val) for cfvo in arrows.iconSet.cfvo] == [
        ("percent", 0),
        ("percent", 20),
        ("percent", 40),
        ("percent", 60),
        ("percent", 80),
    ]
    lights = rules["3TrafficLights1"]
    assert lights.priority == 2
    assert lights.iconSet.percent is False
    assert [(cfvo.type, cfvo.val) for cfvo in lights.iconSet.cfvo] == [
        ("formula", "=A1"),
        ("formula", "=A2"),
        ("formula", "=A3"),
    ]


def test_gradient_fill_in_differential_style(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")
    path = tmp_path / "dxf-gradient.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(
            [["v"]],
            conditional_formats=[
                FormulaRule(
                    "A1:A5",
                    "A1>0",
                    dxf=DifferentialStyle(
                        fill=GradientFill(
                            "linear",
                            degree=45,
                            stop=[Stop("FF0000", 0), Stop("0000FF", 1)],
                        ),
                    ),
                ),
            ],
        )
    sheet = openpyxl.load_workbook(path)["Data"]
    rules = [
        rule
        for range_string in sheet.conditional_formatting
        for rule in sheet.conditional_formatting[range_string]
    ]
    fill = rules[0].dxf.fill
    assert fill.type == "linear"
    assert fill.degree == 45
    assert [(s.position, s.color.rgb) for s in fill.stop] == [
        (0, "FFFF0000"),
        (1, "FF0000FF"),
    ]
