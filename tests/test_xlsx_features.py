"""P0/P1 XLSX features: formulas, validation, hyperlinks, images, printing, charts.

Each feature is checked two ways: the XML that xlspy writes, and what openpyxl
reads back. openpyxl is the reference reader for XLSX here.
"""

import io
import struct
import zipfile
import zlib

import pytest

from xlspy import (
    Chart,
    ChartAxis,
    ChartLegend,
    ChartMarker,
    ChartSeries,
    DataLabels,
    DataValidation,
    ExcelTable,
    Formula,
    Hyperlink,
    Image,
    PageMargins,
    PrintSettings,
    Style,
    XlsxWriter,
)

openpyxl = pytest.importorskip("openpyxl")


def _png_bytes(width, height):
    raw = b"".join(b"\x00" + b"\xff\x00\x00" * width for _ in range(height))

    def chunk(kind, data):
        return (
            struct.pack(">I", len(data)) + kind + data
            + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF)
        )

    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
        + chunk(b"IDAT", zlib.compress(raw))
        + chunk(b"IEND", b"")
    )


def _jpeg_bytes(width, height):
    # Minimal header-only JPEG: SOI, APP0, SOF0 with the dimensions, EOI.
    app0 = b"\xff\xe0" + struct.pack(">H", 16) + b"JFIF\x00\x01\x01\x00\x00\x01\x00\x01\x00\x00"
    sof = (
        b"\xff\xc0" + struct.pack(">H", 17) + b"\x08"
        + struct.pack(">HH", height, width) + b"\x03"
        + b"\x01\x11\x00\x02\x11\x00\x03\x11\x00"
    )
    return b"\xff\xd8" + app0 + sof + b"\xff\xd9"


def _read(path, sheet="Sheet1"):
    return openpyxl.load_workbook(path)[sheet]


def _zip_text(path, name):
    with zipfile.ZipFile(path) as archive:
        return archive.read(name).decode("utf-8")


# ---------------------------------------------------------------- formulas


def test_formula_model_validates_and_normalizes():
    assert Formula("=SUM(A1:A3)").text == "SUM(A1:A3)"
    assert Formula("  A2-B2 ").text == "A2-B2"
    with pytest.raises(ValueError, match="empty"):
        Formula("=")
    with pytest.raises(ValueError, match="unbalanced parentheses"):
        Formula("SUM(A1:A3")
    with pytest.raises(ValueError, match="unterminated string"):
        Formula('IF(A1="x,1)')
    with pytest.raises(ValueError, match="control characters"):
        Formula("A1\n+B1")
    with pytest.raises(TypeError, match="cached value"):
        Formula("A1", value=True)
    with pytest.raises(ValueError, match="A1 range|invalid A1"):
        Formula("A1", array_ref="nope")
    with pytest.raises(ValueError, match="_xlpm"):
        Formula("LET(x,1,x)")


def test_future_functions_get_xlfn_prefix_but_text_is_untouched():
    formula = Formula('XLOOKUP(A2,Lookup!A:A,Lookup!B:B)&"XLOOKUP("')
    assert formula.text == '_xlfn.XLOOKUP(A2,Lookup!A:A,Lookup!B:B)&"XLOOKUP("'
    assert Formula("SUM(A1:A3)").text == "SUM(A1:A3)"
    assert Formula("_xlfn.XLOOKUP(A1,B:B,C:C)").text == "_xlfn.XLOOKUP(A1,B:B,C:C)"
    assert Formula("STDEV.S(A1:A3)").text == "_xlfn.STDEV.S(A1:A3)"


def test_formula_cells_round_trip_with_cached_values(tmp_path):
    path = tmp_path / "formulas.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Report")
        writer.write_sheet([
            ["Revenue", "Cost", "Profit", "Label"],
            [1000, 700, Formula("A2-B2", value=300), Formula('"Gain: "&C2', value="Gain: 300")],
            [2000, 900, Formula("A3-B3"), None],
            ["Total", None, Formula("SUM(C2:C3)"), None],
        ], styles={"C2:C4": Style(number_format="#,##0")})

    sheet = _read(path, "Report")
    assert sheet["C2"].value == "=A2-B2"
    assert sheet["C4"].value == "=SUM(C2:C3)"
    assert sheet["D2"].value == '="Gain: "&C2'
    assert sheet["C2"].number_format == "#,##0"

    cached = openpyxl.load_workbook(path, data_only=True)["Report"]
    assert cached["C2"].value == 300
    assert cached["D2"].value == "Gain: 300"
    assert cached["C3"].value is None  # no cached value written, Excel recalculates


def test_strings_that_start_with_equals_stay_text(tmp_path):
    path = tmp_path / "text.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Sheet1")
        writer.write_sheet([["=SUM(A1:A3)", "=1+1"]])
    assert "<f>" not in _zip_text(path, "xl/worksheets/sheet1.xml")
    sheet = _read(path, "Sheet1")
    assert sheet["A1"].value == "=SUM(A1:A3)"
    assert sheet["A1"].data_type == "s"


def test_cross_sheet_references_keep_quoted_sheet_names(tmp_path):
    path = tmp_path / "quoted.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Q1 Sales")
        writer.write_sheet([[1], [2]])
        writer.add_sheet("O'Brien")
        writer.write_sheet([[Formula("SUM('Q1 Sales'!A1:A2)")]])
    assert _read(path, "O'Brien")["A1"].value == "=SUM('Q1 Sales'!A1:A2)"


def test_formulas_use_recalculation_and_escape_xml(tmp_path):
    path = tmp_path / "escape.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Sheet1")
        writer.write_sheet([[Formula("IF(A1<0,\"<neg>\",A1&\"x\")")]])
    sheet = _zip_text(path, "xl/worksheets/sheet1.xml")
    assert "<f>IF(A1&lt;0,&quot;&lt;neg&gt;&quot;" in sheet or "&lt;neg&gt;" in sheet
    assert 'fullCalcOnLoad="1"' in _zip_text(path, "xl/workbook.xml")


def test_formula_in_table_and_array_formula(tmp_path):
    path = tmp_path / "table.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(
            [["Qty", "Price", "Total"], [2, 3, Formula("A2*B2")], [4, 5, Formula("A3*B3")]],
            tables=[ExcelTable("Sales", "A1:C3")],
        )
        writer.add_sheet("Arrays")
        writer.write_sheet([[1, 2], [3, 4], [Formula("A1:A2*B1:B2", array_ref="C1:C2")]])
    assert _read(path, "Data")["C2"].value == "=A2*B2"
    arrays = _zip_text(path, "xl/worksheets/sheet2.xml")
    assert '<f t="array" ref="C1:C2">A1:A2*B1:B2</f>' in arrays


def test_xlfn_prefix_reaches_the_file_and_openpyxl(tmp_path):
    path = tmp_path / "xlfn.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Sheet1")
        writer.write_sheet([["k", "v"], [1, "a"], [Formula("XLOOKUP(1,A2:A2,B2:B2)")]])
    assert _read(path)["A3"].value == "=_xlfn.XLOOKUP(1,A2:A2,B2:B2)"


def test_formula_in_dataframe_object_column(tmp_path):
    pd = pytest.importorskip("pandas")
    path = tmp_path / "df_formulas.xlsx"
    df = pd.DataFrame({"Revenue": [10, 20], "Tax": [Formula("A2*0.2"), Formula("A3*0.2")]})
    with XlsxWriter(path) as writer:
        writer.add_sheet("Sheet1")
        writer.write_dataframe(df)
    assert _read(path)["B2"].value == "=A2*0.2"


# -------------------------------------------------------------- validation


def test_data_validation_model_rejects_invalid_configurations():
    with pytest.raises(ValueError, match="invalid validation type"):
        DataValidation("any", "A1")
    with pytest.raises(ValueError, match="needs formula2"):
        DataValidation("whole", "A1", formula1=1, operator="between")
    with pytest.raises(ValueError, match="formula2 is only used"):
        DataValidation("whole", "A1", formula1=1, formula2=5, operator="greater_than")
    with pytest.raises(ValueError, match="operator does not apply"):
        DataValidation("list", "A1", formula1='"a,b"', operator="equal")
    with pytest.raises(ValueError, match="show_dropdown only applies"):
        DataValidation("whole", "A1", formula1=1, show_dropdown=False)
    with pytest.raises(ValueError, match="wrapped in double quotes"):
        DataValidation("list", "A1", formula1="a,b")
    with pytest.raises(ValueError, match="at most 255"):
        DataValidation("list", "A1", formula1='"' + "x" * 300 + '"')
    with pytest.raises(ValueError, match="prompt text is set"):
        DataValidation("list", "A1", formula1='"a"', prompt="hi")
    with pytest.raises(ValueError, match="invalid error_style"):
        DataValidation("list", "A1", formula1='"a"', error_style="fatal")
    with pytest.raises(ValueError, match="invalid A1"):
        DataValidation("list", "not a range", formula1='"a"')
    with pytest.raises(ValueError, match="needs formula1"):
        DataValidation("custom", "A1")


def test_list_dropdown_with_literal_and_range_sources(tmp_path):
    path = tmp_path / "dv.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Form")
        writer.write_sheet(
            [["Status"], ["New"]],
            data_validations=[
                DataValidation(
                    "list", "B2:B1000", formula1='"New,Active,Closed"', allow_blank=True,
                    show_input_message=True, prompt_title="Pick", prompt="One of three",
                    error_title="Bad", error="Choose from the list", error_style="warning",
                ),
                DataValidation("list", ["D2:D5", "F2"], formula1="Lists!$A$1:$A$9", show_dropdown=False),
            ],
        )
        writer.add_sheet("Lists")
        writer.write_sheet([["a"]])

    sheet = _read(path, "Form")
    rules = {str(dv.sqref): dv for dv in sheet.data_validations.dataValidation}
    literal = rules["B2:B1000"]
    assert literal.type == "list"
    assert literal.formula1 == '"New,Active,Closed"'
    assert literal.allow_blank is True
    assert literal.showInputMessage is True
    assert literal.promptTitle == "Pick"
    assert literal.errorStyle == "warning"
    assert literal.error == "Choose from the list"
    ranged = rules["D2:D5 F2"]
    assert ranged.formula1 == "Lists!$A$1:$A$9"
    # OOXML's showDropDown="1" hides the arrow, so the flag is written only then.
    assert ranged.showDropDown is True


@pytest.mark.parametrize("kind,operator,values,expected", [
    ("whole", "between", (1, 10), ("1", "10")),
    ("decimal", "not_between", (0.5, 2.5), ("0.5", "2.5")),
    ("text_length", "less_than_or_equal", (20, None), ("20", None)),
    ("custom", None, ("ISNUMBER(A1)", None), ("ISNUMBER(A1)", None)),
])
def test_numeric_and_custom_validation_operators(tmp_path, kind, operator, values, expected):
    path = tmp_path / f"dv_{kind}.xlsx"
    rule = DataValidation(kind, "A1:A10", formula1=values[0], formula2=values[1], operator=operator)
    with XlsxWriter(path) as writer:
        writer.add_sheet("Sheet1")
        writer.write_sheet([["x"]], data_validations=[rule])
    dv = _read(path).data_validations.dataValidation[0]
    assert dv.type == {"text_length": "textLength"}.get(kind, kind)
    assert dv.formula1 == expected[0]
    assert dv.formula2 == expected[1]
    expected_operator = {
        "not_between": "notBetween", "less_than_or_equal": "lessThanOrEqual",
    }.get(operator, operator)
    # OOXML omits the operator when it is the default "between".
    assert dv.operator == (None if operator in (None, "between") else expected_operator)


def test_date_validation_serializes_dates_as_serials(tmp_path):
    import datetime

    path = tmp_path / "dv_date.xlsx"
    rule = DataValidation("date", "C2", formula1=datetime.date(2026, 1, 1),
                          operator="greater_than_or_equal")
    with XlsxWriter(path) as writer:
        writer.add_sheet("Sheet1")
        writer.write_sheet([["d"]], data_validations=[rule])
    assert _read(path).data_validations.dataValidation[0].formula1 == "46023"


# --------------------------------------------------------------- hyperlinks


def test_hyperlinks_external_internal_and_styled(tmp_path):
    path = tmp_path / "links.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Index")
        writer.write_sheet([
            ["Name", "Link"],
            ["Report", Hyperlink("https://example.com/report?a=1&b=2", text="Open report", tooltip="Q3 & Q4")],
            ["Sales", Hyperlink("#'Sales Data'!B2", text="Go to sales")],
            ["Mail", Hyperlink("mailto:team@example.com")],
            ["Styled", Hyperlink("https://example.org", text="Žółw")],
        ], styles={"B5": Style(bold=True)})
        writer.add_sheet("Sales Data")
        writer.write_sheet([["x"], [1]])

    sheet = _read(path, "Index")
    assert sheet["B2"].value == "Open report"
    assert sheet["B2"].hyperlink.target == "https://example.com/report?a=1&b=2"
    assert sheet["B2"].hyperlink.tooltip == "Q3 & Q4"
    assert sheet["B3"].hyperlink.location == "'Sales Data'!B2"
    assert sheet["B4"].hyperlink.target == "mailto:team@example.com"
    assert sheet["B5"].value == "Žółw"
    assert sheet["B5"].font.bold is True
    assert sheet["B2"].font.underline == "single"

    rels = _zip_text(path, "xl/worksheets/_rels/sheet1.xml.rels")
    assert 'TargetMode="External"' in rels
    assert "https://example.com/report?a=1&amp;b=2" in rels


@pytest.mark.parametrize("target", [
    "javascript:alert(1)",
    "file:///C:/secret.txt",
    "data:text/html,hi",
    "relative/report.xlsx",
    "https://example.com/\nbad",
    "",
])
def test_unsafe_or_unsupported_link_targets_are_refused(target):
    with pytest.raises((ValueError, TypeError)):
        Hyperlink(target)


def test_internal_link_to_missing_sheet_fails_on_save(tmp_path):
    path = tmp_path / "missing.xlsx"
    with pytest.raises(ValueError, match="unknown sheet 'Nowhere'"):
        with XlsxWriter(path) as writer:
            writer.add_sheet("Sheet1")
            writer.write_sheet([[Hyperlink("#Nowhere!A1", text="x")]])


def test_hyperlink_cell_rejects_number_format(tmp_path):
    with pytest.raises(TypeError, match="cannot carry a number format"):
        with XlsxWriter(tmp_path / "x.xlsx") as writer:
            writer.add_sheet("Sheet1")
            writer.write_sheet([[(Hyperlink("https://example.com"), "0.00")]])


# ------------------------------------------------------------------- images


def test_png_and_jpeg_images_with_sizes_and_charts(tmp_path):
    pytest.importorskip("PIL")
    png = tmp_path / "logo.png"
    png.write_bytes(_png_bytes(40, 20))
    jpg = tmp_path / "photo.jpg"
    # A real, decodable JPEG: openpyxl decodes pictures with Pillow and drops
    # any it cannot read, so the header-only helper is not enough here.
    from PIL import Image as PILImage

    PILImage.new("RGB", (300, 150), (200, 30, 30)).save(jpg, "JPEG")

    path = tmp_path / "images.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Brand")
        writer.write_sheet(
            [["Metric", "Value"], ["A", 1], ["B", 2]],
            images=[
                Image(png, anchor="D1", width=180, height=60, description="Company logo"),
                Image(jpg, anchor="D10"),
            ],
            charts=[Chart("bar", [ChartSeries("B2:B3", categories="A2:A3")], anchor="F2")],
        )
        writer.add_sheet("Second")
        writer.write_sheet([["x"]], images=[Image(png, anchor="B2", width=20)])

    names = zipfile.ZipFile(path).namelist()
    assert "xl/media/image1.png" in names
    assert "xl/media/image2.jpeg" in names
    assert "xl/media/image1.png" in names and names.count("xl/media/image1.png") == 1
    rels = _zip_text(path, "xl/drawings/_rels/drawing1.xml.rels")
    assert 'Target="../media/image1.png"' in rels
    assert 'Target="../media/image2.jpeg"' in rels
    assert 'Target="../charts/chart1.xml"' in rels
    content_types = _zip_text(path, "[Content_Types].xml")
    assert 'Extension="png" ContentType="image/png"' in content_types
    assert 'Extension="jpeg" ContentType="image/jpeg"' in content_types

    ws = openpyxl.load_workbook(path)["Brand"]
    assert len(ws._images) == 2
    # openpyxl reports the picture's natural size; the anchor carries the requested one.
    assert ws._images[0].anchor.ext.width == 180 * 9525
    assert ws._images[0].anchor.ext.height == 60 * 9525
    assert len(ws._charts) == 1
    assert len(openpyxl.load_workbook(path)["Second"]._images) == 1


def test_image_sizes_come_from_the_file_when_not_given(tmp_path):
    png = tmp_path / "p.png"
    png.write_bytes(_png_bytes(64, 32))
    image = Image(png)
    assert (image.pixel_width, image.pixel_height) == (64, 32)
    proportional = Image(png, width=32)
    assert (proportional.pixel_width, proportional.pixel_height) == (32, 16)


# A one-pixel red GIF.
GIF_1X1 = (
    b"GIF89a\x01\x00\x01\x00\x80\x00\x00\xff\x00\x00\x00\x00\x00"
    b"!\xf9\x04\x00\x00\x00\x00\x00,\x00\x00\x00\x00\x01\x00\x01\x00\x00\x02\x02D\x01\x00;"
)


def bmp_bytes(width, height):
    """A 24-bit bottom-up BMP filled with blue."""
    row = b"\xff\x00\x00" * width
    row += b"\x00" * (-len(row) % 4)
    pixels = row * height
    return (
        b"BM" + struct.pack("<IHHI", 54 + len(pixels), 0, 0, 54)
        + struct.pack("<IiiHHIIiiII", 40, width, height, 1, 24, 0, len(pixels), 2835, 2835, 0, 0)
        + pixels
    )


def test_images_from_memory_gif_and_bmp(tmp_path):
    png = _png_bytes(20, 10)
    images = [
        Image(png, anchor="A1"),
        Image(bytearray(png), anchor="A5", width=40),
        Image(io.BytesIO(png), anchor="A10"),
        Image(GIF_1X1, anchor="C1", width=30),
        Image(bmp_bytes(6, 3), anchor="C5"),
    ]
    assert [(i.extension, i.pixel_width, i.pixel_height) for i in images] == [
        ("png", 20, 10), ("png", 40, 20), ("png", 20, 10), ("gif", 30, 30), ("bmp", 6, 3),
    ]
    path = tmp_path / "memory_images.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Pictures")
        writer.write_sheet([["x"]], images=images)
    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
        types = archive.read("[Content_Types].xml").decode()
        drawing = archive.read("xl/drawings/drawing1.xml").decode()
    # Identical bytes are stored once.
    assert sorted(n for n in names if n.startswith("xl/media/")) == [
        "xl/media/image1.png", "xl/media/image2.gif", "xl/media/image3.bmp",
    ]
    assert 'Extension="gif" ContentType="image/gif"' in types
    assert 'Extension="bmp" ContentType="image/bmp"' in types
    # openpyxl needs Pillow to load pictures, which CI does not install.
    assert drawing.count("</pic>") == 5


def test_image_refuses_unsupported_formats_and_bad_input(tmp_path):
    tiff = tmp_path / "x.tif"
    tiff.write_bytes(b"II*\x00" + b"\x00" * 20)
    with pytest.raises(ValueError, match="not a PNG, JPEG, GIF or BMP"):
        Image(tiff)
    gif = tmp_path / "x.gif"
    gif.write_bytes(b"GIF89a" + b"\x00" * 20)
    with pytest.raises(ValueError, match="empty image size"):
        Image(gif)
    with pytest.raises(TypeError, match="binary mode"):
        Image(io.StringIO("text"))
    with pytest.raises(TypeError, match="image bytes"):
        Image(42)
    truncated = tmp_path / "t.png"
    truncated.write_bytes(b"\x89PNG\r\n\x1a\n" + b"\x00" * 4)
    with pytest.raises(ValueError, match="truncated|IHDR"):
        Image(truncated)
    png = tmp_path / "ok.png"
    png.write_bytes(_png_bytes(4, 4))
    with pytest.raises(ValueError, match="single cell"):
        Image(png, anchor="A1:B2")
    with pytest.raises(TypeError, match="pixels"):
        Image(png, width=1.5)
    with pytest.raises(ValueError, match="between 1"):
        Image(png, width=0)


def test_identical_images_share_one_media_part(tmp_path):
    png = tmp_path / "same.png"
    png.write_bytes(_png_bytes(8, 8))
    path = tmp_path / "shared.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("A")
        writer.write_sheet([["a"]], images=[Image(png, anchor="B2"), Image(png, anchor="B5")])
        writer.add_sheet("B")
        writer.write_sheet([["b"]], images=[Image(png)])
    media = [n for n in zipfile.ZipFile(path).namelist() if n.startswith("xl/media/")]
    assert media == ["xl/media/image1.png"]


# ------------------------------------------------------------------ printing


def test_print_settings_are_written_and_read_back(tmp_path):
    path = tmp_path / "print.xlsx"
    settings = PrintSettings(
        orientation="landscape",
        paper_size="A4",
        margins=PageMargins(left=0.5, right=0.5, top=1, bottom=1),
        print_area="A1:H100",
        repeat_rows="1:2",
        repeat_columns="A:B",
        fit_to_width=1,
        fit_to_height=0,
        center_horizontally=True,
        header="&LQuarterly&CReport&RPage &P of &N",
        footer="&CDraft && confidential",
        row_breaks=[10, 20],
        col_breaks=[3],
    )
    with XlsxWriter(path) as writer:
        writer.add_sheet("Print")
        writer.write_sheet([["h", "x"], [1, 2]], print_settings=settings)
        writer.add_sheet("Plain")
        writer.write_sheet([["y"]])

    ws = openpyxl.load_workbook(path)["Print"]
    assert ws.page_setup.orientation == "landscape"
    assert ws.page_setup.paperSize == 9
    assert ws.sheet_properties.pageSetUpPr.fitToPage is True
    assert ws.page_setup.fitToWidth == 1
    assert ws.page_setup.fitToHeight == 0
    assert ws.print_options.horizontalCentered is True
    assert ws.page_margins.left == 0.5
    assert ws.print_area == "'Print'!$A$1:$H$100"
    assert ws.print_title_rows == "$1:$2"
    assert ws.print_title_cols == "$A:$B"
    assert ws.oddHeader.left.text == "Quarterly"
    assert ws.oddFooter.center.text == "Draft && confidential"
    assert [b.id for b in ws.row_breaks.brk] == [10, 20]
    assert [b.id for b in ws.col_breaks.brk] == [3]
    assert not openpyxl.load_workbook(path)["Plain"].print_area


def test_print_setting_validation():
    with pytest.raises(ValueError, match="orientation"):
        PrintSettings(orientation="sideways")
    with pytest.raises(ValueError, match="paper_size"):
        PrintSettings(paper_size="A9")
    with pytest.raises(ValueError, match="repeat_rows"):
        PrintSettings(repeat_rows="3:1")
    with pytest.raises(ValueError, match="repeat_columns"):
        PrintSettings(repeat_columns="1:2")
    with pytest.raises(ValueError, match="scale cannot be combined"):
        PrintSettings(scale=80, fit_to_width=1)
    with pytest.raises(ValueError, match="between 10 and 400"):
        PrintSettings(scale=5)
    with pytest.raises(ValueError, match="255"):
        PrintSettings(header="x" * 256)
    with pytest.raises(ValueError, match="row_breaks"):
        PrintSettings(row_breaks=[0])
    with pytest.raises(ValueError, match="margin"):
        PageMargins(left=-1)


def test_scale_is_written_when_not_fitting(tmp_path):
    path = tmp_path / "scale.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Sheet1")
        writer.write_sheet([["a"]], print_settings=PrintSettings(scale=80, paper_size="letter"))
    xml = _zip_text(path, "xl/worksheets/sheet1.xml")
    assert '<pageSetup paperSize="1" scale="80"/>' in xml
    assert "fitToPage" not in xml


# ------------------------------------------------------------- chart types


def test_horizontal_bar_chart_keeps_column_bar_default(tmp_path):
    path = tmp_path / "bars.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(
            [["Region", "Sales"], ["North", 10], ["South", 20]],
            charts=[
                Chart("bar", [ChartSeries("B2:B3", categories="A2:A3")], anchor="D2"),
                Chart("barh", [ChartSeries("B2:B3", categories="A2:A3", title="Sales")],
                      anchor="D20", y_axis=ChartAxis(title="Units", minimum=0),
                      x_axis=ChartAxis(reverse=True)),
            ],
        )
    ws = openpyxl.load_workbook(path)["Data"]
    column, horizontal = ws._charts
    assert column.type == "col"
    assert horizontal.type == "bar"
    assert horizontal.y_axis.scaling.min == 0
    assert horizontal.x_axis.scaling.orientation == "maxMin"
    assert horizontal.y_axis.title.tx.rich.p[0].r[0].t == "Units"


def test_scatter_chart_with_explicit_x_values_and_markers(tmp_path):
    path = tmp_path / "scatter.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(
            [["Temp", "Sales", "Other"], [10, 100, 1], [20, 180, 2], [30, 260, 3]],
            charts=[Chart(
                "scatter",
                [
                    ChartSeries("B2:B4", x_values="A2:A4", title="Sales",
                                color="#C00000", marker=ChartMarker("diamond", 9), line=True),
                    ChartSeries("C2:C4", x_values="A2:A4", title="Other"),
                ],
                title="Sales vs temperature",
                x_axis=ChartAxis(title="Temperature", minimum=0, maximum=40, major_unit=10),
                y_axis=ChartAxis(title="Units", number_format="#,##0"),
                legend=ChartLegend("bottom"),
                show_data_labels=True,
            )],
        )
    chart = openpyxl.load_workbook(path)["Data"]._charts[0]
    assert type(chart).__name__ == "ScatterChart"
    assert len(chart.series) == 2
    first = chart.series[0]
    assert first.xVal.numRef.f == "'Data'!$A$2:$A$4"
    assert first.yVal.numRef.f == "'Data'!$B$2:$B$4"
    assert first.marker.symbol == "diamond"
    assert first.marker.size == 9
    assert chart.x_axis.scaling.max == 40
    assert chart.x_axis.majorUnit == 10
    assert chart.legend.position == "b"
    assert chart.dataLabels.showVal is True


def test_scatter_without_x_values_and_default_markers(tmp_path):
    path = tmp_path / "scatter_default.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([["v"], [1], [4]], charts=[Chart("scatter", [ChartSeries("A2:A3")])])
    xml = _zip_text(path, "xl/charts/chart1.xml")
    assert "<c:xVal>" not in xml and "<xVal>" not in xml
    assert '<scatterStyle val="lineMarker"/>' in xml
    assert '<symbol val="circle"/>' in xml


def test_axis_and_legend_defaults_stay_byte_compatible(tmp_path):
    path = tmp_path / "defaults.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([["a", "b"], ["x", 1]],
                           charts=[Chart("bar", [ChartSeries("B2:B2", categories="A2:A2")])])
    xml = _zip_text(path, "xl/charts/chart1.xml")
    # The legend says explicitly that it does not overlay the plot area.
    assert "<legend><legendPos val=\"r\"/><overlay val=\"0\"/></legend>" in xml
    assert "<majorGridlines/>" in xml


def test_chart_text_and_number_settings_are_serialized(tmp_path):
    path = tmp_path / "labels.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(
            [["Month", "Revenue"], ["Jan", 100], ["Feb", 150]],
            charts=[Chart(
                "line",
                [ChartSeries("B2:B3", categories="A2:A3", title="Revenue",
                             color="#00B050", marker=ChartMarker("circle", 6))],
                title="Revenue trend",
                title_size=16,
                x_axis=ChartAxis(title="Month"),
                y_axis=ChartAxis(title="Revenue", minimum=0, number_format="#,##0"),
                data_labels=DataLabels(number_format="#,##0", position="above"),
                legend=False,
            )],
        )
    xml = _zip_text(path, "xl/charts/chart1.xml")
    assert '<numFmt formatCode="#,##0" sourceLinked="0"/>' in xml
    assert '<dLblPos val="t"/>' in xml
    assert 'sz="1600"' in xml
    assert "<legend>" not in xml
    assert 'srgbClr val="00B050"' in xml
    chart = openpyxl.load_workbook(path)["Data"]._charts[0]
    assert chart.y_axis.scaling.min == 0
    assert chart.y_axis.title.tx.rich.p[0].r[0].t == "Revenue"
    assert chart.legend is None


@pytest.mark.parametrize("builder,message", [
    (lambda: Chart("pie", [ChartSeries("B2:B3")], x_axis=ChartAxis()), "no axes"),
    (lambda: Chart("bar", [ChartSeries("B2:B3", x_values="A2:A3")]), "only valid for scatter"),
    (lambda: Chart("scatter", [ChartSeries("B2:B3", categories="A2:A3")]), "take x_values"),
    (lambda: Chart("bar", [ChartSeries("B2:B3", marker=ChartMarker())]), "only valid for line and scatter"),
    (lambda: Chart("pie", [ChartSeries("B2:B3", color="#FF0000")]), "colour slices"),
    (lambda: Chart("bar", [ChartSeries("B2:B3")], data_labels=DataLabels(position="inside_end"), title_size=0), "title_size"),
    (lambda: Chart("bar", [ChartSeries("B2:B3")], data_labels=DataLabels(position="best_fit")), "not valid"),
    (lambda: ChartAxis(minimum=5, maximum=1), "less than axis maximum"),
    (lambda: ChartAxis(major_unit=0), "positive"),
    (lambda: ChartLegend("middle"), "invalid legend position"),
    (lambda: ChartMarker("heart"), "invalid marker symbol"),
    (lambda: ChartSeries("B2:B3", color="blue"), "#RRGGBB"),
    (lambda: Chart("bar", [ChartSeries("B2:B3")], legend="yes"), "legend must be"),
])
def test_chart_model_rejects_invalid_settings(builder, message):
    with pytest.raises((ValueError, TypeError), match=message):
        builder()
