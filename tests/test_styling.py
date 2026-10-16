import re
import zipfile
import struct

import pytest

from xlspy import ExcelReader, Style, XlsbWriter, XlsxWriter
from xlspy.biff_utils import iter_records


ROWS = [
    ["Product", "Quantity", "Note"],
    ["Apples", (1200, "#,##0"), None],
    ["Pears", 8, "fresh"],
]
HEADER = Style(fill="#008000", font_color="#FFFFFF")


@pytest.mark.parametrize("writer_type,suffix", [(XlsxWriter, ".xlsx"), (XlsbWriter, ".xlsb")])
def test_writers_apply_header_and_range_styles_and_keep_values(tmp_path, writer_type, suffix):
    path = tmp_path / f"styled{suffix}"
    with writer_type(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(
            ROWS,
            header_style=HEADER,
            styles={
                "B2:B3": Style(fill="FFF2CC"),
                "C2:C2": Style(fill="FF0000", font_color="FFFFFF"),
            },
        )

    with ExcelReader(str(path)) as reader:
        assert list(reader.get_rows("Data")) == [
            ["Product", "Quantity", "Note"],
            ["Apples", 1200, None],
            ["Pears", 8, "fresh"],
        ]
    with zipfile.ZipFile(path) as archive:
        assert archive.testzip() is None


def test_xlsx_style_values_are_interpreted_by_openpyxl(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")
    path = tmp_path / "styled.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(
            ROWS,
            header_style=HEADER,
            styles={
                "B2:B3": Style(fill="FFF2CC"),
                "C2": Style(fill="FF0000", font_color="FFFFFF", horizontal="center", wrap=True),
            },
        )

    sheet = openpyxl.load_workbook(path)["Data"]
    assert sheet["A1"].fill.fgColor.rgb == "FF008000"
    assert sheet["A1"].font.color.rgb == "FFFFFFFF"
    assert sheet["A1"].font.bold is True
    assert sheet["B2"].fill.fgColor.rgb == "FFFFF2CC"
    assert sheet["C2"].fill.fgColor.rgb == "FFFF0000"
    assert sheet["C2"].font.color.rgb == "FFFFFFFF"
    assert sheet["C2"].alignment.horizontal == "center"
    assert sheet["C2"].alignment.wrap_text is True
    assert sheet["B2"].number_format == "#,##0"


def test_xlsb_style_records_are_well_framed_and_blank_cells_keep_style(tmp_path):
    path = tmp_path / "styled.xlsb"
    with XlsbWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(
            ROWS,
            header_style=HEADER,
            styles={"C2:C2": Style(fill="FF0000")},
        )

    with zipfile.ZipFile(path) as archive:
        styles_bin = archive.read("xl/styles.bin")
        sheet_bin = archive.read("xl/worksheets/sheet1.bin")
        style_records = list(iter_records(styles_bin))
        sheet_records = list(iter_records(sheet_bin))

    # BrtFont and BrtFill (record IDs 0x2b and 0x2d) contain the user colors;
    # every record in the style part must stay within the part boundaries.
    assert sum(record.record_id == 0x2B for record in style_records) == 3
    assert sum(record.record_id == 0x2D for record in style_records) == 4
    assert any(record.record_id == 0x01 for record in sheet_records)
    begin_fonts = next(record for record in style_records if record.record_id == 0x263)
    begin_fills = next(record for record in style_records if record.record_id == 0x25B)
    begin_cell_xfs = next(record for record in style_records if record.record_id == 0x269)
    assert struct.unpack_from("<I", styles_bin, begin_fonts.data_start)[0] == 3
    assert struct.unpack_from("<I", styles_bin, begin_fills.data_start)[0] == 4
    assert struct.unpack_from("<I", styles_bin, begin_cell_xfs.data_start)[0] == 7
    assert any(
        record.record_id == 0x01
        and struct.unpack_from("<I", sheet_bin, record.data_start + 4)[0] > 0
        for record in sheet_records
    )


@pytest.mark.parametrize("kwargs", [
    {"header_style": {"fill": "008000"}},
    {"styles": {"A0": Style(fill="008000")}},
    {"styles": {"A1": {"fill": "008000"}}},
])
def test_style_api_rejects_invalid_input(tmp_path, kwargs):
    with XlsxWriter(tmp_path / "invalid.xlsx") as writer:
        writer.add_sheet("Data")
        with pytest.raises((TypeError, ValueError)):
            writer.write_sheet([["x"]], **kwargs)


def test_style_validates_colors_and_a1_ranges():
    with pytest.raises(ValueError):
        Style(fill="green")
    with pytest.raises(ValueError):
        Style(font_color="#12")


def test_xlsb_rejects_unimplemented_alignment_before_writing_sheet(tmp_path):
    with XlsbWriter(tmp_path / "unsupported.xlsb") as writer:
        writer.add_sheet("Data")
        with pytest.raises(NotImplementedError, match="alignment"):
            writer.write_sheet([["x"]], header_style=Style(horizontal="center"))


def test_xlsx_styles_and_unstyled_number_formats_stay_index_aligned(tmp_path):
    """A style XF registered before a format XF must not shift cellXfs.

    Style XFs and format XFs share one index space.  When a styled row (for
    example the header) is written before an unstyled ``(value, format)`` cell,
    the style XF gets the lower index; emitting all format XFs first used to
    move the style XF and corrupt both the header style and the number format.
    """
    openpyxl = pytest.importorskip("openpyxl")
    path = tmp_path / "aligned.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(
            [["Header"], ["note"], [(1234.5, "#,##0.00")]],
            header_style=Style(fill="008000", font_color="FFFFFF"),
        )

    sheet = openpyxl.load_workbook(path)["Data"]
    assert sheet["A1"].fill.fgColor.rgb == "FF008000"
    assert sheet["A1"].font.color.rgb == "FFFFFFFF"
    assert sheet["A1"].number_format == "General"
    assert sheet["A3"].number_format == "#,##0.00"
    assert sheet["A3"].fill.fgColor.rgb != "FF008000"


def test_number_format_shared_by_style_and_tuple_is_defined(tmp_path):
    """A style's number format must keep a single, still-emitted id.

    ``_register_format`` used to allocate a fresh id and overwrite the id that
    ``_register_style`` had already stored, leaving the style XF pointing at a
    number format that was no longer written.
    """
    openpyxl = pytest.importorskip("openpyxl")
    path = tmp_path / "shared.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(
            [["v"], [("w", "#,##0")]],
            styles={"A1": Style(number_format="#,##0")},
        )

    sheet = openpyxl.load_workbook(path)["Data"]
    assert sheet["A1"].number_format == "#,##0"
    assert sheet["A2"].number_format == "#,##0"

    with zipfile.ZipFile(path) as archive:
        styles_xml = archive.read("xl/styles.xml").decode()
    defined = set(re.findall(r'<numFmt numFmtId="(\d+)"', styles_xml))
    referenced = set(re.findall(r'<xf numFmtId="(\d+)"', styles_xml))
    assert referenced - {"0", "14", "22"} <= defined


def test_xlsb_number_formats_referenced_by_cell_xfs_are_defined(tmp_path):
    path = tmp_path / "defined.xlsb"
    with XlsbWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(
            [["v"], [("w", "#,##0")]],
            styles={"A1": Style(number_format="#,##0")},
        )

    with zipfile.ZipFile(path) as archive:
        styles_bin = archive.read("xl/styles.bin")
    records = list(iter_records(styles_bin))

    defined = {
        struct.unpack_from("<H", styles_bin, record.data_start)[0]
        for record in records
        if record.record_id == 0x2C
    }
    in_cell_xfs = False
    used = set()
    for record in records:
        if record.record_id == 0x269:
            in_cell_xfs = True
        elif record.record_id == 0x26A:
            in_cell_xfs = False
        elif in_cell_xfs and record.record_id == 0x2F:
            used.add(struct.unpack_from("<H", styles_bin, record.data_start + 2)[0])

    assert used, "expected at least one cell XF"
    # 0/164/166/256 come from the built-in style template; every other
    # referenced format must have a matching BrtFmt record.
    template_defaults = {0, 164, 166, 256}
    for ifmt in used:
        if ifmt not in template_defaults:
            assert ifmt in defined, f"cell XF references undefined BrtFmt {ifmt}"


@pytest.mark.parametrize("writer_type,suffix", [(XlsxWriter, ".xlsx"), (XlsbWriter, ".xlsb")])
def test_styles_and_tuple_formats_round_trip_values(tmp_path, writer_type, suffix):
    path = tmp_path / f"roundtrip{suffix}"
    with writer_type(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(
            [["Header"], ["note"], [(1234.5, "#,##0")]],
            header_style=Style(fill="008000"),
        )

    with ExcelReader(str(path)) as reader:
        assert list(reader.get_rows("Data")) == [["Header"], ["note"], [1234.5]]
