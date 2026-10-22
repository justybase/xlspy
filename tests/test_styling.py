import re
import zipfile
import struct
import datetime

import pytest

from xlspy import (
    Alignment,
    Border,
    ExcelReader,
    Font,
    GradientFill,
    PatternFill,
    Protection,
    Side,
    Stop,
    Style,
    XlsbWriter,
    XlsxWriter,
)
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


@pytest.mark.parametrize("use_shared_strings", [True, False])
def test_xlsx_c_encoder_applies_range_styles_to_cell_kinds(tmp_path, use_shared_strings):
    from xlspy._accel import HAVE_C_EXT
    if not HAVE_C_EXT:
        pytest.skip("requires the optional C encoder")
    openpyxl = pytest.importorskip("openpyxl")
    path = tmp_path / f"styled-fast-path-{use_shared_strings}.xlsx"
    with XlsxWriter(path, useSharedStrings=use_shared_strings) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(
            [
                ["Header", "Text", "Number", "Flag", "Date", "Blank"],
                ["row", "value", 3.5, True, datetime.date(2024, 1, 2), None],
            ],
            header_style=Style(
                fill="008000", font_color="FFFFFF", horizontal="center",
            ),
            styles={
                "B2:F2": Style(fill="FFFF00", horizontal="center"),
                "E2": Style(number_format="yyyy-mm-dd"),
            },
        )

    sheet = openpyxl.load_workbook(path)["Data"]
    assert sheet["A1"].fill.fgColor.rgb == "FF008000"
    assert sheet["A1"].alignment.horizontal == "center"
    for ref in ("B2", "C2", "D2", "E2", "F2"):
        assert sheet[ref].fill.fgColor.rgb == "FFFFFF00"
        assert sheet[ref].alignment.horizontal == "center"
    assert sheet["E2"].number_format == "yyyy-mm-dd"


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


def _xlsb_payloads(styles_bin, record_id):
    return [
        styles_bin[record.data_start:record.data_end]
        for record in iter_records(styles_bin)
        if record.record_id == record_id
    ]


def _xlsb_cell_xfs(styles_bin):
    """Decoded BrtXF records of the CELLXFS collection."""
    records = list(iter_records(styles_bin))
    begin = next(record for record in records if record.record_id == 0x0269)
    end = next(record for record in records if record.record_id == 0x026A)
    return [
        struct.unpack("<HHHHHBBBBH", styles_bin[record.data_start:record.data_end])
        for record in records
        if record.record_id == 0x002F
        and begin.data_end <= record.header_start < end.header_start
    ]


def test_xlsb_cell_xf_encodes_alignment_and_protection(tmp_path):
    path = tmp_path / "xf.xlsb"
    style = Style(
        alignment=Alignment(
            horizontal="center",
            vertical="top",
            wrap_text=True,
            shrink_to_fit=True,
            indent=2,
            text_rotation=45,
            justify_last_line=True,
            reading_order=2,
        ),
        protection=Protection(locked=False, hidden=True),
    )
    with XlsbWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([["header"], ["x"]], styles={"A2": style})

    with zipfile.ZipFile(path) as archive:
        styles_bin = archive.read("xl/styles.bin")
    xfs = _xlsb_cell_xfs(styles_bin)
    assert len(xfs) == 5  # four built-in XFs plus the styled cell
    ixfe_parent, ifmt, font, fill, border, trot, indent, align, flags, grbit = xfs[-1]
    assert (ixfe_parent, ifmt, font, fill, border) == (0, 0, 0, 0, 0)
    assert (trot, indent) == (45, 2)
    assert align == 0xC2  # centre, top, wrap, justify last line
    assert flags == 0x29  # shrink to fit, right-to-left, unlocked, hidden
    assert grbit == 0x24  # applyAlignment | applyProtection


def test_xlsb_structured_font_fill_and_border_records(tmp_path):
    path = tmp_path / "records.xlsb"
    style = Style(
        font=Font(
            name="Arial",
            size=14,
            bold=True,
            italic=True,
            strike=True,
            underline="double",
            vertAlign="superscript",
            color="FF0000",
            scheme="major",
        ),
        fill=PatternFill(
            fill_type="gray125", start_color="00FF00", end_color="0000FF"
        ),
        border=Border(
            left=Side(style="thin", color="FF0000"),
            top=Side(style="double"),
            diagonal=Side(style="dashed"),
            diagonal_up=True,
        ),
    )
    with XlsbWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([["header"], ["x"]], styles={"A2": style})

    with zipfile.ZipFile(path) as archive:
        styles_bin = archive.read("xl/styles.bin")

    fonts = {
        payload[25:].decode("utf-16-le"): payload
        for payload in _xlsb_payloads(styles_bin, 0x002B)
    }
    font = fonts["Arial"]
    height, flags, weight, baseline = struct.unpack_from("<HHHH", font, 0)
    assert (height, flags, weight, baseline) == (280, 0x0B, 700, 1)
    assert font[8] == 0x02  # double underline
    assert font[20] == 0x01  # major scheme
    assert font[12:20] == bytes((0x05, 0, 0, 0, 0xFF, 0x00, 0x00, 0xFF))
    assert "Calibri" in fonts  # the built-in fonts are still present

    fills = _xlsb_payloads(styles_bin, 0x002D)
    assert len(fills) == 3  # none, gray125 and the custom pattern
    assert struct.unpack_from("<I", fills[1], 0)[0] == 17  # gray125
    assert fills[2][4:12] == bytes((0x05, 0, 0, 0, 0x00, 0xFF, 0x00, 0xFF))
    assert fills[2][12:20] == bytes((0x05, 0, 0, 0, 0x00, 0x00, 0xFF, 0xFF))

    borders = _xlsb_payloads(styles_bin, 0x002E)
    assert len(borders) == 2  # the built-in empty border plus the custom one
    border = borders[1]
    assert border[0] == 0x02  # diagonal up
    border_sides = [border[1 + index * 10 : 11 + index * 10] for index in range(5)]
    top, _bottom, left, _right, diagonal = border_sides
    assert top[0] == 0x06  # double
    assert left[0] == 0x01  # thin
    assert left[2:10] == bytes((0x05, 0, 0, 0, 0xFF, 0x00, 0x00, 0xFF))
    assert diagonal[0] == 0x03  # dashed

    with ExcelReader(str(path)) as reader:
        assert list(reader.get_rows("Data")) == [["header"], ["x"]]


def test_xlsb_gradient_fill_record(tmp_path):
    path = tmp_path / "gradient.xlsb"
    with XlsbWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(
            [["x"]],
            header_style=Style(
                fill=GradientFill(
                    fill_type="linear",
                    degree=90,
                    stop=[Stop("FF0000", 0.0), Stop("0000FF", 1.0)],
                )
            ),
        )

    with zipfile.ZipFile(path) as archive:
        styles_bin = archive.read("xl/styles.bin")
    fills = _xlsb_payloads(styles_bin, 0x002D)
    gradient = next(
        payload for payload in fills if struct.unpack_from("<I", payload, 0)[0] == 0x28
    )
    gradient_type, degree = struct.unpack_from("<Id", gradient, 20)
    assert gradient_type == 0  # linear
    assert degree == 270.0  # BIFF12 angles run counter-clockwise
    assert struct.unpack_from("<I", gradient, 64)[0] == 2  # two stops
    assert gradient[68:76] == bytes((0x05, 0, 0, 0, 0xFF, 0x00, 0x00, 0xFF))
    assert struct.unpack_from("<d", gradient, 76)[0] == 0.0
    assert gradient[84:92] == bytes((0x05, 0, 0, 0, 0x00, 0x00, 0xFF, 0xFF))
    assert struct.unpack_from("<d", gradient, 92)[0] == 1.0


@pytest.mark.parametrize("style", [
    Style(alignment=Alignment(relative_indent=1)),
    Style(border=Border(vertical=Side(style="thin"))),
    Style(border=Border(outline=False)),
])
def test_xlsb_rejects_unrepresentable_style_properties(tmp_path, style):
    with XlsbWriter(tmp_path / "unsupported.xlsb") as writer:
        writer.add_sheet("Data")
        with pytest.raises(NotImplementedError):
            writer.write_sheet([["x"]], header_style=style)


def test_xlsb_rejects_values_outside_record_limits(tmp_path):
    with XlsbWriter(tmp_path / "long-name.xlsb") as writer:
        writer.add_sheet("Data")
        with pytest.raises(ValueError, match="31 characters"):
            writer.write_sheet([["x"]], header_style=Style(font=Font(name="A" * 32)))
    with XlsbWriter(tmp_path / "long-format.xlsb") as writer:
        writer.add_sheet("Data")
        with pytest.raises(ValueError, match="255 characters"):
            writer.write_sheet([["x"]], header_style=Style(number_format="0" * 256))


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


def _xlsb_sheet_payloads(path, record_id, part="xl/worksheets/sheet1.bin"):
    with zipfile.ZipFile(path) as archive:
        blob = archive.read(part)
    records = list(iter_records(blob))
    assert records[-1].data_end == len(blob), "sheet records must tile the buffer"
    return [blob[record.data_start:record.data_end] for record in records if record.record_id == record_id]


def test_xlsb_sheet_args_are_encoded_in_worksheet_records(tmp_path):
    path = tmp_path / "sheet_args.xlsb"
    with XlsbWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(
            [["A", "B", "C"], [1, 2, 3], [4, 5, 6]],
            merges=["A1:B1", "C2:C3"],
            column_widths={"A": 20, 2: 30.5},
            hidden_columns=["B"],
            row_heights={1: 30},
            hidden_rows=[2],
            freeze_panes="B2",
        )

    pane = _xlsb_sheet_payloads(path, 0x0097)[0]
    assert struct.unpack("<ddIIIB", pane) == (1.0, 1.0, 1, 1, 2, 0x03)
    selection = _xlsb_sheet_payloads(path, 0x0098)[0]
    assert struct.unpack("<IIII", selection[:16]) == (0, 1, 1, 0)
    assert struct.unpack("<IIII", selection[20:36]) == (1, 1, 1, 1)

    columns = _xlsb_sheet_payloads(path, 0x003C)
    assert [struct.unpack("<IIIIH", payload) for payload in columns] == [
        (0, 0, 20 * 256, 0, 0x02),
        # Automatic widths never go below Excel's default 8.43 characters.
        (1, 1, round((8.43 + 5 / 7) * 256), 0, 0x03),
        (2, 2, round(30.5 * 256), 0, 0x02),
    ]

    rows = _xlsb_sheet_payloads(path, 0x0000)
    assert len(rows) == 3
    assert struct.unpack("<H", rows[0][8:10])[0] == 300
    assert (rows[0][11], rows[0][12]) == (0x00, 0x00)
    assert struct.unpack("<H", rows[1][8:10])[0] == 600
    assert rows[1][11] == 0x20  # fUnsynced: custom height
    assert struct.unpack("<H", rows[2][8:10])[0] == 300
    assert rows[2][11] == 0x10  # fDyZero: hidden

    begins = _xlsb_sheet_payloads(path, 0x00B1)
    assert [struct.unpack("<I", payload)[0] for payload in begins] == [2]
    merges = _xlsb_sheet_payloads(path, 0x00B0)
    assert [struct.unpack("<IIII", payload) for payload in merges] == [
        (0, 0, 0, 1),
        (1, 2, 2, 2),
    ]
    assert all(len(payload) == 0 for payload in _xlsb_sheet_payloads(path, 0x00B2))

    with ExcelReader(str(path)) as reader:
        assert list(reader.get_rows("Data")) == [["A", "B", "C"], [1, 2, 3], [4, 5, 6]]


def test_xlsb_freeze_and_filter_defaults(tmp_path):
    path = tmp_path / "defaults.xlsb"
    with XlsbWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([["A", "B"], [1, 2]])

    pane = _xlsb_sheet_payloads(path, 0x0097)[0]
    assert pane == XlsbWriter._stickHeaderA1bytes[3:]
    assert _xlsb_sheet_payloads(path, 0x00A1)  # autofilter present by default
    columns = _xlsb_sheet_payloads(path, 0x003C)
    assert [struct.unpack("<IIIIH", payload) for payload in columns] == [
        (0, 0, round((8.43 + 5 / 7) * 256), 0, 0x02),
        (1, 1, round((8.43 + 5 / 7) * 256), 0, 0x02),
    ]

    path = tmp_path / "plain.xlsb"
    with XlsbWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([["A"], [1]], freeze_panes="", auto_filter=False)

    pane = _xlsb_sheet_payloads(path, 0x0097)[0]
    assert struct.unpack("<ddIIIB", pane) == (0.0, 0.0, 0, 0, 3, 0x00)
    selection = _xlsb_sheet_payloads(path, 0x0098)[0]
    assert struct.unpack("<IIII", selection[:16]) == (3, 0, 0, 0)
    assert not _xlsb_sheet_payloads(path, 0x00A1)
    assert not _xlsb_sheet_payloads(path, 0x00A2)

    with ExcelReader(str(path)) as reader:
        assert list(reader.get_rows("Data")) == [["A"], [1]]


def test_xlsb_sheet_args_reject_invalid_input(tmp_path):
    with XlsbWriter(tmp_path / "bad.xlsb") as writer:
        writer.add_sheet("Data")
        with pytest.raises(ValueError, match="overlapping merged ranges"):
            writer.write_sheet([["x"]], merges=["A1:B2", "B2:C3"])
    with XlsbWriter(tmp_path / "bad.xlsb") as writer:
        writer.add_sheet("Data")
        with pytest.raises(ValueError, match="freeze_panes must be a single cell"):
            writer.write_sheet([["x"]], freeze_panes="A1:B2")
    with XlsbWriter(tmp_path / "bad.xlsb") as writer:
        writer.add_sheet("Data")
        with pytest.raises(TypeError, match="column_widths must be a mapping"):
            writer.write_sheet([["x"]], column_widths=[("A", 20)])
    with XlsbWriter(tmp_path / "bad.xlsb") as writer:
        writer.add_sheet("Data")
        with pytest.raises(ValueError, match="column_widths width must be 0.5..255"):
            writer.write_sheet([["x"]], column_widths={"A": 500})
    with XlsbWriter(tmp_path / "bad.xlsb") as writer:
        writer.add_sheet("Data")
        with pytest.raises(TypeError, match="hidden_columns must be a collection"):
            writer.write_sheet([["x"]], hidden_columns="B")
    with XlsbWriter(tmp_path / "bad.xlsb") as writer:
        writer.add_sheet("Data")
        with pytest.raises(TypeError, match="row_heights keys must be 0-based row ints"):
            writer.write_sheet([["x"]], row_heights={"1": 20})
    with XlsbWriter(tmp_path / "bad.xlsb") as writer:
        writer.add_sheet("Data")
        with pytest.raises(ValueError, match="row height must be > 0"):
            writer.write_sheet([["x"]], row_heights={0: 0})
    with XlsbWriter(tmp_path / "bad.xlsb") as writer:
        writer.add_sheet("Data")
        with pytest.raises(TypeError, match="hidden_rows entries must be 0-based row ints"):
            writer.write_sheet([["x"]], hidden_rows=["1"])
    with XlsbWriter(tmp_path / "bad.xlsb") as writer:
        writer.add_sheet("Data")
        with pytest.raises(ValueError, match="hidden column out of range"):
            writer.write_sheet([["x"]], hidden_columns=[16384])


def _ws_view_flags(path, sheet=1):
    """First flag byte of BrtBeginWsView (record 0x89) in an XLSB sheet."""
    with zipfile.ZipFile(path) as archive:
        data = archive.read(f"xl/worksheets/sheet{sheet}.bin")
    view = next(record for record in iter_records(data) if record.record_id == 0x89)
    return data[view.data_start]


@pytest.mark.parametrize("show", [True, False])
def test_show_grid_lines_xlsx(tmp_path, show):
    path = tmp_path / "grid.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("First")
        writer.write_sheet([["a"]], show_grid_lines=show)
        writer.add_sheet("Second")
        writer.write_sheet([["b"]], freeze_panes="B2", show_grid_lines=show)
    with zipfile.ZipFile(path) as archive:
        for part in ("xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml"):
            xml = archive.read(part).decode()
            assert ('showGridLines="0"' in xml) is (not show)
    openpyxl = pytest.importorskip("openpyxl")
    workbook = openpyxl.load_workbook(path)
    for sheet in workbook.worksheets:
        # openpyxl reads a missing attribute as None, Excel's default (shown).
        assert (sheet.sheet_view.showGridLines is not False) is show
    from xlspy import load_workbook
    loaded = load_workbook(path)
    try:
        assert loaded.active.sheet_view.showGridLines is show
    finally:
        loaded.close()


@pytest.mark.parametrize("show", [True, False])
def test_show_grid_lines_xlsb(tmp_path, show):
    path = tmp_path / "grid.xlsb"
    with XlsbWriter(path) as writer:
        writer.add_sheet("First")
        writer.write_sheet([["a"]], show_grid_lines=show)
        writer.add_sheet("Second")
        writer.write_sheet([["b"]], freeze_panes="B2", show_grid_lines=show)
    grid = 0x04 if show else 0
    # The first sheet is also selected (0x40); only fDspGrid changes.
    assert _ws_view_flags(path, 1) == 0xD8 | grid
    assert _ws_view_flags(path, 2) == 0x98 | grid


@pytest.mark.parametrize("ext", ["xlsx", "xlsb"])
def test_show_grid_lines_from_workbook_api(tmp_path, ext):
    from xlspy import Workbook
    path = tmp_path / f"grid.{ext}"
    workbook = Workbook()
    workbook.active.append(["x"])
    workbook.active.sheet_view.showGridLines = False
    workbook.create_sheet("Grid").append(["y"])
    workbook.save(path)
    if ext == "xlsb":
        assert not _ws_view_flags(path, 1) & 0x04
        assert _ws_view_flags(path, 2) & 0x04
    else:
        with zipfile.ZipFile(path) as archive:
            assert 'showGridLines="0"' in archive.read("xl/worksheets/sheet1.xml").decode()
            assert "showGridLines" not in archive.read("xl/worksheets/sheet2.xml").decode()


@pytest.mark.parametrize("writer_cls,ext", [(XlsxWriter, "xlsx"), (XlsbWriter, "xlsb")])
def test_show_grid_lines_must_be_bool(tmp_path, writer_cls, ext):
    with writer_cls(tmp_path / f"bad.{ext}") as writer:
        writer.add_sheet("Data")
        with pytest.raises(TypeError, match="show_grid_lines must be a bool"):
            writer.write_sheet([["x"]], show_grid_lines=0)


def _cell_styles_via_openpyxl(path):
    openpyxl = pytest.importorskip("openpyxl")
    return openpyxl.load_workbook(path).active


@pytest.mark.parametrize("use_c", [True, False])
def test_value_style_cells_xlsx(tmp_path, monkeypatch, use_c):
    import xlspy.xlsx_writer
    if not use_c:
        monkeypatch.setattr(xlspy.xlsx_writer, "accel_encode_xlsx_row", lambda *a, **k: None)
    bold = Style(font=Font(bold=True))
    red = Style(fill="FF0000", number_format="0.0")
    rows = [["h1", "h2", "h3"]]
    rows += [[(r, bold), (r / 2, red), (f"t{r}", "@")] for r in range(1, 30)]
    rows.append([("plain", Style()), None, (datetime.date(2024, 5, 1), red)])
    path = tmp_path / "cells.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(rows, styles={"A2:C10": Style(italic=True)})
    sheet = _cell_styles_via_openpyxl(path)
    assert sheet["A2"].value == 1 and sheet["A2"].font.b and sheet["A2"].font.i
    assert sheet["A20"].font.b and not sheet["A20"].font.i
    assert sheet["B3"].value == 1.0 and sheet["B3"].number_format == "0.0"
    assert sheet["B3"].fill.fgColor.rgb == "FFFF0000" and sheet["B3"].font.i
    assert sheet["C25"].number_format == "@" and sheet["C25"].value == "t24"
    assert sheet["A31"].value == "plain"
    # A date under the cell's "0.0" format is a plain serial to openpyxl.
    assert sheet["C31"].number_format == "0.0" and sheet["C31"].value == 45413


@pytest.mark.parametrize("use_c", [True, False])
def test_value_style_cells_xlsb(tmp_path, monkeypatch, use_c):
    import xlspy.writer
    if not use_c:
        monkeypatch.setattr(xlspy.writer, "encode_xlsb_row", lambda *a, **k: None)
    bold = Style(bold=True)
    rows = [["h"], *[[(r, bold)] for r in range(1, 20)], [("x", Style(number_format="0.00"))]]
    path = tmp_path / "cells.xlsb"
    with XlsbWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(rows, styles={"A2:A5": Style(italic=True)})
    with ExcelReader(str(path)) as reader:
        values = [row[0] for row in reader.get_rows("Data")]
    assert values == ["h", *range(1, 20), "x"]
    with zipfile.ZipFile(path) as archive:
        sheet_bin = archive.read("xl/worksheets/sheet1.bin")
    # Every (value, Style) cell carries a custom xf: none of A2..A21 uses xf 0.
    cells = [record for record in iter_records(sheet_bin) if record.record_id in (0x02, 0x05, 0x07)]
    xfs = [struct.unpack_from("<I", sheet_bin, record.data_start + 4)[0] for record in cells][1:]
    assert len(xfs) == 20 and 0 not in xfs
    assert len(set(xfs[:4])) == 1 and len(set(xfs[4:19])) == 1 and xfs[0] != xfs[4]


def test_value_style_cell_column_widths_use_the_value(tmp_path):
    path = tmp_path / "widths.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([[("a long cell text", Style(bold=True)), "a long cell text"]])
    with zipfile.ZipFile(path) as archive:
        xml = archive.read("xl/worksheets/sheet1.xml").decode()
    widths = re.findall(r'<col min="\d+" max="\d+" width="([0-9.]+)"', xml)
    assert len(widths) == 2 and widths[0] == widths[1]


def test_value_style_rejected_rich_style_in_xlsb(tmp_path):
    with pytest.raises(NotImplementedError):
        with XlsbWriter(tmp_path / "bad.xlsb") as writer:
            writer.add_sheet("Data")
            writer.write_sheet([[(1, Style(border=Border(start=Side(style="thin"))))]])


@pytest.mark.parametrize("writer_cls,ext", [(XlsxWriter, "xlsx"), (XlsbWriter, "xlsb")])
def test_open_style_ranges(tmp_path, writer_cls, ext):
    from xlspy.styles import normalize_style_ranges
    ranges = normalize_style_ranges({"B:C": Style(bold=True), "A3:A": Style(italic=True),
                                     "1:1": Style(fill="00FF00")})
    assert [r[:4] for r in ranges] == [(1, 0, 2, 1048575), (0, 2, 0, 1048575), (0, 0, 16383, 0)]
    path = tmp_path / f"open.{ext}"
    with writer_cls(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([[r, r, r, r] for r in range(50)],
                           styles={"B:C": Style(bold=True), "A3:A": Style(italic=True)})
    with zipfile.ZipFile(path) as archive:
        sheet = archive.read(f"xl/worksheets/sheet1.{'xml' if ext == 'xlsx' else 'bin'}")
    # Only the written rows are styled: an open range adds no empty rows.
    if ext == "xlsx":
        assert sheet.count(b"<row ") == 50
    if ext == "xlsx":
        loaded = _cell_styles_via_openpyxl(path)
        assert loaded["B50"].font.b and loaded["C1"].font.b and not loaded["D1"].font.b
        assert loaded["A3"].font.i and loaded["A50"].font.i and not loaded["A2"].font.i


@pytest.mark.parametrize("reference", ["D:B", "5:2", "B2:A", "XFE:XFE", "0:1", "B:"])
def test_open_style_ranges_reject_bad_references(reference):
    from xlspy.styles import parse_style_range
    with pytest.raises(ValueError):
        parse_style_range(reference)
