"""Scenarios for the fixes from the review of the outline/table/validation work.

Each section names the problem it guards against. XLSX parts are checked as
XML (and opened with openpyxl where that adds something); XLSB parts are
checked record by record.
"""

import re
import struct
import zipfile
import zlib
from decimal import Decimal

import pytest

from xlspy import (
    Chart,
    ChartSeries,
    DataValidation,
    ExcelReader,
    ExcelTable,
    Formula,
    Group,
    Hyperlink,
    IconSetRule,
    Image,
    XlsbWriter,
    XlsxWriter,
)
from xlspy.biff_utils import iter_records
from xlspy.outline import build_outline

openpyxl = pytest.importorskip("openpyxl")

DATA = [["h"], [1], [2], [3]]


def _png_bytes(width=2, height=2):
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


def _png_file(tmp_path):
    path = tmp_path / "pixel.png"
    path.write_bytes(_png_bytes())
    return path


def _zip_text(path, name):
    with zipfile.ZipFile(path) as archive:
        return archive.read(name).decode("utf-8")


def _xlsx_rows(path, part="xl/worksheets/sheet1.xml"):
    """Map 1-based row number to the attribute string of its ``<row>``."""
    xml = _zip_text(path, part)
    return {int(m.group(1)): m.group(2) for m in re.finditer(r'<row r="(\d+)"([^>]*?)/?>', xml)}


def _xlsb_row_flags(path, part="xl/worksheets/sheet1.bin"):
    """Map 0-based row index to the BrtRowHdr flags byte."""
    with zipfile.ZipFile(path) as archive:
        blob = archive.read(part)
    records = list(iter_records(blob))
    assert records[-1].data_end == len(blob), "sheet records must tile the buffer"
    return {
        struct.unpack_from("<I", blob, r.data_start)[0]: blob[r.data_start + 11]
        for r in records if r.record_id == 0x0000
    }


def _check_package_consistency(path):
    """Every content-type override names a part, and every part a rels file targets exists.

    In particular a theme part must be referenced from the workbook rels.
    """
    with zipfile.ZipFile(path) as archive:
        names = set(archive.namelist())
        content_types = archive.read("[Content_Types].xml").decode()
        for part in re.findall(r'PartName="/([^"]+)"', content_types):
            assert part in names, f"content type for missing part {part}"
        targeted = set()
        for rels_name in (n for n in names if n.endswith(".rels")):
            base = rels_name.replace("_rels/", "").removesuffix(".rels")
            folder = base.rpartition("/")[0]
            for target, mode in re.findall(
                r'Target="([^"]+)"(\s+TargetMode="External")?', archive.read(rels_name).decode()
            ):
                if mode:
                    continue
                parts = (folder.split("/") if folder else []) + target.split("/")
                resolved = []
                for piece in parts:
                    if piece == "..":
                        resolved.pop()
                    elif piece:
                        resolved.append(piece)
                full = "/".join(resolved)
                assert full in names, f"{rels_name} targets missing part {full}"
                targeted.add(full)
        for name in names:
            if name.startswith("xl/theme/"):
                assert name in targeted, f"{name} is in the package but no rels file targets it"


# ------------------------------------------------- 1. outline rows past the data


def test_indices_from_lists_grouped_and_summary_rows_in_order():
    outline = build_outline([Group(1, 3, True), Group(6, 8), Group(7, 7, True)], "row")

    assert list(outline.indices_from(0)) == [1, 2, 3, 4, 6, 7, 8]
    assert list(outline.indices_from(4)) == [4, 6, 7, 8]
    assert list(outline.indices_from(9)) == []


def test_group_on_the_last_row_has_no_summary_past_the_sheet():
    outline = build_outline([Group(1048570, 1048575, True)], "row")

    assert not outline.collapsed_at(1048576)
    assert list(outline.indices_from(1048574)) == [1048574, 1048575]
    columns = build_outline([Group("XFA", "XFD", True)], "column")
    assert 16384 not in columns.indices()


def test_xlsx_collapsed_group_ending_on_last_data_row_writes_summary_row(tmp_path):
    path = tmp_path / "summary.xlsx"
    with XlsxWriter(path) as writer:
        writer.write_sheet(DATA, row_groups=[Group(1, 3, collapsed=True)])

    rows = _xlsx_rows(path)
    for r in (2, 3, 4):
        assert 'outlineLevel="1"' in rows[r] and 'hidden="1"' in rows[r]
    assert 'collapsed="1"' in rows[5]
    assert "outlineLevel" not in rows[5] and "hidden" not in rows[5]
    sheet = openpyxl.load_workbook(path).active
    assert sheet.row_dimensions[3].outline_level == 1
    assert sheet.row_dimensions[3].hidden


def test_xlsx_group_reaching_past_the_data_keeps_its_level(tmp_path):
    path = tmp_path / "past.xlsx"
    with XlsxWriter(path) as writer:
        writer.write_sheet(
            DATA,
            row_groups=[Group(2, 8), Group(5, 6, collapsed=True)],
            row_heights={6: 30},
        )

    rows = _xlsx_rows(path)
    # Group(2, 8) is open, so it has no summary row: rows stop at r9.
    assert set(rows) == {1, 2, 3, 4, 5, 6, 7, 8, 9}
    assert 'outlineLevel="1"' in rows[5]
    assert 'outlineLevel="2"' in rows[6] and 'hidden="1"' in rows[6]
    assert 'ht="30.0"' in rows[7] and 'outlineLevel="2"' in rows[7]
    assert 'collapsed="1"' in rows[8] and 'outlineLevel="1"' in rows[8]
    assert 'outlineLevel="1"' in rows[9] and 'collapsed="1"' not in rows[9]
    # The dimension still covers only the cells that hold data.
    assert '<dimension ref="A1:A4"/>' in _zip_text(path, "xl/worksheets/sheet1.xml")
    sheet = openpyxl.load_workbook(path).active
    assert sheet.max_row in (4, 9)
    assert [c.value for c in sheet["A"][:4]] == ["h", 1, 2, 3]


def test_xlsx_outline_only_sheet_without_data(tmp_path):
    path = tmp_path / "empty.xlsx"
    with XlsxWriter(path) as writer:
        writer.write_sheet([], row_groups=[Group(0, 1, collapsed=True)])

    rows = _xlsx_rows(path)
    assert sorted(rows) == [1, 2, 3]
    assert 'collapsed="1"' in rows[3]
    openpyxl.load_workbook(path)


def test_xlsb_collapsed_group_ending_on_last_data_row_writes_summary_row(tmp_path):
    path = tmp_path / "summary.xlsb"
    with XlsbWriter(path) as writer:
        writer.write_sheet(DATA, row_groups=[Group(1, 3, collapsed=True)])

    flags = _xlsb_row_flags(path)
    for row in (1, 2, 3):
        assert flags[row] & 0x07 == 1 and flags[row] & 0x10
    assert flags[4] == 0x08  # collapsed, level 0, visible

    with ExcelReader(str(path)) as reader:
        rows = [row for row in reader.get_rows("Sheet1") if any(cell not in (None, "") for cell in row)]
    assert [row[0] for row in rows] == ["h", 1, 2, 3]


def test_xlsb_group_reaching_past_the_data_keeps_its_level(tmp_path):
    path = tmp_path / "past.xlsb"
    with XlsbWriter(path) as writer:
        writer.write_sheet(DATA, row_groups=[Group(2, 8), Group(5, 6, collapsed=True)])

    flags = _xlsb_row_flags(path)
    assert sorted(flags) == list(range(9))
    assert flags[4] & 0x07 == 1
    assert flags[5] & 0x07 == 2 and flags[5] & 0x10
    assert flags[7] & 0x08 and flags[7] & 0x07 == 1
    assert flags[8] == 0x01


@pytest.mark.parametrize("writer_cls,suffix", [(XlsxWriter, "xlsx"), (XlsbWriter, "xlsb")])
def test_large_group_past_data_is_streamed(tmp_path, writer_cls, suffix):
    path = tmp_path / f"big.{suffix}"
    with writer_cls(path) as writer:
        writer.write_sheet(DATA, row_groups=[Group(1, 20000, collapsed=True)])
    assert path.stat().st_size > 0
    if suffix == "xlsx":
        rows = _xlsx_rows(path)
        assert len(rows) == 20002 and 'collapsed="1"' in rows[20002]
    else:
        flags = _xlsb_row_flags(path)
        assert len(flags) == 20002 and flags[20001] == 0x08


# ----------------------------------------------- 2. non-finite formula caches


@pytest.mark.parametrize(
    "value",
    [float("inf"), float("-inf"), float("nan"), Decimal("Infinity"), Decimal("-Infinity"), Decimal("NaN")],
)
def test_formula_rejects_non_finite_cached_values(value):
    with pytest.raises(ValueError, match="finite"):
        Formula("1/0", value=value)


def test_formula_keeps_finite_cached_values(tmp_path):
    path = tmp_path / "finite.xlsx"
    with XlsxWriter(path) as writer:
        writer.write_sheet([[Formula("1/3", value=Decimal("0.25")), Formula("2*2", value=1e308)]])

    xml = _zip_text(path, "xl/worksheets/sheet1.xml")
    assert "<v>0.25</v>" in xml
    assert "inf" not in xml.lower()


# ------------------------------------------------- 3. theme part and its rels


def test_image_only_workbook_references_its_theme(tmp_path):
    path = tmp_path / "image.xlsx"
    with XlsxWriter(path) as writer:
        writer.write_sheet([["x"]], images=[Image(_png_file(tmp_path), anchor="B2")])

    rels = _zip_text(path, "xl/_rels/workbook.xml.rels")
    assert 'Target="theme/theme1.xml"' in rels
    _check_package_consistency(path)
    openpyxl.load_workbook(path)


def test_chart_and_image_on_different_sheets_share_one_theme(tmp_path):
    path = tmp_path / "mixed.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Pictures")
        writer.write_sheet([["x"]], images=[Image(_png_file(tmp_path), anchor="B2")])
        writer.add_sheet("Chart")
        writer.write_sheet(
            [["k", "v"], ["a", 1], ["b", 2]],
            charts=[Chart("bar", [ChartSeries("B2:B3", categories="A2:A3")])],
        )

    assert _zip_text(path, "xl/_rels/workbook.xml.rels").count("theme/theme1.xml") == 1
    _check_package_consistency(path)


def test_workbook_without_drawings_has_no_theme(tmp_path):
    path = tmp_path / "plain.xlsx"
    with XlsxWriter(path) as writer:
        writer.write_sheet(DATA)

    with zipfile.ZipFile(path) as archive:
        assert "xl/theme/theme1.xml" not in archive.namelist()
    assert "theme" not in _zip_text(path, "xl/_rels/workbook.xml.rels")
    _check_package_consistency(path)


# --------------------------------------------- 4. dynamic list validations


@pytest.mark.parametrize(
    "formula",
    [
        "OFFSET($B$1,0,0,5,1)",
        "INDEX($B:$B,1):INDEX($B:$B,COUNTA($B:$B))",
        'INDIRECT("Lists!A1:A5")',
        "'Lists, 2024'!$A$1:$A$5",
        "$B$1:$B$5",
        "MyNamedRange",
    ],
)
def test_list_validation_accepts_ranges_and_dynamic_formulas(formula):
    rule = DataValidation("list", "A1:A5", formula1=formula)
    assert rule.formula1 == formula


@pytest.mark.parametrize("formula", ["Yes,No", "$B$1:$B$2,$C$1", "OFFSET($B$1,0,0),X"])
def test_list_validation_still_rejects_unquoted_literal_lists(formula):
    with pytest.raises(ValueError, match="double quotes"):
        DataValidation("list", "A1:A5", formula1=formula)


def test_dynamic_list_validation_round_trips(tmp_path):
    path = tmp_path / "dynamic.xlsx"
    with XlsxWriter(path) as writer:
        writer.write_sheet(
            [["pick", "options"], [None, "a"], [None, "b"]],
            data_validations=[DataValidation("list", "A2:A3", formula1="OFFSET($B$2,0,0,COUNTA($B:$B)-1,1)")],
        )

    sheet = openpyxl.load_workbook(path).active
    (rule,) = sheet.data_validations.dataValidation
    assert rule.type == "list"
    assert rule.formula1 == "OFFSET($B$2,0,0,COUNTA($B:$B)-1,1)"


# ---------------------------------------- 5. internal hyperlink sheet names


@pytest.mark.parametrize(
    "sheet_name,location",
    [
        ("it's", "#'it''s'!B2"),
        ("O'Brien Q1", "#'o''brien q1'!A1"),
        ("Sheet1", "#sheet1!A1"),
        ("Sales Data", "#'SALES DATA'!B2"),
        ("Data", "#Data!A1:B3"),
    ],
)
def test_internal_links_resolve_quoted_and_differently_cased_names(tmp_path, sheet_name, location):
    path = tmp_path / "links.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet(sheet_name)
        writer.write_sheet([[Hyperlink(location, text="go")]])

    assert "<hyperlink " in _zip_text(path, "xl/worksheets/sheet1.xml")


def test_internal_link_to_a_later_sheet_is_accepted(tmp_path):
    path = tmp_path / "later.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Index")
        writer.write_sheet([[Hyperlink("#'Details 2'!A1", text="details")]])
        writer.add_sheet("Details 2")
        writer.write_sheet([["here"]])
    openpyxl.load_workbook(path)


def test_internal_link_with_apostrophe_but_wrong_name_still_fails(tmp_path):
    # Only the enclosing quotes are removed: "'its'" names "its", not "it's".
    with pytest.raises(ValueError, match="unknown sheet 'its'"):
        with XlsxWriter(tmp_path / "bad.xlsx") as writer:
            writer.add_sheet("it's")
            writer.write_sheet([[Hyperlink("#'its'!A1", text="go")]])


# ---------------------------------- 6. a rejected write_sheet changes nothing


def test_xlsb_rejected_sheet_leaves_no_charts_tables_or_priorities(tmp_path):
    path = tmp_path / "retry.xlsb"
    data = [["k", "v"], ["a", 1], ["b", 2]]
    chart = Chart("bar", [ChartSeries("B2:B3", categories="A2:A3")])
    with XlsbWriter(path) as writer:
        writer.add_sheet("Data")
        with pytest.raises(ValueError, match="duplicate table display name"):
            writer.write_sheet(
                data,
                charts=[chart],
                tables=[ExcelTable("Sales", "A1:B3"), ExcelTable("Sales", "D1:E3")],
            )
        cf = [IconSetRule("B2:B3", "3TrafficLights1", [0, 33, 67], priority=5)]
        with pytest.raises(ValueError, match="duplicate conditional formatting priority 5"):
            writer.write_sheet(data, conditional_formats=cf + cf)
        # The same name, chart-free and with priority 5, is accepted afterwards.
        writer.write_sheet(data, tables=[ExcelTable("Sales", "A1:B3")], conditional_formats=cf)

    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
    assert not any(name.startswith(("xl/charts/", "xl/drawings/")) for name in names)
    assert "xl/tables/table1.bin" in names and "xl/tables/table2.bin" not in names
    with ExcelReader(str(path)) as reader:
        assert [row[:2] for row in reader.get_rows("Data")][:3] == data


def test_xlsx_rejected_sheet_does_not_reserve_table_names_or_priorities(tmp_path):
    path = tmp_path / "retry.xlsx"
    data = [["k", "v"], ["a", 1], ["b", 2]]
    cf = [IconSetRule("B2:B3", "3TrafficLights1", [0, 33, 67], priority=3)]
    with XlsxWriter(path) as writer:
        with pytest.raises(TypeError, match="images must be Image instances"):
            writer.write_sheet(data, tables=[ExcelTable("Sales", "A1:B3")], images=["logo.png"])
        with pytest.raises(ValueError, match="duplicate conditional formatting priority 3"):
            writer.write_sheet(data, conditional_formats=cf + cf)
        writer.write_sheet(data, tables=[ExcelTable("Sales", "A1:B3")], conditional_formats=cf)

    sheet = openpyxl.load_workbook(path).active
    assert list(sheet.tables) == ["Sales"]


# ------------------------------------------ 7. table headers checked early


def _counting_rows(rows, consumed):
    for row in rows:
        consumed.append(row)
        yield row


@pytest.mark.parametrize("writer_cls,suffix", [(XlsxWriter, "xlsx"), (XlsbWriter, "xlsb")])
def test_bad_table_header_fails_before_the_rest_is_read(tmp_path, writer_cls, suffix):
    consumed = []
    rows = _counting_rows([["Name", 7]] + [[f"r{i}", i] for i in range(10000)], consumed)
    # The failed sheet aborts the writer, so leaving the block raises too.
    with pytest.raises(RuntimeError, match="previous write failed"):
        with writer_cls(tmp_path / f"bad.{suffix}") as writer:
            with pytest.raises(ValueError, match="header cell B1 must hold text"):
                writer.write_sheet(rows, tables=[ExcelTable("Tbl", "A1:B10001")])
    assert len(consumed) == 1


@pytest.mark.parametrize("writer_cls,suffix", [(XlsxWriter, "xlsx"), (XlsbWriter, "xlsb")])
def test_mismatched_table_column_name_fails_at_header_row(tmp_path, writer_cls, suffix):
    consumed = []
    rows = _counting_rows([["intro"], ["Name", "Qty"]] + [["x", 1]] * 500, consumed)
    with pytest.raises(RuntimeError, match="previous write failed"):
        with writer_cls(tmp_path / f"bad.{suffix}") as writer:
            with pytest.raises(ValueError, match="does not match"):
                writer.write_sheet(rows, tables=[ExcelTable("Tbl", "A2:B502", columns=["Name", "Amount"])])
    assert len(consumed) == 2


def test_hyperlink_header_cell_names_the_table_column(tmp_path):
    path = tmp_path / "link_header.xlsx"
    with XlsxWriter(path) as writer:
        writer.write_sheet(
            [[Hyperlink("https://example.com", text="Site"), "Visits"], ["a", 1], ["b", 2]],
            tables=[ExcelTable("Traffic", "A1:B3")],
        )

    table = openpyxl.load_workbook(path).active.tables["Traffic"]
    assert [column.name for column in table.tableColumns] == ["Site", "Visits"]


def test_formula_header_cell_is_still_rejected(tmp_path):
    with pytest.raises(RuntimeError, match="previous write failed"):
        with XlsxWriter(tmp_path / "f.xlsx") as writer:
            with pytest.raises(ValueError, match="must hold text"):
                writer.write_sheet([[Formula('"A"'), "B"], [1, 2]], tables=[ExcelTable("Tbl", "A1:B2")])


# ---------------------------------- 8/9. one rule set for hidden rows/columns


@pytest.mark.parametrize("writer_cls,suffix", [(XlsxWriter, "xlsx"), (XlsbWriter, "xlsb")])
@pytest.mark.parametrize("bad", [-1, 16384, "XFE"])
def test_hidden_columns_out_of_range_rejected_by_both_writers(tmp_path, writer_cls, suffix, bad):
    with writer_cls(tmp_path / f"bad.{suffix}") as writer:
        with pytest.raises(ValueError, match="out of range"):
            writer.write_sheet([["x"]], hidden_columns=[bad])
        writer.write_sheet([["x"]], hidden_columns=["XFD", 0])


@pytest.mark.parametrize("writer_cls,suffix", [(XlsxWriter, "xlsx"), (XlsbWriter, "xlsb")])
@pytest.mark.parametrize("bad,error", [([-1], TypeError), (["2"], TypeError), ([True], TypeError), (5, TypeError)])
def test_hidden_rows_rules_match_between_writers(tmp_path, writer_cls, suffix, bad, error):
    with writer_cls(tmp_path / f"bad.{suffix}") as writer:
        with pytest.raises(error, match="hidden_rows"):
            writer.write_sheet([["x"]], hidden_rows=bad)


@pytest.mark.parametrize("writer_cls,suffix", [(XlsxWriter, "xlsx"), (XlsbWriter, "xlsb")])
def test_hidden_rows_and_columns_written_by_both_writers(tmp_path, writer_cls, suffix):
    path = tmp_path / f"hidden.{suffix}"
    with writer_cls(path) as writer:
        writer.write_sheet([["a", "b", "c"], [1, 2, 3], [4, 5, 6]], hidden_rows={1}, hidden_columns=["B"])
    if suffix == "xlsx":
        sheet = openpyxl.load_workbook(path).active
        assert sheet.row_dimensions[2].hidden
        assert sheet.column_dimensions["B"].hidden
    else:
        assert _xlsb_row_flags(path)[1] & 0x10


# ----------------------------- 10. tables covering the data skip autoFilter


def test_table_over_the_whole_data_gets_no_sheet_autofilter(tmp_path):
    path = tmp_path / "table.xlsx"
    with XlsxWriter(path) as writer:
        writer.write_sheet(
            [["k", "v"], ["a", 1], ["b", 2]],
            tables=[ExcelTable("Full", "A1:B3")],
            auto_filter=True,
        )

    assert "<autoFilter" not in _zip_text(path, "xl/worksheets/sheet1.xml")
    assert "<autoFilter" in _zip_text(path, "xl/tables/table1.xml")

