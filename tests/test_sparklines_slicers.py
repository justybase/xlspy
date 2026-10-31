"""Sparklines and table slicers in XLSX output, and their XLSB guard."""

import os
import re
import xml.etree.ElementTree as ET
import zipfile

import pytest

from xlspy import ExcelTable, XlsbWriter, XlsxWriter
from xlspy.slicers import Slicer, normalize_slicers
from xlspy.sparklines import SparklineGroup, _expand, normalize_sparklines, sparklines_ext

X14 = "{http://schemas.microsoft.com/office/spreadsheetml/2009/9/main}"
EXT_SPARKLINES = "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}"

ROWS = [
    ["Region", "Q1", "Q2", "Q3", "Trend"],
    ["North", 10, 12, 9, None],
    ["South", 7, 11, 13, None],
    ["East", 5, 6, 4, None],
]
COLUMNS = ["Region", "Q1", "Q2", "Q3", "Trend"]


def _write_sample(path, *, sparklines=None, slicers=None):
    with XlsxWriter(path) as writer:
        writer.write_sheet(
            ROWS,
            tables=[ExcelTable("Sales", "A1:E4", columns=COLUMNS)],
            sparklines=sparklines,
            slicers=slicers,
        )
    return path


def _parts(path):
    with zipfile.ZipFile(path) as archive:
        return {name: archive.read(name).decode("utf-8") for name in archive.namelist()}


def _parse_all(parts):
    for name, text in parts.items():
        if name.endswith((".xml", ".rels")):
            ET.fromstring(text.encode("utf-8"))  # raises if not well-formed


# --- sparklines ------------------------------------------------------------

def test_single_cell_location_draws_one_sparkline_from_all_data():
    expanded = _expand(SparklineGroup(data="B2:D2", location="E2"), "Sheet1")
    assert expanded == [("E2", "'Sheet1'!B2:D2")]


def test_column_of_locations_takes_one_data_row_each():
    group = SparklineGroup(data="B2:D4", location="E2:E4")
    assert _expand(group, "Sales") == [
        ("E2", "'Sales'!B2:D2"),
        ("E3", "'Sales'!B3:D3"),
        ("E4", "'Sales'!B4:D4"),
    ]


def test_row_of_locations_takes_one_data_column_each():
    group = SparklineGroup(data="B2:D4", location="B6:D6")
    assert _expand(group, "Sheet1") == [
        ("B6", "'Sheet1'!B2:B4"),
        ("C6", "'Sheet1'!C2:C4"),
        ("D6", "'Sheet1'!D2:D4"),
    ]


def test_data_on_another_sheet_is_qualified_and_quoted_where_needed():
    group = SparklineGroup(data="'Sales data'!B2:D2", location="E2")
    assert _expand(group, "Dashboard") == [("E2", "'Sales data'!B2:D2")]


@pytest.mark.parametrize(
    "kwargs, message",
    [
        ({"data": "B2:D2", "location": "E2:E4"}, "one data row each"),
        ({"data": "B2:B4", "location": "B6:E6"}, "one data column each"),
        ({"data": "B2:D2", "location": "E2:F3"}, "must be one cell"),
        ({"data": "B2:D2", "location": "E2", "kind": "bar"}, "kind"),
        ({"data": "B2:D2", "location": "E2", "color": "green"}, "RGB hex"),
        ({"data": "B2:D2", "location": "E2", "empty_cells": "span"}, "empty_cells"),
        ({"data": "B2:D2", "location": "E2", "line_width": 0}, "line_width"),
        ({"data": "B2:D2", "location": "E2", "axis_min": True}, "axis_min"),
    ],
)
def test_invalid_sparkline_groups_are_rejected(kwargs, message):
    with pytest.raises(ValueError, match=message):
        normalize_sparklines([SparklineGroup(**kwargs)])


def test_sparkline_xml_matches_excel_layout():
    group = SparklineGroup(data="B2:D2", location="E2", kind="win_loss", negative_points=True,
                           markers=True)
    ext = sparklines_ext([group], "Sheet1")
    root = ET.fromstring(
        f'<root xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" '
        f'xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">{ext}</root>'
    )
    ext_el = root.find("ext")
    assert ext_el.get("uri") == EXT_SPARKLINES
    spark_group = ext_el.find(f".//{X14}sparklineGroup")
    assert spark_group.get("type") == "stacked"
    assert spark_group.get("negative") == "1"
    assert spark_group.get("markers") == "1"
    assert spark_group.get("displayEmptyCellsAs") == "gap"
    colors = [child.tag.removeprefix(X14) for child in spark_group if child.tag != f"{X14}sparklines"]
    assert colors == [
        "colorSeries", "colorNegative", "colorAxis", "colorMarkers",
        "colorFirst", "colorLast", "colorHigh", "colorLow",
    ]
    formula = spark_group.find(f".//{{http://schemas.microsoft.com/office/excel/2006/main}}f")
    assert formula.text == "'Sheet1'!B2:D2"


def test_line_sparkline_has_no_type_attribute():
    ext = sparklines_ext([SparklineGroup(data="B2:D2", location="E2")], "Sheet1")
    assert "type=" not in ext


def test_no_sparklines_gives_empty_ext():
    assert sparklines_ext([], "Sheet1") == ""


def test_sheet_written_with_sparklines_has_extension(tmp_path):
    path = _write_sample(tmp_path / "spark.xlsx",
                         sparklines=[SparklineGroup(data="B2:D2", location="E2", kind="column")])
    sheet = _parts(path)["xl/worksheets/sheet1.xml"]
    assert EXT_SPARKLINES in sheet
    assert "<xm:f>'Sheet1'!B2:D2</xm:f>" in sheet
    assert sheet.rstrip().endswith("</worksheet>")
    _parse_all(_parts(path))


# --- slicers -----------------------------------------------------------------

def _tables():
    return [ExcelTable("Sales", "A1:E4", columns=COLUMNS)]


def test_slicer_name_and_cache_derive_from_column():
    slicer = Slicer(table="Sales", column="Sales Region")
    assert slicer.display_name == "Sales Region"
    assert slicer.cache_name == "Slicer_Sales_Region"


def test_slicer_must_name_a_table_on_the_same_sheet():
    with pytest.raises(ValueError, match="not a table written to this sheet"):
        normalize_slicers([Slicer(table="Other", column="Region")], _tables(), set(), set(), set())


def test_slicer_must_name_an_existing_column():
    with pytest.raises(ValueError, match="not a column"):
        normalize_slicers([Slicer(table="Sales", column="Missing")], _tables(), set(), set(), set())


def test_one_slicer_per_column():
    slicers = [Slicer(table="Sales", column="Region"), Slicer(table="Sales", column="Region", name="B")]
    with pytest.raises(ValueError, match="already has a slicer"):
        normalize_slicers(slicers, _tables(), set(), set(), set())


def test_duplicate_slicer_names_are_rejected_across_the_workbook():
    with pytest.raises(ValueError, match="duplicate slicer name"):
        normalize_slicers([Slicer(table="Sales", column="Region")], _tables(), {"region"}, set(), set())


def test_slicer_table_needs_a_header_row():
    table = ExcelTable("Sales", "A1:E4", columns=COLUMNS, header_row=False)
    with pytest.raises(ValueError, match="header row"):
        normalize_slicers([Slicer(table="Sales", column="Region")], [table], set(), set(), set())


def test_slicer_anchor_must_be_one_cell():
    with pytest.raises(ValueError, match="single cell"):
        Slicer(table="Sales", column="Region", anchor="G2:H4")


def test_sheet_written_with_slicer_has_all_parts(tmp_path):
    path = _write_sample(tmp_path / "slicer.xlsx", slicers=[Slicer(table="Sales", column="Region", anchor="G2")])
    parts = _parts(path)
    _parse_all(parts)

    assert "xl/slicers/slicer1.xml" in parts
    assert "xl/slicerCaches/slicerCache1.xml" in parts
    assert 'name="Slicer_Region"' in parts["xl/slicerCaches/slicerCache1.xml"]
    assert 'sourceName="Region"' in parts["xl/slicerCaches/slicerCache1.xml"]
    assert 'tableId="1" column="1"' in parts["xl/slicerCaches/slicerCache1.xml"]
    assert 'cache="Slicer_Region"' in parts["xl/slicers/slicer1.xml"]

    content_types = parts["[Content_Types].xml"]
    assert 'PartName="/xl/slicers/slicer1.xml"' in content_types
    assert "application/vnd.ms-excel.slicer+xml" in content_types
    assert "application/vnd.ms-excel.slicerCache+xml" in content_types

    workbook = parts["xl/workbook.xml"]
    assert '<definedName name="Slicer_Region">#N/A</definedName>' in workbook
    cache_rid = re.search(r'<x14:slicerCache r:id="(rId\d+)"/>', workbook)
    assert cache_rid is not None
    workbook_rels = parts["xl/_rels/workbook.xml.rels"]
    assert f'Id="{cache_rid[1]}"' in workbook_rels
    assert 'Target="slicerCaches/slicerCache1.xml"' in workbook_rels

    sheet_rels = parts["xl/worksheets/_rels/sheet1.xml.rels"]
    slicer_rid = re.search(r'<x14:slicer r:id="(rId\d+)"/>', parts["xl/worksheets/sheet1.xml"])
    assert slicer_rid is not None
    assert f'Id="{slicer_rid[1]}"' in sheet_rels
    assert 'Target="../slicers/slicer1.xml"' in sheet_rels

    drawing = parts["xl/drawings/drawing1.xml"]
    assert 'name="Region"' in drawing
    assert "http://schemas.microsoft.com/office/drawing/2010/slicer" in drawing


def test_two_slicers_get_separate_parts(tmp_path):
    path = _write_sample(tmp_path / "two.xlsx", slicers=[
        Slicer(table="Sales", column="Region", anchor="G2"),
        Slicer(table="Sales", column="Q1", anchor="G20"),
    ])
    parts = _parts(path)
    assert "xl/slicers/slicer2.xml" in parts
    assert "xl/slicerCaches/slicerCache2.xml" in parts
    assert 'tableId="1" column="2"' in parts["xl/slicerCaches/slicerCache2.xml"]
    assert parts["xl/workbook.xml"].count("<definedName name=\"Slicer_") == 2
    _parse_all(parts)


def test_plain_sheet_has_no_slicer_or_sparkline_parts(tmp_path):
    path = _write_sample(tmp_path / "plain.xlsx")
    parts = _parts(path)
    assert not [name for name in parts if "slicer" in name]
    assert "extLst" not in parts["xl/workbook.xml"]
    assert EXT_SPARKLINES not in parts["xl/worksheets/sheet1.xml"]


# --- XLSB guard ----------------------------------------------------------------

def test_xlsb_rejects_sparklines_until_supported(tmp_path):
    with XlsbWriter(tmp_path / "x.xlsb") as writer:
        writer.add_sheet("Sheet1")
        with pytest.raises(NotImplementedError, match="sparklines"):
            writer.write_sheet(ROWS, sparklines=[SparklineGroup(data="B2:D2", location="E2")])


def test_xlsb_rejects_slicers_until_supported(tmp_path):
    with XlsbWriter(tmp_path / "x.xlsb") as writer:
        writer.add_sheet("Sheet1")
        with pytest.raises(NotImplementedError, match="slicers"):
            writer.write_sheet(ROWS, tables=[ExcelTable("Sales", "A1:E4", columns=COLUMNS)],
                               slicers=[Slicer(table="Sales", column="Region")])


# --- Excel COM -----------------------------------------------------------------

@pytest.mark.skipif(os.environ.get("XLSPY_RUN_EXCEL_COM") != "1",
                    reason="Set XLSPY_RUN_EXCEL_COM=1 to run Microsoft Excel COM validation")
def test_excel_reads_sparklines_and_slicer(tmp_path):
    win32 = pytest.importorskip("win32com.client")
    path = _write_sample(
        tmp_path / "com.xlsx",
        sparklines=[SparklineGroup(data="B2:D2", location="E2", kind="column")],
        slicers=[Slicer(table="Sales", column="Region", anchor="G2")],
    )
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    try:
        book = excel.Workbooks.Open(str(path))
        sheet = book.Worksheets(1)
        sheet.Activate()
        groups = sheet.Range("E2").SparklineGroups
        assert groups.Count == 1
        assert groups.Item(1).Type == 2  # xlSparkColumn
        assert book.SlicerCaches.Count == 1
        cache = book.SlicerCaches(1)
        assert cache.Name == "Slicer_Region"
        assert cache.SourceName == "Region"
        assert cache.Slicers(1).Caption == "Region"
        book.Close(SaveChanges=False)
    finally:
        excel.Quit()


# --- loading and editing an existing file -----------------------------------------

def _two_sheet_file(path):
    with XlsxWriter(path) as writer:
        writer.write_sheet([["Note"]])
        writer.add_sheet("Data")
        writer.write_sheet(
            ROWS,
            tables=[ExcelTable("Sales", "A1:E4", columns=COLUMNS)],
            sparklines=[SparklineGroup(data="B2:D2", location="E2")],
            slicers=[Slicer(table="Sales", column="Region", anchor="G2")],
        )
    return path


def test_renaming_a_loaded_sheet_updates_sparkline_formulas(tmp_path):
    from xlspy.workbook import load_workbook

    source = _write_sample(tmp_path / "source.xlsx",
                           sparklines=[SparklineGroup(data="B2:D2", location="E2")])
    workbook = load_workbook(source)
    workbook.worksheets[0].title = "Dashboard"
    out = tmp_path / "renamed.xlsx"
    workbook.save(out)
    workbook.close()
    formulas = re.findall(r"<xm:f>(.*?)</xm:f>", _parts(out)["xl/worksheets/sheet1.xml"])
    assert formulas == ["Dashboard!B2:D2"]


def test_removing_a_sheet_with_slicers_is_refused_before_writing(tmp_path):
    from xlspy.workbook import load_workbook

    source = _two_sheet_file(tmp_path / "two.xlsx")
    workbook = load_workbook(source)
    workbook.remove(workbook.worksheets[1])
    out = tmp_path / "removed.xlsx"
    with pytest.raises(NotImplementedError, match="table slicers"):
        workbook.save(out)
    workbook.close()
    assert not out.exists()


def test_loaded_workbook_round_trip_keeps_sparklines_and_slicers(tmp_path):
    from xlspy.workbook import load_workbook

    source = _write_sample(
        tmp_path / "source.xlsx",
        sparklines=[SparklineGroup(data="B2:D2", location="E2")],
        slicers=[Slicer(table="Sales", column="Region", anchor="G2")],
    )
    workbook = load_workbook(source)
    out = tmp_path / "copy.xlsx"
    workbook.save(out)
    workbook.close()
    parts = _parts(out)
    assert EXT_SPARKLINES in parts["xl/worksheets/sheet1.xml"]
    assert "xl/slicers/slicer1.xml" in parts
    assert "Slicer_Region" in parts["xl/workbook.xml"]
