"""Model edits keep sparklines, slicers and shapes that the model does not read.

A loaded sheet with a table slicer, sparklines, a KPI shape and a chart gets a new
data validation and conditional format and loses its chart; the save must keep
the slicer, the sparklines and the shape, in XLSX and XLSB.
"""

import re
import zipfile

import pytest

from xlspy import (
    CellIsRule, Chart, ChartSeries, DataValidation, DifferentialStyle, ExcelTable, Font, Shape,
    Slicer, SparklineGroup, XlsbWriter, XlsxWriter,
)
from xlspy.biff_utils import iter_records
from xlspy.workbook import load_workbook

BOLD = DifferentialStyle(font=Font(bold=True))

ROWS = [
    ["Region", "Q1", "Q2", "Q3", "Trend"],
    ["North", 10, 12, 9, None],
    ["South", 7, 11, 13, None],
]


def _source(path, writer_class):
    with writer_class(path) as writer:
        writer.add_sheet("Sheet1")
        writer.write_sheet(
            ROWS,
            tables=[ExcelTable("Sales", "A1:E3")],
            slicers=[Slicer("Sales", "Region", anchor="H2")],
            sparklines=[SparklineGroup(data="B2:D3", location="E2:E3")],
            shapes=[Shape(anchor="H14", text="KPI", link_cell="B2", name="KPI")],
            charts=[Chart("bar", [ChartSeries("B2:B3", categories="A2:A3")], anchor="L2", title="Sales")],
        )
    return path


def _edit_and_save(source, out):
    workbook = load_workbook(source)
    try:
        sheet = workbook.active
        assert [chart.title for chart in sheet.charts] == ["Sales"]  # the slicer is not a chart
        sheet.add_data_validation(DataValidation(type="list", sqref="A5", formula1='"North,South"'))
        sheet.conditional_formatting.add("B2:D3", CellIsRule("B2", "greaterThan", ["10"], dxf=BOLD))
        del sheet._charts[0]
        workbook.save(out)
    finally:
        workbook.close()
    return out


def _parts(path):
    with zipfile.ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def test_xlsx_edits_keep_the_extension_objects(tmp_path):
    out = _edit_and_save(_source(tmp_path / "s.xlsx", XlsxWriter), tmp_path / "out.xlsx")
    parts = _parts(out)
    sheet = parts["xl/worksheets/sheet1.xml"].decode()
    drawing = parts["xl/drawings/drawing1.xml"].decode()
    assert "<x14:sparklineGroups" in sheet and "<x14:slicerList" in sheet
    assert sheet.rindex("<extLst>") > sheet.rindex("</dataValidations>")  # extLst stays last
    assert "<dataValidation " in sheet and "<conditionalFormatting " in sheet
    assert "<sle:slicer " in drawing and 'textlink="$B$2"' in drawing
    assert not re.search(r"<c:chart\b|<xdr:graphicFrame[^>]*>.*chart", drawing)
    assert not [name for name in parts if name.startswith("xl/charts/")]
    assert "xl/slicers/slicer1.xml" in parts and "xl/slicerCaches/slicerCache1.xml" in parts


def test_xlsb_edits_keep_the_extension_objects(tmp_path):
    out = _edit_and_save(_source(tmp_path / "s.xlsb", XlsbWriter), tmp_path / "out.xlsb")
    parts = _parts(out)
    ids = [record.record_id for record in iter_records(parts["xl/worksheets/sheet1.bin"])]
    assert ids.count(0x413) == 2 and 0x439 in ids  # sparklines and the slicer list
    drawing = parts["xl/drawings/drawing1.xml"].decode()
    assert "<sle:slicer " in drawing and 'textlink="$B$2"' in drawing
    assert not [name for name in parts if name.startswith("xl/charts/")]


@pytest.mark.com
@pytest.mark.parametrize("suffix", ["xlsx", "xlsb"])
def test_excel_opens_the_edited_file(tmp_path, suffix):
    win32 = pytest.importorskip("win32com.client")
    writer_class = XlsxWriter if suffix == "xlsx" else XlsbWriter
    out = _edit_and_save(_source(tmp_path / f"s.{suffix}", writer_class), tmp_path / f"out.{suffix}")
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    try:
        book = excel.Workbooks.Open(str(out))
        sheet = book.Worksheets(1)
        sheet.Activate()
        assert book.SlicerCaches.Count == 1
        assert sheet.Range("E2").SparklineGroups.Count == 1
        assert sheet.ChartObjects().Count == 0
        assert sheet.Shapes("KPI").TextFrame2.TextRange.Text == "10"
        # Excel shows the list with the locale's separator (";" in Polish).
        assert sheet.Range("A5").Validation.Formula1.replace(";", ",") == "North,South"
        book.Close(SaveChanges=False)
    finally:
        excel.Quit()
