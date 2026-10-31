"""Shapes and text boxes (KPI cards) in the drawing of XLSX and XLSB output."""

import xml.etree.ElementTree as ET
import zipfile

import pytest

from xlspy import XlsbWriter, XlsxWriter
from xlspy.shapes import Shape, normalize_shapes, shape_frame_xml

ROWS = [["Metric", "Value"], ["Revenue", 1234]]


def _drawing(path, name):
    with zipfile.ZipFile(path) as archive:
        return archive.read(name).decode("utf-8")


def _wrap(fragment):
    return (
        '<wsDr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">'
        f"{fragment}</wsDr>"
    )


@pytest.mark.parametrize(
    "kwargs, message",
    [
        ({"anchor": "B2:C3"}, "single cell"),
        ({"anchor": "B2", "width": 0}, "width"),
        ({"anchor": "B2", "height": 1.5}, "height"),
        ({"anchor": "B2", "kind": "star"}, "kind"),
        ({"anchor": "B2", "align": "justify"}, "align"),
        ({"anchor": "B2", "valign": "middle"}, "valign"),
        ({"anchor": "B2", "fill": "red"}, "fill"),
        ({"anchor": "B2", "line": "#12345"}, "line"),
        ({"anchor": "B2", "font_color": "#GGGGGG"}, "font_color"),
        ({"anchor": "B2", "font_size": 0}, "font_size"),
        ({"anchor": "B2", "link_cell": "B2:B3"}, "single cell"),
        ({"anchor": "B2", "name": ""}, "name"),
    ],
)
def test_invalid_shapes_are_rejected(kwargs, message):
    with pytest.raises(ValueError, match=message):
        Shape(**kwargs)


def test_normalize_rejects_non_shapes():
    with pytest.raises(TypeError):
        normalize_shapes([object()])


def test_text_is_escaped_and_split_into_paragraphs():
    xml = shape_frame_xml(Shape(anchor="B2", text="A & B\n<C>"), 3)
    root = ET.fromstring(_wrap(xml))
    texts = [node.text for node in root.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}t")]
    assert texts == ["A & B", "<C>"]


def test_link_cell_writes_textlink_and_a_linked_field():
    xml = shape_frame_xml(Shape(anchor="B2", text="1 234", link_cell="C5"), 4)
    root = ET.fromstring(_wrap(xml))
    sp = root.find(".//{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}sp")
    assert sp.get("textlink") == "$C$5"
    fields = root.findall(".//{http://schemas.openxmlformats.org/drawingml/2006/main}fld")
    assert len(fields) == 1
    assert fields[0].get("type") == "TxLink"
    assert fields[0].find("{http://schemas.openxmlformats.org/drawingml/2006/main}t").text == "1 234"


def test_plain_shape_has_no_textlink():
    xml = shape_frame_xml(Shape(anchor="B2"), 2)
    assert "textlink" not in xml


def test_textbox_kind_sets_txbox_flag_and_ellipse_geometry():
    textbox = shape_frame_xml(Shape(anchor="B2", kind="textbox"), 2)
    ellipse = shape_frame_xml(Shape(anchor="B2", kind="ellipse", fill="#1F4E79", line="#FFFFFF"), 3)
    assert 'txBox="1"' in textbox
    assert 'prst="ellipse"' in ellipse
    assert '<a:srgbClr val="1F4E79"/>' in ellipse


def test_xlsx_sheet_with_shapes_has_drawing(tmp_path):
    path = tmp_path / "kpi.xlsx"
    with XlsxWriter(path) as writer:
        writer.write_sheet(ROWS, shapes=[Shape(anchor="D2", width=180, height=90, fill="#F2F2F2",
                                               text="Revenue", link_cell="B2")])
    drawing = _drawing(path, "xl/drawings/drawing1.xml")
    assert 'textlink="$B$2"' in drawing
    assert 'name="Shape 1"' in drawing  # the first object on a sheet with no charts or pictures


def test_xlsb_sheet_with_shapes_has_drawing(tmp_path):
    path = tmp_path / "kpi.xlsb"
    with XlsbWriter(path) as writer:
        writer.add_sheet("Sheet1")
        writer.write_sheet(ROWS, shapes=[Shape(anchor="D2", text="Revenue")])
    drawing = _drawing(path, "xl/drawings/drawing1.xml")
    assert "<sp " in drawing
    assert "Revenue" in drawing


def test_shapes_on_one_sheet_get_unique_ids(tmp_path):
    path = tmp_path / "mixed.xlsx"
    with XlsxWriter(path) as writer:
        writer.write_sheet(ROWS, shapes=[Shape(anchor="D2"), Shape(anchor="D8")])
    drawing = _drawing(path, "xl/drawings/drawing1.xml")
    ids = [node.get("id") for node in ET.fromstring(drawing).iter(
        "{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}cNvPr")]
    assert len(ids) == len(set(ids))


# --- Excel COM -----------------------------------------------------------------

@pytest.mark.com
def test_excel_reads_kpi_shape_text_and_link(tmp_path):
    win32 = pytest.importorskip("win32com.client")
    path = tmp_path / "kpi_com.xlsx"
    with XlsxWriter(path) as writer:
        writer.write_sheet(ROWS, shapes=[Shape(anchor="D2", text="1 234", link_cell="B2", name="KPI")])
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    try:
        book = excel.Workbooks.Open(str(path))
        sheet = book.Worksheets(1)
        sheet.Activate()
        shape = sheet.Shapes("KPI")
        assert shape.DrawingObject.Formula.strip() == "$B$2"  # the cell the card is linked to
        assert shape.TextFrame2.TextRange.Text == "1234"
        book.Close(SaveChanges=False)
    finally:
        excel.Quit()
