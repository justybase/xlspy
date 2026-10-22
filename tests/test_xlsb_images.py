"""Pictures in XLSB output: the drawing part, media, relationships and content types."""

import re
import zipfile

import pytest

from xlspy import Chart, ChartSeries, Comment, Image, XlsbWriter, XlsxWriter
from xlspy.biff_utils import iter_records
from xlspy.writer import _BRT_DRAWING
from test_xlsx_features import _png_bytes


def _parts(path):
    with zipfile.ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def _drawing_rel(sheet_bin: bytes) -> str:
    (record,) = [record for record in iter_records(sheet_bin) if record.record_id == _BRT_DRAWING]
    payload = sheet_bin[record.data_start:record.data_end]
    return payload[4:].decode("utf-16-le")


def test_pictures_share_media_and_link_from_the_drawing(tmp_path):
    png = _png_bytes(40, 20)
    path = tmp_path / "pictures.xlsb"
    with XlsbWriter(path) as writer:
        writer.add_sheet("Logo")
        writer.write_sheet([["x"]], images=[Image(png, anchor="B2"), Image(png, anchor="D2", width=80)])
        writer.add_sheet("Again")
        writer.write_sheet([["y"]], images=[Image(png, anchor="A3", description="Company logo")])

    parts = _parts(path)
    # One media part for the three identical pictures.
    assert [name for name in parts if name.startswith("xl/media/")] == ["xl/media/image1.png"]
    assert parts["xl/media/image1.png"] == png
    content_types = parts["[Content_Types].xml"].decode()
    assert '<Default Extension="png" ContentType="image/png"/>' in content_types
    assert content_types.count("drawing+xml") == 2

    for number, pictures in ((1, 2), (2, 1)):
        drawing = parts[f"xl/drawings/drawing{number}.xml"].decode()
        assert drawing.count("<pic>") == pictures
        rels = parts[f"xl/drawings/_rels/drawing{number}.xml.rels"].decode()
        assert rels.count('Target="../media/image1.png"') == pictures
        sheet_rels = parts[f"xl/worksheets/_rels/sheet{number}.bin.rels"].decode()
        rel_id = _drawing_rel(parts[f"xl/worksheets/sheet{number}.bin"])
        assert f'Id="{rel_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"' in sheet_rels
    assert 'descr="Company logo"' in parts["xl/drawings/drawing2.xml"].decode()
    # Every XLSB carries the theme, pictures or not.
    assert parts["xl/theme/theme1.xml"]


def test_pictures_and_charts_share_one_drawing(tmp_path):
    path = tmp_path / "mixed.xlsb"
    with XlsbWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(
            [["k", "v"], ["a", 1], ["b", 2]],
            charts=[Chart("bar", [ChartSeries("B2:B3")], anchor="D2")],
            images=[Image(_png_bytes(10, 10), anchor="D20")],
            comments={"A1": Comment("note", "Author")},
        )

    parts = _parts(path)
    drawing = parts["xl/drawings/drawing1.xml"].decode()
    assert drawing.count("<c:chart ") == 1 and drawing.count("<pic>") == 1
    rels = parts["xl/drawings/_rels/drawing1.xml.rels"].decode()
    assert 'Id="rId1"' in rels and 'Target="../charts/chart1.xml"' in rels
    assert 'Id="rId2"' in rels and 'Target="../media/image1.png"' in rels
    sheet_rels = parts["xl/worksheets/_rels/sheet1.bin.rels"].decode()
    ids = re.findall(r'Id="(rId\d+)"', sheet_rels)
    assert len(ids) == len(set(ids)) == 4  # binary index, drawing, comments, VML


def test_drawing_part_matches_xlsx(tmp_path):
    images = [Image(_png_bytes(40, 20), anchor="B2", width=60)]
    for writer_class, suffix in ((XlsxWriter, "xlsx"), (XlsbWriter, "xlsb")):
        with writer_class(tmp_path / f"p.{suffix}") as writer:
            writer.add_sheet("S")
            writer.write_sheet([["x"]], images=images)
    xlsx, xlsb = _parts(tmp_path / "p.xlsx"), _parts(tmp_path / "p.xlsb")
    assert xlsb["xl/drawings/drawing1.xml"] == xlsx["xl/drawings/drawing1.xml"]
    assert xlsb["xl/drawings/_rels/drawing1.xml.rels"] == xlsx["xl/drawings/_rels/drawing1.xml.rels"]


def test_images_must_be_image_instances(tmp_path):
    with XlsbWriter(tmp_path / "bad.xlsb") as writer:
        writer.add_sheet("S")
        with pytest.raises(TypeError, match="Image instances"):
            writer.write_sheet([["x"]], images=["logo.png"])
        writer.write_sheet([["x"]])
