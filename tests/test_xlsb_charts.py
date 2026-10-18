import struct
import xml.etree.ElementTree as ET
import zipfile

import pytest

from xlspy import Chart, ChartSeries, ExcelReader, XlsbWriter, XlsxWriter
from xlspy.biff_utils import iter_records
from xlspy.writer import _BRT_DRAWING


def _sample_charts():
    return [
        Chart(
            "bar",
            [
                ChartSeries("B2:B4", title_cell="B1", categories="A2:A4"),
                ChartSeries("C2:C4", title="Lit", categories="A2:A4"),
            ],
            title="Sales & <Costs>",
            anchor="E2",
            style=10,
        ),
        Chart(
            "pie",
            [ChartSeries("B2:B4", title_cell="B1")],
            title="Share",
            anchor="L2",
        ),
    ]


def _sample_data():
    return [
        ["Q", "A", "B"],
        ["Q1", 10, 20],
        ["Q2", 30, 40],
        ["Q3", 50, 60],
    ]


def _rel_ids(payload: bytes) -> str:
    (count,) = struct.unpack("<I", payload[:4])
    return payload[4:4 + 2 * count].decode("utf-16-le")


def _drawing_records(sheet_bin: bytes):
    return [r for r in iter_records(sheet_bin) if r.record_id == _BRT_DRAWING]


def test_xlsb_rejects_non_charts(tmp_path):
    with XlsbWriter(tmp_path / "bad.xlsb") as writer:
        writer.add_sheet("Data")
        with pytest.raises(TypeError, match="Chart instances"):
            writer.write_sheet([["x"]], charts=["B2:B4"])


def test_xlsb_chart_xml_matches_xlsx_except_caches(tmp_path):
    """Same structure as XLSX, plus Excel-style caches in the XLSB parts."""
    xlsx_path = tmp_path / "same.xlsx"
    with XlsxWriter(xlsx_path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(_sample_data(), charts=_sample_charts())

    xlsb_path = tmp_path / "same.xlsb"
    with XlsbWriter(xlsb_path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(_sample_data(), charts=_sample_charts())

    with zipfile.ZipFile(xlsx_path) as xlsx, zipfile.ZipFile(xlsb_path) as xlsb:
        # The drawing (anchors, sizes, rels) is shared verbatim.
        assert xlsb.read("xl/drawings/drawing1.xml") == xlsx.read("xl/drawings/drawing1.xml")
        assert (
            xlsb.read("xl/drawings/_rels/drawing1.xml.rels")
            == xlsx.read("xl/drawings/_rels/drawing1.xml.rels")
        )
        bar_ref = xlsx.read("xl/charts/chart1.xml").decode()
        assert "strCache" not in bar_ref and "numCache" not in bar_ref
        assert "varyColors" not in bar_ref
        pie_ref = xlsx.read("xl/charts/chart2.xml").decode()
        assert "strCache" not in pie_ref and "numCache" not in pie_ref
        assert '<pieChart><varyColors val="1"/>' in pie_ref
        for name in ("xl/charts/chart1.xml", "xl/charts/chart2.xml"):
            assert xlsb.read(name).decode() != xlsx.read(name).decode()
        # Both are well-formed XML.
        for name in ("xl/charts/chart1.xml", "xl/drawings/drawing1.xml"):
            ET.fromstring(xlsb.read(name))


def test_xlsb_chart_caches_carry_cell_values(tmp_path):
    path = tmp_path / "caches.xlsb"
    with XlsbWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(_sample_data(), charts=_sample_charts())

    with zipfile.ZipFile(path) as archive:
        bar = archive.read("xl/charts/chart1.xml").decode()
        # Series values are cached as numbers...
        assert "<numCache><formatCode>General</formatCode>" in bar
        assert "<pt idx=\"0\"><v>10</v></pt>" in bar
        assert "<pt idx=\"2\"><v>60</v></pt>" in bar
        # ...titles via strRef + strCache...
        assert "<f>'Data'!B1</f><strCache>" in bar
        assert "<pt idx=\"0\"><v>A</v></pt>" in bar
        assert "<tx><v>Lit</v></tx>" in bar
        # ...and text categories via strRef (like Excel), not numRef.
        assert "<cat><strRef><f>'Data'!$A$2:$A$4</f>" in bar
        assert "<cat><numRef>" not in bar
        assert "<pt idx=\"0\"><v>Q1</v></pt>" in bar
        # Explicit single-series colors, matching Excel-generated files.
        assert "<barChart><barDir val=\"col\"/><grouping val=\"clustered\"/>" in bar
        assert "<varyColors val=\"0\"/>" in bar
        # Explicit axis attributes: the XLSB loader hides labels without them.
        assert "<delete val=\"0\"/><axPos val=\"b\"/>" in bar
        assert "<delete val=\"0\"/><axPos val=\"l\"/><majorGridlines/>" in bar
        assert "<tickLblPos val=\"nextTo\"/>" in bar
        pie = archive.read("xl/charts/chart2.xml").decode()
        assert "<pieChart><varyColors val=\"1\"/>" in pie


def test_xlsb_chart_style_uses_c14_extension(tmp_path):
    """Numeric styles go through mc:AlternateContent, like Excel's own files."""
    path = tmp_path / "styled.xlsb"
    with XlsbWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(_sample_data(), charts=_sample_charts())
        writer.add_sheet("Second")
        writer.write_sheet([["x", 1]], charts=[
            Chart("bar", [ChartSeries("B1:B1")], style=2, anchor="C3"),
        ])

    with zipfile.ZipFile(path) as archive:
        bar = archive.read("xl/charts/chart1.xml").decode()
        # Style 10 becomes c14:style 110 (legacy + 100, [MS-ODRAWXML]
        # 2.7.3.5) with the legacy element as fallback; a bare legacy
        # element alone is not applied by Excel's XLSB loader.
        expected = (
            "<mc:AlternateContent"
            " xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\">"
            "<mc:Choice Requires=\"c14\""
            " xmlns:c14=\"http://schemas.microsoft.com/office/drawing/2007/8/2/chart\">"
            "<c14:style val=\"110\"/></mc:Choice>"
            "<mc:Fallback><style val=\"10\"/></mc:Fallback>"
            "</mc:AlternateContent>"
        )
        assert expected in bar
        assert "<style val=\"10\"/><chart>" not in bar
        ET.fromstring(bar)
        # The lowest style follows the same rule (2 -> 102)...
        second = archive.read("xl/charts/chart3.xml").decode()
        assert "<c14:style val=\"102\"/>" in second
        assert "<mc:Fallback><style val=\"2\"/></mc:Fallback>" in second
        # ...and charts without a style keep the implicit default: no block.
        pie = archive.read("xl/charts/chart2.xml").decode()
        assert "AlternateContent" not in pie


def test_xlsb_brt_drawing_record(tmp_path):
    from xlspy.writer import XlsbWriter as W

    path = tmp_path / "drawing.xlsb"
    with XlsbWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(_sample_data(), charts=_sample_charts()[:1])

    with zipfile.ZipFile(path) as archive:
        sheet_bin = archive.read("xl/worksheets/sheet1.bin")

    records = list(iter_records(sheet_bin))
    drawings = [r for r in records if r.record_id == _BRT_DRAWING]
    assert len(drawings) == 1
    assert _rel_ids(sheet_bin[drawings[0].data_start:drawings[0].data_end]) == "rId2"

    # BrtDrawing sits between BrtMargins (476) and BrtEndSheet (130),
    # directly before the tail FRT alternative-content block (0x25).
    # (The template also has an earlier 0x25 for the ACCELLTABLE block,
    # so the assertion uses record neighbours, not first occurrence.)
    ids = [r.record_id for r in records]
    pos = ids.index(_BRT_DRAWING)
    assert ids[pos - 1] == 476
    assert ids[pos + 1] == 0x25
    assert ids.index(_BRT_DRAWING) < ids.index(130)

    # Template split offsets used by the writer: 290 starts BrtPrintOptions
    # (477), 346 starts the tail ACBegin (0x25) block.
    template_records = list(iter_records(W._sheet1Bytes))
    by_offset = {r.header_start: r for r in template_records}
    assert by_offset[290].record_id == 477
    assert by_offset[346].record_id == 0x25


def test_xlsb_sheet_rels_and_content_types(tmp_path):
    path = tmp_path / "rels.xlsb"
    with XlsbWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(_sample_data(), charts=_sample_charts())
        writer.add_sheet("Plain")
        writer.write_sheet([["x", 1]])

    with zipfile.ZipFile(path) as archive:
        sheet_rels = archive.read("xl/worksheets/_rels/sheet1.bin.rels").decode()
        assert 'Id="rId1"' in sheet_rels and "xlBinaryIndex" in sheet_rels
        assert 'Id="rId2"' in sheet_rels
        assert 'Target="../drawings/drawing1.xml"' in sheet_rels
        plain_rels = archive.read("xl/worksheets/_rels/sheet2.bin.rels").decode()
        assert "drawing" not in plain_rels
        content_types = archive.read("[Content_Types].xml").decode()
        assert "/xl/drawings/drawing1.xml" in content_types
        assert "/xl/charts/chart1.xml" in content_types
        assert "/xl/charts/chart2.xml" in content_types
        assert "/xl/theme/theme1.xml" in content_types
        assert "drawing2" not in content_types
        assert archive.read("xl/theme/theme1.xml").startswith(b"<?xml")
        workbook_rels = archive.read("xl/_rels/workbook.bin.rels").decode()
        assert 'Target="theme/theme1.xml"' in workbook_rels


def test_xlsb_charts_multi_sheet(tmp_path):
    path = tmp_path / "multi.xlsb"
    with XlsbWriter(path) as writer:
        writer.add_sheet("First")
        writer.write_sheet(_sample_data(), charts=_sample_charts())
        writer.add_sheet("Second")
        writer.write_sheet([["x", 1]], charts=[
            Chart("bar", [ChartSeries("B1:B1")], anchor="C3"),
        ])

    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
        assert "xl/drawings/drawing1.xml" in names
        assert "xl/drawings/drawing2.xml" in names
        assert "xl/charts/chart3.xml" in names
        for sheet_name in ("sheet1.bin", "sheet2.bin"):
            sheet_bin = archive.read(f"xl/worksheets/{sheet_name}")
            drawings = _drawing_records(sheet_bin)
            assert len(drawings) == 1
            assert _rel_ids(sheet_bin[drawings[0].data_start:drawings[0].data_end]) == "rId2"


def test_xlsb_plain_write_has_no_drawing(tmp_path):
    plain_path = tmp_path / "plain.xlsb"
    with XlsbWriter(plain_path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([[f"h{j}"] + [i] * 3 for j in range(1) for i in range(11)])

    chart_path = tmp_path / "chart.xlsb"
    with XlsbWriter(chart_path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([[f"h{j}"] + [i] * 3 for j in range(1) for i in range(11)],
                           charts=_sample_charts()[:1])

    with zipfile.ZipFile(plain_path) as plain, zipfile.ZipFile(chart_path) as charted:
        names = plain.namelist()
        assert not [n for n in names if "/drawings/" in n or "/charts/" in n]
        assert "drawing" not in plain.read("[Content_Types].xml").decode()
        assert "theme" not in plain.read("[Content_Types].xml").decode()
        assert "theme" not in plain.read("xl/_rels/workbook.bin.rels").decode()
        plain_bin = plain.read("xl/worksheets/sheet1.bin")
        assert _drawing_records(plain_bin) == []
        # Charts do not touch the workbook binary or the styles.
        assert plain.read("xl/workbook.bin") == charted.read("xl/workbook.bin")
        assert plain.read("xl/styles.bin") == charted.read("xl/styles.bin")


def test_xlsb_charts_keep_cell_data_readable(tmp_path):
    path = tmp_path / "readable.xlsb"
    with XlsbWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(_sample_data(), charts=_sample_charts()[:1])

    with ExcelReader(str(path)) as reader:
        assert list(reader.get_rows("Data"))[1] == ["Q1", 10, 20]


def test_xlsb_charts_via_dataframe(tmp_path):
    pd = pytest.importorskip("pandas")
    from xlspy import to_xlsb

    df = pd.DataFrame({"Q": ["Q1", "Q2"], "A": [10, 30]})
    path = to_xlsb(
        df, tmp_path / "df.xlsb",
        charts=[Chart("bar", [ChartSeries("B2:B3", title_cell="B1")], anchor="D2")],
    )
    with zipfile.ZipFile(path) as archive:
        assert "xl/charts/chart1.xml" in archive.namelist()
        sheet_bin = archive.read("xl/worksheets/sheet1.bin")
        assert len(_drawing_records(sheet_bin)) == 1
