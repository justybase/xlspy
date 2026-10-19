"""Smaller checks for XLSX feature corners and the XLSB refusals."""

import zipfile

import pytest

from xlspy import PrintSettings, XlsbWriter, XlsxWriter

openpyxl = pytest.importorskip("openpyxl")


def _workbook_xml(path):
    with zipfile.ZipFile(path) as archive:
        return archive.read("xl/workbook.xml").decode("utf-8")


def test_multi_range_print_area_names_each_part(tmp_path):
    path = tmp_path / "areas.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Print Me")
        writer.write_sheet([["a", "b"]], print_settings=PrintSettings(print_area="A1:B2,D4:E5"))
    xml = _workbook_xml(path)
    assert "&apos;Print Me&apos;!$A$1:$B$2,&apos;Print Me&apos;!$D$4:$E$5" in xml or \
        "'Print Me'!$A$1:$B$2,'Print Me'!$D$4:$E$5" in xml


def test_filter_database_name_quotes_sheet_names_with_spaces(tmp_path):
    path = tmp_path / "filter.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Q1 Sales")
        writer.write_sheet([["a", "b"], [1, 2], [3, 4]])
    xml = _workbook_xml(path)
    assert "_xlnm._FilterDatabase" in xml
    assert "&apos;Q1 Sales&apos;!$A$1:$B$3" in xml or "'Q1 Sales'!$A$1:$B$3" in xml
    ws = openpyxl.load_workbook(path)["Q1 Sales"]
    assert ws.auto_filter.ref == "A1:B3"


def test_xlsb_refuses_xlsx_only_arguments(tmp_path):
    from xlspy import Image

    png = tmp_path / "x.png"
    png.write_bytes(b"\x89PNG\r\n\x1a\n" + b"\x00\x00\x00\rIHDR" + (1).to_bytes(4, "big") * 2
                    + b"\x08\x02\x00\x00\x00" + b"\x00" * 8)
    with pytest.raises(TypeError, match="unexpected keyword argument 'images'"):
        with XlsbWriter(tmp_path / "x.xlsb") as writer:
            writer.add_sheet("Sheet1")
            writer.write_sheet([["a"]], images=[Image(png)])
    with pytest.raises(TypeError, match="unexpected keyword argument 'print_settings'"):
        with XlsbWriter(tmp_path / "y.xlsb") as writer:
            writer.add_sheet("Sheet1")
            writer.write_sheet([["a"]], print_settings=PrintSettings(orientation="landscape"))


def test_explicit_column_width_beyond_the_data_is_written(tmp_path):
    path = tmp_path / "widths.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Sheet1")
        writer.write_sheet([["a", "b"]], column_widths={"F": 22})
    ws = openpyxl.load_workbook(path)["Sheet1"]
    assert ws.column_dimensions["F"].width == pytest.approx(22.0)


def test_xlsx_uses_hyperlink_style_only_when_the_cell_has_none(tmp_path):
    from xlspy import Hyperlink, Style

    path = tmp_path / "style.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Sheet1")
        writer.write_sheet(
            [[Hyperlink("https://example.com", text="plain"), Hyperlink("https://example.org", text="own")]],
            styles={"B1": Style(font_color="#FF0000", underline=False)},
        )
    ws = openpyxl.load_workbook(path)["Sheet1"]
    assert ws["A1"].font.underline == "single"
    assert ws["B1"].font.underline is None
    assert ws["B1"].font.color.rgb.endswith("FF0000")
