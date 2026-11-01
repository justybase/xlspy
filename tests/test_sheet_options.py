"""write_sheet takes its settings as a SheetOptions object or as keyword arguments."""

import dataclasses
import zipfile

import pytest

from xlspy import SheetOptions, XlsbWriter, XlsxWriter


def _part(path, name):
    with zipfile.ZipFile(path) as archive:
        return archive.read(name)


def test_options_object_and_keywords_write_the_same_sheet(tmp_path):
    rows = [["a", 1], ["b", 2]]
    by_keyword = tmp_path / "keyword.xlsx"
    by_object = tmp_path / "object.xlsx"
    with XlsxWriter(by_keyword) as writer:
        writer.add_sheet("S")
        writer.write_sheet(rows, zoom=80, freeze_panes="A2")
    with XlsxWriter(by_object) as writer:
        writer.add_sheet("S")
        writer.write_sheet(rows, options=SheetOptions(zoom=80, freeze_panes="A2"))
    sheet = "xl/worksheets/sheet1.xml"
    assert _part(by_keyword, sheet) == _part(by_object, sheet)


def test_keyword_overrides_the_options_object(tmp_path):
    path = tmp_path / "override.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("S")
        writer.write_sheet([[1]], options=SheetOptions(zoom=50), zoom=90)
    assert b'zoomScale="90"' in _part(path, "xl/worksheets/sheet1.xml")


def test_one_options_object_serves_several_sheets(tmp_path):
    path = tmp_path / "several.xlsx"
    options = SheetOptions(zoom=70)
    with XlsxWriter(path) as writer:
        writer.add_sheet("A")
        writer.write_sheet([[1]], options=options)
        writer.add_sheet("B")
        writer.write_sheet([[2]], options=options)
    for part in ("xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml"):
        assert b'zoomScale="70"' in _part(path, part)


def test_sheet_options_is_frozen():
    options = SheetOptions(zoom=70)
    with pytest.raises(dataclasses.FrozenInstanceError):
        options.zoom = 90


def test_unknown_keyword_is_refused(tmp_path):
    with XlsxWriter(tmp_path / "unknown.xlsx") as writer:
        writer.add_sheet("S")
        with pytest.raises(TypeError, match="'zooom'"):
            writer.write_sheet([[1]], zooom=80)


def test_options_must_be_a_sheet_options(tmp_path):
    with XlsxWriter(tmp_path / "type.xlsx") as writer:
        writer.add_sheet("S")
        with pytest.raises(TypeError, match="SheetOptions"):
            writer.write_sheet([[1]], options={"zoom": 80})


def test_xlsb_refuses_an_xlsx_only_option(tmp_path):
    with XlsbWriter(tmp_path / "only.xlsb") as writer:
        writer.add_sheet("S")
        with pytest.raises(TypeError, match="ignored_errors"):
            writer.write_sheet([[1]], options=SheetOptions(ignored_errors=[]))
        with pytest.raises(TypeError, match="ignored_errors"):
            writer.write_sheet([[1]], ignored_errors=[])
