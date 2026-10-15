"""XLSB UTF-16 character-count and surrogate-pair regressions."""

import pytest

from xlspy import ExcelReader, XlsbUpdater, XlsbWriter
from xlspy import reader as xlspy_reader


ASTRAL = "𐀀𝄞"


@pytest.mark.parametrize("force_python", [False, True], ids=["default", "python"])
def test_xlsb_roundtrips_supplementary_unicode(tmp_path, monkeypatch, force_python):
    if force_python:
        monkeypatch.setattr(xlspy_reader, "HAVE_C_EXT", False)

    path = tmp_path / "unicode.xlsb"
    sheet_name = f"Sheet {ASTRAL}"
    value = f"before {ASTRAL} after"
    with XlsbWriter(str(path)) as writer:
        writer.add_sheet(sheet_name)
        writer.write_sheet([["header"], [value]])

    with ExcelReader(str(path)) as reader:
        assert reader.get_sheet_names() == [sheet_name]
        assert reader.read_all(sheet_name) == [["header"], [value]]


@pytest.mark.parametrize("force_python", [False, True], ids=["default", "python"])
def test_xlsb_updater_preserves_supplementary_unicode(tmp_path, monkeypatch, force_python):
    if force_python:
        monkeypatch.setattr(xlspy_reader, "HAVE_C_EXT", False)

    path = tmp_path / "updated.xlsb"
    with XlsbWriter(str(path)) as writer:
        writer.add_sheet("Sheet1")
        writer.write_sheet([["original"]])

    replacement = f"updated {ASTRAL} value"
    updater = XlsbUpdater(path)
    try:
        updater.replace_sheet_data("Sheet1", [[replacement]])
        updater.save()
    finally:
        updater.close()

    with ExcelReader(str(path)) as reader:
        assert reader.read_all("Sheet1") == [[replacement]]
