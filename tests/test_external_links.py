"""External workbook links are listed (read only) and survive an edit."""

import zipfile
from pathlib import Path

import pytest

from xlspy import ExternalLink, load_workbook

DATA = Path(__file__).parent / "data"  # both made by Excel


@pytest.mark.parametrize("name,part", [
    ("external_ref.xlsx", "xl/externalLinks/externalLink1.xml"),
    ("external_ref.xlsb", "xl/externalLinks/externalLink1.bin"),
])
def test_external_links_are_listed(name, part):
    workbook = load_workbook(DATA / name)
    try:
        assert workbook.external_links == [ExternalLink(1, "ext.xlsx", ["Data"], part)]
    finally:
        workbook.close()


@pytest.mark.parametrize("name", ["external_ref.xlsx", "external_ref.xlsb"])
def test_external_links_survive_a_save(tmp_path, name):
    out = tmp_path / name
    workbook = load_workbook(DATA / name)
    try:
        workbook.active["A3"] = 9
        workbook.save(out)
    finally:
        workbook.close()
    reloaded = load_workbook(out)
    try:
        assert [link.sheet_names for link in reloaded.external_links] == [["Data"]]
    finally:
        reloaded.close()
    with zipfile.ZipFile(out) as archive:
        assert any(n.startswith("xl/externalLinks/externalLink1") for n in archive.namelist())


def test_workbook_without_links_lists_none():
    workbook = load_workbook(DATA / "chartsheet.xlsb")
    try:
        assert workbook.external_links == []
    finally:
        workbook.close()
