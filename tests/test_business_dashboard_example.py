"""The business dashboard example writes both formats, and the files load back."""

import importlib.util
import zipfile
from pathlib import Path

import pytest

from xlspy.biff_utils import iter_records
from xlspy.workbook import load_workbook

EXAMPLE = Path(__file__).parents[1] / "examples" / "business_dashboard.py"


@pytest.fixture(scope="module")
def dashboard():
    spec = importlib.util.spec_from_file_location("business_dashboard", EXAMPLE)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _names(path):
    with zipfile.ZipFile(path) as archive:
        return archive.namelist()


def test_xlsx_dashboard_has_slicer_sparklines_cards_and_charts(dashboard, tmp_path):
    out = tmp_path / "dashboard.xlsx"
    dashboard.main(["--out", str(out)])
    names = _names(out)
    assert "xl/slicers/slicer1.xml" in names
    assert "xl/charts/chart3.xml" in names
    assert "xl/drawings/drawing1.xml" in names
    with zipfile.ZipFile(out) as archive:
        sheet = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")
        drawing = archive.read("xl/drawings/drawing1.xml").decode("utf-8")
    assert "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" in sheet  # sparklines
    assert 'textlink="$T$5"' in drawing  # a KPI card linked to its helper cell
    workbook = load_workbook(out)
    try:
        assert workbook.worksheets[0].title == "Dashboard"
    finally:
        workbook.close()


def test_xlsb_dashboard_has_slicer_sparklines_cards_and_charts(dashboard, tmp_path):
    out = tmp_path / "dashboard.xlsb"
    dashboard.main(["--out", str(out)])
    names = _names(out)
    assert "xl/charts/chart3.xml" in names
    assert {"xl/slicers/slicer1.bin", "xl/slicerCaches/slicerCache1.bin"} <= set(names)
    with zipfile.ZipFile(out) as archive:
        drawing = archive.read("xl/drawings/drawing1.xml").decode("utf-8")
        sheet = archive.read("xl/worksheets/sheet1.bin")
    assert "textlink" in drawing
    sparklines = [record for record in iter_records(sheet) if record.record_id == 0x413]
    assert len(sparklines) == 4  # one per region
    workbook = load_workbook(out, data_only=False)
    try:
        sheet = workbook.worksheets[0]
        assert sheet["T5"].value == "=SUMIFS($E$51:$E$266,$H$51:$H$266,1)"
        assert sheet["H51"].value == "=SUBTOTAL(103,A51)"
    finally:
        workbook.close()


def test_dashboard_totals_match_the_rows(dashboard):
    rows = dashboard.sample_rows()
    totals = dashboard._totals(rows)
    assert totals["orders"] == len(rows) == 216
    assert totals["sales"] == sum(row[4] for row in rows)
    assert sum(totals["product"].values()) == totals["sales"]
    assert sum(totals["channel"].values()) == totals["sales"]
