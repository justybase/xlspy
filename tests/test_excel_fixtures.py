"""Read workbooks written by desktop Excel 365, without Excel on the test machine.

``tools/make_excel_fixtures.py`` produces the files and the manifest; CI only
reads the committed corpus.  The manifest keeps the Excel build and the file
hashes visible, so a refresh is an explicit commit rather than a silent drift.
"""

import datetime as dt
import hashlib
import json
from pathlib import Path

import pytest

from xlspy import ExcelReader, load_workbook

FIXTURES = Path(__file__).with_name("data") / "excel_365"
MANIFEST = FIXTURES / "manifest.json"


def _manifest():
    return json.loads(MANIFEST.read_text(encoding="utf-8"))


@pytest.mark.parametrize("suffix", ["xlsx", "xlsb"])
def test_the_corpus_matches_its_manifest(suffix):
    manifest = _manifest()
    assert manifest["excel"]["version"].startswith("16.")
    name = f"features.{suffix}"
    digest = hashlib.sha256((FIXTURES / name).read_bytes()).hexdigest()
    assert manifest["files"][name] == digest


@pytest.mark.parametrize("suffix", ["xlsx", "xlsb"])
def test_excel_rows_read_with_values_and_dates(suffix):
    path = FIXTURES / f"features.{suffix}"
    with ExcelReader(str(path)) as reader:
        rows = reader.read_all("Data")
    assert rows[0] == ["name", "value", "date", "link", "merged", None]
    assert rows[1] == [
        "alpha",
        12500,
        dt.datetime(2024, 1, 1),
        "https://example.com/excel-fixture",
        None,
        None,
    ]
    assert rows[2][0] == "beta"
    assert rows[2][1] == 8400
    assert rows[2][2] == dt.datetime(2024, 2, 6)
    # Excel's cached formula result: the reader shows it, the model shows the text.
    assert rows[3][0] == "total"
    assert rows[3][1] == 20900


@pytest.mark.parametrize("suffix", ["xlsx", "xlsb"])
def test_excel_model_reads_styles_objects_and_metadata(suffix):
    workbook = load_workbook(FIXTURES / f"features.{suffix}")
    try:
        sheet = workbook["Data"]
        assert sheet.max_row == 4
        assert sheet.max_column == 6
        assert sheet["B1"].font.bold is True
        assert sheet["B1"].fill.fill_type == "solid"
        assert sheet["B1"].fill.start_color.rgb == "FF183C52"
        assert sheet["B2"].number_format == "#,##0"
        # Excel escapes the y's and the dash in the stored format code.
        assert sheet["C2"].number_format.replace("\\", "") == "yyyy-mm-dd"
        assert sheet["B4"].formula == "SUM(B2:B3)"
        assert sheet["D2"].hyperlink.target == "https://example.com/excel-fixture"
        assert sheet["A1"].comment.text == "Written by Excel for the read-only corpus"
        assert list(sheet.merged_cells) == ["E1:F1"]
        assert [table.ref for table in sheet.tables] == ["A1:C3"]
    finally:
        workbook.close()
