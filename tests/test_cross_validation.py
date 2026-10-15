"""Cross-validation against openpyxl, the reference implementation.

Two directions are covered:

* the goldens in ``tests/data`` (produced by openpyxl and xlsxwriter, see
  ``tools/make_goldens.py``) must read exactly as recorded, and openpyxl must
  see the same values in them;
* files produced by the xlspy writers and updaters must read back in openpyxl
  as the values that were written, and both readers must agree on them.

openpyxl cannot read XLSB; those files are covered by the goldens (when an
Excel-produced one is added) and by the writer/reader round-trip tests.
"""

import datetime as dt
import json
import pathlib

import pytest

from xlspy import ExcelReader, XlsbWriter, XlsxUpdater, XlsxWriter

from golden_data import SHEETS, encoded_sheets, writer_rows

DATA_DIR = pathlib.Path(__file__).parent / "data"
GOLDEN_FILES = sorted(
    path for path in DATA_DIR.glob("golden_*") if path.suffix in {".xlsx", ".xlsb"}
)
GOLDEN_IDS = [path.name for path in GOLDEN_FILES]


def _write_xlsx(path, rows, sheet_name="Main"):
    with XlsxWriter(str(path)) as writer:
        writer.add_sheet(sheet_name)
        writer.write_sheet(rows)
    return path


def _read_with_xlspy(path):
    with ExcelReader(str(path)) as reader:
        return {
            name: [list(row) for row in reader.get_rows(name)]
            for name in reader.get_sheet_names()
        }


def _read_with_openpyxl(path):
    openpyxl = pytest.importorskip("openpyxl")
    workbook = openpyxl.load_workbook(str(path), data_only=True)
    try:
        return {
            name: [list(row) for row in workbook[name].iter_rows(values_only=True)]
            for name in workbook.sheetnames
        }
    finally:
        workbook.close()


def _decode_cell(value):
    if isinstance(value, dict) and "datetime" in value:
        return dt.datetime.fromisoformat(value["datetime"])
    return value


def _decode_sheets(encoded):
    return {
        name: [[_decode_cell(cell) for cell in row] for row in rows]
        for name, rows in encoded.items()
    }


def _normalize(value):
    if isinstance(value, bool) or value is None:
        return value
    if isinstance(value, (int, float)):
        return float(value)
    return value


def _assert_same_grid(actual, expected, context):
    assert set(actual) == set(expected), f"{context}: sheets {set(actual)} != {set(expected)}"
    for sheet in expected:
        actual_rows = actual[sheet]
        expected_rows = expected[sheet]
        assert len(actual_rows) == len(expected_rows), (
            f"{context} [{sheet}]: {len(actual_rows)} rows != {len(expected_rows)}"
        )
        for row_index, (actual_row, expected_row) in enumerate(zip(actual_rows, expected_rows)):
            assert len(actual_row) == len(expected_row), (
                f"{context} [{sheet}] row {row_index}: width {len(actual_row)} != {len(expected_row)}"
            )
            for column_index, (actual_cell, expected_cell) in enumerate(zip(actual_row, expected_row)):
                where = f"{context} [{sheet}] row {row_index} column {column_index}"
                assert _normalize(actual_cell) == _normalize(expected_cell), (
                    f"{where}: {actual_cell!r} != {expected_cell!r}"
                )


def test_golden_files_are_present():
    assert GOLDEN_FILES, "no golden workbooks found in tests/data"


@pytest.mark.parametrize("golden", GOLDEN_FILES, ids=GOLDEN_IDS)
def test_golden_workbook_matches_sidecar(golden):
    sidecar = json.loads(golden.with_suffix(golden.suffix + ".json").read_text(encoding="utf-8"))
    expected = _decode_sheets(sidecar["sheets"])

    _assert_same_grid(_read_with_xlspy(golden), expected, f"xlspy {golden.name}")
    _assert_same_grid(_read_with_openpyxl(golden), expected, f"openpyxl {golden.name}")


def test_openpyxl_reads_our_writer_output(tmp_path):
    path = _write_xlsx(tmp_path / "writers.xlsx", writer_rows())
    # The xlspy writer stores the exact IEEE double (no 15-digit rounding),
    # so no float overrides apply here.
    expected = {"Main": _decode_sheets(encoded_sheets({"Main": SHEETS["Main"]}, overrides=False))["Main"]}
    _assert_same_grid(_read_with_xlspy(path), expected, "xlspy on our writer")
    _assert_same_grid(_read_with_openpyxl(path), expected, "openpyxl on our writer")


def test_readers_agree_on_our_writer_output(tmp_path):
    path = _write_xlsx(tmp_path / "agree.xlsx", writer_rows())
    _assert_same_grid(_read_with_xlspy(path), _read_with_openpyxl(path), "xlspy vs openpyxl")


@pytest.mark.parametrize("streaming", [False, True], ids=["regular", "streaming"])
def test_openpyxl_reads_xlsx_updater_output(tmp_path, streaming):
    source = _write_xlsx(tmp_path / "update.xlsx", [["name", "when"], ["old", 1]])
    replacement = [["early", dt.date(1900, 1, 1)], ["unicode ⚡", dt.datetime(2024, 1, 15, 10, 30)]]

    updater = XlsxUpdater(source)
    try:
        if streaming:
            updater.replace_sheet_data_stream("Main", iter(replacement), headers=["name", "when"])
        else:
            updater.replace_sheet_data("Main", replacement, headers=["name", "when"])
        updater.save()
    finally:
        updater.close()

    expected = {
        "Main": [
            ["name", "when"],
            ["early", dt.datetime(1900, 1, 1)],
            ["unicode ⚡", dt.datetime(2024, 1, 15, 10, 30)],
        ]
    }
    _assert_same_grid(_read_with_xlspy(source), expected, "xlspy on updater output")
    _assert_same_grid(_read_with_openpyxl(source), expected, "openpyxl on updater output")
