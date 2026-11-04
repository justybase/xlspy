"""Formulas of an XLSB that refer to tables (PtgList) read as Excel writes them to a file."""

import json
import pathlib

import pytest

from xlspy import ExcelReader, UnsupportedFeatureError, load_workbook
from xlspy._xlsb_formulas import structured_reference

DATA = pathlib.Path(__file__).parent / "data"
EXPECTED = json.loads((DATA / "structured_refs.json").read_text(encoding="utf-8"))
COLUMNS = ["name", "ins", "qty", "my col", "z"]


def test_every_formula_excel_wrote_reads_as_excel_reports_it():
    # tests/data/structured_refs.xlsb is made by tools/make_structured_refs_xlsb.py in Excel.
    workbook = load_workbook(DATA / "structured_refs.xlsb")
    try:
        for sheet_name, cells in EXPECTED.items():
            sheet = workbook[sheet_name]
            assert {address: sheet[address].value for address in cells} == cells
    finally:
        workbook.close()


def test_the_reader_and_the_streaming_rows_agree_with_the_model():
    with ExcelReader(DATA / "structured_refs.xlsb") as reader:
        rows = list(reader.get_rows("Data", data_only=False))
    formulas = [value for row in rows for value in row if isinstance(value, str) and value.startswith("=")]
    assert sorted(formulas) == sorted(EXPECTED["Data"].values())


def test_the_values_are_the_ones_excel_calculated():
    workbook = load_workbook(DATA / "structured_refs.xlsb", data_only=True)
    try:
        sheet = workbook["Data"]
        # =SUM(Tbl[qty]) over qty = 2..5
        assert sheet["AD21"].value == 14
    finally:
        workbook.close()


@pytest.mark.parametrize(("flags", "first", "last", "expected"), [
    (0x0000, 0, 0, "Tbl"),
    (0x0001, 2, 2, "Tbl[qty]"),
    (0x0002, 2, 3, "Tbl[[qty]:[my col]]"),
    (0x0004, 0, 0, "Tbl[#All]"),
    (0x0011, 2, 2, "Tbl[[#Data],[qty]]"),
    (0x0012, 1, 3, "Tbl[[#Data],[ins]:[my col]]"),
    (0x0019, 2, 2, "Tbl[[#Headers],[#Data],[qty]]"),
    (0x0030, 0, 0, "Tbl[[#Data],[#Totals]]"),
    (0x0640, 0, 0, "Tbl[#This Row]"),
    (0x0641, 2, 2, "Tbl[[#This Row],[qty]]"),
    (0x0642, 2, 3, "Tbl[[#This Row],[qty]:[my col]]"),
    (0x0021, 4, 4, "Tbl[[#Totals],[z]]"),
])
def test_the_flags_of_a_reference(flags, first, last, expected):
    assert structured_reference("Tbl", COLUMNS, flags, first, last) == expected


@pytest.mark.parametrize(("name", "expected"), [
    ("a[b]", "Tbl[a'[b']]"), ("a#b", "Tbl[a'#b]"), ("it's", "Tbl[it''s]"), ("a,b", "Tbl[a,b]"),
    (" lead", "Tbl[[ lead]]"), ("trail ", "Tbl[[trail ]]"), ("two words", "Tbl[two words]"),
])
def test_column_names_with_special_characters(name, expected):
    assert structured_reference("Tbl", [name], 0x0001, 0, 0) == expected


@pytest.mark.parametrize(("flags", "first", "last"), [(0x0003, 0, 0), (0x0001, 7, 7), (0x0002, 3, 1)])
def test_a_reference_to_a_column_that_is_not_there_is_refused(flags, first, last):
    with pytest.raises(ValueError, match="column|reference"):
        structured_reference("Tbl", COLUMNS, flags, first, last)


def test_a_table_that_formulas_use_cannot_be_removed(tmp_path):
    workbook = load_workbook(DATA / "structured_refs.xlsb")
    try:
        workbook["Data"].tables.remove("Other")
        with pytest.raises(UnsupportedFeatureError, match="used by a formula"):
            workbook.save(tmp_path / "out.xlsb")
    finally:
        workbook.close()
    assert not (tmp_path / "out.xlsb").exists()


def test_a_table_can_be_renamed_because_formulas_name_it_by_id(tmp_path):
    workbook = load_workbook(DATA / "structured_refs.xlsb")
    try:
        table = workbook["Data"].tables["Other"]
        table.displayName = table.name = "Renamed"
        workbook.save(tmp_path / "out.xlsb")
    finally:
        workbook.close()
    reloaded = load_workbook(tmp_path / "out.xlsb")
    try:
        assert reloaded["Data"]["AF20"].value == "=Tbl[[#Totals],[qty]]+Renamed[[#Totals],[v]]*2"
    finally:
        reloaded.close()


def test_an_edit_of_another_cell_keeps_the_formulas(tmp_path):
    workbook = load_workbook(DATA / "structured_refs.xlsb")
    try:
        workbook["Data"]["B2"] = 99
        workbook.save(tmp_path / "out.xlsb")
    finally:
        workbook.close()
    reloaded = load_workbook(tmp_path / "out.xlsb")
    try:
        assert {a: reloaded["Data"][a].value for a in EXPECTED["Data"]} == EXPECTED["Data"]
    finally:
        reloaded.close()


def test_a_new_reference_to_a_table_is_refused_with_a_reason(tmp_path):
    workbook = load_workbook(DATA / "structured_refs.xlsb")
    try:
        workbook["Data"]["Z2"] = "=SUM(Tbl[qty])"
        with pytest.raises(UnsupportedFeatureError, match="cannot be written yet"):
            workbook.save(tmp_path / "out.xlsb")
    finally:
        workbook.close()


def test_values_load_without_reading_formulas():
    workbook = load_workbook(DATA / "structured_refs.xlsb", data_only=True)
    try:
        assert workbook["Data"]["E2"].value == 2 * 2 + 4
    finally:
        workbook.close()
