"""Functions Excel 2007 made built-in (Analysis ToolPak) and the newer _xlfn ones, in XLSB and XLSX."""

import zipfile

import pytest

from xlspy import Workbook, load_workbook
from xlspy._xlsb_formulas import BY_ID, FUNCTIONS, _VOLATILE_FUNCTIONS

# (formula, Excel's function id): ids as Excel writes them in a saved .xlsb.
BUILT_IN = [
    ("=EOMONTH(A1,1)", 450),
    ("=WORKDAY(A1,5)", 471),
    ("=WORKDAY(A1,5,B1:B3)", 471),
    ("=NETWORKDAYS(A1,B1)", 472),
    ("=GCD(A1,6,9)", 473),
    ("=MROUND(A1,5)", 422),
    ("=ISEVEN(A1)", 420),
    ("=XIRR(A1:A3,B1:B3)", 429),
    ("=CUMIPMT(0.05/12,60,10000,1,12,0)", None),
    ("=RANDBETWEEN(1,6)", 464),
    ("=COMPLEX(1,2)", None),
    ("=YEARFRAC(A1,B1,1)", None),
]


@pytest.mark.parametrize("formula,function_id", BUILT_IN)
def test_built_in_function_round_trips_in_xlsb(tmp_path, formula, function_id):
    workbook = Workbook()
    workbook.active["D1"] = formula
    path = tmp_path / "f.xlsb"
    workbook.save(path)
    workbook.close()
    loaded = load_workbook(path)
    try:
        assert loaded.active["D1"].value == formula
    finally:
        loaded.close()
    if function_id is not None:
        assert BY_ID[function_id][0] == formula[1:].split("(")[0]


def test_function_table_holds_the_toolpak_functions():
    for name in ("EOMONTH", "WORKDAY", "NETWORKDAYS", "WEEKNUM", "CONVERT", "DEC2BIN", "IMSUM", "PRICE"):
        assert name in FUNCTIONS
    assert "RANDBETWEEN" in _VOLATILE_FUNCTIONS
    assert FUNCTIONS["EOMONTH"][:4] == (450, 2, 2, 2)


def test_wrong_argument_count_is_refused(tmp_path):
    workbook = Workbook()
    workbook.active["A1"] = "=EOMONTH(A1)"
    with pytest.raises(ValueError, match="expects"):
        workbook.save(tmp_path / "f.xlsb")
    workbook.close()


@pytest.mark.parametrize("name", ["BITAND", "NUMBERVALUE", "UNICHAR", "WORKDAY.INTL", "RRI", "ISFORMULA"])
def test_newer_functions_get_the_xlfn_prefix_in_xlsx(tmp_path, name):
    workbook = Workbook()
    workbook.active["A1"] = f"={name}(1,2)"
    path = tmp_path / "f.xlsx"
    workbook.save(path)
    workbook.close()
    with zipfile.ZipFile(path) as archive:
        sheet = archive.read("xl/worksheets/sheet1.xml").decode()
    assert f"_xlfn.{name}(1,2)" in sheet


def test_newer_functions_are_written_to_xlsb(tmp_path):
    workbook = Workbook()
    workbook.active["A1"] = "=BITAND(12,10)"
    workbook.active["A2"] = "=NUMBERVALUE(\"1.5\",\".\")"
    path = tmp_path / "f.xlsb"
    workbook.save(path)
    workbook.close()
    loaded = load_workbook(path)
    try:
        assert loaded.active["A1"].value == "=_xlfn.BITAND(12,10)"
        assert loaded.active["A2"].value == "=_xlfn.NUMBERVALUE(\"1.5\",\".\")"
    finally:
        loaded.close()


EXTERNAL = __import__("pathlib").Path(__file__).parent / "data" / "external_ref.xlsb"  # made by Excel


def test_formulas_with_external_workbook_references_load():
    loaded = load_workbook(EXTERNAL)
    try:
        sheet = loaded.active
        assert sheet["A1"].value == "='[1]Data'!$A$1"
        assert sheet["A2"].value == "=SUM('[1]Data'!$A$1:$A$3)"
    finally:
        loaded.close()


def test_editing_a_workbook_with_external_references_keeps_them(tmp_path):
    out = tmp_path / "out.xlsb"
    loaded = load_workbook(EXTERNAL)
    try:
        loaded.active["A3"] = 5
        loaded.save(out)
    finally:
        loaded.close()
    with zipfile.ZipFile(out) as archive:
        assert "xl/externalLinks/externalLink1.bin" in archive.namelist()
    reloaded = load_workbook(out)
    try:
        assert reloaded.active["A1"].value == "='[1]Data'!$A$1" and reloaded.active["A3"].value == 5
    finally:
        reloaded.close()


def test_streaming_reader_reads_the_cached_value_of_an_external_reference():
    from xlspy import ExcelReader

    with ExcelReader(str(EXTERNAL)) as reader:
        rows = list(reader.get_rows(reader.get_sheet_names()[0]))
    assert rows[0][0] == 7


def test_a_formula_that_points_at_a_missing_group_is_an_invalid_file():
    from xlspy import InvalidFileError
    from xlspy._xlsb_read import formula_text

    # PtgExp (0x01) naming a shared formula anchored at A1 that the sheet never defined.
    ptg_exp = bytes((0x01,)) + (0).to_bytes(4, "little")
    formula = len(ptg_exp).to_bytes(4, "little") + ptg_exp + (4).to_bytes(4, "little") + (0).to_bytes(4, "little")
    with pytest.raises(InvalidFileError, match="group definition was not found"):
        formula_text(formula, {}, None, "S", 0, 0)
