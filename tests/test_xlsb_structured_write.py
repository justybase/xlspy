"""Formulas that refer to tables are written to an XLSB as PtgList tokens, the way Excel writes them."""

import json
import pathlib
import posixpath
import struct
import zipfile

import pytest

from xlspy import DefinedName, ExcelTable, UnsupportedFeatureError, Workbook, XlsbWriter, load_workbook
from xlspy._safe_xml import fromstring
from xlspy._xlsb_formulas import FormulaContext, compile_formula, decode_formula, read_parsed
from xlspy._xlsb_metadata import read_wide, records
from xlspy._xlsb_table_part import table_catalog

DATA = pathlib.Path(__file__).parent / "data"
FIXTURE = DATA / "structured_refs.xlsb"
EXPECTED = json.loads((DATA / "structured_refs.json").read_text(encoding="utf-8"))


def _fixture_context():
    """A formula context over the fixture, as the reader of a loaded workbook builds it."""
    archive = zipfile.ZipFile(FIXTURE)

    def read(name):
        return archive.read(name) if name in archive.namelist() else None

    workbook = archive.read("xl/workbook.bin")
    relations = {node.get("Id"): node.get("Target") for node in fromstring(archive.read("xl/_rels/workbook.bin.rels"))}
    sheets = {}
    for record, payload in records(workbook):
        if record.record_id == 0x9C:  # BrtBundleSh: state, sheet id, relationship id, name
            relation, offset = read_wide(payload, 8)
            name, _ = read_wide(payload, offset)
            target = relations[relation].lstrip("/")
            sheets[name] = posixpath.normpath(target if target.startswith("xl/") else "xl/" + target)
    catalog = table_catalog(archive.namelist(), read, sheets)
    archive.close()
    return FormulaContext(list(sheets), workbook, tables=lambda: catalog)


def _ptg_list(compiled):
    """The PtgList tokens of a compiled formula: (ixti, flags, table id, first column, last column)."""
    tokens = read_parsed(compiled)[0]
    found = []
    index = tokens.find(b"\x18\x19")
    while index >= 0:
        found.append(struct.unpack_from("<HHIHH", tokens, index + 2))
        index = tokens.find(b"\x18\x19", index + 2)
    return found


# The PtgList tokens Excel wrote for these formulas in the fixture (ixti, flags, table id, first, last).
EXCEL_TOKENS = [
    ("Data", "=SUM(Tbl)", [(0, 0x0000, 2, 0, 0)]),
    ("Data", "=SUM(Tbl[qty])", [(0, 0x0001, 2, 2, 2)]),
    ("Data", "=SUM(Tbl[[qty]:[my col]])", [(0, 0x0002, 2, 2, 3)]),
    ("Data", "=SUM(Tbl[[#Data],[qty]])", [(0, 0x0011, 2, 2, 2)]),
    ("Data", "=SUM(Tbl[[#Headers],[#Data],[qty]])", [(0, 0x0019, 2, 2, 2)]),
    ("Data", "=SUM(Tbl[[#Totals],[qty]])", [(0, 0x0021, 2, 2, 2)]),
    ("Data", "=SUM(Other[[#Totals],[v]])", [(0, 0x0021, 3, 1, 1)]),
    ("Data", "=Tbl[[#Totals],[qty]]+Other[[#Totals],[v]]*2", [(0, 0x0421, 2, 2, 2), (0, 0x0421, 3, 1, 1)]),
    ("Data", "=Tbl[[#This Row],[qty]]*2+Tbl[[#This Row],[my col]]", [(0, 0x0641, 2, 2, 2), (0, 0x0641, 2, 3, 3)]),
    ("Names", "=SUM(Sp[plain])", [(1, 0x0001, 4, 0, 0)]),
]


@pytest.mark.parametrize(("sheet", "text", "expected"), EXCEL_TOKENS)
def test_the_reference_tokens_are_the_ones_excel_wrote(sheet, text, expected):
    compiled = compile_formula(text, _fixture_context(), sheet=sheet)
    assert _ptg_list(compiled) == expected


@pytest.mark.parametrize("sheet", ["Data", "Names"])
def test_every_formula_of_the_fixture_compiles_back_to_its_text(sheet):
    context = _fixture_context()
    for address, text in EXPECTED[sheet].items():
        compiled = compile_formula(text, context, sheet=sheet)
        assert decode_formula(compiled, context, sheet=sheet) == text.removeprefix("="), address


@pytest.mark.parametrize(("text", "error"), [
    ("=SUM([qty])", UnsupportedFeatureError),  # a bare column needs its table's name
    ("=SUM(Nope[qty])", UnsupportedFeatureError),  # no such table
    ("=SUM(Tbl[nope])", ValueError),  # no such column
    ("=SUM(Tbl[#Bogus])", ValueError),  # no such row keyword
    ("=SUM(Tbl[[qty]:[name]])", ValueError),  # the range runs backwards
    ("=SUM(Tbl[[qty],[name]])", ValueError),  # two columns are not a range
    ("=SUM(Tbl[a'x])", ValueError),  # an escape that escapes nothing
    ("=SUM(Tbl[qty)", UnsupportedFeatureError),  # the bracket is not closed
])
def test_a_reference_that_cannot_be_written_is_refused(text, error):
    with pytest.raises(error):
        compile_formula(text, _fixture_context(), sheet="Data")


def test_a_table_name_alone_and_the_display_form_are_written_as_the_file_form():
    context = _fixture_context()
    assert decode_formula(compile_formula("=Tbl", context, sheet="Data"), context, sheet="Data") == "Tbl"
    assert decode_formula(compile_formula("=SUM(Tbl[@qty])", context, sheet="Data"), context, sheet="Data") == \
        "SUM(Tbl[[#This Row],[qty]])"
    assert decode_formula(compile_formula("=Tbl[@[my col]]", context, sheet="Data"), context, sheet="Data") == \
        "Tbl[[#This Row],[my col]]"


def _sales_workbook(path, *, formula_above_header=False, table_on_later_sheet=False):
    """A sheet with table Tbl and formulas that refer to it; Calc, a later sheet, refers to it too."""
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    if table_on_later_sheet:
        sheet["A1"] = "=SUM(Tbl[qty])"
        later = workbook.create_sheet("Later")
        for offset, row in enumerate([["name", "qty"], ["x", 1]]):
            for column, value in enumerate(row):
                later.cell(row=offset + 1, column=column + 1, value=value)
        later.add_table(ExcelTable("Tbl", "A1:B2"))
    else:
        top = 3 if formula_above_header else 1
        for offset, row in enumerate([["name", "qty"], ["a", 2], ["b", 3]]):
            for column, value in enumerate(row):
                sheet.cell(row=top + offset, column=column + 1, value=value)
        sheet.add_table(ExcelTable("Tbl", f"A{top}:B{top + 2}"))
        if formula_above_header:
            sheet["A1"] = "=SUM(Tbl[qty])"
        else:
            sheet["D1"] = "=SUM(Tbl[qty])"
            sheet["C2"] = "=Tbl[@qty]*2"
    workbook.create_sheet("Calc")["A1"] = "=SUM(Tbl[qty])+1"
    workbook.save(path)
    workbook.close()


def test_a_new_workbook_holds_formulas_that_refer_to_its_own_and_an_earlier_table(tmp_path):
    path = tmp_path / "new.xlsb"
    _sales_workbook(path)
    reloaded = load_workbook(path)
    try:
        assert reloaded["Data"]["D1"].value == "=SUM(Tbl[qty])"
        assert reloaded["Data"]["C2"].value == "=Tbl[[#This Row],[qty]]*2"
        assert reloaded["Calc"]["A1"].value == "=SUM(Tbl[qty])+1"
    finally:
        reloaded.close()


def test_a_formula_above_the_header_of_its_table_is_refused_with_a_reason(tmp_path):
    with pytest.raises(UnsupportedFeatureError, match="header row is not written yet"):
        _sales_workbook(tmp_path / "above.xlsb", formula_above_header=True)


def test_a_formula_cannot_refer_to_a_table_of_a_later_sheet(tmp_path):
    with pytest.raises(UnsupportedFeatureError, match="not in the workbook"):
        _sales_workbook(tmp_path / "later.xlsb", table_on_later_sheet=True)


def test_a_loaded_workbook_keeps_formulas_that_refer_to_its_tables(tmp_path):
    workbook = load_workbook(FIXTURE)
    try:
        workbook["Data"]["Z2"] = "=SUM(Tbl[qty])"
        workbook["Data"]["Z3"] = "=SUM(Other[[#Totals],[v]])*Tbl[[#Totals],[qty]]"
        workbook.save(tmp_path / "out.xlsb")
    finally:
        workbook.close()
    reloaded = load_workbook(tmp_path / "out.xlsb")
    try:
        assert reloaded["Data"]["Z2"].value == "=SUM(Tbl[qty])"
        assert reloaded["Data"]["Z3"].value == "=SUM(Other[[#Totals],[v]])*Tbl[[#Totals],[qty]]"
        assert {a: reloaded["Data"][a].value for a in EXPECTED["Data"]} == EXPECTED["Data"]
    finally:
        reloaded.close()


def test_a_table_that_a_new_formula_refers_to_cannot_be_removed(tmp_path):
    workbook = load_workbook(FIXTURE)
    try:
        workbook["Data"]["Z2"] = "=SUM(Other[v])"
        workbook["Data"].tables.remove("Other")
        with pytest.raises(UnsupportedFeatureError, match="used by a formula"):
            workbook.save(tmp_path / "out.xlsb")
    finally:
        workbook.close()


def test_a_defined_name_cannot_refer_to_a_table(tmp_path):
    # The refusal may come from set_defined_name or from save; either way it is raised
    # inside the block, so the writer closes without publishing.
    with pytest.raises(UnsupportedFeatureError, match="defined name cannot refer to a table"):
        with XlsbWriter(tmp_path / "name.xlsb") as writer:
            writer.write_sheet([["name", "qty"], ["a", 1]], tables=[ExcelTable("Tbl", "A1:B2")])
            writer.set_defined_name(DefinedName("Qty", "Tbl[qty]"))
            writer.save()
    assert not (tmp_path / "name.xlsb").exists()


def test_a_formula_refers_to_a_table_added_in_the_same_loaded_session(tmp_path):
    workbook = load_workbook(FIXTURE)
    try:
        sheet = workbook["Data"]
        sheet["Y1"], sheet["Z1"], sheet["Y2"], sheet["Z2"] = "name", "qty", "a", 1
        sheet.add_table(ExcelTable("New", "Y1:Z2"))
        sheet["AB2"] = "=SUM(New[qty])+SUM(Tbl[qty])"
        workbook.save(tmp_path / "out.xlsb")
    finally:
        workbook.close()
    reloaded = load_workbook(tmp_path / "out.xlsb")
    try:
        assert reloaded["Data"]["AB2"].value == "=SUM(New[qty])+SUM(Tbl[qty])"
        assert [table.displayName for table in reloaded["Data"].tables] == ["Tbl", "Other", "New"]
    finally:
        reloaded.close()
