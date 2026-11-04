"""The table part of an XLSB table with a totals row, as Excel reads it."""

import zipfile

import pytest

from xlspy import ExcelTable, Formula, Workbook
from xlspy._xlsb_formulas import FormulaContext, decode_formula
from xlspy._xlsb_metadata import read_wide
from xlspy.biff_utils import iter_record_payloads
from xlspy.rich_text import CellRichText, InlineFont, TextBlock


def _make(path, value):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    for row, values in enumerate([["name", "qty"], ["a", 2], ["b", 3], ["Total", value]], start=1):
        for column, cell in enumerate(values, start=1):
            sheet.cell(row=row, column=column, value=cell)
    sheet.add_table(ExcelTable("Tbl", "A1:B4", totals_row=True))
    workbook.save(path)
    workbook.close()


def _table(path):
    with zipfile.ZipFile(path) as archive:
        data = archive.read("xl/tables/table1.bin")
    return list(iter_record_payloads(data))


def _column_label(payload):
    _name, offset = read_wide(payload, 24, nullable=True)
    _caption, offset = read_wide(payload, offset, nullable=True)
    label, _offset = read_wide(payload, offset, nullable=True)
    return label


@pytest.mark.parametrize(
    ("value", "ilta", "label"),
    [
        (5, 6, None),
        ("5", 0, "5"),
        (None, 0, None),
        (Formula("SUM(B2:B3)", 5), 9, None),
        (CellRichText("To", TextBlock(InlineFont(b=True), "tal")), 0, "Total"),
    ],
)
def test_totals_metadata_follows_the_cells(tmp_path, value, ilta, label):
    path = tmp_path / "totals.xlsb"
    _make(path, value)
    records = _table(path)

    begin = next(payload for record_id, payload in records if record_id == 0x157)
    assert int.from_bytes(begin[28:32], "little") == 1
    assert int.from_bytes(begin[32:36], "little") & 1

    auto_filter = next(payload for record_id, payload in records if record_id == 0xA1)
    assert tuple(int.from_bytes(auto_filter[i:i + 4], "little") for i in range(0, 16, 4)) == (0, 2, 0, 1)

    columns = [payload for record_id, payload in records if record_id == 0x15B]
    assert int.from_bytes(columns[0][4:8], "little") == 0
    assert _column_label(columns[0]) == "Total"
    assert int.from_bytes(columns[1][4:8], "little") == ilta
    assert _column_label(columns[1]) == label

    formulas = [payload for record_id, payload in records if record_id == 0x160]
    if ilta == 9:
        assert len(formulas) == 1
        assert formulas[0][0] == 0
        assert decode_formula(
            formulas[0][1:], FormulaContext(["Data"]), sheet="Data", origin=(3, 1),
        ) == "SUM(B2:B3)"
    else:
        assert formulas == []


def test_a_totals_table_needs_a_data_row():
    with pytest.raises(ValueError, match="at least one data row"):
        ExcelTable("Bad", "A1:B2", totals_row=True)


def test_explicit_totals_functions_are_written(tmp_path):
    path = tmp_path / "totals.xlsb"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    for row, values in enumerate([["name", "qty"], ["a", 2], ["b", 3], ["Total", 5]], start=1):
        for column, cell in enumerate(values, start=1):
            sheet.cell(row=row, column=column, value=cell)
    sheet.add_table(ExcelTable(
        "Tbl", "A1:B4", totals_row=True, totals_row_functions=("label", "average"),
    ))
    workbook.save(path)
    workbook.close()

    columns = [payload for record_id, payload in _table(path) if record_id == 0x15B]
    assert int.from_bytes(columns[0][4:8], "little") == 0
    assert _column_label(columns[0]) == "Total"
    assert int.from_bytes(columns[1][4:8], "little") == 1
    assert _column_label(columns[1]) is None


def test_a_headerless_table_has_no_auto_filter(tmp_path):
    path = tmp_path / "headerless.xlsb"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    for row, values in enumerate([["a", 2], ["b", 3], ["Total", 5]], start=1):
        for column, cell in enumerate(values, start=1):
            sheet.cell(row=row, column=column, value=cell)
    sheet.add_table(ExcelTable("Tbl", "A1:B3", header_row=False, totals_row=True))
    workbook.save(path)
    workbook.close()

    records = _table(path)
    begin = next(payload for record_id, payload in records if record_id == 0x157)
    assert int.from_bytes(begin[28:32], "little") == 1
    assert int.from_bytes(begin[32:36], "little") & 1
    assert not any(record_id in (0xA1, 0xA2) for record_id, _payload in records)


def test_an_explicit_totals_function_must_fit_the_cell(tmp_path):
    path = tmp_path / "bad.xlsb"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    for row, values in enumerate([["name", "qty"], ["a", 2], ["b", 3], ["Total", 5]], start=1):
        for column, cell in enumerate(values, start=1):
            sheet.cell(row=row, column=column, value=cell)
    sheet.add_table(ExcelTable(
        "Tbl", "A1:B4", totals_row=True, totals_row_functions=("label", "label"),
    ))
    with pytest.raises(ValueError, match="label totals cell"):
        workbook.save(path)
    workbook.close()


def test_totals_row_functions_are_validated():
    with pytest.raises(ValueError, match="len\\(totals_row_functions\\)"):
        ExcelTable("Bad", "A1:B4", totals_row=True, totals_row_functions=("sum",))
    with pytest.raises(ValueError, match="totals_row=True"):
        ExcelTable("Bad", "A1:B4", totals_row_functions=("sum", "sum"))
    with pytest.raises(ValueError, match="invalid totals function"):
        ExcelTable("Bad", "A1:B4", totals_row=True, totals_row_functions=("sum", "nope"))
