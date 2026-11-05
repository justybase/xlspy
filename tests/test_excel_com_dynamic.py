"""Excel checks LET, LAMBDA and dynamic-array spill written by xlspy."""

import pytest

from xlspy import Formula, XlsbWriter, XlsxWriter, load_workbook

pytestmark = pytest.mark.com  # skipped unless XLSPY_RUN_EXCEL_COM=1 (conftest.py)

WRITERS = [(XlsxWriter, "xlsx"), (XlsbWriter, "xlsb")]


@pytest.mark.parametrize("writer_class, suffix", WRITERS)
def test_spill_in_excel(tmp_path, open_workbook, writer_class, suffix):
    path = tmp_path / f"spill.{suffix}"
    with writer_class(path) as writer:
        writer.add_sheet("S")
        writer.write_sheet([[Formula("SEQUENCE(3,2)", spill_ref="A1:B3")]])
    sheet = open_workbook(path).Worksheets("S")
    sheet.Application.Calculate()
    assert sheet.Range("A1").HasSpill
    assert sheet.Range("A1").SpillingToRange.Address == "$A$1:$B$3"
    assert [[sheet.Cells(row, column).Value for column in (1, 2)] for row in (1, 2, 3)] == [
        [1.0, 2.0], [3.0, 4.0], [5.0, 6.0],
    ]


def test_let_and_lambda_in_excel(tmp_path, open_workbook):
    path = tmp_path / "let.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("S")
        writer.write_sheet([[
            Formula("LET(x,1,x+1)"),
            Formula("LAMBDA(x,x*2)(3)"),
            Formula("REDUCE(0,A2:A4,LAMBDA(a,b,a+b))"),
        ], [1], [2], [3]])
    sheet = open_workbook(path).Worksheets("S")
    sheet.Application.Calculate()
    assert [sheet.Range(cell).Value for cell in ("A1", "B1", "C1")] == [2.0, 6.0, 6.0]


def test_let_in_xlsb_in_excel(tmp_path, open_workbook):
    path = tmp_path / "let.xlsb"
    with XlsbWriter(path) as writer:
        writer.add_sheet("S")
        writer.write_sheet([[
            Formula("LET(x,1,x+1)"),
            Formula("REDUCE(0,A2:A4,LAMBDA(a,b,a+b))"),
        ], [1], [2], [3]])
    sheet = open_workbook(path).Worksheets("S")
    sheet.Application.Calculate()
    assert [sheet.Range(cell).Value for cell in ("A1", "B1")] == [2.0, 6.0]


@pytest.mark.parametrize("writer_class, suffix", WRITERS)
def test_loaded_spill_survives_an_unrelated_edit(tmp_path, open_workbook, writer_class, suffix):
    source = tmp_path / f"spill.{suffix}"
    with writer_class(source) as writer:
        writer.add_sheet("S")
        writer.write_sheet([[Formula("SEQUENCE(3,2)", spill_ref="A1:B3")]])
    loaded = load_workbook(source, data_only=False)
    try:
        loaded["S"]["D10"] = 1
        out = tmp_path / f"out.{suffix}"
        loaded.save(out)
    finally:
        loaded.close()
    sheet = open_workbook(out).Worksheets("S")
    sheet.Application.Calculate()
    assert sheet.Range("A1").HasSpill
    assert sheet.Range("A1").SpillingToRange.Address == "$A$1:$B$3"
