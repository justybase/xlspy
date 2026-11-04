"""Excel reads the comment of a defined name that xlspy wrote, and xlspy reads the one Excel wrote."""

import pytest

from xlspy import DefinedName, XlsbWriter, XlsxWriter, load_workbook

pytestmark = pytest.mark.com  # skipped unless XLSPY_RUN_EXCEL_COM=1 (conftest.py)

XL_XLSB = 50


@pytest.mark.parametrize("suffix", ["xlsx", "xlsb"])
def test_excel_reads_the_comment_xlspy_wrote(tmp_path, open_workbook, suffix):
    path = tmp_path / f"names.{suffix}"
    with (XlsxWriter if suffix == "xlsx" else XlsbWriter)(path) as writer:
        writer.set_defined_name(DefinedName("Rate", "Data!$A$1", comment="the rate, <as> agreed"))
        writer.add_sheet("Data")
        writer.write_sheet([[1]])
    book = open_workbook(path)
    assert book.Names("Rate").Comment == "the rate, <as> agreed"


def test_xlspy_reads_the_comment_excel_wrote_to_an_xlsb(tmp_path, excel):
    path = tmp_path / "excel.xlsb"
    book = excel.Workbooks.Add()
    try:
        book.Worksheets(1).Name = "Data"
        name = book.Names.Add("Rate", "=Data!$A$1")
        name.Comment = "written by Excel"
        book.SaveAs(str(path), XL_XLSB)
    finally:
        book.Close(False)
    loaded = load_workbook(path)
    try:
        assert loaded.defined_names["Rate"].comment == "written by Excel"
    finally:
        loaded.close()
