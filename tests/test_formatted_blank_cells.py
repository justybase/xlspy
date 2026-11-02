"""A ``(None, format)`` or ``(None, Style)`` row entry is a formatted cell without a value.

Both writers used to fall through to ``str(None)`` and store the text "None". The cell is now
written like a blank with a range style: no value, the format or style kept.
"""
import zipfile

import pytest

from xlspy import Style, XlsbWriter, XlsxWriter, load_workbook

openpyxl = pytest.importorskip("openpyxl")

ROWS = [
    ["head", "x", "y", "z"],
    [1, (None, "0.00"), (None, Style(bold=True)), (None, Style(fill="#FFFF00", number_format="0%"))],
]


def _write(path, writer_class, rows=ROWS, **options):
    with writer_class(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(rows, **options)


@pytest.fixture(params=[XlsxWriter, XlsbWriter], ids=["xlsx", "xlsb"])
def written(request, tmp_path):
    extension = "xlsx" if request.param is XlsxWriter else "xlsb"
    path = tmp_path / f"blank.{extension}"
    _write(path, request.param)
    return path


def test_no_value_and_the_format_is_kept(written):
    workbook = load_workbook(written)
    try:
        sheet = workbook["Data"]
        assert [cell.value for cell in sheet[2]] == [1, None, None, None]
        assert sheet["B2"].number_format == "0.00"
        assert sheet["C2"].font.bold
        assert sheet["D2"].fill.start_color.rgb == "FFFFFF00" and sheet["D2"].number_format == "0%"
    finally:
        workbook.close()


def test_the_text_none_is_not_stored(written):
    # XLSX text is UTF-8, XLSB text UTF-16-LE; neither part may hold "None".
    with zipfile.ZipFile(written) as archive:
        for name in archive.namelist():
            if name.startswith(("xl/sharedStrings", "xl/worksheets/sheet")):
                data = archive.read(name)
                assert b"None" not in data and "None".encode("utf-16-le") not in data, name


def test_openpyxl_reads_a_styled_blank(tmp_path):
    path = tmp_path / "blank.xlsx"
    _write(path, XlsxWriter)
    workbook = openpyxl.load_workbook(path)
    try:
        sheet = workbook["Data"]
        assert sheet["B2"].value is None and sheet["B2"].number_format == "0.00"
        assert sheet["C2"].value is None and sheet["C2"].font.b
        assert sheet.max_column == 4
    finally:
        workbook.close()


@pytest.mark.parametrize("use_shared_strings", [True, False])
def test_xlsb_inline_and_shared_strings(tmp_path, use_shared_strings):
    path = tmp_path / "blank.xlsb"
    with XlsbWriter(path, use_shared_strings=use_shared_strings) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([["a", (None, "0.00"), "b"]])
    workbook = load_workbook(path)
    try:
        assert [cell.value for cell in workbook["Data"][1]] == ["a", None, "b"]
        assert workbook["Data"]["B1"].number_format == "0.00"
    finally:
        workbook.close()


def test_a_trailing_formatted_blank_counts_in_the_xlsb_dimension(tmp_path):
    from test_xlsb_interop import sheet_dimension

    path = tmp_path / "trailing.xlsb"
    _write(path, XlsbWriter, [[None, "a", (None, Style(bold=True))]])
    # As Excel counts a styled blank: B1:C1.
    assert sheet_dimension(path) == (0, 0, 1, 2)


@pytest.mark.com
@pytest.mark.parametrize("writer_class", [XlsxWriter, XlsbWriter], ids=["xlsx", "xlsb"])
def test_excel_sees_an_empty_formatted_cell(tmp_path, open_workbook, writer_class):
    extension = "xlsx" if writer_class is XlsxWriter else "xlsb"
    path = tmp_path / f"blank_com.{extension}"
    _write(path, writer_class)
    sheet = open_workbook(path).Worksheets("Data")
    for ref in ("B2", "C2", "D2"):
        assert sheet.Range(ref).Value is None and sheet.Range(ref).Formula == "", ref
    assert sheet.Range("B2").NumberFormat == "0.00"
    assert sheet.Range("C2").Font.Bold
    assert sheet.Range("D2").Interior.Color == 0x00FFFF  # BGR yellow
    assert sheet.UsedRange.Address == "$A$1:$D$2"
