"""Where a value is in the rows of ``ExcelReader``, and what edits do to a sheet's end.

Each case here was a silent loss found by ``test_edit_sequences_fuzz.py``.
"""

import gc
import re
import zipfile

import pytest

from xlspy import ExcelReader, Hyperlink, XlsbUpdater, XlsbWriter, XlsxUpdater, XlsxWriter, load_workbook

FORMATS = {"xlsx": (XlsxWriter, XlsxUpdater), "xlsb": (XlsbWriter, XlsbUpdater)}
PATHS = [{}, {"streaming": True}, {"data_only": False}]

LAYOUTS = [
    # A title over a table that starts further left.
    ([[None, "title"], [], ["id", "v"], [1, 2]], [[None, "title"], [None, None], ["id", "v"], [1, 2]]),
    # A sheet that starts away from A1.
    ([[], [], [None, None, 5]], [[], [], [None, None, 5]]),
    # A wider row does not widen the rows after it, a narrower one is padded.
    ([[1], [2, 3, 4], [5]], [[1], [2, 3, 4], [5]]),
    ([[1, 2, 3], [4], [None, None, None, 5]], [[1, 2, 3], [4, None, None], [None, None, None, 5]]),
]


def _write(path, rows):
    with FORMATS[path.suffix[1:]][0](str(path)) as writer:
        writer.add_sheet("S")
        writer.write_sheet(rows, auto_filter=False)


def _plain(rows):
    return [[float(v) if isinstance(v, int) else v for v in row] for row in rows]


@pytest.mark.parametrize("ext", FORMATS)
@pytest.mark.parametrize("options", PATHS, ids=["bulk", "streaming", "formulas"])
@pytest.mark.parametrize("rows, expected", LAYOUTS)
def test_rows_keep_their_values_in_place(tmp_path, ext, options, rows, expected):
    path = tmp_path / f"a.{ext}"
    _write(path, rows)
    with ExcelReader(str(path)) as reader:
        assert _plain(reader.get_rows("S", **options)) == _plain(expected)


def test_xlsb_replace_on_a_sheet_without_rows_stays_in_the_sheet(tmp_path):
    path = tmp_path / "a.xlsb"
    _write(path, [])
    updater = XlsbUpdater(str(path))
    updater.replace_sheet_data("S", [[1, "two"]])
    updater.save(str(path))
    updater.close()
    with ExcelReader(str(path)) as reader:
        assert _plain(reader.read_all("S")) == _plain([[1, "two"]])


@pytest.mark.parametrize("ext", FORMATS)
@pytest.mark.parametrize("number", [0.1 + 0.2, 1 / 3, 1234567890123456, 5e-324, 1e22, -0.25])
def test_replace_keeps_every_digit(tmp_path, ext, number):
    path = tmp_path / f"a.{ext}"
    _write(path, [[1]])
    updater = FORMATS[ext][1](str(path))
    updater.replace_sheet_data("S", [[number]])
    updater.save(str(path))
    updater.close()
    with ExcelReader(str(path)) as reader:
        assert reader.read_all("S")[0][0] == number


@pytest.mark.parametrize("ext", FORMATS)
@pytest.mark.parametrize("value", [5, "text", None])
def test_a_value_keeps_the_link_of_its_cell(tmp_path, ext, value):
    path = tmp_path / f"a.{ext}"
    _write(path, [[1]])
    link = Hyperlink("https://a.example/")
    wb = load_workbook(str(path))
    wb["S"]["C2"].hyperlink = link
    wb["S"]["C2"].value = value
    assert wb["S"]["C2"].hyperlink is link
    wb.save(str(path))
    wb.close()
    wb = load_workbook(str(path))
    assert wb["S"]["C2"].hyperlink.target == link.target
    # And one that was already in the file.
    wb["S"]["C2"].value = "after"
    assert wb["S"]["C2"].hyperlink.target == link.target
    wb.save(str(path))
    wb.close()
    wb = load_workbook(str(path))
    assert wb["S"]["C2"].value == "after"
    assert wb["S"]["C2"].hyperlink.target == link.target
    wb.close()


@pytest.mark.parametrize("ext", FORMATS)
def test_append_goes_below_the_last_cell_in_both_formats(tmp_path, ext):
    path = tmp_path / f"a.{ext}"
    _write(path, [[1], [2], [], [], []])
    wb = load_workbook(str(path))
    assert wb["S"].max_row == 2
    wb["S"].append(["new"])
    assert wb["S"]["A3"].value == "new"
    wb.close()


def test_a_stale_dimension_does_not_move_append(tmp_path):
    path = tmp_path / "a.xlsx"
    _write(path, [[1], [2]])
    with zipfile.ZipFile(path) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    sheet = parts["xl/worksheets/sheet1.xml"]
    stale, count = re.subn(rb'<dimension ref="[^"]*"', b'<dimension ref="A1:Z500"', sheet)
    assert count == 1
    parts["xl/worksheets/sheet1.xml"] = stale
    with zipfile.ZipFile(path, "w") as archive:
        for name, body in parts.items():
            archive.writestr(name, body)
    wb = load_workbook(str(path))
    assert (wb["S"].max_row, wb["S"].max_column) == (2, 1)
    wb.close()


@pytest.mark.parametrize("shared", [True, False])
def test_text_keeps_its_spaces_for_excel(tmp_path, shared):
    """Excel drops the spaces at the ends of a text whose <t> has no xml:space="preserve"."""
    texts = [" lead", "trail ", "tab\t", "a\nb", "　wide", "plain"]
    path = tmp_path / "a.xlsx"
    with XlsxWriter(str(path), use_shared_strings=shared) as writer:
        writer.add_sheet("S")
        writer.write_sheet([texts], auto_filter=False)
    with zipfile.ZipFile(path) as archive:
        sheet = archive.read("xl/worksheets/sheet1.xml").decode()
        table = archive.read("xl/sharedStrings.xml").decode() if shared else sheet
    for text in texts[:-1]:
        assert re.search(r'<t xml:space="preserve">' + re.escape(text) + "</t>", table), text
    assert "<t>plain</t>" in table
    with ExcelReader(str(path)) as reader:
        assert reader.read_all("S")[0] == texts


@pytest.mark.com
def test_excel_written_xlsb_reads_like_its_xlsx(excel, tmp_path):
    """Excel writes each row's own cells: a title over a table further left, a cell far right."""
    source = tmp_path / "source.xlsx"
    _write(source, [[None, "title"], [], ["id", "v"], [1, 2], [], [None, None, None, None, "far"]])
    binary = tmp_path / "excel.xlsb"
    book = excel.Workbooks.Open(str(source))
    try:
        book.SaveAs(str(tmp_path / "excel.xlsx"), FileFormat=51)  # xlOpenXMLWorkbook
        book.SaveAs(str(binary), FileFormat=50)  # xlExcel12
    finally:
        book.Close(SaveChanges=False)
        del book
        gc.collect()
    for options in PATHS:
        with ExcelReader(str(binary)) as reader:
            from_xlsb = list(reader.get_rows("S", **options))
        with ExcelReader(str(tmp_path / "excel.xlsx")) as reader:
            from_xlsx = list(reader.get_rows("S", **options))
        assert _plain(from_xlsb) == _plain(from_xlsx)
        assert from_xlsb[0][1] == "title" and from_xlsb[2][0] == "id" and from_xlsb[5][4] == "far"
