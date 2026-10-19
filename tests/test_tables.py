"""Excel tables: column names must match the header cells.

Excel 16.0 refuses a whole workbook (verified over COM) when a table's header
cell is empty, is not text, differs from the column name (other than in
case), or repeats another column's name, and when the table name reads as a
cell reference. The writers derive names from the header row, or check the
given ones, and raise instead of writing such a file.
"""

import io
import zipfile

import pytest

from xlspy import ExcelReader, ExcelTable, XlsbWriter, XlsxWriter

WRITERS = [pytest.param(XlsxWriter, "xlsx", id="xlsx"), pytest.param(XlsbWriter, "xlsb", id="xlsb")]


def _write(writer_cls, rows, table):
    buffer = io.BytesIO()
    with writer_cls(buffer) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(rows, tables=[table])
    return buffer.getvalue()


def _table_part(data, ext):
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return archive.read(f"xl/tables/table1.{'xml' if ext == 'xlsx' else 'bin'}")


def _contains_name(part, ext, name):
    encoded = name.encode("utf-8") if ext == "xlsx" else name.encode("utf-16-le")
    return encoded in part


@pytest.mark.parametrize("writer_cls,ext", WRITERS)
def test_names_come_from_the_header_cells(writer_cls, ext):
    rows = [["Product", "Units", "Unit price"], ["Tea", 3, 1.5], ["Cocoa", 1, 2.0]]
    part = _table_part(_write(writer_cls, rows, ExcelTable("Sales", "A1:C3")), ext)
    for name in ("Product", "Units", "Unit price"):
        assert _contains_name(part, ext, name)
    assert not _contains_name(part, ext, "Column1")


@pytest.mark.parametrize("writer_cls,ext", WRITERS)
def test_header_row_below_the_first_rows_is_found(tmp_path, writer_cls, ext):
    # The header is far below the 100 rows kept for column widths.
    rows = [[f"filler {i}", i] for i in range(150)] + [["Name", "Qty"], ["a", 1], ["b", 2]]
    data = _write(writer_cls, rows, ExcelTable("Late", "A151:B153"))
    assert _contains_name(_table_part(data, ext), ext, "Qty")
    path = tmp_path / f"late.{ext}"
    path.write_bytes(data)
    with ExcelReader(str(path)) as reader:
        assert list(reader.get_rows("Data"))[150] == ["Name", "Qty"]


@pytest.mark.parametrize("writer_cls,ext", WRITERS)
def test_explicit_names_may_differ_only_in_case(writer_cls, ext):
    rows = [["name", "QTY"], ["a", 1]]
    part = _table_part(_write(writer_cls, rows, ExcelTable("Sales", "A1:B2", columns=["Name", "Qty"])), ext)
    assert _contains_name(part, ext, "Name") and _contains_name(part, ext, "Qty")


@pytest.mark.parametrize("writer_cls,ext", WRITERS)
@pytest.mark.parametrize("header,columns,match", [
    pytest.param(["Name", "Qty"], ["Name", "Amount"], "does not match header cell B1", id="mismatch"),
    pytest.param(["Name", None], None, "header cell B1 must hold text", id="empty"),
    pytest.param(["Name", ""], None, "header cell B1 must hold text", id="empty-string"),
    pytest.param(["Name", 2024], ["Name", "2024"], "header cell B1 must hold text", id="numeric"),
    pytest.param(["Name", "name"], None, "duplicate column name", id="duplicate"),
    pytest.param(["Name"], None, "header cell B1 must hold text", id="short-row"),
])
def test_headers_excel_refuses_raise_and_publish_nothing(writer_cls, ext, header, columns, match):
    buffer = io.BytesIO()
    with pytest.raises(ValueError, match=match):
        with writer_cls(buffer) as writer:
            writer.add_sheet("Data")
            writer.write_sheet([header, ["a", 1]], tables=[ExcelTable("Sales", "A1:B2", columns=columns)])
    assert buffer.getvalue() == b""


@pytest.mark.parametrize("writer_cls,ext", WRITERS)
def test_table_without_header_row_keeps_default_names(writer_cls, ext):
    part = _table_part(_write(writer_cls, [[1, 2], [3, 4]], ExcelTable("Plain", "A1:B2", header_row=False)), ext)
    assert _contains_name(part, ext, "Column1") and _contains_name(part, ext, "Column2")


def test_xlsx_names_are_escaped_like_the_header_cells():
    rows = [["lit _x0041_", "two\nlines", 'quote "q"'], [1, 2, 3]]
    part = _table_part(_write(XlsxWriter, rows, ExcelTable("Esc", "A1:C2")), "xlsx").decode("utf-8")
    assert 'name="lit _x005F_x0041_"' in part
    assert 'name="two&#10;lines"' in part
    assert 'name="quote &quot;q&quot;"' in part


@pytest.mark.parametrize("name", ["T1", "t1", "AB12", "XFD1048576", "R1C1", "RC", "R", "c", "R12", "C3"])
def test_table_names_that_read_as_cell_references_are_refused(name):
    with pytest.raises(ValueError, match="reads as a cell reference"):
        ExcelTable(name, "A1:B2")


@pytest.mark.parametrize("name", ["Tbl", "Sales", "T1A", "XFE1", "A1048577", "Rows1", "Col_2", "_T1"])
def test_ordinary_table_names_are_accepted(name):
    assert ExcelTable(name, "A1:B2").display_name == name
