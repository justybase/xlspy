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
from helpers import part_text
from xlspy import Formula
from xlspy import Hyperlink
from xlspy import XlsbUpdater
from xlspy import XlsxUpdater
from xlspy import load_workbook
from xlspy.exceptions import UnsupportedFeatureError

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


openpyxl = pytest.importorskip("openpyxl")


def _counting_rows(rows, consumed):
    for row in rows:
        consumed.append(row)
        yield row


@pytest.mark.parametrize("writer_cls,suffix", [(XlsxWriter, "xlsx"), (XlsbWriter, "xlsb")])
def test_bad_table_header_fails_before_the_rest_is_read(tmp_path, writer_cls, suffix):
    consumed = []
    rows = _counting_rows([["Name", 7]] + [[f"r{i}", i] for i in range(10000)], consumed)
    # The failed sheet aborts the writer, so leaving the block raises too.
    with pytest.raises(RuntimeError, match="previous write failed"):
        with writer_cls(tmp_path / f"bad.{suffix}") as writer:
            with pytest.raises(ValueError, match="header cell B1 must hold text"):
                writer.write_sheet(rows, tables=[ExcelTable("Tbl", "A1:B10001")])
    assert len(consumed) == 1


@pytest.mark.parametrize("writer_cls,suffix", [(XlsxWriter, "xlsx"), (XlsbWriter, "xlsb")])
def test_mismatched_table_column_name_fails_at_header_row(tmp_path, writer_cls, suffix):
    consumed = []
    rows = _counting_rows([["intro"], ["Name", "Qty"]] + [["x", 1]] * 500, consumed)
    with pytest.raises(RuntimeError, match="previous write failed"):
        with writer_cls(tmp_path / f"bad.{suffix}") as writer:
            with pytest.raises(ValueError, match="does not match"):
                writer.write_sheet(rows, tables=[ExcelTable("Tbl", "A2:B502", columns=["Name", "Amount"])])
    assert len(consumed) == 2


def test_hyperlink_header_cell_names_the_table_column(tmp_path):
    path = tmp_path / "link_header.xlsx"
    with XlsxWriter(path) as writer:
        writer.write_sheet(
            [[Hyperlink("https://example.com", text="Site"), "Visits"], ["a", 1], ["b", 2]],
            tables=[ExcelTable("Traffic", "A1:B3")],
        )

    table = openpyxl.load_workbook(path).active.tables["Traffic"]
    assert [column.name for column in table.tableColumns] == ["Site", "Visits"]


def test_formula_header_cell_is_still_rejected(tmp_path):
    with pytest.raises(RuntimeError, match="previous write failed"):
        with XlsxWriter(tmp_path / "f.xlsx") as writer:
            with pytest.raises(ValueError, match="must hold text"):
                writer.write_sheet([[Formula('"A"'), "B"], [1, 2]], tables=[ExcelTable("Tbl", "A1:B2")])


def test_table_over_the_whole_data_gets_no_sheet_autofilter(tmp_path):
    path = tmp_path / "table.xlsx"
    with XlsxWriter(path) as writer:
        writer.write_sheet(
            [["k", "v"], ["a", 1], ["b", 2]],
            tables=[ExcelTable("Full", "A1:B3")],
            auto_filter=True,
        )

    assert "<autoFilter" not in part_text(path, "xl/worksheets/sheet1.xml")
    assert "<autoFilter" in part_text(path, "xl/tables/table1.xml")


def parts(path):
    with zipfile.ZipFile(path) as zf:
        return {name: zf.read(name) for name in zf.namelist()}


@pytest.mark.parametrize("writer", [XlsxWriter, XlsbWriter])
def test_table_names_differing_only_in_case_are_rejected(tmp_path, writer):
    rows = [["h", "h2", None, "d1", "d2"], [1, 2, None, 3, 4], [5, 6, None, 7, 8]]
    with pytest.raises(ValueError, match="Excel ignores case"):
        with writer(tmp_path / "tables") as wb:
            wb.add_sheet("S")
            wb.write_sheet(
                rows,
                tables=[ExcelTable("Sales", "A1:B3"), ExcelTable("sales", "D1:E3")],
            )


FORMATS = pytest.mark.parametrize("suffix", ["xlsx", "xlsb"])


STREAMING = pytest.mark.parametrize("streaming", [False, True])


TABLE_ROWS = [["Name", "Qty"], ["a", 1], ["b", 2]]


def _table_base(tmp_path, suffix, rows=TABLE_ROWS, ref="A1:B3"):
    base = tmp_path / f"base.{suffix}"
    writer = XlsxWriter if suffix == "xlsx" else XlsbWriter
    with writer(base) as wb:
        wb.add_sheet("S")
        wb.write_sheet(rows, tables=[ExcelTable("Tbl", ref)])
    return base


def _updater(path, suffix):
    return XlsxUpdater(path) if suffix == "xlsx" else XlsbUpdater(path)


def _replace(updater, rows, streaming, headers=None):
    if streaming:
        updater.replace_sheet_data_stream("S", rows, headers=headers)
    else:
        updater.replace_sheet_data("S", rows, headers=headers)


@FORMATS
@STREAMING
def test_table_header_mismatch_is_refused_and_package_unchanged(tmp_path, suffix, streaming):
    base = _table_base(tmp_path, suffix)
    before = parts(base)
    updater = _updater(base, suffix)
    with pytest.raises(UnsupportedFeatureError, match="header cell A1"):
        _replace(updater, [["Nome", "Qty"], ["c", 3]], streaming)
    out = tmp_path / f"out.{suffix}"
    updater.save(out)
    assert parts(out) == before


@FORMATS
@STREAMING
def test_table_header_mismatch_with_headers_argument_is_refused(tmp_path, suffix, streaming):
    base = _table_base(tmp_path, suffix)
    updater = _updater(base, suffix)
    with pytest.raises(UnsupportedFeatureError, match="header cell B1"):
        _replace(updater, [["c", 3]], streaming, headers=["Name", "Amount"])


@FORMATS
@STREAMING
def test_table_header_mismatch_in_the_middle_of_the_data_is_refused(tmp_path, suffix, streaming):
    # The table's header is row 4; the new data changes it on the way down.
    rows = [["intro", None], [None, None], [None, None], ["Name", "Qty"], ["c", 3]]
    base = _table_base(tmp_path, suffix, rows=rows, ref="A4:B5")
    updater = _updater(base, suffix)
    new_rows = [["x"], ["y"], ["z"], ["Name", "Other"], ["c", 3]]
    with pytest.raises(UnsupportedFeatureError, match="header cell B4"):
        _replace(updater, new_rows, streaming)


@FORMATS
@STREAMING
def test_data_too_short_to_reach_the_table_header_is_refused(tmp_path, suffix, streaming):
    rows = [["intro", None], [None, None], [None, None], ["Name", "Qty"], ["c", 3]]
    base = _table_base(tmp_path, suffix, rows=rows, ref="A4:B5")
    updater = _updater(base, suffix)
    with pytest.raises(UnsupportedFeatureError, match="header cell A4"):
        _replace(updater, [["only", 1]], streaming)


@FORMATS
@STREAMING
def test_matching_headers_are_accepted_ignoring_case_and_length(tmp_path, suffix, streaming):
    base = _table_base(tmp_path, suffix)
    out = tmp_path / f"out.{suffix}"
    updater = _updater(base, suffix)
    # Excel only needs the header cells to match; a shorter range is not corruption.
    _replace(updater, [["NAME", "qty"], ["c", 3]], streaming)
    updater.save(out)

    sheet = load_workbook(out)["S"]
    assert [sheet.cell(row=2, column=c).value for c in (1, 2)] == ["c", 3]


# -- ws.tables: openpyxl's name-keyed operations ---------------------------------

def _sheet_with_tables(tmp_path):
    from xlspy import Workbook
    from xlspy.workbook import Table

    wb = Workbook()
    ws = wb.active
    for row in [["Name", "Qty", "", "Code"], ["a", 1, None, "x"], ["b", 2, None, "y"]]:
        ws.append(row)
    ws.add_table(Table(displayName="Sales", ref="A1:B3"))
    ws.add_table(ExcelTable("Codes", "D1:D3"))
    path = tmp_path / "tables.xlsx"
    wb.save(path)
    return path


def test_tables_work_like_openpyxls_name_keyed_list(tmp_path):
    wb = load_workbook(_sheet_with_tables(tmp_path))
    tables = wb.active.tables
    assert tables.keys() == ["Sales", "Codes"]
    assert tables.items() == [("Sales", "A1:B3"), ("Codes", "D1:D3")]
    assert tables["Codes"] is tables.values()[1] is tables.get("Codes")
    assert tables.get("Missing") is None and "Missing" not in tables
    assert tables[0] is tables["Sales"]  # an int index still works like a list
    with pytest.raises(KeyError):
        del tables["Missing"]
    wb.close()


def test_writer_tables_are_found_by_name():
    from xlspy import Workbook

    ws = Workbook().active
    ws.append(["Code"])
    ws.add_table(ExcelTable("Codes", "A1:A2"))
    assert "Codes" in ws.tables and ws.tables.keys() == ["Codes"]
    ws.tables.remove("Codes")
    assert not ws.tables


def test_deleting_a_table_by_name_is_saved(tmp_path):
    wb = load_workbook(_sheet_with_tables(tmp_path))
    del wb.active.tables["Sales"]
    out = tmp_path / "out.xlsx"
    wb.save(out)
    wb.close()
    wb = load_workbook(out)
    assert wb.active.tables.keys() == ["Codes"]
    wb.close()
