"""Pivot tables written with ``XlsbWriter``: the parts, the records Excel reads
and the pre-rendered report.

The record layout was mapped from pivot tables Excel saved as XLSB; the same
pivots are checked in Excel by ``test_excel_com_pivots.py``.
"""

import io
import struct
import zipfile

import openpyxl
import pytest

from xlspy import ExcelTable, PivotTable, PivotValue, XlsbWriter, XlsxWriter, load_workbook
from xlspy.biff_utils import iter_records

ROWS = [
    ["Region", "Product", "Quarter", "Sales", "Units"],
    ["North", "Coffee", 1, 1200, 40],
    ["North", "Tea", 1, 800, 50],
    ["South", "Coffee", 1, 1500, 55],
    ["South", "Tea", 2, 700, 30],
    ["East", "Coffee", 2, 900, 33],
    ["East", "Tea", 2, 1100, 60],
    ["North", "Coffee", 2, 1300, 44],
    ["South", "Coffee", 2, 1600, 58],
]


def pivot(values="Sales", *, location="A12", **kwargs):
    kwargs.setdefault("rows", "Region")
    return PivotTable(source="SalesTable", values=values, location=location, **kwargs)


def make(writer, pivots, *, rows=ROWS, ref="A1:E9", summary=False):
    out = io.BytesIO()
    with writer(out) as w:
        w.add_sheet("Data")
        w.write_sheet(rows, tables=[ExcelTable("SalesTable", ref)], pivots=pivots)
        if summary:
            w.add_sheet("Summary")
            w.write_sheet([])
    return out.getvalue()


def xlsb_cells(data, tmp_path, sheet, ref):
    path = tmp_path / "book.xlsb"
    path.write_bytes(data)
    book = load_workbook(path)
    try:
        return [[cell.value for cell in row] for row in book[sheet][ref]]
    finally:
        book.close()


def xlsx_cells(data, sheet, ref):
    book = openpyxl.load_workbook(io.BytesIO(data))
    return [[cell.value for cell in row] for row in book[sheet][ref]]


def part_records(data, name):
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        part = zf.read(name)
    return [(r.record_id, part[r.data_start:r.data_end]) for r in iter_records(part)]


def records_of(data, name, record_id):
    return [payload for rid, payload in part_records(data, name) if rid == record_id]


TABLE = "xl/pivotTables/pivotTable1.bin"
CACHE = "xl/pivotCache/pivotCacheDefinition1.bin"

SHAPES = {
    "row_one": dict(),
    "row_two": dict(values=["Sales", "Units"]),
    "row_col": dict(columns="Quarter"),
    "row_col_two": dict(columns="Quarter", values=["Sales", "Units"]),
    "text_column": dict(columns="Product"),
    "filter": dict(filters="Product"),
    "average": dict(values=[PivotValue("Units", "average")]),
    "count": dict(values=[PivotValue("Sales", "count")]),
    "max": dict(values=[PivotValue("Sales", "max", number_format="#,##0")]),
}
RANGES = {
    "row_one": "A12:B16", "row_two": "A12:C16", "row_col": "A12:D17",
    "row_col_two": "A12:G18", "text_column": "A12:D17", "filter": "A10:B16",
    "average": "A12:B16", "count": "A12:B16", "max": "A12:B16",
}


@pytest.mark.parametrize("shape", sorted(SHAPES))
def test_rendered_report_is_the_same_as_in_xlsx(tmp_path, shape):
    pivots = [pivot(**SHAPES[shape])]
    expected = xlsx_cells(make(XlsxWriter, pivots), "Data", RANGES[shape])
    assert xlsb_cells(make(XlsbWriter, pivots), tmp_path, "Data", RANGES[shape]) == expected


def test_pivot_on_another_sheet_renders_there(tmp_path):
    out = io.BytesIO()
    with XlsbWriter(out) as w:
        w.add_sheet("Data")
        w.write_sheet(
            ROWS, tables=[ExcelTable("SalesTable", "A1:E9")],
            pivots=[pivot(columns="Quarter", location="B2", sheet="Summary")],
        )
        w.add_sheet("Summary")
        w.write_sheet([])
    cells = xlsb_cells(out.getvalue(), tmp_path, "Summary", "B2:E7")
    assert cells[0][0] == "Sum of Sales"
    assert cells[4] == ["South", 1500, 2300, 3800]
    assert cells[5] == ["Grand Total", 3500, 5600, 9100]


def test_pivot_parts_are_binary_and_linked():
    data = make(XlsbWriter, [pivot(), pivot(["Units"], location="A20", name="Units")])
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        names = set(zf.namelist())
        content_types = zf.read("[Content_Types].xml").decode()
        workbook_rels = zf.read("xl/_rels/workbook.bin.rels").decode()
        sheet_rels = zf.read("xl/worksheets/_rels/sheet1.bin.rels").decode()
        pivot_rels = zf.read("xl/pivotTables/_rels/pivotTable2.bin.rels").decode()
    for number in (1, 2):
        assert f"xl/pivotTables/pivotTable{number}.bin" in names
        assert f"xl/pivotCache/pivotCacheDefinition{number}.bin" in names
        assert f'PartName="/xl/pivotTables/pivotTable{number}.bin" ContentType="application/vnd.ms-excel.pivotTable"' in content_types
        assert f'PartName="/xl/pivotCache/pivotCacheDefinition{number}.bin" ContentType="application/vnd.ms-excel.pivotCacheDefinition"' in content_types
        assert f'Target="pivotCache/pivotCacheDefinition{number}.bin"' in workbook_rels
        assert f'Target="../pivotTables/pivotTable{number}.bin"' in sheet_rels
    assert 'Target="../pivotCache/pivotCacheDefinition2.bin"' in pivot_rels
    assert not any(name.endswith(".xml") and "pivot" in name for name in names)


def test_sheet_relationship_ids_are_unique():
    data = make(XlsbWriter, [pivot(), pivot(["Units"], location="A20", name="Units")])
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        rels = zf.read("xl/worksheets/_rels/sheet1.bin.rels").decode()
    ids = [part.split('"')[0] for part in rels.split('Id="')[1:]]
    assert len(ids) == len(set(ids)) == 4


def test_workbook_lists_each_cache_before_file_recover():
    data = make(XlsbWriter, [pivot(), pivot(["Units"], location="A20", name="Units")])
    records = part_records(data, "xl/workbook.bin")
    ids = [rid for rid, _payload in records]
    start = ids.index(0x180)
    assert ids[start:start + 6] == [0x180, 0x182, 0x183, 0x182, 0x183, 0x181]
    assert ids.index(0x9B) == start + 6
    assert ids.index(0x9D) < start  # after BrtCalcProp
    first, second = (p for rid, p in records if rid == 0x182)
    assert struct.unpack_from("<I", first)[0] == 1 and struct.unpack_from("<I", second)[0] == 2


def test_workbook_without_pivots_has_no_pivot_records():
    data = make(XlsbWriter, [])
    assert not any(rid in (0x180, 0x182) for rid, _p in part_records(data, "xl/workbook.bin"))
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        assert not any("pivot" in name.lower() for name in zf.namelist())


def test_location_matches_the_rendered_range():
    data = make(XlsbWriter, [pivot(columns="Quarter", values=["Sales", "Units"])])
    (location,) = records_of(data, TABLE, 0x13A)
    first_row, last_row, first_col, last_col, head, first_data, data_col, row_pages, col_pages = (
        struct.unpack("<9I", location)
    )
    assert (first_row, last_row, first_col, last_col) == (11, 17, 0, 6)  # A12:G18
    assert (head, first_data, data_col) == (12, 14, 1)
    assert (row_pages, col_pages) == (0, 0)


def test_location_is_absolute_and_counts_page_fields():
    data = make(XlsbWriter, [pivot(filters="Product", location="C14")])
    (location,) = records_of(data, TABLE, 0x13A)
    values = struct.unpack("<9I", location)
    assert values[:4] == (13, 17, 2, 3)  # C14:D18
    assert values[4:] == (14, 14, 3, 1, 1)


@pytest.mark.parametrize(
    ("function", "code"), [("sum", 0), ("count", 1), ("average", 2), ("max", 3), ("min", 4)],
)
def test_data_field_function_codes(function, code):
    data = make(XlsbWriter, [pivot([PivotValue("Sales", function)])])
    (field,) = records_of(data, TABLE, 0x125)
    column, subtotal = struct.unpack_from("<II", field)
    assert (column, subtotal) == (3, code)


def test_data_field_number_format_is_registered_in_styles():
    data = make(XlsbWriter, [pivot([PivotValue("Sales", "sum", number_format="#,##0.0")])])
    (field,) = records_of(data, TABLE, 0x125)
    numfmt = struct.unpack_from("<I", field, 20)[0]
    assert numfmt >= 164
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        styles = zf.read("xl/styles.bin")
    assert "#,##0.0".encode("utf-16le") in styles


def test_axis_flags_and_items_per_field():
    data = make(XlsbWriter, [pivot(columns="Quarter", filters="Product")])
    axes = [payload[0] for payload in records_of(data, TABLE, 0x11D)]
    assert axes == [0x01, 0x04, 0x02, 0x08, 0x00]  # Region row, Product page, Quarter col, Sales data
    counts = [struct.unpack("<I", p)[0] for p in records_of(data, TABLE, 0x11B)]
    assert counts == [4, 3, 3]  # three regions, two products, two quarters, plus the default item
    (page,) = records_of(data, TABLE, 0x121)
    assert struct.unpack_from("<I", page)[0] == 1


def test_row_items_follow_the_display_order():
    data = make(XlsbWriter, [pivot()])
    # Items are listed by cache index in display order: East(2), North(0), South(1).
    items = [struct.unpack_from("<I", p, 3)[0] for p in records_of(data, TABLE, 0x11A)]
    assert items == [2, 0, 1, 0xFFFFFFFF]


def test_cache_records_the_table_source_and_field_types():
    data = make(XlsbWriter, [pivot(columns="Quarter")])
    (source,) = records_of(data, CACHE, 0xBB)
    assert source == b"\x01\x00\x00" + struct.pack("<I", 10) + "SalesTable".encode("utf-16le")
    (header,) = records_of(data, CACHE, 0xB3)
    assert struct.unpack_from("<I", header, 17)[0] == 8  # records
    flags = [struct.unpack_from("<HI", p) for p in records_of(data, CACHE, 0xBD)]
    # Region (3 items), Product (unused text), Quarter (2 items), Sales, Units (numbers)
    assert flags == [(0x040B, 3), (0x040B, 0), (0x05C2, 2), (0x05C2, 0), (0x05C2, 0)]
    ranges = [struct.unpack_from("<dd", p, 6) for p in records_of(data, CACHE, 0xBD) if len(p) == 22]
    assert ranges == [(1.0, 2.0), (700.0, 1600.0), (30.0, 60.0)]
    numbers = [struct.unpack("<d", p)[0] for p in records_of(data, CACHE, 0x15)]
    assert numbers == [1.0, 2.0]
    strings = [p[4:].decode("utf-16le") for p in records_of(data, CACHE, 0x18)]
    assert strings == ["North", "South", "East"]


def test_blank_and_decimal_cells_set_the_type_flags():
    rows = [
        ["Region", "Amount", "Qty"],
        ["North", 1.5, 1],
        ["South", 2, 2],
        [None, 3.25, None],
        ["North", None, 3],
    ]
    data = make(
        XlsbWriter,
        [PivotTable(source="SalesTable", rows="Region", columns="Qty", values="Amount", location="A8")],
        rows=rows, ref="A1:C5",
    )
    flags = [struct.unpack_from("<HI", p) for p in records_of(data, CACHE, 0xBD)]
    assert flags == [(0x041B, 3), (0x0553, 0), (0x05D3, 4)]  # text+blank, decimals+blank, ints+blank
    assert len(records_of(data, CACHE, 0x14)) == 2  # one blank item on Region, one on Qty


def test_table_name_and_display_name_are_equal():
    # A pivot cache finds its table by display name, so the two must match.
    data = make(XlsbWriter, [pivot()])
    (begin,) = records_of(data, "xl/tables/table1.bin", 0x157)
    name = "SalesTable".encode("utf-16le")
    assert begin.count(struct.pack("<I", 10) + name) == 2


def test_output_is_deterministic():
    pivots = [pivot(columns="Quarter")]
    first, second = make(XlsbWriter, pivots), make(XlsbWriter, pivots)
    with zipfile.ZipFile(io.BytesIO(first)) as a, zipfile.ZipFile(io.BytesIO(second)) as b:
        for name in (TABLE, CACHE):
            assert a.read(name) == b.read(name)


# -- refusals -----------------------------------------------------------------


def test_dates_in_a_pivot_field_are_refused():
    import datetime

    rows = [["Day", "Sales"], [datetime.date(2024, 1, 1), 1], [datetime.date(2024, 1, 2), 2]]
    with pytest.raises(ValueError, match="dates, booleans or NaN"):
        make(XlsbWriter, [PivotTable(source="SalesTable", rows="Day", values="Sales", location="A5")],
             rows=rows, ref="A1:B3")


def test_booleans_in_a_value_field_are_refused():
    rows = [["Region", "Flag"], ["North", True], ["South", False]]
    with pytest.raises(ValueError, match="dates, booleans or NaN"):
        make(XlsbWriter, [PivotTable(source="SalesTable", rows="Region", values="Flag", location="A5")],
             rows=rows, ref="A1:B3")


def test_pivot_target_sheet_must_be_written():
    out = io.BytesIO()
    with pytest.raises(ValueError, match="never written"):
        with XlsbWriter(out) as w:
            w.add_sheet("Data")
            w.write_sheet(ROWS, tables=[ExcelTable("SalesTable", "A1:E9")],
                          pivots=[pivot(location="A3", sheet="Summary")])
            w.add_sheet("Summary")


def test_target_sheet_written_before_its_source_is_refused():
    out = io.BytesIO()
    with pytest.raises(ValueError, match="written before its source"):
        with XlsbWriter(out) as w:
            w.add_sheet("Summary")
            w.write_sheet([])
            w.add_sheet("Data")
            w.write_sheet(ROWS, tables=[ExcelTable("SalesTable", "A1:E9")],
                          pivots=[pivot(location="A3", sheet="Summary")])


def test_pivot_on_the_source_sheet_must_sit_below_the_data():
    with pytest.raises(ValueError, match="below the sheet's data"):
        make(XlsbWriter, [pivot(location="A5")])


def test_overlapping_pivots_are_refused():
    with pytest.raises(ValueError, match="overlap"):
        make(XlsbWriter, [pivot(location="A12"), pivot(["Units"], location="A13")])


def test_unknown_field_is_refused():
    with pytest.raises(ValueError, match="not a column"):
        make(XlsbWriter, [pivot("Revenue")])


@pytest.mark.parametrize(("writer", "extension"), [(XlsxWriter, "xlsx"), (XlsbWriter, "xlsb")])
def test_removing_a_loaded_sheet_with_a_pivot_is_refused(tmp_path, writer, extension):
    source = tmp_path / f"pivot.{extension}"
    source.write_bytes(make(writer, [pivot()], summary=True))
    book = load_workbook(source)
    try:
        book.remove(book["Data"])
        out = tmp_path / f"removed.{extension}"
        with pytest.raises(NotImplementedError, match="pivot tables"):
            book.save(out)
        assert not out.exists()
    finally:
        book.close()


def test_loaded_xlsb_keeps_its_pivot_parts_through_a_save(tmp_path):
    source = tmp_path / "pivot.xlsb"
    source.write_bytes(make(XlsbWriter, [pivot()]))
    book = load_workbook(source)
    try:
        book["Data"]["H1"] = "note"
        out = tmp_path / "copy.xlsb"
        book.save(out)
    finally:
        book.close()
    with zipfile.ZipFile(out) as zf:
        names = set(zf.namelist())
    assert {TABLE, CACHE} <= names


def test_write_dataframe_passes_pivots_through(tmp_path):
    pd = pytest.importorskip("pandas")
    frame = pd.DataFrame(ROWS[1:], columns=ROWS[0])
    out = io.BytesIO()
    with XlsbWriter(out) as w:
        w.add_sheet("Data")
        w.write_dataframe(frame, tables=[ExcelTable("SalesTable", "A1:E9")], pivots=[pivot()])
    assert xlsb_cells(out.getvalue(), tmp_path, "Data", "A12:B16")[-1] == ["Grand Total", 9100]
