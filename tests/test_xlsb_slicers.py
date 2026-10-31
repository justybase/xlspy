"""Table slicers in XLSB output, checked against the records Excel writes.

The expected bytes come from workbooks Excel 16 saved as XLSB
(``tools/capture_xlsb.py --slicers``): a table ``Sales`` with a slicer on
``Region``, then one option changed at a time.
"""

import os
import struct
import zipfile

import pytest

from xlspy import ExcelTable, Formula, Slicer, XlsbWriter
from xlspy._xlsb_metadata import read_wide
from xlspy._xlsb_slicers import (
    sheet_slicers_bin, slicer_cache_bin, slicers_part_bin, workbook_slicer_caches_bin,
)
from xlspy.biff_utils import iter_records
from xlspy.workbook import load_workbook

ROWS = [
    ["Region", "Month", "Sales", "Visible", None, "Total"],
    ["North", "Jan", 10, Formula("SUBTOTAL(103,A2)", 1), None, Formula("SUMIFS(C2:C5,D2:D5,1)", 40)],
    ["North", "Feb", 14, Formula("SUBTOTAL(103,A3)", 1)],
    ["South", "Jan", 7, Formula("SUBTOTAL(103,A4)", 1)],
    ["South", "Feb", 9, Formula("SUBTOTAL(103,A5)", 1)],
]
TABLE = ExcelTable("Sales", "A1:D5")

# Excel's records: the slicer header (flags, first item, columns, level; the row
# height that follows depends on the font), the table cache link and the
# relationship-id records of the sheet and the workbook.
EXCEL_SLICER_HEAD = bytes.fromhex("03 00000000 01000000 00000000")
EXCEL_SLICER_HEAD_TWO_COLUMNS = bytes.fromhex("03 00000000 02000000 00000000")
EXCEL_TABLE_CACHE = bytes.fromhex("00000000 01000000 01000000 15")
EXCEL_TABLE_CACHE_SECOND_TABLE = bytes.fromhex("00000000 01000000 02000000 15")
EXCEL_SHEET_SLICER_REL = bytes.fromhex("08000000 0400 72004900640034 00")
EXCEL_BOOK_CACHE_REL = bytes.fromhex("08000000 0400 72004900640032 00")


def _records(data):
    return [(record.record_id, data[record.data_start:record.data_end]) for record in iter_records(data)]


def _write(path, slicers, rows=ROWS, tables=(TABLE,)):
    with XlsbWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(rows, tables=list(tables), slicers=slicers)
    return path


def _part(path, name):
    with zipfile.ZipFile(path) as archive:
        return archive.read(name)


def test_slicer_record_matches_excel_layout():
    slicer = Slicer("Sales", "Region", caption="Wybierz region", name="RegionSlicer")
    records = _records(slicers_part_bin([slicer]))
    assert [record_id for record_id, _ in records] == [0x45B, 0x43B, 0x43C, 0x45C]
    payload = records[1][1]
    assert payload[:13] == EXCEL_SLICER_HEAD
    assert struct.unpack_from("<I", payload, 13)[0] == 241300
    name, offset = read_wide(payload, 17)
    cache, offset = read_wide(payload, offset)
    caption, offset = read_wide(payload, offset)
    assert (name, cache, caption) == ("RegionSlicer", "Slicer_Region", "Wybierz region")
    assert offset == len(payload)
    wide = _records(slicers_part_bin([Slicer("Sales", "Region", columns=2)]))[1][1]
    assert wide[:13] == EXCEL_SLICER_HEAD_TWO_COLUMNS


def test_cache_records_match_excel():
    records = _records(slicer_cache_bin(Slicer("Sales", "Region"), table_id=1, column_id=1))
    assert [record_id for record_id, _ in records] == [
        0x433, 0x435, 0x43D, 0x23, 0x81D, 0x81E, 0x24, 0x436, 0x434]
    payloads = dict(records)
    assert payloads[0x23] == b"\x03\x0f\x00\x00"
    assert payloads[0x81D] == EXCEL_TABLE_CACHE
    assert payloads[0x43D] == b"\x00" * 4
    cache, offset = read_wide(payloads[0x435])
    assert (cache, read_wide(payloads[0x435], offset)[0]) == ("Slicer_Region", "Region")
    second = dict(_records(slicer_cache_bin(Slicer("Costs", "Region"), table_id=2, column_id=1)))
    assert second[0x81D] == EXCEL_TABLE_CACHE_SECOND_TABLE


def test_relationship_records_match_excel():
    sheet = _records(sheet_slicers_bin("rId4"))
    assert [record_id for record_id, _ in sheet] == [0x23, 0x437, 0x439, 0x43A, 0x438, 0x24]
    assert dict(sheet)[0x439] == EXCEL_SHEET_SLICER_REL
    book = _records(workbook_slicer_caches_bin(["rId2", "rId3"]))
    assert [record_id for record_id, _ in book] == [0x23, 0x81B, 0x81C, 0x81C, 0x24]
    assert book[2][1] == EXCEL_BOOK_CACHE_REL
    assert workbook_slicer_caches_bin([]) == b""


def test_writer_links_parts_relationships_and_names(tmp_path):
    path = _write(tmp_path / "s.xlsb", [Slicer("Sales", "Region", anchor="H2"),
                                         Slicer("Sales", "Month", anchor="K2", name="Months")])
    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
        types = archive.read("[Content_Types].xml").decode()
        book_rels = archive.read("xl/_rels/workbook.bin.rels").decode()
        sheet_rels = archive.read("xl/worksheets/_rels/sheet1.bin.rels").decode()
        drawing = archive.read("xl/drawings/drawing1.xml").decode()
    assert {"xl/slicers/slicer1.bin", "xl/slicerCaches/slicerCache1.bin",
            "xl/slicerCaches/slicerCache2.bin"} <= set(names)
    assert 'PartName="/xl/slicers/slicer1.bin" ContentType="application/vnd.ms-excel.slicer"' in types
    assert types.count("application/vnd.ms-excel.slicerCache") == 2
    assert 'Target="slicerCaches/slicerCache2.bin"' in book_rels
    assert 'Target="../slicers/slicer1.bin"' in sheet_rels
    assert drawing.count("<sle:slicer ") == 2 and 'name="Months"' in drawing

    sheet = _records(_part(path, "xl/worksheets/sheet1.bin"))
    ids = [record_id for record_id, _ in sheet]
    assert ids.index(0x437) > ids.index(0x296)  # after LISTPARTS
    rel_id = dict(sheet)[0x439][6:].decode("utf-16le")
    assert f'Id="{rel_id}"' in sheet_rels

    book = _records(_part(path, "xl/workbook.bin"))
    ids = [record_id for record_id, _ in book]
    assert ids.index(0x81B) == ids.index(0x9B) + 2  # BrtFileRecover, BrtFRTBegin, the list
    defined = [read_wide(payload, 9)[0] for record_id, payload in book if record_id == 0x27]
    assert defined == ["Slicer_Region", "Slicer_Month"]
    cache_rels = [payload[6:].decode("utf-16le") for record_id, payload in book if record_id == 0x81C]
    assert all(f'Id="{rel_id}"' in book_rels for rel_id in cache_rels)
    tables = dict(_records(_part(path, "xl/slicerCaches/slicerCache2.bin")))
    assert struct.unpack_from("<II", tables[0x81D], 4) == (2, 1)  # Month is column 2 of table 1


def test_hyperlinks_keep_their_own_relationship_ids(tmp_path):
    from xlspy import Hyperlink

    rows = [list(row) for row in ROWS]
    rows[2] = rows[2] + [None, Hyperlink("https://example.com", "site")]
    path = _write(tmp_path / "s.xlsb", [Slicer("Sales", "Region")], rows=rows)
    rels = _part(path, "xl/worksheets/_rels/sheet1.bin.rels").decode()
    ids = [chunk.split('"')[0] for chunk in rels.split('Id="')[1:]]
    assert len(ids) == len(set(ids))
    assert "slicers/slicer1.bin" in rels and "https://example.com" in rels


def test_unknown_column_is_rejected_before_the_rows_stream(tmp_path):
    with XlsbWriter(tmp_path / "s.xlsb") as writer:
        writer.add_sheet("Data")
        with pytest.raises(ValueError, match="not a column"):
            writer.write_sheet(ROWS, tables=[ExcelTable("Sales", "A1:D5",
                                                        columns=["Region", "Month", "Sales", "Visible"])],
                               slicers=[Slicer("Sales", "Country")])
        writer.write_sheet(ROWS, tables=[TABLE], slicers=[Slicer("Sales", "Region")])


def test_unknown_header_found_while_streaming_fails_the_sheet(tmp_path):
    # Without ``columns`` the names come from the header row, read as it streams.
    with pytest.raises(ValueError, match="not a column"):
        with XlsbWriter(tmp_path / "s.xlsb") as writer:
            writer.add_sheet("Data")
            writer.write_sheet(ROWS, tables=[TABLE], slicers=[Slicer("Sales", "Country")])


def test_slicer_names_must_be_unique_across_sheets(tmp_path):
    with XlsbWriter(tmp_path / "s.xlsb") as writer:
        writer.add_sheet("A")
        writer.write_sheet(ROWS, tables=[TABLE], slicers=[Slicer("Sales", "Region")])
        writer.add_sheet("B")
        with pytest.raises(ValueError, match="cache name"):
            writer.write_sheet(ROWS, tables=[ExcelTable("Costs", "A1:D5")],
                               slicers=[Slicer("Costs", "Region", name="Other")])


def test_loaded_workbook_keeps_the_slicers_on_save(tmp_path):
    source = _write(tmp_path / "s.xlsb", [Slicer("Sales", "Region")])
    workbook = load_workbook(source)
    try:
        workbook.worksheets[0].cell(row=8, column=1).value = "edited"
        workbook.save(tmp_path / "saved.xlsb")
    finally:
        workbook.close()
    saved = tmp_path / "saved.xlsb"
    for part in ("xl/slicers/slicer1.bin", "xl/slicerCaches/slicerCache1.bin"):
        assert _part(saved, part) == _part(source, part)
    book_ids = [record_id for record_id, _ in _records(_part(saved, "xl/workbook.bin"))]
    assert 0x81C in book_ids
    sheet_ids = [record_id for record_id, _ in _records(_part(saved, "xl/worksheets/sheet1.bin"))]
    assert 0x439 in sheet_ids


def test_removing_a_sheet_with_slicers_is_refused(tmp_path):
    source = tmp_path / "two.xlsb"
    with XlsbWriter(source) as writer:
        writer.add_sheet("Plain")
        writer.write_sheet([["x"]])
        writer.add_sheet("Data")
        writer.write_sheet(ROWS, tables=[TABLE], slicers=[Slicer("Sales", "Region")])
    workbook = load_workbook(source)
    try:
        workbook.remove(workbook.worksheets[1])
        with pytest.raises(NotImplementedError, match="table slicers"):
            workbook.save(tmp_path / "removed.xlsb")
    finally:
        workbook.close()
    assert not (tmp_path / "removed.xlsb").exists()


@pytest.mark.skipif(os.environ.get("XLSPY_RUN_EXCEL_COM") != "1",
                    reason="Set XLSPY_RUN_EXCEL_COM=1 to run Microsoft Excel COM validation")
def test_excel_filters_the_table_with_an_xlsb_slicer(tmp_path):
    win32 = pytest.importorskip("win32com.client")
    path = _write(tmp_path / "com.xlsb", [
        Slicer("Sales", "Region", anchor="H2", caption="Wybierz region"),
        Slicer("Sales", "Month", anchor="K2", columns=2, name="Months"),
    ])
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    try:
        book = excel.Workbooks.Open(str(path))
        sheet = book.Worksheets(1)
        sheet.Activate()
        assert book.SlicerCaches.Count == 2
        region = book.SlicerCaches("Slicer_Region")
        assert region.SourceName == "Region"
        assert region.ListObject.Name == "Sales"
        assert region.Slicers(1).Caption == "Wybierz region"
        months = book.SlicerCaches("Slicer_Month").Slicers(1)
        assert (months.Name, months.NumberOfColumns) == ("Months", 2)
        assert sheet.Range("F2").Value == 40
        region.SlicerItems("South").Selected = False
        excel.Calculate()
        assert sheet.Range("F2").Value == 24
        book.Close(SaveChanges=False)
    finally:
        excel.Quit()
