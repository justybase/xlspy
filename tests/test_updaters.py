import datetime
import struct
import zipfile

import pytest

from xlspy import ExcelReader, F, XlsbUpdater, XlsbWriter, XlsxUpdater, XlsxWriter
from xlspy.biff_utils import build_record, iter_records, read_utf16

from zip_utils import rewrite_workbook_members
from xlspy import Workbook
from xlspy import load_workbook
from xlspy.formulas import Formula
from xml.etree import ElementTree as ET
import math
import re


def _write_xlsx(path, sheets, *, use_shared_strings=True):
    with XlsxWriter(path, use_shared_strings=use_shared_strings) as writer:
        for name, rows in sheets:
            writer.add_sheet(name)
            writer.write_sheet(rows)


def _write_xlsb(path, sheets):
    with XlsbWriter(path) as writer:
        for name, rows in sheets:
            writer.add_sheet(name)
            writer.write_sheet(rows)


def _read_sheet(path, name):
    reader = ExcelReader()
    reader.open(str(path))
    try:
        return list(reader.get_rows(name))
    finally:
        reader.close()


def _add_zip_members(path, members):
    rewrite_workbook_members(path, members)


def _make_xlsx_pivot(path):
    _add_zip_members(
        path,
        {
            "xl/pivotCache/pivotCacheDefinition1.xml": (
                b'<pivotCacheDefinition recordCount="1">'
                b'<cacheSource type="worksheet">'
                b'<worksheetSource ref="A1:B2" sheet="Data"/>'
                b"</cacheSource></pivotCacheDefinition>"
            ),
            "xl/pivotTables/pivotTable1.xml": b'<pivotTableDefinition refreshOnLoad="0"/>',
        },
    )


def _make_xlsb_pivot(path):
    sheet_name = "Data"
    encoded_name = sheet_name.encode("utf-16le")
    source_payload = (
        b"\x00\x00\x00"
        + struct.pack("<I", len(sheet_name))
        + encoded_name
        + struct.pack("<iiii", 0, 1, 0, 1)
    )
    cache = build_record(0xB3, b"\x00\x00\x00\x00") + build_record(0xBB, source_payload)
    _add_zip_members(path, {"xl/pivotCache/pivotCacheDefinition1.bin": cache})


def _xlsx_sheet_xml(path):
    with zipfile.ZipFile(path) as archive:
        return archive.read("xl/worksheets/sheet1.xml").decode("utf-8")


def _xlsb_sheet_bytes(path):
    with zipfile.ZipFile(path) as archive:
        return archive.read("xl/worksheets/sheet1.bin")


def _set_xlsx_date_system(path, enabled=True):
    with zipfile.ZipFile(path) as archive:
        workbook = archive.read("xl/workbook.xml").decode("utf-8")
    value = "1" if enabled else "0"
    workbook = workbook.replace(
        "<workbookPr ", f'<workbookPr date1904="{value}" '
    )
    if "<workbookPr" not in workbook:
        workbook = workbook.replace(
            "<workbook ", f'<workbookPr date1904="{value}"/><workbook '
        )
    _add_zip_members(path, {"xl/workbook.xml": workbook.encode("utf-8")})


def _set_xlsb_date_system(path, enabled=True):
    with zipfile.ZipFile(path) as archive:
        workbook = bytearray(archive.read("xl/workbook.bin"))
    for record in iter_records(bytes(workbook)):
        if record.record_id == 0x99 and record.length >= 4:
            flags = struct.unpack_from("<I", workbook, record.data_start)[0]
            struct.pack_into(
                "<I",
                workbook,
                record.data_start,
                flags | 0x01 if enabled else flags & ~0x01,
            )
            break
    _add_zip_members(path, {"xl/workbook.bin": bytes(workbook)})


def _set_xlsb_trailing_row_header(path):
    with zipfile.ZipFile(path) as archive:
        sheet = archive.read("xl/worksheets/sheet1.bin")
    insertion = next(
        record.header_start for record in iter_records(sheet) if record.record_id == 0x92
    )
    payload = bytearray(25)
    struct.pack_into("<i", payload, 0, 4)
    struct.pack_into("<i", payload, 8, 300)
    payload[13] = 1
    struct.pack_into("<i", payload, 17, 0)
    struct.pack_into("<i", payload, 21, 0)
    sheet = sheet[:insertion] + build_record(0x00, bytes(payload)) + sheet[insertion:]
    _add_zip_members(path, {"xl/worksheets/sheet1.bin": sheet})


def test_xlsx_updater_replaces_data_and_preserves_other_sheet(tmp_path):
    source = tmp_path / "source.xlsx"
    output = tmp_path / "output.xlsx"
    _write_xlsx(
        source,
        [("Data", [["old", 1], ["remove", 2]]), ("Keep", [["untouched", 9]])],
    )

    updater = XlsxUpdater(source)
    assert updater.get_sheet_names() == ["Data", "Keep"]
    updater.replace_sheet_data(
        "Data",
        ((row[0], row[1]) for row in [["new & <value>", 42], [None, True], ["middle", None], [None, None]]),
        headers=["Name", "Value"],
    )
    output.write_bytes(updater.to_bytes())
    updater.save(output)

    assert _read_sheet(output, "Data") == [
        ["Name", "Value"],
        ["new & <value>", 42],
        [None, True],
        ["middle", None],
    ]
    assert _read_sheet(output, "Keep") == [["untouched", 9]]
    sheet_xml = _xlsx_sheet_xml(output)
    assert 'ref="A1:B4"' in sheet_xml
    assert "old" not in sheet_xml


def test_xlsx_updater_supports_inline_strings_and_general_fallback(tmp_path):
    path = tmp_path / "inline.xlsx"
    _write_xlsx(path, [("Sheet1", [["old", 10], ["old2", 20]])], use_shared_strings=False)

    updater = XlsxUpdater(path)
    updater.replace_sheet_data("Sheet1", [["<&\"'\x01", 99]], style_fallback="general")
    path.write_bytes(updater.to_bytes())

    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
        xml = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")
        assert "xl/sharedStrings.xml" not in names
        assert 't="inlineStr"' in xml
        assert "&lt;&amp;&quot;&apos;" in xml
        assert "_x0001_" in xml
        assert ' s="' not in xml
    assert _read_sheet(path, "Sheet1")[0][0] == "<&\"'\x01"


def test_xlsx_updater_inherits_date_and_header_styles(tmp_path):
    path = tmp_path / "styles.xlsx"
    _write_xlsx(
        path,
        [
            (
                "Sheet1",
                [["Date", "Amount"], [(datetime.date(2024, 1, 2), F.DATE_SHORT), (1.25, F.TWO_DECIMALS)]],
            )
        ],
    )
    updater = XlsxUpdater(path)
    updater.replace_sheet_data(
        "Sheet1",
        [[datetime.date(2025, 2, 3), 2.5]],
        headers=["Date", "Amount"],
    )
    path.write_bytes(updater.to_bytes())
    xml = _xlsx_sheet_xml(path)
    assert xml.count('s="3"') >= 2
    assert 's="1"' in xml or 's="4"' in xml
    assert isinstance(_read_sheet(path, "Sheet1")[1][0], (datetime.date, datetime.datetime))


def test_xlsx_updater_updates_pivot_metadata_and_refresh_flag(tmp_path):
    path = tmp_path / "pivot.xlsx"
    _write_xlsx(path, [("Data", [["old", 1], ["old2", 2]])])
    _make_xlsx_pivot(path)

    updater = XlsxUpdater(path)
    updater.replace_sheet_data("Data", [["one", 1], ["two", 2], ["three", 3]])
    path.write_bytes(updater.to_bytes())

    with zipfile.ZipFile(path) as archive:
        cache = archive.read("xl/pivotCache/pivotCacheDefinition1.xml").decode()
        table = archive.read("xl/pivotTables/pivotTable1.xml").decode()
        assert 'recordCount="3"' in cache
        assert 'ref="A1:B3"' in cache
        assert 'refreshOnLoad="1"' in cache
        assert 'refreshOnLoad="1"' in table


@pytest.mark.parametrize("binary", [False, True])
def test_updater_honors_1904_date_system(binary, tmp_path):
    path = tmp_path / ("date1904.xlsb" if binary else "date1904.xlsx")
    if binary:
        _write_xlsb(path, [("Sheet1", [[1]])])
        _set_xlsb_date_system(path)
        updater = XlsbUpdater(path)
    else:
        _write_xlsx(path, [("Sheet1", [[1]])])
        _set_xlsx_date_system(path)
        updater = XlsxUpdater(path)

    updater.replace_sheet_data("Sheet1", [[datetime.date(2024, 1, 1)]])
    path.write_bytes(updater.to_bytes())

    expected_serial = 43830.0
    if binary:
        sheet = _xlsb_sheet_bytes(path)
        date_values = [
            struct.unpack_from("<d", sheet, record.data_start + 8)[0]
            for record in iter_records(sheet)
            if record.record_id == 0x05
        ]
        assert date_values == [expected_serial]
    else:
        assert f"<v>{expected_serial:.15g}</v>" in _xlsx_sheet_xml(path)


def test_xlsb_updater_detects_custom_date_format_for_unstyled_column(tmp_path):
    path = tmp_path / "custom-date.xlsb"
    _write_xlsb(
        path,
        [("Sheet1", [[None, None, (datetime.date(2024, 1, 1), F.DATE_SHORT)]])],
    )

    updater = XlsbUpdater(path)
    updater.replace_sheet_data("Sheet1", [[datetime.date(2025, 2, 3)]])
    path.write_bytes(updater.to_bytes())

    sheet = _xlsb_sheet_bytes(path)
    styles = [
        struct.unpack_from("<I", sheet, record.data_start + 4)[0] & 0xFFFFFF
        for record in iter_records(sheet)
        if record.record_id == 0x05
    ]
    assert styles and styles[0] > 0


def test_xlsx_updater_preserves_worksheet_and_sst_prefixes(tmp_path):
    path = tmp_path / "prefix.xlsx"
    _write_xlsx(path, [("Sheet1", [["old"]])])
    with zipfile.ZipFile(path) as archive:
        sheet = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")
        strings = archive.read("xl/sharedStrings.xml").decode("utf-8")
    sheet = sheet.replace(
        "<worksheet ", '<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    ).replace("</worksheet>", "</x:worksheet>")
    sheet = sheet.replace("<sheetData", "<x:sheetData").replace(
        "</sheetData>", "</x:sheetData>"
    )
    strings = strings.replace(
        "<sst ", '<x:sst xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    ).replace("</sst>", "</x:sst>")
    strings = strings.replace("<si>", "<x:si>").replace("</si>", "</x:si>")
    strings = strings.replace("<t", "<x:t").replace("</t>", "</x:t>")
    _add_zip_members(
        path,
        {
            "xl/worksheets/sheet1.xml": sheet.encode("utf-8"),
            "xl/sharedStrings.xml": strings.encode("utf-8"),
        },
    )

    updater = XlsxUpdater(path)
    updater.replace_sheet_data("Sheet1", [["new"]])
    path.write_bytes(updater.to_bytes())

    with zipfile.ZipFile(path) as archive:
        updated_sheet = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")
        updated_strings = archive.read("xl/sharedStrings.xml").decode("utf-8")
    assert "<x:row" in updated_sheet
    assert "<x:c" in updated_sheet
    assert "<x:v" in updated_sheet
    assert "<x:si>" in updated_strings
    assert "<x:t>new</x:t>" in updated_strings


@pytest.mark.parametrize("binary", [False, True])
def test_updater_recalculates_shared_string_reference_count(binary, tmp_path):
    path = tmp_path / ("count.xlsb" if binary else "count.xlsx")
    if binary:
        _write_xlsb(path, [("Data", [["old"]]), ("Keep", [["keep"]])])
        updater = XlsbUpdater(path)
    else:
        _write_xlsx(path, [("Data", [["old"]]), ("Keep", [["keep"]])])
        updater = XlsxUpdater(path)

    updater.replace_sheet_data("Data", [])
    path.write_bytes(updater.to_bytes())

    member = "xl/sharedStrings.bin" if binary else "xl/sharedStrings.xml"
    with zipfile.ZipFile(path) as archive:
        data = archive.read(member)
    if binary:
        header = next(record for record in iter_records(data) if record.record_id == 0x9F)
        assert struct.unpack_from("<I", data, header.data_start)[0] == 1
    else:
        assert 'count="1"' in data.decode("utf-8")


@pytest.mark.parametrize("binary", [False, True])
def test_updater_uses_first_row_style_as_data_without_headers(binary, tmp_path):
    path = tmp_path / ("stylefirst.xlsb" if binary else "stylefirst.xlsx")
    if binary:
        _write_xlsb(path, [("Sheet1", [[(1, F.TWO_DECIMALS)]])])
        updater = XlsbUpdater(path)
    else:
        _write_xlsx(path, [("Sheet1", [[(1, F.TWO_DECIMALS)]])])
        updater = XlsxUpdater(path)

    updater.replace_sheet_data("Sheet1", [[2.5]])
    path.write_bytes(updater.to_bytes())

    if binary:
        sheet = _xlsb_sheet_bytes(path)
        styles = [
            struct.unpack_from("<I", sheet, record.data_start + 4)[0] & 0xFFFFFF
            for record in iter_records(sheet)
            if record.record_id == 0x05
        ]
        assert styles == [4]
    else:
        assert 's="' in _xlsx_sheet_xml(path)


def test_xlsb_updater_replaces_trailing_row_headers(tmp_path):
    path = tmp_path / "trailing.xlsb"
    _write_xlsb(path, [("Sheet1", [[1]])])
    _set_xlsb_trailing_row_header(path)

    updater = XlsbUpdater(path)
    updater.replace_sheet_data("Sheet1", [[2]])
    path.write_bytes(updater.to_bytes())

    sheet = _xlsb_sheet_bytes(path)
    assert sum(record.record_id == 0x00 for record in iter_records(sheet)) == 1


def test_xlsb_updater_only_refreshes_matching_pivot_cache(tmp_path):
    path = tmp_path / "multiple-pivots.xlsb"
    _write_xlsb(path, [("Data", [["old"]]), ("Other", [["old"]])])
    _make_xlsb_pivot(path)
    other_name = "Other".encode("utf-16le")
    other_payload = (
        b"\x00\x00\x00"
        + struct.pack("<I", len("Other"))
        + other_name
        + struct.pack("<iiii", 0, 1, 0, 0)
    )
    _add_zip_members(
        path,
        {
            "xl/pivotCache/pivotCacheDefinition2.bin": build_record(
                0xB3, b"\x00\x00\x00\x00"
            )
            + build_record(0xBB, other_payload)
        },
    )

    updater = XlsbUpdater(path)
    updater.replace_sheet_data("Data", [["new"]])
    path.write_bytes(updater.to_bytes())

    with zipfile.ZipFile(path) as archive:
        other_cache = archive.read("xl/pivotCache/pivotCacheDefinition2.bin")
    other_header = next(
        record for record in iter_records(other_cache) if record.record_id == 0xB3
    )
    assert other_cache[other_header.data_start + 3] & 0x04 == 0


def test_xlsb_updater_replaces_data_and_preserves_binary_wrappers(tmp_path):
    source = tmp_path / "source.xlsb"
    output = tmp_path / "output.xlsb"
    _write_xlsb(source, [("Data", [["old", 1], ["remove", 2]]), ("Keep", [["yes", 8]])])
    before = _xlsb_sheet_bytes(source)

    updater = XlsbUpdater(source)
    assert updater.get_sheet_names() == ["Data", "Keep"]
    updater.replace_sheet_data(
        "Data",
        [["new", 42], [None, True], ["middle", None], [None, None]],
        headers=["Name", "Value"],
    )
    output.write_bytes(updater.to_bytes())

    assert _read_sheet(output, "Data") == [
        ["Name", "Value"],
        ["new", 42],
        [None, True],
        ["middle", None],
    ]
    assert _read_sheet(output, "Keep") == [["yes", 8]]
    after = _xlsb_sheet_bytes(output)
    assert len(after) != len(before)
    assert any(record.record_id in (0x92, 0x217, 0x1DD, 0x1DC) for record in iter_records(after))


def test_xlsb_updater_supports_dates_styles_and_shared_strings(tmp_path):
    path = tmp_path / "styles.xlsb"
    _write_xlsb(
        path,
        [("Sheet1", [["Date", "Amount"], [(datetime.date(2024, 1, 2), F.DATE_SHORT), (1.25, F.TWO_DECIMALS)]])],
    )
    updater = XlsbUpdater(path)
    updater.replace_sheet_data(
        "Sheet1",
        [[datetime.date(2025, 2, 3), 2.5]],
        headers=["Date", "Amount"],
    )
    path.write_bytes(updater.to_bytes())

    with zipfile.ZipFile(path) as archive:
        sheet = archive.read("xl/worksheets/sheet1.bin")
        cell_styles = []
        for record in iter_records(sheet):
            if 0x01 <= record.record_id <= 0x0B and record.length >= 8:
                cell_styles.append(struct.unpack_from("<I", sheet, record.data_start + 4)[0] & 0xFFFFFF)
        assert 3 in cell_styles
        assert any(style in (1, 4) for style in cell_styles)
    assert isinstance(_read_sheet(path, "Sheet1")[1][0], (datetime.date, datetime.datetime))


def test_xlsb_updater_updates_pivot_cache_source_and_refresh_flag(tmp_path):
    path = tmp_path / "pivot.xlsb"
    _write_xlsb(path, [("Data", [["old", 1], ["old2", 2]])])
    _make_xlsb_pivot(path)

    updater = XlsbUpdater(path)
    updater.replace_sheet_data("Data", [["one", 1], ["two", 2], ["three", 3]])
    path.write_bytes(updater.to_bytes())

    with zipfile.ZipFile(path) as archive:
        cache = archive.read("xl/pivotCache/pivotCacheDefinition1.bin")
    refresh_flags = []
    source_ranges = []
    for record in iter_records(cache):
        payload = cache[record.data_start : record.data_end]
        if record.record_id == 0xB3:
            refresh_flags.append(payload[3] & 0x04)
        elif record.record_id == 0xBB:
            count = struct.unpack_from("<I", payload, 3)[0]
            name, name_end = read_utf16(payload, 7, count)
            if name == "Data":
                source_ranges.append(struct.unpack_from("<iiii", payload, name_end))
    assert refresh_flags == [4]
    assert source_ranges == [(0, 2, 0, 1)]


@pytest.mark.parametrize("binary", [False, True])
def test_updater_can_clear_a_sheet(binary, tmp_path):
    path = tmp_path / ("clear.xlsb" if binary else "clear.xlsx")
    if binary:
        _write_xlsb(path, [("Sheet1", [[1], [2]])])
        updater = XlsbUpdater(path)
    else:
        _write_xlsx(path, [("Sheet1", [[1], [2]])])
        updater = XlsxUpdater(path)

    updater.replace_sheet_data("Sheet1", [])
    updater.save()
    assert _read_sheet(path, "Sheet1") == []


@pytest.mark.parametrize("binary", [False, True])
def test_updater_preserves_non_target_zip_payloads(binary, tmp_path):
    path = tmp_path / ("preserve.xlsb" if binary else "preserve.xlsx")
    if binary:
        _write_xlsb(path, [("Data", [["old", 1]]), ("Keep", [["keep", 2]])])
        updater = XlsbUpdater(path)
    else:
        _write_xlsx(path, [("Data", [["old", 1]]), ("Keep", [["keep", 2]])])
        updater = XlsxUpdater(path)

    with zipfile.ZipFile(path) as archive:
        before = {name: archive.read(name) for name in archive.namelist()}
    updater.replace_sheet_data("Data", [["new", 10]])
    output = tmp_path / ("preserve-output.xlsb" if binary else "preserve-output.xlsx")
    updater.save(output)
    with zipfile.ZipFile(output) as archive:
        after = {name: archive.read(name) for name in archive.namelist()}

    changed_parts = {"xl/worksheets/sheet1.bin" if binary else "xl/worksheets/sheet1.xml"}
    changed_parts.add("xl/sharedStrings.bin" if binary else "xl/sharedStrings.xml")
    for name, payload in before.items():
        if name not in changed_parts:
            assert after[name] == payload, name


@pytest.mark.parametrize("updater_type", [XlsxUpdater, XlsbUpdater])
def test_updater_errors_are_explicit(tmp_path, updater_type):
    path = tmp_path / ("bad.xlsx" if updater_type is XlsxUpdater else "bad.xlsb")
    path.write_bytes(b"not a zip")
    with pytest.raises(zipfile.BadZipFile):
        updater_type(path)


def test_updater_rejects_wrong_sheet_and_style_option(tmp_path):
    path = tmp_path / "errors.xlsx"
    _write_xlsx(path, [("Sheet1", [[1]])])
    updater = XlsxUpdater(path)
    with pytest.raises(KeyError):
        updater.replace_sheet_data("Missing", [[1]])
    with pytest.raises(ValueError):
        updater.replace_sheet_data("Sheet1", [[1]], style_fallback="bad")


def _first_cell(path):
    workbook = load_workbook(path)
    try:
        return workbook.active["A1"].value
    finally:
        workbook.close()


def _save_plain(path):
    workbook = Workbook()
    workbook.active["A1"] = "x"
    workbook.save(path)
    workbook.close()
    return path


def test_xlsx_updater_escapes_control_characters_and_round_trips(tmp_path):
    path = _save_plain(tmp_path / "ctrl.xlsx")
    updater = XlsxUpdater(path)
    try:
        updater.replace_sheet_data("Sheet", [["a\x01b"]])
        path.write_bytes(updater.to_bytes())
    finally:
        updater.close()
    with zipfile.ZipFile(path) as archive:
        package = "".join(archive.read(name).decode("utf-8") for name in archive.namelist() if name.endswith(".xml"))
    assert "\x01" not in package
    assert "_x0001_" in package
    assert _first_cell(path) == "a\x01b"


def test_xlsx_updater_rejects_text_over_excel_limit(tmp_path):
    path = _save_plain(tmp_path / "long.xlsx")
    updater = XlsxUpdater(path)
    try:
        with pytest.raises(ValueError):
            updater.replace_sheet_data("Sheet", [["x" * 40000]])
    finally:
        updater.close()


@pytest.mark.parametrize("value", [math.nan, math.inf, -math.inf])
def test_xlsx_updater_rejects_non_finite_numbers(tmp_path, value):
    path = _save_plain(tmp_path / "nan.xlsx")
    updater = XlsxUpdater(path)
    try:
        with pytest.raises(ValueError):
            updater.replace_sheet_data("Sheet", [[value]])
    finally:
        updater.close()


def test_xlsb_updater_keeps_control_characters_and_rejects_nan(tmp_path):
    path = tmp_path / "ctrl.xlsb"
    workbook = Workbook()
    workbook.active["A1"] = "x"
    workbook.save(path)
    workbook.close()

    updater = XlsbUpdater(path)
    try:
        with pytest.raises(ValueError):
            updater.replace_sheet_data("Sheet", [[math.nan]])
        updater.replace_sheet_data("Sheet", [["a\x01b"]])
        path.write_bytes(updater.to_bytes())
    finally:
        updater.close()
    assert _first_cell(path) == "a\x01b"


MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def parts(path):
    with zipfile.ZipFile(path) as zf:
        return {name: zf.read(name) for name in zf.namelist()}


def xml_root(data: bytes):
    return ET.fromstring(data)


def xlsb_records(data: bytes):
    """Yield ``(record_id, payload)`` from a BIFF12 stream."""

    def varint(i):
        value = shift = 0
        while True:
            byte = data[i]
            i += 1
            value |= (byte & 0x7F) << shift
            shift += 7
            if not byte & 0x80:
                return value, i

    i = 0
    while i < len(data):
        record_id, i = varint(i)
        size, i = varint(i)
        yield record_id, data[i:i + size]
        i += size


def xlsb_sheet(path, number=1):
    return parts(path)[f"xl/worksheets/sheet{number}.bin"]


ROWS = [["h1", "h2", "h3"], [1, 2, 3], [4, 5, 6]]


def _xlsb_base(tmp_path):
    base = tmp_path / "base.xlsb"
    with XlsbWriter(base) as wb:
        wb.add_sheet("S")
        wb.write_sheet(ROWS)
    return base


def _dimension_and_selection(data):
    dimension = selection = None
    for record_id, payload in xlsb_records(data):
        if record_id == 0x94:
            dimension = struct.unpack("<iiii", payload[:16])
        if record_id == 0x98:
            selection = payload
    return dimension, selection


@pytest.mark.parametrize("streaming", [False, True])
def test_xlsb_replace_writes_dimension_and_leaves_selection(tmp_path, streaming):
    base = _xlsb_base(tmp_path)
    _, selection_before = _dimension_and_selection(xlsb_sheet(base))
    new_rows = [[r * 10 + c for c in range(6)] for r in range(10)]
    out = tmp_path / "out.xlsb"
    updater = XlsbUpdater(base)
    if streaming:
        updater.replace_sheet_data_stream("S", new_rows)
    else:
        updater.replace_sheet_data("S", new_rows)
    updater.save(out)

    dimension, selection_after = _dimension_and_selection(xlsb_sheet(out))
    assert dimension == (0, 9, 0, 5)
    assert selection_after == selection_before


CALC_CHAIN = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    b'<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    b'<c r="B2" i="1"/><c r="B3" i="1"/></calcChain>'
)


def _with_calc_chain(source, target):
    with zipfile.ZipFile(source) as zin, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/_rels/workbook.xml.rels":
                data = data.replace(
                    b"</Relationships>",
                    b'<Relationship Id="rIdCalc" Type="http://schemas.openxmlformats.org/'
                    b'officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/></Relationships>',
                )
            if item.filename == "[Content_Types].xml":
                data = data.replace(
                    b"</Types>",
                    b'<Override PartName="/xl/calcChain.xml" ContentType="application/'
                    b'vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/></Types>',
                )
            zout.writestr(item.filename, data)
        zout.writestr("xl/calcChain.xml", CALC_CHAIN)


@pytest.mark.parametrize("streaming", [False, True])
def test_replacing_formulas_drops_the_calc_chain(tmp_path, streaming):
    base = tmp_path / "base.xlsx"
    with XlsxWriter(base) as wb:
        wb.add_sheet("S")
        wb.write_sheet([["h1", "h2"], [1, Formula("A2*2")], [2, Formula("A3*2")]])
    chained = tmp_path / "chained.xlsx"
    _with_calc_chain(base, chained)

    out = tmp_path / "out.xlsx"
    updater = XlsxUpdater(chained)
    if streaming:
        updater.replace_sheet_data_stream("S", [["h1", "h2"], [5, 6]])
    else:
        updater.replace_sheet_data("S", [["h1", "h2"], [5, 6]])
    updater.save(out)

    package = parts(out)
    assert "xl/calcChain.xml" not in package
    assert b"calcChain" not in package["xl/_rels/workbook.xml.rels"]
    assert b"<f>" not in package["xl/worksheets/sheet1.xml"]


@pytest.mark.parametrize("streaming", [False, True])
def test_xlsx_replace_writes_formula_with_cached_value(tmp_path, streaming):
    base = tmp_path / "base.xlsx"
    with XlsxWriter(base) as wb:
        wb.add_sheet("S")
        wb.write_sheet([["h1", "h2"], [1, 2]])
    out = tmp_path / "out.xlsx"
    updater = XlsxUpdater(base)
    rows = [["h1", "h2"], [7, Formula("A2*3", value=21)], [8, Formula('"x"&"y"', value="xy")]]
    if streaming:
        updater.replace_sheet_data_stream("S", rows)
    else:
        updater.replace_sheet_data("S", rows)
    updater.save(out)

    # openpyxl reads it back; the wheel smoke job installs only the wheel and pytest.
    openpyxl = pytest.importorskip("openpyxl")
    formulas = openpyxl.load_workbook(out)["S"]
    cached = openpyxl.load_workbook(out, data_only=True)["S"]
    assert formulas["B2"].value == "=A2*3"
    assert cached["B2"].value == 21
    assert formulas["B3"].value == '="x"&"y"'
    assert cached["B3"].value == "xy"


def test_xlsb_replace_writes_formulas_instead_of_their_text(tmp_path):
    base = _xlsb_base(tmp_path)
    updater = XlsbUpdater(base)
    try:
        updater.replace_sheet_data("S", [["a", Formula("1+1", 2)]])
        updater.save(base)
    finally:
        updater.close()
    with ExcelReader(base) as reader:
        assert reader.read_all("S") == [["a", 2]]
    with ExcelReader(base) as reader:
        assert reader.read_all("S", data_only=False) == [["a", "=1+1"]]


def test_rows_of_a_self_closing_sheetdata_go_inside_it(tmp_path):
    path = tmp_path / "blank.xlsx"
    with XlsxWriter(path) as wb:
        wb.add_sheet("Blank")  # never written: <sheetData/>
    wb = load_workbook(path)
    wb["Blank"].row_dimensions[2].height = 30
    out = tmp_path / "blank_out.xlsx"
    wb.save(out)
    root = xml_root(parts(out)["xl/worksheets/sheet1.xml"])
    sheet_data = root.find(f"{MAIN}sheetData")
    assert [row.get("r") for row in sheet_data] == ["2"]


def test_dimension_follows_sheetpr_when_the_updater_adds_it(tmp_path):
    base = tmp_path / "base.xlsx"
    with XlsxWriter(base) as wb:
        wb.add_sheet("S")
        wb.write_sheet([["a", "b"], [1, 2]], tab_color="#FF0000")
    stripped = tmp_path / "no_dimension.xlsx"
    with zipfile.ZipFile(base) as zin, zipfile.ZipFile(stripped, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                data = re.sub(rb"<dimension[^>]*/>", b"", data)
            zout.writestr(item.filename, data)

    for method in ("replace_sheet_data", "replace_sheet_data_stream"):
        out = tmp_path / f"{method}.xlsx"
        updater = XlsxUpdater(stripped)
        getattr(updater, method)("S", [["a", "b"], [3, 4]])
        updater.save(out)
        names = [child.tag.replace(MAIN, "") for child in xml_root(parts(out)["xl/worksheets/sheet1.xml"])]
        assert names.index("sheetPr") < names.index("dimension") < names.index("sheetData"), method
