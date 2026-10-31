"""Sparklines in XLSB output, checked against the records Excel writes.

The expected payloads come from workbooks Excel 16 saved as XLSB
(``tools/capture_xlsb.py``): a line sparkline in E2 over Sheet1!B2:D2 with
Excel's defaults, then one option changed at a time.
"""

import os
import struct
import zipfile

import pytest

from xlspy import SparklineGroup, XlsbWriter
from xlspy._xlsb_sparklines import _group_payload
from xlspy.biff_utils import iter_records
from xlspy.workbook import load_workbook

ROWS = [
    ["Region", "Q1", "Q2", "Q3", "Trend"],
    ["North", 3, -1, 4, None],
    ["South", 1, 5, -9, None],
]

# BrtBeginSparklineGroup and BrtSparkline exactly as Excel saved the default group.
EXCEL_GROUP = bytes.fromhex(
    "00000000 0218"
    "05ff0000376092ff" "05ff0000d00000ff" "05ff0000000000ff" "05ff0000d00000ff"
    "05ff0000d00000ff" "05ff0000d00000ff" "05ff0000d00000ff" "05ff0000d00000ff"
    "0000000000000000" "0000000000000000" "000000000000e83f" "0000" "0000"
)
EXCEL_SPARKLINE = bytes.fromhex(
    "06000000 01000000 02000000 01000000"
    "01000000 01000000 04000000 04000000"
    "01000000 02000000 0f000000 00000000"
    "3b0000 01000000 01000000 01c0 03c0"
)


def _write(path, groups, rows=ROWS, sheet="Sheet1"):
    with XlsbWriter(path) as writer:
        writer.add_sheet(sheet)
        writer.write_sheet(rows, sparklines=groups)
    return path


def _records(path, part="xl/worksheets/sheet1.bin"):
    with zipfile.ZipFile(path) as archive:
        data = archive.read(part)
    return [(record.record_id, data[record.data_start:record.data_end]) for record in iter_records(data)]


def test_default_group_matches_excel_byte_for_byte(tmp_path):
    records = _records(_write(tmp_path / "s.xlsb", [SparklineGroup(data="B2:D2", location="E2")]))
    payloads = dict(records)
    assert payloads[0x411] == EXCEL_GROUP
    assert payloads[0x413] == EXCEL_SPARKLINE


def test_block_order_and_position(tmp_path):
    records = _records(_write(tmp_path / "s.xlsb", [SparklineGroup(data="B2:D3", location="E2:E3")]))
    ids = [record_id for record_id, _ in records]
    start = ids.index(0x422)
    assert records[start - 1] == (0x23, b"\x02\x0e\x00\x00")  # BrtFRTBegin, Excel 2010 block
    assert ids[start:start + 9] == [0x422, 0x411, 0x420, 0x413, 0x413, 0x421, 0x412, 0x423, 0x24]
    assert ids[-1] == 0x82  # BrtEndSheet still closes the sheet
    second = records[start + 4][1]
    assert struct.unpack_from("<II", second, 16) == (2, 2)  # E3: row index 2
    assert second[51] == 2  # data row 3


@pytest.mark.parametrize(("options", "offset", "value"), [
    ({"kind": "column"}, 94, 1),
    ({"kind": "win_loss"}, 94, 2),
    ({"markers": True}, 4, 0x0A),
    ({"high_point": True}, 4, 0x12),
    ({"low_point": True}, 4, 0x22),
    ({"first_point": True}, 4, 0x42),
    ({"last_point": True}, 4, 0x82),
    ({"empty_cells": "zero"}, 4, 0x00),
    ({"empty_cells": "connect"}, 4, 0x04),
    ({"negative_points": True}, 5, 0x19),
    ({"show_axis": True}, 5, 0x1A),
    ({"show_hidden": True}, 5, 0x1C),
    ({"right_to_left": True}, 5, 0x98),
    ({"axis_min": -5}, 5, 0x08),
    ({"axis_min": "group", "axis_max": "group"}, 5, 0x60),
])
def test_options_set_the_bytes_excel_sets(options, offset, value):
    payload = _group_payload(SparklineGroup(data="B2:D2", location="E2", **options))
    assert len(payload) == 98
    assert payload[offset] == value
    # Bytes 70..93 hold the manual axis values and line weight, checked below.
    changed = [index for index in range(98)
               if not 70 <= index < 94 and payload[index] != EXCEL_GROUP[index]]
    expected = [] if EXCEL_GROUP[offset] == value else [offset]
    assert changed == expected


def test_manual_axis_values_colour_and_line_weight():
    payload = _group_payload(SparklineGroup(data="B2:D2", location="E2", axis_min=-5, axis_max=50,
                                            color="#112233", line_width=2.25))
    assert payload[5] == 0x00
    assert struct.unpack_from("<ddd", payload, 70) == (50.0, -5.0, 2.25)
    assert payload[6:14] == bytes.fromhex("05ff0000112233ff")


def test_data_on_another_sheet_uses_the_extern_sheet_table(tmp_path):
    path = tmp_path / "s.xlsb"
    with XlsbWriter(path) as writer:
        writer.add_sheet("Data sheet")
        writer.write_sheet(ROWS)
        writer.add_sheet("Summary")
        writer.write_sheet([["x"]], sparklines=[SparklineGroup(data="'Data sheet'!B2:D2", location="B2")])
    sparkline = dict(_records(path, "xl/worksheets/sheet2.bin"))[0x413]
    ixti = struct.unpack_from("<H", sparkline, 49)[0]
    externs = dict(_records(path, "xl/workbook.bin"))[0x16A]
    count = struct.unpack_from("<I", externs)[0]
    assert ixti < count
    assert struct.unpack_from("<Iii", externs, 4 + ixti * 12)[1:] == (0, 0)  # the first sheet


def test_bad_shape_is_rejected_before_the_sheet_is_written(tmp_path):
    with XlsbWriter(tmp_path / "s.xlsb") as writer:
        writer.add_sheet("Sheet1")
        with pytest.raises(ValueError, match="one data row each"):
            writer.write_sheet(ROWS, sparklines=[SparklineGroup(data="B2:D3", location="E2:E4")])


def test_loaded_workbook_keeps_the_sparklines_on_save(tmp_path):
    source = _write(tmp_path / "s.xlsb", [SparklineGroup(data="B2:D3", location="E2:E3", markers=True)])
    workbook = load_workbook(source)
    try:
        workbook.worksheets[0].cell(row=4, column=1).value = "edited"
        workbook.save(tmp_path / "saved.xlsb")
    finally:
        workbook.close()
    before = [record for record in _records(source) if record[0] in (0x411, 0x413)]
    after = [record for record in _records(tmp_path / "saved.xlsb") if record[0] in (0x411, 0x413)]
    assert after == before


@pytest.mark.skipif(os.environ.get("XLSPY_RUN_EXCEL_COM") != "1",
                    reason="Set XLSPY_RUN_EXCEL_COM=1 to run Microsoft Excel COM validation")
def test_excel_reads_xlsb_sparklines(tmp_path):
    win32 = pytest.importorskip("win32com.client")
    path = tmp_path / "com.xlsb"
    with XlsbWriter(path) as writer:
        writer.add_sheet("Data sheet")
        writer.write_sheet(ROWS, sparklines=[
            SparklineGroup(data="B2:D3", location="E2:E3", markers=True, color="#112233",
                           first_color="#550000", low_color="#880000", line_width=2.25,
                           empty_cells="zero", axis_min=-5, axis_max="group", show_axis=True),
            SparklineGroup(data="B2:D2", location="F2", kind="win_loss", right_to_left=True),
        ])
        writer.add_sheet("Second")
        writer.write_sheet([["x"]], sparklines=[SparklineGroup(data="'Data sheet'!B2:D2", location="B2")])
    excel = win32.DispatchEx("Excel.Application")
    excel.DisplayAlerts = False
    try:
        book = excel.Workbooks.Open(str(path))
        group = book.Worksheets(1).Range("E2").SparklineGroups.Item(1)
        assert (group.Type, group.Count, group.SourceData) == (1, 2, "B2:D3")
        assert group.SeriesColor.Color == 0x332211  # BGR
        assert group.Points.Firstpoint.Color.Color == 0x000055
        assert group.Points.Lowpoint.Color.Color == 0x000088
        assert group.Points.Markers.Visible
        assert group.LineWeight == 2.25
        assert group.DisplayBlanksAs == 2  # xlZero
        assert group.Axes.Horizontal.Axis.Visible
        assert group.Axes.Vertical.MinScaleType == 3  # xlSparkScaleCustom
        assert group.Axes.Vertical.CustomMinScaleValue == -5
        assert group.Axes.Vertical.MaxScaleType == 1  # xlSparkScaleGroup
        win_loss = book.Worksheets(1).Range("F2").SparklineGroups.Item(1)
        assert win_loss.Type == 3
        assert win_loss.Axes.Horizontal.RightToLeftPlotOrder
        other = book.Worksheets(2).Range("B2").SparklineGroups.Item(1)
        assert other.SourceData == "'Data sheet'!B2:D2"
        book.Close(SaveChanges=False)
    finally:
        excel.Quit()
