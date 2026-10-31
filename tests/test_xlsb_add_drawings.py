"""Charts and images can be added to a loaded XLSB workbook."""

import io
import struct
import zipfile
import zlib

import pytest

from xlspy import Chart, ChartSeries, Image, Workbook, load_workbook


def _png(width=8, height=4):
    raw = b"".join(b"\x00" + bytes((0x10, 0x80, 0x10)) * width for _ in range(height))

    def chunk(kind, data):
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data))

    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))


def _chart(title):
    return Chart("bar", [ChartSeries("B2:B4", categories="A2:A4")], title=title)


@pytest.fixture(params=["plain", "with_chart"])
def source(request, tmp_path):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Report"
    for row in (["name", "value"], ["a", 1], ["b", 2], ["c", 3]):
        sheet.append(row)
    if request.param == "with_chart":
        sheet.add_chart(_chart("Existing"), "D2")
    path = tmp_path / "source.xlsb"
    workbook.save(path)
    workbook.close()
    return path, request.param == "with_chart"


def test_chart_and_image_added_to_loaded_xlsb(source, tmp_path):
    path, has_chart = source
    out = tmp_path / "out.xlsb"
    workbook = load_workbook(path)
    try:
        sheet = workbook.active
        sheet.add_chart(_chart("Added"), "L2")
        sheet.add_image(Image(io.BytesIO(_png())), "D20")
        workbook.save(out)
    finally:
        workbook.close()
    with zipfile.ZipFile(out) as archive:
        assert archive.testzip() is None
        names = archive.namelist()
    assert sum(name.startswith("xl/charts/chart") for name in names) == 1 + has_chart
    assert any(name.startswith("xl/media/") for name in names)
    reloaded = load_workbook(out)
    try:
        sheet = reloaded.active
        titles = sorted(chart.title for chart in sheet.charts)
        assert titles == sorted(["Added"] + (["Existing"] if has_chart else []))
        assert len(sheet.images) == 1
        assert sheet["B3"].value == 2
    finally:
        reloaded.close()


def test_chart_on_a_new_sheet(source, tmp_path):
    path, _ = source
    out = tmp_path / "new.xlsb"
    workbook = load_workbook(path)
    try:
        sheet = workbook.create_sheet("Extra")
        sheet.append(["k", "v"])
        sheet.append(["x", 1])
        sheet.append(["y", 2])
        sheet.add_chart(Chart("line", [ChartSeries("B2:B3", categories="A2:A3")], title="Fresh"), "D2")
        workbook.save(out)
    finally:
        workbook.close()
    reloaded = load_workbook(out)
    try:
        assert [chart.title for chart in reloaded["Extra"].charts] == ["Fresh"]
    finally:
        reloaded.close()


def test_added_objects_are_not_added_twice(source, tmp_path):
    path, _ = source
    workbook = load_workbook(path)
    try:
        workbook.active.add_chart(_chart("Once"), "L2")
        workbook.save(tmp_path / "a.xlsb")
        workbook.save(tmp_path / "b.xlsb")
    finally:
        workbook.close()
    reloaded = load_workbook(tmp_path / "b.xlsb")
    try:
        assert [c.title for c in reloaded.active.charts].count("Once") == 1
    finally:
        reloaded.close()
