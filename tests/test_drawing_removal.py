"""Loaded charts and images can be removed from XLSX and XLSB files."""

import io
import struct
import zipfile
import zlib

import pytest

from xlspy import Chart, ChartSeries, Image, Workbook, load_workbook


def _png(width, height, shade=0x80):
    raw = b"".join(b"\x00" + bytes((0x10, shade, 0x10)) * width for _ in range(height))

    def chunk(kind, data):
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data))

    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))


@pytest.fixture(params=["xlsx", "xlsb"])
def source(request, tmp_path):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Report"
    for row in (["name", "value"], ["a", 1], ["b", 2], ["c", 3]):
        sheet.append(row)
    sheet.add_chart(Chart("bar", [ChartSeries("B2:B4", categories="A2:A4")], title="First"), "D2")
    sheet.add_image(Image(io.BytesIO(_png(8, 4))), "D20")
    sheet.add_chart(Chart("line", [ChartSeries("B2:B4", categories="A2:A4")], title="Second"), "L2")
    path = tmp_path / f"source.{request.param}"
    workbook.save(path)
    workbook.close()
    return path


def _parts(path):
    with zipfile.ZipFile(path) as archive:
        return archive.namelist()


def test_removed_objects_leave_the_file(source, tmp_path):
    out = tmp_path / f"out{source.suffix}"
    workbook = load_workbook(source)
    try:
        sheet = workbook.active
        first = next(chart for chart in sheet.charts if chart.title == "First")
        sheet._charts.remove(first)
        sheet._images.clear()
        workbook.save(out)
    finally:
        workbook.close()
    reloaded = load_workbook(out)
    try:
        assert [chart.title for chart in reloaded.active.charts] == ["Second"]
        assert reloaded.active.images == []
        assert reloaded.active["B3"].value == 2
    finally:
        reloaded.close()
    parts = _parts(out)
    assert sum(name.startswith("xl/charts/chart") for name in parts) == 1
    assert not any(name.startswith("xl/media/") for name in parts)
    with zipfile.ZipFile(out) as archive:
        types = archive.read("[Content_Types].xml").decode()
    for name in parts:
        assert name.startswith("[") or "/" + name in types or name.rsplit(".", 1)[-1] in types


def test_remove_then_add_in_one_save(tmp_path):
    workbook = Workbook()
    workbook.active.append(["a", 1])
    workbook.active.add_image(Image(io.BytesIO(_png(4, 4))), "C3")
    workbook.save(tmp_path / "one.xlsx")
    workbook.close()

    loaded = load_workbook(tmp_path / "one.xlsx")
    try:
        loaded.active._images.clear()
        loaded.active.add_image(Image(io.BytesIO(_png(6, 2, 0x20))), "E5")
        loaded.save(tmp_path / "two.xlsx")
    finally:
        loaded.close()
    reloaded = load_workbook(tmp_path / "two.xlsx")
    try:
        images = reloaded.active.images
        assert len(images) == 1 and images[0].anchor == "E5"
    finally:
        reloaded.close()


def test_images_saved_earlier_in_the_session_cannot_be_removed(tmp_path):
    workbook = Workbook()
    workbook.active["A1"] = 1
    workbook.save(tmp_path / "plain.xlsx")
    workbook.close()
    loaded = load_workbook(tmp_path / "plain.xlsx")
    try:
        loaded.active.add_image(Image(io.BytesIO(_png(4, 4))), "C3")
        loaded.save(tmp_path / "first.xlsx")
        loaded.active._images.clear()
        with pytest.raises(NotImplementedError, match="load the saved workbook"):
            loaded.save(tmp_path / "second.xlsx")
    finally:
        loaded.close()


@pytest.mark.com
@pytest.mark.skipif(
    __import__("os").environ.get("XLSPY_RUN_EXCEL_COM") != "1", reason="Set XLSPY_RUN_EXCEL_COM=1")
@pytest.mark.parametrize("file_format", [51, 50])  # xlOpenXMLWorkbook, xlExcel12
def test_removal_from_excel_files(tmp_path, file_format):
    import win32com.client

    picture = tmp_path / "pixel.png"
    picture.write_bytes(_png(16, 16))
    suffix = ".xlsx" if file_format == 51 else ".xlsb"
    native = tmp_path / f"native{suffix}"
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.DisplayAlerts = False
    try:
        book = excel.Workbooks.Add()
        sheet = book.Worksheets(1)
        for row in range(1, 6):
            sheet.Cells(row, 1).Value = row
        sheet.Shapes.AddChart2(201, 51, 200, 10, 300, 200).Chart.SetSourceData(sheet.Range("A1:A5"))
        sheet.Shapes.AddChart2(227, 4, 200, 250, 300, 200).Chart.SetSourceData(sheet.Range("A1:A5"))
        sheet.Shapes.AddPicture(str(picture), False, True, 10, 120, 32, 32)
        book.SaveAs(str(native), file_format)
        book.Close(False)

        workbook = load_workbook(native)
        try:
            assert len(workbook.active.charts) == 2 and len(workbook.active.images) == 1
            del workbook.active._charts[0]
            workbook.active._images.clear()
            workbook.save(tmp_path / f"edited{suffix}")
        finally:
            workbook.close()

        book = excel.Workbooks.Open(str(tmp_path / f"edited{suffix}"), CorruptLoad=0)
        try:
            shapes = book.Worksheets(1).Shapes
            assert shapes.Count == 1
            assert shapes(1).HasChart
            assert book.Worksheets(1).Range("A5").Value == 5
        finally:
            book.Close(False)
    finally:
        excel.Quit()
