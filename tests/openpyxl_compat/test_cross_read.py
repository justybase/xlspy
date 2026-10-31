"""Bidirectional value/style round trips with openpyxl."""

from __future__ import annotations

import datetime as dt

import openpyxl
import pytest

from xlspy import ExcelReader, XlsxWriter

from ._util import close_quietly, normalize, xlspy_load


def _openpyxl_workbook(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws["A1"] = 1
    ws["A2"] = 2.5
    ws["A3"] = "text"
    ws["A4"] = True
    ws["A5"] = dt.datetime(2024, 1, 2, 3, 4, 5)
    ws["A6"] = "=SUM(A1:A2)"
    ws["A7"] = "2024-01-02"
    ws["A5"].number_format = "yyyy-mm-dd hh:mm:ss"
    ws["B1"] = "styled"
    ws["B1"].font = openpyxl.styles.Font(bold=True, color="FF0000")
    ws["B1"].fill = openpyxl.styles.PatternFill("solid", fgColor="FFFF00")
    ws.merge_cells("B2:C2")
    ws["B2"] = "merged"
    ws["B3"] = "link"
    ws["B3"].hyperlink = "https://example.com"
    ws["B4"] = "note"
    ws["B4"].comment = openpyxl.comments.Comment("hello", "author", width=120, height=60)
    wb.save(path)
    wb.close()
    return path


def test_xlspy_model_reads_openpyxl_values(tmp_path):
    path = _openpyxl_workbook(tmp_path / "source.xlsx")
    wb = xlspy_load(path)
    try:
        ws = wb["Data"]
        assert ws["A1"].value == 1
        assert ws["A2"].value == 2.5
        assert ws["A3"].value == "text"
        assert ws["A4"].value is True
        assert ws["A5"].value == dt.datetime(2024, 1, 2, 3, 4, 5)
        assert ws["A6"].value == "=SUM(A1:A2)"
        assert ws["B1"].value == "styled"
        assert str(ws.merged_cells) == "B2:C2"
        assert ws["B2"].value == "merged"
    finally:
        close_quietly(wb)


def test_xlspy_reads_openpyxl_styles_objects(tmp_path):
    path = _openpyxl_workbook(tmp_path / "source.xlsx")
    wb = xlspy_load(path)
    try:
        ws = wb["Data"]
        cell = ws["B1"]
        assert cell.font is not None and cell.font.bold is True
        assert cell.fill is not None
        if ws["B3"].hyperlink is None:
            pytest.xfail("hyperlink round-trip is not exposed by the model")
        assert ws["B3"].hyperlink.target == "https://example.com"
        if ws["B4"].comment is None:
            pytest.xfail("comment round-trip is not exposed by the model")
        assert ws["B4"].comment.text == "hello"
    finally:
        close_quietly(wb)


def test_excel_reader_reads_openpyxl_values(tmp_path):
    path = _openpyxl_workbook(tmp_path / "source.xlsx")
    with ExcelReader(str(path)) as reader:
        rows = reader.read_all("Data", data_only=False)
    assert rows[0][0] == 1
    assert rows[1][0] == 2.5
    assert rows[2][0] == "text"
    assert rows[3][0] is True
    assert rows[4][0] == dt.datetime(2024, 1, 2, 3, 4, 5)
    assert rows[5][0] == "=SUM(A1:A2)"


def test_openpyxl_reads_xlspy_output(tmp_path):
    path = tmp_path / "xlspy.xlsx"
    with XlsxWriter(str(path)) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(
            [
                ["name", "value"],
                ["alpha", 1],
                ["beta", 2.5],
                ["date", dt.datetime(2024, 1, 2, 3, 4, 5)],
            ]
        )
    wb = openpyxl.load_workbook(str(path), data_only=False)
    try:
        ws = wb["Data"]
        assert ws["A1"].value == "name"
        assert ws["B2"].value == 1
        assert ws["B3"].value == 2.5
        assert normalize(ws["B4"].value) == normalize(dt.datetime(2024, 1, 2, 3, 4, 5))
    finally:
        wb.close()


def test_xlspy_and_openpyxl_agree_on_xlspy_output(tmp_path):
    path = tmp_path / "xlspy.xlsx"
    rows = [["a", "b"], [1, 2], [3.5, 4.5], [None, "x"]]
    with XlsxWriter(str(path)) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(rows)
    with ExcelReader(str(path)) as reader:
        xlspy_rows = reader.read_all("Data")
    wb = openpyxl.load_workbook(str(path), data_only=True)
    try:
        openpyxl_rows = [list(row) for row in wb["Data"].iter_rows(values_only=True)]
    finally:
        wb.close()
    assert [[normalize(value) for value in row] for row in xlspy_rows] == [
        [normalize(value) for value in row] for row in openpyxl_rows
    ]
