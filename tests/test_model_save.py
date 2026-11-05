"""Saving the workbook model: what edits and new sheets keep."""
from __future__ import annotations

import re
import zipfile

import openpyxl
import pytest
from helpers import part_text

from xlspy import Font, IgnoredError, NamedStyle, Table, Workbook, load_workbook
from xlspy.styles.numbers import BUILTIN_FORMATS
from xlspy.workbook._model import Break, RowBreak


def _save_plain(path):
    workbook = Workbook()
    workbook.active["A1"] = "x"
    workbook.save(path)
    workbook.close()
    return path


def test_custom_autofilter_survives_column_width_edit(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")
    from openpyxl.worksheet.filters import CustomFilter, CustomFilters, FilterColumn

    source = tmp_path / "filters.xlsx"
    book = openpyxl.Workbook()
    sheet = book.active
    sheet.append(["n"])
    for value in range(1, 6):
        sheet.append([value])
    sheet.auto_filter.ref = "A1:A6"
    sheet.auto_filter.filterColumn.append(FilterColumn(
        colId=0,
        customFilters=CustomFilters(customFilter=[CustomFilter(operator="greaterThan", val="2")]),
    ))
    book.save(source)

    out = tmp_path / "out.xlsx"
    workbook = load_workbook(source)
    try:
        workbook.active.column_dimensions["A"].width = 20
        workbook.save(out)
    finally:
        workbook.close()

    xml = part_text(out, "xl/worksheets/sheet1.xml")
    assert "customFilter" in xml
    assert 'operator="greaterThan"' in xml


def test_style_only_column_survives_column_width_edit(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")
    from openpyxl.styles import Font

    source = tmp_path / "col_style.xlsx"
    book = openpyxl.Workbook()
    book.active["A1"] = "a"
    book.active["B1"] = "b"
    book.active.column_dimensions["B"].font = Font(bold=True)
    book.save(source)

    out = tmp_path / "out.xlsx"
    workbook = load_workbook(source)
    try:
        workbook.active.column_dimensions["A"].width = 25
        workbook.save(out)
    finally:
        workbook.close()

    xml = part_text(out, "xl/worksheets/sheet1.xml")
    cols = xml[xml.index("<cols>"):xml.index("</cols>")]
    assert 'min="2" max="2"' in cols
    assert ' style="' in cols


def test_new_sheet_in_loaded_workbook_keeps_protection_breaks_and_ignored_errors(tmp_path):
    source = _save_plain(tmp_path / "base.xlsx")
    out = tmp_path / "out.xlsx"

    workbook = load_workbook(source)
    try:
        sheet = workbook.create_sheet("New")
        sheet["A1"] = 1
        sheet.protection.sheet = True
        sheet.ignored_errors.append(IgnoredError("A1", numberStoredAsText=True))
        sheet.row_breaks = RowBreak([Break(id=3)])
        sheet.sheet_format.defaultRowHeight = 20
        workbook.save(out)
    finally:
        workbook.close()

    reloaded = load_workbook(out)
    try:
        sheet = reloaded["New"]
        assert sheet.protection.sheet is True
        assert [item.sqref for item in sheet.ignored_errors] == ["A1"]
        assert [item.id for item in sheet.row_breaks.brk] == [3]
        assert sheet.sheet_format.defaultRowHeight == 20
    finally:
        reloaded.close()


def test_new_workbook_saves_table_objects(tmp_path):
    path = tmp_path / "table.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    sheet["A1"] = "Item"
    sheet["B1"] = "Qty"
    sheet["A2"] = "x"
    sheet["B2"] = 1
    sheet.add_table(Table(displayName="Sales1", ref="A1:B2"))
    workbook.save(path)
    workbook.close()

    reloaded = load_workbook(path)
    try:
        assert "Sales1" in [table.displayName for table in reloaded["Data"].tables]
    finally:
        reloaded.close()


@pytest.mark.parametrize("extension", ["xlsx", "xlsb"])
def test_named_style_on_new_workbook_is_written_and_applied(tmp_path, extension):
    path = tmp_path / f"new.{extension}"
    workbook = Workbook()
    workbook.active["A1"] = 1.25
    workbook.add_named_style(
        NamedStyle(name="Mine", font=Font(bold=True), number_format="0.00")
    )
    workbook.add_named_style(NamedStyle(name="Unused", font=Font(italic=True)))
    workbook.active["A1"].style = "Mine"
    workbook.save(path)
    workbook.close()

    reloaded = load_workbook(path)
    try:
        assert {"Mine", "Unused"} <= {style.name for style in reloaded.named_styles}
        assert reloaded.active["A1"].font.bold is True
        assert reloaded.active["A1"].number_format == "0.00"
    finally:
        reloaded.close()
    if extension == "xlsx":
        reference = openpyxl.load_workbook(path)
        assert {"Mine", "Unused"} <= set(reference.named_styles)
        assert reference.active["A1"].font.bold is True
        assert reference.active["A1"].number_format == "0.00"


def test_named_style_on_loaded_workbook_is_written(tmp_path):
    source = _save_plain(tmp_path / "src.xlsx")
    out = tmp_path / "out.xlsx"
    workbook = load_workbook(source)
    try:
        workbook.add_named_style(NamedStyle(name="Mine"))
        workbook.save(out)
    finally:
        workbook.close()
    assert "Mine" in part_text(out, "xl/styles.xml")


def test_builtin_accounting_format_survives_a_font_edit(tmp_path):
    fmt = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'
    base = tmp_path / "base.xlsx"
    book = openpyxl.Workbook()
    book.active["A1"] = 1234.5
    book.active["A1"].number_format = fmt
    book.save(base)
    # Re-point the custom format at built-in id 44, as Excel writes accounting formats.
    builtin = tmp_path / "builtin.xlsx"
    with zipfile.ZipFile(base) as zin, zipfile.ZipFile(builtin, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/styles.xml":
                text = re.sub(rb"<numFmts[^>]*>.*?</numFmts>", b"", data, flags=re.S)
                data = text.replace(b'numFmtId="164"', b'numFmtId="44"')
            zout.writestr(item.filename, data)

    # The built-in spelling of id 44 (openpyxl's), not the custom string written above.
    builtin_44 = BUILTIN_FORMATS[44]
    loaded = load_workbook(builtin)
    assert loaded.active["A1"].number_format == builtin_44
    loaded.active["A1"].font = Font(bold=True)
    out = tmp_path / "out.xlsx"
    loaded.save(out)
    assert openpyxl.load_workbook(out).active["A1"].number_format == builtin_44
