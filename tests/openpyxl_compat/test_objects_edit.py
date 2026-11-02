"""Editing and removing tables, validations, CF and named styles."""

from __future__ import annotations

from dataclasses import replace

import openpyxl

from xlspy import (
    DataValidation,
    DifferentialStyle,
    Font,
    FormulaRule,
    NamedStyle,
    PatternFill,
)

from ._util import close_quietly, golden_path, xlspy_load


def test_table_style_edit(tmp_path):
    workbook = xlspy_load(golden_path("objects.xlsx"))
    output = tmp_path / "table_edit.xlsx"
    try:
        sheet = workbook["Objects"]
        table = sheet.tables["ObjectsTable"]
        table.tableStyleInfo.showRowStripes = False
        workbook.save(output)
    finally:
        close_quietly(workbook)

    book = openpyxl.load_workbook(str(output), data_only=False)
    try:
        table = book["Objects"].tables["ObjectsTable"]
        assert table.tableStyleInfo.showRowStripes is False
    finally:
        book.close()


def test_table_remove(tmp_path):
    workbook = xlspy_load(golden_path("objects.xlsx"))
    output = tmp_path / "tables.xlsx"
    try:
        sheet = workbook["Objects"]
        sheet.tables.remove("ObjectsTable")
        workbook.save(output)
    finally:
        close_quietly(workbook)

    book = openpyxl.load_workbook(str(output), data_only=False)
    try:
        assert "ObjectsTable" not in book["Objects"].tables
    finally:
        book.close()


def test_validation_edit_add_remove(tmp_path):
    workbook = xlspy_load(golden_path("objects.xlsx"))
    output = tmp_path / "validations.xlsx"
    try:
        sheet = workbook["Objects"]
        sheet.data_validations[0] = replace(
            sheet.data_validations[0], error_title="Changed title"
        )
        sheet.data_validations.remove(sheet.data_validations[1])
        sheet.add_data_validation(
            DataValidation("list", "F2:F10", '"Yes,No"', allow_blank=True)
        )
        workbook.save(output)
    finally:
        close_quietly(workbook)

    book = openpyxl.load_workbook(str(output), data_only=False)
    try:
        rules = list(book["Objects"].data_validations.dataValidation)
        assert len(rules) == 3
        assert any(rule.errorTitle == "Changed title" for rule in rules)
        assert any(str(rule.sqref) == "F2:F10" for rule in rules)
    finally:
        book.close()


def test_conditional_formatting_edit_add_remove(tmp_path):
    workbook = xlspy_load(golden_path("objects.xlsx"))
    output = tmp_path / "formatting.xlsx"
    try:
        sheet = workbook["Objects"]
        sheet.conditional_formatting.remove("B2:B5")
        sheet.conditional_formatting.add(
            "E2:E5",
            FormulaRule(
                "E2:E5", ["E2>0"],
                dxf=DifferentialStyle(font=Font(bold=True)),
            ),
        )
        workbook.save(output)
    finally:
        close_quietly(workbook)

    book = openpyxl.load_workbook(str(output), data_only=False)
    try:
        ranges = {str(item.sqref) for item in book["Objects"].conditional_formatting}
        assert "B2:B5" not in ranges
        assert "E2:E5" in ranges
        rules = [
            rule
            for item in book["Objects"].conditional_formatting
            for rule in item.rules
            if str(item.sqref) == "E2:E5"
        ]
        assert rules and rules[0].type == "expression"
        # Loaded rules are rewritten from the model: their formulas must survive.
        for item in book["Objects"].conditional_formatting:
            for rule in item.rules:
                if rule.type in ("cellIs", "expression"):
                    assert rule.formula, (str(item.sqref), rule.type)
    finally:
        book.close()


def test_loaded_named_style_can_be_applied(tmp_path):
    workbook = xlspy_load(golden_path("objects.xlsx"))
    output = tmp_path / "loaded_style.xlsx"
    try:
        sheet = workbook["Objects"]
        sheet["B2"].style = "Fancy"
        workbook.save(output)
    finally:
        close_quietly(workbook)

    book = openpyxl.load_workbook(str(output), data_only=False)
    try:
        assert book["Objects"]["B2"].font.bold is True
    finally:
        book.close()


def test_named_style_is_written_and_applied(tmp_path):
    workbook = xlspy_load(golden_path("objects.xlsx"))
    output = tmp_path / "named_style.xlsx"
    try:
        sheet = workbook["Objects"]
        style = NamedStyle(
            name="MyStyle",
            font=Font(bold=True, color="FF0000"),
            fill=PatternFill(fill_type="solid", start_color="FFFF00"),
        )
        workbook.add_named_style(style)
        sheet["A1"].style = "MyStyle"
        workbook.save(output)
    finally:
        close_quietly(workbook)

    book = openpyxl.load_workbook(str(output), data_only=False)
    try:
        assert "MyStyle" in [item if isinstance(item, str) else item.name for item in book.named_styles]
        cell = book["Objects"]["A1"]
        assert cell.font.bold is True
        assert cell.font.color.rgb == "00FF0000" or cell.font.color.rgb == "FFFF0000"
        assert cell.fill.fill_type == "solid"
    finally:
        book.close()

    reloaded = xlspy_load(output)
    try:
        assert "MyStyle" in reloaded.named_styles.names
        assert reloaded["Objects"]["A1"].font.bold is True
    finally:
        close_quietly(reloaded)
