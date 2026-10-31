"""Editing worksheet/workbook metadata through the model and reading it back."""

from __future__ import annotations

import openpyxl
import pytest

from ._util import close_quietly, golden_path, xlspy_load


def test_metadata_edits_survive_save(tmp_path):
    workbook = xlspy_load(golden_path("metadata.xlsx"))
    output = tmp_path / "edited.xlsx"
    try:
        sheet = workbook["Meta"]
        sheet.sheet_format.defaultColWidth = 20
        sheet.sheet_properties.codeName = "EditedCode"
        sheet.sheet_view.showZeros = True
        sheet.page_setup.firstPageNumber = 3
        sheet.print_options.headings = False
        sheet.HeaderFooter.differentFirst = True
        sheet.evenHeader.center.text = "Even header"
        sheet.protection.formatCells = False
        initial_breaks = len(sheet.row_breaks.brk)
        sheet.row_breaks.append()
        sheet.column_dimensions["A"].bestFit = True
        sheet.auto_filter.add_filter_column(1, ["10"])
        sheet.auto_filter.add_sort_condition("C2:C5", descending=True)

        workbook.calculation.calcMode = "manual"
        workbook.security.lockWindows = False
        workbook.views[0].showSheetTabs = False
        workbook.save(output)
    finally:
        close_quietly(workbook)

    book = openpyxl.load_workbook(str(output), data_only=False)
    try:
        sheet = book["Meta"]
        assert sheet.sheet_format.defaultColWidth == 20
        assert sheet.sheet_properties.codeName == "EditedCode"
        assert sheet.sheet_view.showZeros is True
        assert sheet.page_setup.firstPageNumber == 3
        assert sheet.print_options.headings is False
        assert sheet.HeaderFooter.differentFirst is True
        assert sheet.evenHeader.center.text == "Even header"
        assert sheet.protection.formatCells is False
        assert len(sheet.row_breaks.brk) == initial_breaks + 1
        assert sheet.column_dimensions["A"].bestFit is True
        added_column = next(
            item for item in sheet.auto_filter.filterColumn if item.colId == 1
        )
        filter_values = [
            getattr(item, "val", item) for item in added_column.filters.filter
        ]
        assert filter_values == ["10"]
        assert sheet.auto_filter.sortState.sortCondition[0].descending is True
        assert book.calculation.calcMode == "manual"
        assert book.security.lockWindows is False
        assert book.views[0].showSheetTabs is False
    finally:
        book.close()

    reloaded = xlspy_load(output)
    try:
        sheet = reloaded["Meta"]
        assert sheet.sheet_format.defaultColWidth == 20
        assert sheet.sheet_properties.codeName == "EditedCode"
        assert sheet.sheet_view.showZeros is True
        assert sheet.page_setup.firstPageNumber == 3
        assert sheet.print_options.headings is False
        assert sheet.HeaderFooter.differentFirst is True
        assert sheet.evenHeader.center.text == "Even header"
        assert sheet.protection.formatCells is False
        assert len(sheet.row_breaks.brk) == initial_breaks + 1
        assert sheet.column_dimensions["A"].bestFit is True
        added_column = next(
            item for item in sheet.auto_filter.filterColumn if item.colId == 1
        )
        assert added_column.vals == ["10"]
        assert sheet.auto_filter.sortState.sortCondition[0].descending is True
        assert reloaded.calculation.calcMode == "manual"
        assert reloaded.security.lockWindows is False
        assert reloaded.views[0].showSheetTabs is False
    finally:
        close_quietly(reloaded)


def test_defined_name_edit_survives_save(tmp_path):
    workbook = xlspy_load(golden_path("metadata.xlsx"))
    output = tmp_path / "names.xlsx"
    try:
        workbook.defined_names.add(xlspy_defined_name("AddedName", "'Meta'!$D$4"))
        workbook.save(output)
    finally:
        close_quietly(workbook)

    book = openpyxl.load_workbook(str(output), data_only=False)
    try:
        assert book.defined_names["AddedName"].attr_text == "'Meta'!$D$4"
    finally:
        book.close()


def xlspy_defined_name(name, formula):
    from xlspy import DefinedName

    return DefinedName(name, formula)
