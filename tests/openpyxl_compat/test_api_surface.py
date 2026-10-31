"""Attribute-level API surface checks, grouped by gap ID.

Each missing attribute is reported as xfail with the ID from the parity
matrix, so this file doubles as a live gap report.  Implementing a gap turns
the corresponding check into a pass without editing the test.
"""

from __future__ import annotations

import importlib

import pytest

import xlspy

from ._util import close_quietly, golden_path, require_path, xlspy_load


def _check(obj, path: str, gap: str):
    return require_path(obj, path, gap)


def test_workbook_metadata_surface():
    wb = xlspy.Workbook()
    try:
        security = _check(wb, "security", "A3")
        for name in (
            "lockStructure", "lockWindows", "lockRevision",
            "workbookPassword", "revisionsPassword",
        ):
            _check(security, name, "A3/B11")

        calculation = _check(wb, "calculation", "A3")
        for name in (
            "calcId", "calcMode", "fullCalcOnLoad", "refMode", "iterate",
            "iterateCount", "iterateDelta", "fullPrecision", "calcCompleted",
            "calcOnSave", "concurrentCalc", "forceFullCalc",
        ):
            _check(calculation, name, "A3/B11")

        _check(wb, "views", "A3")
        _check(wb, "epoch", "A3")
        _check(wb, "template", "A5")
        _check(wb, "vba_archive", "A5")
        _check(wb, "custom_doc_props", "A11")
        _check(wb, "read_only", "A10")
    finally:
        close_quietly(wb)


def test_document_properties_surface():
    wb = xlspy.Workbook()
    try:
        _check(wb.properties, "lastPrinted", "A12")
    finally:
        close_quietly(wb)


def test_worksheet_metadata_surface():
    wb = xlspy.Workbook()
    ws = wb.active
    try:
        sheet_format = _check(ws, "sheet_format", "B2")
        for name in (
            "baseColWidth", "defaultColWidth", "defaultRowHeight",
            "customHeight", "zeroHeight", "thickTop", "thickBottom",
            "outlineLevelRow", "outlineLevelCol",
        ):
            _check(sheet_format, name, "B2")

        properties = _check(ws, "sheet_properties", "B3")
        for name in (
            "codeName", "enableFormatConditionsCalculation", "filterMode",
            "published", "syncHorizontal", "syncRef", "syncVertical",
            "transitionEvaluation", "transitionEntry", "tabColor",
            "outlinePr", "pageSetUpPr",
        ):
            _check(properties, name, "B3")
        for name in ("applyStyles", "summaryBelow", "summaryRight"):
            _check(properties.outlinePr, name, "B3")

        view = _check(ws, "sheet_view", "B1")
        for name in (
            "windowProtection", "showFormulas", "showGridLines",
            "showRowColHeaders", "showZeros", "rightToLeft", "tabSelected",
            "showRuler", "showOutlineSymbols", "defaultGridColor",
            "showWhiteSpace", "view", "topLeftCell", "colorId", "zoomScale",
            "zoomScaleNormal", "zoomScaleSheetLayoutView",
            "zoomScalePageLayoutView", "workbookViewId", "selection", "pane",
        ):
            _check(view, name, "B1")

        _check(ws, "active_cell", "B1")
        _check(ws, "selected_cell", "B1")
        _check(ws, "show_gridlines", "B1")

        setup = _check(ws, "page_setup", "B5")
        for name in (
            "orientation", "paperSize", "scale", "firstPageNumber",
            "fitToHeight", "fitToWidth", "draft", "horizontalDpi",
            "verticalDpi", "copies", "pageOrder", "useFirstPageNumber",
            "paperHeight", "paperWidth", "blackAndWhite", "cellComments",
            "errors",
        ):
            _check(setup, name, "B5")

        options = _check(ws, "print_options", "B6")
        for name in (
            "horizontalCentered", "verticalCentered", "headings",
            "gridLines", "gridLinesSet",
        ):
            _check(options, name, "B6")

        footer = _check(ws, "HeaderFooter", "B7")
        for name in (
            "differentFirst", "differentOddEven", "scaleWithDoc",
            "alignWithMargins", "oddHeader", "oddFooter", "evenHeader",
            "evenFooter", "firstHeader", "firstFooter",
        ):
            _check(footer, name, "B7")

        _check(ws, "row_breaks", "B8")
        _check(ws, "col_breaks", "B8")
        _check(ws, "print_titles", "B9")

        protection = _check(ws, "protection", "B10")
        for name in (
            "sheet", "objects", "scenarios", "formatCells", "formatRows",
            "formatColumns", "insertColumns", "insertRows", "insertHyperlinks",
            "deleteColumns", "deleteRows", "selectLockedCells", "sort",
            "autoFilter", "pivotTables", "password",
        ):
            _check(protection, name, "B10")

        auto_filter = _check(ws, "auto_filter", "B4")
        _check(auto_filter, "filterColumn", "B4")
        _check(auto_filter, "sortState", "B4")
        _check(auto_filter, "add_filter_column", "B4")
        _check(auto_filter, "add_sort_condition", "B4")

        for holder_name in ("column_dimensions", "row_dimensions"):
            holder = _check(ws, holder_name, "B13")
            _check(holder, "group", "B13")
    finally:
        close_quietly(wb)


def test_cell_surface():
    wb = xlspy.Workbook()
    cell = wb.active["A1"]
    try:
        for name in ("is_date", "encoding", "base_date", "row", "column",
                     "coordinate", "column_letter", "data_type", "internal_value"):
            _check(cell, name, "A10/D8")
    finally:
        close_quietly(wb)


def test_write_only_surface():
    try:
        from xlspy.workbook import WriteOnlyCell
    except ImportError:
        pytest.xfail("A10: WriteOnlyCell is missing")
    wb = xlspy.Workbook(write_only=True)
    ws = wb.create_sheet()
    try:
        cell = WriteOnlyCell(ws, value=1)
        cell.font = None
        ws.append([cell])
    finally:
        close_quietly(wb)


def test_read_only_cell_surface():
    try:
        wb = xlspy.load_workbook(golden_path("objects.xlsx"), read_only=True)
    except Exception as exc:
        pytest.xfail(f"A10: read-only load failed: {exc}")
    try:
        ws = wb["Objects"]
        cell = ws["B2"]
        _check(cell, "number_format", "A10")
        _check(cell, "is_date", "A10")
    finally:
        close_quietly(wb)


def test_utils_surface():
    for module in (
        "xlspy.utils.cell",
        "xlspy.utils.datetime",
        "xlspy.utils.escape",
        "xlspy.utils.units",
        "xlspy.utils.exceptions",
    ):
        try:
            importlib.import_module(module)
        except ImportError:
            pytest.xfail(f"D2: {module} is missing")


def test_styles_numbers_surface():
    try:
        numbers = importlib.import_module("xlspy.styles.numbers")
    except ImportError:
        pytest.xfail("D1: xlspy.styles.numbers is missing")
    for name in (
        "FORMAT_GENERAL", "FORMAT_TEXT", "FORMAT_DATE_YYYYMMDD2",
        "builtin_format_code", "builtin_format_id", "is_builtin",
        "is_date_format", "is_datetime", "is_timedelta_format",
    ):
        _check(numbers, name, "D1")


def test_formula_surface():
    try:
        formula = importlib.import_module("xlspy.formula")
    except ImportError:
        pytest.xfail("D3: xlspy.formula is missing")
    _check(formula, "Tokenizer", "D3")
    _check(formula, "Translator", "D3")


def test_defined_names_surface():
    wb = xlspy.Workbook()
    try:
        names = _check(wb, "defined_names", "A4")
        for name in ("add", "remove", "__getitem__"):
            _check(names, name, "A4")
    finally:
        close_quietly(wb)


def test_tables_surface():
    wb = xlspy_load(golden_path("objects.xlsx"))
    try:
        ws = wb["Objects"]
        tables = _check(ws, "tables", "C1")
        _check(tables, "__getitem__", "C1")
        try:
            table = tables["ObjectsTable"]
        except (TypeError, KeyError, IndexError):
            pytest.xfail("C1: tables are not a name-keyed collection")
        for name in (
            "displayName", "ref", "tableColumns", "tableStyleInfo",
            "totalsRowCount", "autoFilter", "sortState",
        ):
            _check(table, name, "C1")
        _check(table.tableStyleInfo, "name", "C1")
    finally:
        close_quietly(wb)


def test_data_validation_surface():
    wb = xlspy_load(golden_path("objects.xlsx"))
    try:
        ws = wb["Objects"]
        validations = _check(ws, "data_validations", "C2")
        _check(validations, "add", "C2")
        _check(validations, "remove", "C2")
        rules = list(validations)
        if not rules:
            pytest.xfail("C2: validations are not loaded from the workbook")
        rule = rules[0]
        for name in (
            "allowBlank", "showDropDown", "showErrorMessage",
            "showInputMessage", "errorStyle", "errorTitle", "promptTitle",
            "imeMode", "sqref",
        ):
            _check(rule, name, "C2")
    finally:
        close_quietly(wb)


def test_conditional_formatting_surface():
    wb = xlspy_load(golden_path("objects.xlsx"))
    try:
        ws = wb["Objects"]
        formatting = _check(ws, "conditional_formatting", "C3")
        _check(formatting, "__getitem__", "C3")
        _check(formatting, "remove", "C3")
        if len(formatting) < 3:
            pytest.xfail("C3: conditional formatting rules are not loaded")
        try:
            from xlspy.styles import Rule
        except ImportError:
            pytest.xfail("C3: generic Rule is missing")
        _check(Rule, "type", "C3")
    finally:
        close_quietly(wb)


def test_named_styles_surface():
    wb = xlspy.Workbook()
    try:
        _check(wb, "named_styles", "C4")
        _check(wb, "add_named_style", "C4")
    finally:
        close_quietly(wb)


def test_charts_surface():
    from xlspy.charts import CHART_TYPES

    for chart_type in ("radar", "bubble", "stock", "surface", "surface3d",
                       "bar3d", "line3d", "pie3d", "area3d"):
        if chart_type not in CHART_TYPES:
            pytest.xfail(f"C6: chart type {chart_type!r} is missing")

    wb = xlspy_load(golden_path("charts.xlsx"))
    try:
        counts = {ws.title: len(list(ws.charts)) for ws in wb.worksheets}
        if sum(counts.values()) != 15:
            pytest.xfail(f"C6: loaded charts {counts}")
    finally:
        close_quietly(wb)


def test_images_surface():
    wb = xlspy_load(golden_path("images.xlsx"))
    try:
        images = list(wb["Images"].images)
        if len(images) != 4:
            pytest.xfail(f"C7: loaded images {len(images)}")
    finally:
        close_quietly(wb)


def test_chartsheets_surface():
    wb = xlspy_load(golden_path("chartsheet.xlsx"))
    try:
        chartsheets = _check(wb, "chartsheets", "C8/A8")
        assert [sheet.title for sheet in chartsheets] == ["ChartSheet"]
    finally:
        close_quietly(wb)
