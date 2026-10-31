"""Semantic comparison of the xlspy model with openpyxl on the goldens."""

from __future__ import annotations


from ._util import (
    gap_xfail,
    close_quietly,
    golden_path,
    normalize,
    openpyxl_load,
    require,
    xfail_mismatch,
    xlspy_load,
)


def _pair(name):
    try:
        xwb = xlspy_load(golden_path(name))
    except Exception as exc:
        gap_xfail(f"xlspy cannot load {name}: {type(exc).__name__}: {exc}")
    owb = openpyxl_load(golden_path(name), data_only=False)
    return xwb, owb


def _close(xwb, owb):
    close_quietly(xwb)
    close_quietly(owb)


def _color(value):
    if value is None or isinstance(value, str):
        return value
    return getattr(value, "rgb", None) or getattr(value, "theme", None) or str(value)


def test_sheet_format(manifest):
    xwb, owb = _pair("metadata.xlsx")
    try:
        xs = require(xwb["Meta"], "sheet_format", "B2")
        os = owb["Meta"].sheet_format
        for name in (
            "baseColWidth", "defaultColWidth", "defaultRowHeight",
            "customHeight", "zeroHeight", "thickTop", "thickBottom",
            "outlineLevelRow", "outlineLevelCol",
        ):
            xfail_mismatch("B2", f"sheet_format.{name}", require(xs, name, "B2"), getattr(os, name))
    finally:
        _close(xwb, owb)


def test_sheet_properties(manifest):
    xwb, owb = _pair("metadata.xlsx")
    try:
        xs = xwb["Meta"].sheet_properties
        os = owb["Meta"].sheet_properties
        for name in (
            "codeName", "enableFormatConditionsCalculation", "filterMode",
            "published", "syncHorizontal", "syncRef", "syncVertical",
            "transitionEvaluation", "transitionEntry",
        ):
            xfail_mismatch("B3", f"sheet_properties.{name}", require(xs, name, "B3"), getattr(os, name))
        xfail_mismatch("B3", "tabColor", _color(require(xs, "tabColor", "B3")), _color(os.tabColor))
        xoutline = require(xs, "outlinePr", "B3")
        for name in ("applyStyles", "summaryBelow", "summaryRight"):
            xfail_mismatch(
                "B3", f"outlinePr.{name}",
                require(xoutline, name, "B3"), getattr(os.outlinePr, name),
            )
        xpage = require(xs, "pageSetUpPr", "B3")
        xfail_mismatch(
            "B3", "pageSetUpPr.fitToPage",
            require(xpage, "fitToPage", "B3"), os.pageSetUpPr.fitToPage,
        )
    finally:
        _close(xwb, owb)


def test_sheet_view(manifest):
    xwb, owb = _pair("metadata.xlsx")
    try:
        xs = xwb["Meta"].sheet_view
        os = owb["Meta"].sheet_view
        for name in (
            "showFormulas", "showGridLines", "showRowColHeaders", "showZeros",
            "rightToLeft", "tabSelected", "showOutlineSymbols",
            "defaultGridColor", "showWhiteSpace", "view", "topLeftCell",
            "colorId", "zoomScale", "zoomScaleNormal", "workbookViewId",
        ):
            xfail_mismatch("B1", f"sheet_view.{name}", require(xs, name, "B1"), getattr(os, name))
        xp = require(xs, "pane", "B1")
        op = os.pane
        if op is not None:
            for name in ("xSplit", "ySplit", "topLeftCell", "activePane", "state"):
                xfail_mismatch("B1", f"pane.{name}", require(xp, name, "B1"), getattr(op, name))
        else:
            gap_xfail("B1: openpyxl golden has no pane")
        xfail_mismatch("B1", "show_gridlines", require(xwb["Meta"], "show_gridlines", "B1"), owb["Meta"].show_gridlines)
        xfail_mismatch("B1", "active_cell", require(xwb["Meta"], "active_cell", "B1"), owb["Meta"].active_cell)
        xfail_mismatch("B1", "selected_cell", require(xwb["Meta"], "selected_cell", "B1"), owb["Meta"].selected_cell)
    finally:
        _close(xwb, owb)


def test_page_setup_and_print_options(manifest):
    xwb, owb = _pair("metadata.xlsx")
    try:
        xs = xwb["Meta"].page_setup
        os = owb["Meta"].page_setup
        for name in (
            "orientation", "paperSize", "scale", "firstPageNumber",
            "fitToHeight", "fitToWidth", "draft", "horizontalDpi",
            "verticalDpi", "copies", "pageOrder", "useFirstPageNumber",
            "paperHeight", "paperWidth", "blackAndWhite", "cellComments",
            "errors",
        ):
            xfail_mismatch("B5", f"page_setup.{name}", require(xs, name, "B5"), getattr(os, name))
        xo = xwb["Meta"].print_options
        oo = owb["Meta"].print_options
        for name in ("horizontalCentered", "verticalCentered", "headings", "gridLines", "gridLinesSet"):
            xfail_mismatch("B6", f"print_options.{name}", require(xo, name, "B6"), getattr(oo, name))
    finally:
        _close(xwb, owb)


def test_header_footer(manifest):
    xwb, owb = _pair("metadata.xlsx")
    try:
        xh = xwb["Meta"].HeaderFooter
        oh = owb["Meta"].HeaderFooter
        for name in (
            "differentFirst", "differentOddEven", "scaleWithDoc",
            "alignWithMargins",
        ):
            xfail_mismatch("B7", f"HeaderFooter.{name}", require(xh, name, "B7"), getattr(oh, name))
        for side in ("oddHeader", "oddFooter", "evenHeader", "evenFooter", "firstHeader", "firstFooter"):
            xitem = require(xh, side, "B7")
            oitem = getattr(oh, side)
            for part in ("left", "center", "right"):
                xfail_mismatch(
                    "B7", f"{side}.{part}.text",
                    require(getattr(xitem, part), "text", "B7"),
                    getattr(getattr(oitem, part), "text", None),
                )
    finally:
        _close(xwb, owb)


def test_protection(manifest):
    xwb, owb = _pair("metadata.xlsx")
    try:
        xp = require(xwb["Meta"], "protection", "B10")
        op = owb["Meta"].protection
        for name in (
            "sheet", "objects", "scenarios", "formatCells", "formatRows",
            "formatColumns", "insertColumns", "insertRows", "insertHyperlinks",
            "deleteColumns", "deleteRows", "selectLockedCells", "sort",
            "autoFilter", "pivotTables", "password",
        ):
            xfail_mismatch("B10", f"protection.{name}", require(xp, name, "B10"), getattr(op, name))

        xw = require(xwb, "security", "B11")
        ow = owb.security
        for name in (
            "lockStructure", "lockWindows", "lockRevision",
            "workbookPassword", "revisionsPassword",
        ):
            xfail_mismatch("B11", f"security.{name}", require(xw, name, "B11"), getattr(ow, name))
    finally:
        _close(xwb, owb)


def test_calculation_and_views(manifest):
    xwb, owb = _pair("metadata.xlsx")
    try:
        xc = require(xwb, "calculation", "B11")
        oc = owb.calculation
        for name in (
            "calcId", "calcMode", "fullCalcOnLoad", "refMode", "iterate",
            "iterateCount", "iterateDelta", "fullPrecision", "calcCompleted",
            "calcOnSave", "concurrentCalc", "forceFullCalc",
        ):
            xfail_mismatch("B11", f"calculation.{name}", require(xc, name, "B11"), getattr(oc, name))
        xviews = require(xwb, "views", "B11")
        xv = xviews[0]
        ov = owb.views[0]
        for name in (
            "visibility", "minimized", "showHorizontalScroll", "showVerticalScroll",
            "showSheetTabs", "xWindow", "yWindow", "windowWidth", "windowHeight",
            "tabRatio", "firstSheet", "activeTab", "autoFilterDateGrouping",
        ):
            xfail_mismatch("B11", f"views[0].{name}", require(xv, name, "B11"), getattr(ov, name))
    finally:
        _close(xwb, owb)


def test_properties_and_custom_properties(manifest):
    xwb, owb = _pair("metadata.xlsx")
    try:
        xfail_mismatch("A12", "properties.lastPrinted",
                       require(xwb.properties, "lastPrinted", "A12"), owb.properties.lastPrinted)
        for name in (
            "creator", "title", "subject", "description", "keywords", "category",
            "lastModifiedBy", "created", "modified", "contentStatus", "version",
            "revision", "identifier", "language",
        ):
            xfail_mismatch("A12", f"properties.{name}",
                           require(xwb.properties, name, "A12"), getattr(owb.properties, name))
        xcustom = require(xwb, "custom_doc_props", "A11")
        xnames = sorted(item.name for item in xcustom)
        onames = sorted(item.name for item in owb.custom_doc_props)
        xfail_mismatch("A11", "custom property names", xnames, onames)
    finally:
        _close(xwb, owb)


def test_defined_names(manifest):
    xwb, owb = _pair("metadata.xlsx")
    try:
        xnames = require(xwb, "defined_names", "A4")
        require(xnames, "__getitem__", "A4")
        for name in ("GlobalName", "LocalName", "HiddenName"):
            try:
                entry = xnames[name]
            except Exception as exc:
                gap_xfail(f"A4: defined_names[{name!r}] failed: {exc}")
            try:
                oentry = owb.defined_names[name]
            except KeyError:
                oentry = owb["Meta"].defined_names[name]
            xfail_mismatch("A4", f"{name}.attr_text",
                           require(entry, "attr_text", "A4"), oentry.attr_text)
            xfail_mismatch("A4", f"{name}.localSheetId",
                           require(entry, "localSheetId", "A4"), oentry.localSheetId)
            xfail_mismatch("A4", f"{name}.hidden",
                           require(entry, "hidden", "A4"), oentry.hidden)
    finally:
        _close(xwb, owb)


def test_dimensions_and_auto_filter(manifest):
    xwb, owb = _pair("metadata.xlsx")
    try:
        xa = xwb["Meta"].column_dimensions["A"]
        oa = owb["Meta"].column_dimensions["A"]
        xfail_mismatch("B13", "column A width", require(xa, "width", "B13"), oa.width)
        xb = xwb["Meta"].column_dimensions["B"]
        ob = owb["Meta"].column_dimensions["B"]
        xfail_mismatch("B13", "column B hidden", require(xb, "hidden", "B13"), ob.hidden)
        xr = xwb["Meta"].row_dimensions[1]
        orow = owb["Meta"].row_dimensions[1]
        xfail_mismatch("B13", "row 1 height", require(xr, "height", "B13"), orow.height)
        xr2 = xwb["Meta"].row_dimensions[2]
        xfail_mismatch("B13", "row 2 outlineLevel", require(xr2, "outlineLevel", "B13"), owb["Meta"].row_dimensions[2].outlineLevel)

        xf = xwb["Meta"].auto_filter
        of = owb["Meta"].auto_filter
        xfail_mismatch("B4", "auto_filter.ref", require(xf, "ref", "B4"), of.ref)
        def filter_values(item):
            filters = getattr(item, "filters", None)
            if filters is not None:
                return [getattr(entry, "val", entry) for entry in filters.filter]
            return list(getattr(item, "vals", []) or [])

        xfail_mismatch("B4", "filter columns",
                       [(item.colId, filter_values(item))
                        for item in require(xf, "filterColumn", "B4")],
                       [(item.colId, filter_values(item))
                        for item in of.filterColumn])
        xfail_mismatch("B4", "sort conditions",
                       [item.ref for item in require(xf.sortState, "sortCondition", "B4")],
                       [item.ref for item in of.sortState.sortCondition])
    finally:
        _close(xwb, owb)


def test_tables(manifest):
    xwb, owb = _pair("objects.xlsx")
    try:
        xt = xwb["Objects"].tables
        ot = owb["Objects"].tables
        _check = require(xt, "__getitem__", "C1")
        try:
            xtable = xt["ObjectsTable"]
        except (TypeError, KeyError, IndexError):
            gap_xfail("C1: tables are not a name-keyed collection")
        xfail_mismatch(
            "C1", "table names",
            sorted(getattr(item, "displayName", str(item)) for item in xt),
            sorted(ot),
        )
        otable = ot["ObjectsTable"]
        for name in ("displayName", "ref", "totalsRowCount", "totalsRowShown", "headerRowCount"):
            xfail_mismatch("C1", f"table.{name}", require(xtable, name, "C1"), getattr(otable, name))
        xfail_mismatch("C1", "table columns",
                       [column.name for column in require(xtable, "tableColumns", "C1")],
                       [column.name for column in otable.tableColumns])
        xfail_mismatch("C1", "table style name",
                       require(xtable.tableStyleInfo, "name", "C1"), otable.tableStyleInfo.name)
        xfail_mismatch("C1", "table row stripes",
                       require(xtable.tableStyleInfo, "showRowStripes", "C1"), otable.tableStyleInfo.showRowStripes)
    finally:
        _close(xwb, owb)


def test_data_validations(manifest):
    xwb, owb = _pair("objects.xlsx")
    try:
        xv = list(xwb["Objects"].data_validations)
        ov = list(owb["Objects"].data_validations.dataValidation)
        xfail_mismatch("C2", "validation count", len(xv), len(ov))
        for index, overt in enumerate(ov):
            xret = xv[index]
            for name in ("type", "operator", "formula1", "formula2", "allowBlank",
                         "showDropDown", "showErrorMessage", "showInputMessage",
                         "errorStyle", "errorTitle", "promptTitle", "imeMode"):
                expected = getattr(overt, name)
                if name == "errorStyle" and expected is None:
                    expected = "stop"
                xfail_mismatch(
                    "C2", f"validation[{index}].{name}",
                    require(xret, name, "C2"), expected,
                )
    finally:
        _close(xwb, owb)


def test_conditional_formatting(manifest):
    xwb, owb = _pair("objects.xlsx")
    try:
        xlist = list(xwb["Objects"].conditional_formatting)
        olist = list(owb["Objects"].conditional_formatting)
        xfail_mismatch("C3", "cf range count", len(xlist), len(olist))
        xrules = [rule for item in xlist for rule in getattr(item, "rules", [])]
        orules = [rule for item in olist for rule in item.rules]
        xfail_mismatch("C3", "cf rule count", len(xrules), len(orules))
        xfail_mismatch(
            "C3", "cf rule types",
            sorted(getattr(rule, "rule_type", getattr(rule, "type", None)) for rule in xrules),
            sorted(rule.type for rule in orules),
        )
    finally:
        _close(xwb, owb)


def test_named_styles(manifest):
    xwb, owb = _pair("objects.xlsx")
    try:
        xstyles = require(xwb, "named_styles", "C4")
        if not hasattr(xstyles, "names"):
            gap_xfail("C4: named_styles has no names collection")
        xnames = sorted(xstyles.names)
        onames = sorted(
            item if isinstance(item, str) else item.name
            for item in owb.named_styles
        )
        xfail_mismatch("C4", "named style names", xnames, onames)
    finally:
        _close(xwb, owb)


def test_charts(manifest):
    xwb, owb = _pair("charts.xlsx")
    try:
        for sheet_name in ("Bar", "Bar3D", "Line", "Line3D", "Pie", "Pie3D",
                           "Doughnut", "Area", "Area3D", "Scatter", "Radar",
                           "Bubble", "Stock", "Surface", "Surface3D"):
            xcharts = list(xwb[sheet_name].charts)
            ocharts = list(owb[sheet_name]._charts)
            xfail_mismatch("C6", f"{sheet_name} chart count", len(xcharts), len(ocharts))
    finally:
        _close(xwb, owb)


def test_images(manifest):
    xwb, owb = _pair("images.xlsx")
    try:
        xfail_mismatch("C7", "image count", len(list(xwb["Images"].images)), len(owb["Images"]._images))
    finally:
        _close(xwb, owb)


def test_chartsheets(manifest):
    xwb, owb = _pair("chartsheet.xlsx")
    try:
        xsheets = require(xwb, "chartsheets", "C8/A8")
        xfail_mismatch("C8/A8", "chartsheet names",
                       [sheet.title for sheet in xsheets],
                       [sheet.title for sheet in owb.chartsheets])
    finally:
        _close(xwb, owb)


def test_long_tail_objects(manifest):
    xwb, owb = _pair("long_tail.xlsx")
    try:
        xs = xwb["LongTail"]
        os = owb["LongTail"]
        xfail_mismatch("A2", "comment", getattr(xs["A1"].comment, "text", None), os["A1"].comment.text)
        xfail_mismatch("A2", "hyperlink", getattr(xs["A2"].hyperlink, "target", None), os["A2"].hyperlink.target)
        xfail_mismatch("A2", "internal hyperlink", getattr(xs["A3"].hyperlink, "target", None), os["A3"].hyperlink.target)
        xfail_mismatch("D6", "rich text", normalize(xs["A4"].value), normalize(os["A4"].value))
        xfail_mismatch("B14", "merged cells", str(xs.merged_cells), str(os.merged_cells))
    finally:
        _close(xwb, owb)
