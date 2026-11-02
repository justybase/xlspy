"""Generate openpyxl-authored golden workbooks for the compatibility harness.

Usage:
    python tools/make_openpyxl_goldens.py [--out tests/data/openpyxl_parity]

The files are deliberately small but exercise the model surface that xlspy is
expected to match: workbook/sheet metadata, tables, validations, conditional
formatting, named styles, charts, images and chartsheets.  ``manifest.json``
records the openpyxl/Python versions used, so a compatibility run is
reproducible and a diff can be explained.
"""

from __future__ import annotations

import argparse
import base64
import datetime as dt
import io
import json
import platform
import traceback
from pathlib import Path

import openpyxl
from openpyxl import Workbook
from openpyxl.chart import (
    AreaChart,
    AreaChart3D,
    BarChart,
    BarChart3D,
    BubbleChart,
    DoughnutChart,
    LineChart,
    LineChart3D,
    PieChart,
    PieChart3D,
    RadarChart,
    Reference,
    ScatterChart,
    StockChart,
    SurfaceChart,
    SurfaceChart3D,
)
from openpyxl.chart.data_source import AxDataSource, NumDataSource, NumRef, StrRef
from openpyxl.chart.series import Series, SeriesLabel
from openpyxl.comments import Comment
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import (
    AbsoluteAnchor,
    AnchorMarker,
    OneCellAnchor,
    TwoCellAnchor,
)
from openpyxl.drawing.xdr import XDRPoint2D, XDRPositiveSize2D
from openpyxl.formatting.rule import (
    CellIsRule,
    ColorScaleRule,
    DataBarRule,
    FormulaRule,
    IconSetRule,
    Rule,
)
from openpyxl.styles import Border, Font, PatternFill, Side
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.styles.named_styles import NamedStyle
from openpyxl.utils.units import pixels_to_EMU
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.pagebreak import Break
from openpyxl.worksheet.table import Table, TableColumn, TableStyleInfo
from openpyxl.worksheet.views import Pane, Selection
from openpyxl.packaging.custom import (
    BoolProperty,
    DateTimeProperty,
    FloatProperty,
    IntProperty,
    StringProperty,
)

# A 1x1 transparent PNG; keeps the golden independent of external assets.
_PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9ZQmcAAAAASUVORK5CYII="
)


def _metadata_workbook() -> Workbook:
    wb = Workbook()
    ws = wb.active
    ws.title = "Meta"

    props = wb.properties
    props.creator = "openpyxl golden"
    props.title = "Metadata golden"
    props.subject = "xlspy parity"
    props.description = "Workbook, sheet and print metadata"
    props.keywords = "xlspy, openpyxl, parity"
    props.category = "test"
    props.lastModifiedBy = "golden generator"
    props.created = dt.datetime(2024, 1, 2, 3, 4, 5)
    props.modified = dt.datetime(2024, 1, 2, 3, 4, 6)
    props.lastPrinted = dt.datetime(2024, 1, 2, 3, 4, 7)
    props.contentStatus = "final"
    props.version = "1.0"
    props.revision = "2"
    props.identifier = "golden-metadata"
    props.language = "en-US"

    wb.defined_names.add(DefinedName("GlobalName", attr_text="'Meta'!$A$1"))
    wb.defined_names.add(DefinedName("LocalName", localSheetId=0, attr_text="'Meta'!$B$2"))
    wb.defined_names.add(DefinedName("HiddenName", hidden=True, attr_text="'Meta'!$C$3"))

    ws["A1"] = "Global"
    ws["B2"] = "Local"
    ws["C3"] = "Hidden"

    ws.sheet_properties.tabColor = "FF0000"
    ws.sheet_properties.codeName = "MetaCode"
    ws.sheet_properties.outlinePr.summaryBelow = False
    ws.sheet_properties.outlinePr.summaryRight = False
    ws.sheet_properties.outlinePr.applyStyles = True
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.sheet_properties.syncHorizontal = True
    ws.sheet_properties.syncVertical = True
    ws.sheet_properties.syncRef = "A1"
    ws.sheet_properties.transitionEvaluation = True
    ws.sheet_properties.transitionEntry = True
    ws.sheet_properties.published = True
    ws.sheet_properties.filterMode = True
    ws.sheet_properties.enableFormatConditionsCalculation = True

    ws.sheet_format.baseColWidth = 8
    ws.sheet_format.defaultColWidth = 12.5
    ws.sheet_format.defaultRowHeight = 18.5
    ws.sheet_format.customHeight = True
    ws.sheet_format.zeroHeight = False
    ws.sheet_format.thickTop = True
    ws.sheet_format.thickBottom = True
    ws.sheet_format.outlineLevelRow = 1
    ws.sheet_format.outlineLevelCol = 1

    ws.column_dimensions["A"].width = 15
    ws.column_dimensions["B"].hidden = True
    ws.column_dimensions["C"].bestFit = True
    ws.column_dimensions["D"].outlineLevel = 1
    ws.column_dimensions["D"].collapsed = True
    ws.row_dimensions[1].height = 20
    ws.row_dimensions[2].outlineLevel = 1
    ws.row_dimensions[2].collapsed = True
    ws.row_dimensions[3].hidden = True

    ws.sheet_view.showGridLines = False
    ws.sheet_view.showZeros = False
    ws.sheet_view.showRowColHeaders = False
    ws.sheet_view.zoomScale = 90
    ws.sheet_view.zoomScaleNormal = 100
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.tabSelected = True
    ws.sheet_view.topLeftCell = "B2"
    ws.sheet_view.pane = Pane(
        xSplit=1, ySplit=1, topLeftCell="B2", activePane="bottomRight", state="frozen"
    )
    ws.sheet_view.selection = [Selection(pane="bottomRight", activeCell="B2", sqref="B2")]

    ws.freeze_panes = "B2"
    ws.merge_cells("A1:B1")
    ws.auto_filter.ref = "A3:C10"
    ws.auto_filter.add_filter_column(0, ["A", "B"])
    ws.auto_filter.add_sort_condition("A4:A10", descending=True)

    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.scale = 90
    ws.page_setup.firstPageNumber = 1
    ws.page_setup.useFirstPageNumber = True
    ws.page_setup.draft = True
    ws.page_setup.horizontalDpi = 300
    ws.page_setup.verticalDpi = 300
    ws.page_setup.copies = 2
    ws.page_setup.pageOrder = "overThenDown"
    ws.page_setup.blackAndWhite = True
    ws.page_setup.cellComments = "atEnd"
    ws.page_setup.errors = "dash"
    ws.page_margins.left = 0.5
    ws.page_margins.right = 0.6
    ws.page_margins.top = 0.7
    ws.page_margins.bottom = 0.8
    ws.page_margins.header = 0.25
    ws.page_margins.footer = 0.35
    ws.print_options.horizontalCentered = True
    ws.print_options.verticalCentered = True
    ws.print_options.headings = True
    ws.print_options.gridLines = True
    ws.print_options.gridLinesSet = True
    ws.print_area = "A1:C10"
    ws.print_title_rows = "1:2"
    ws.print_title_cols = "A:B"

    ws.oddHeader.left.text = "Left"
    ws.oddHeader.center.text = "Center"
    ws.oddHeader.right.text = "Right"
    ws.oddFooter.center.text = "Page &P of &N"
    ws.evenHeader.left.text = "Even"
    ws.evenFooter.right.text = "Even footer"
    ws.firstHeader.center.text = "First"
    ws.firstFooter.left.text = "First footer"
    ws.HeaderFooter.differentFirst = True
    ws.HeaderFooter.differentOddEven = True
    ws.HeaderFooter.scaleWithDoc = False
    ws.HeaderFooter.alignWithMargins = False

    ws.row_breaks.append(Break(id=5))
    ws.col_breaks.append(Break(id=3))

    ws.protection.sheet = True
    ws.protection.password = "secret"
    ws.protection.objects = True
    ws.protection.scenarios = True
    ws.protection.formatCells = False
    ws.protection.formatColumns = False
    ws.protection.formatRows = False
    ws.protection.insertColumns = False
    ws.protection.insertRows = False
    ws.protection.insertHyperlinks = False
    ws.protection.deleteColumns = False
    ws.protection.deleteRows = False
    ws.protection.selectLockedCells = True
    ws.protection.selectUnlockedCells = True
    ws.protection.sort = False
    ws.protection.autoFilter = False
    ws.protection.pivotTables = False

    ws.protection.password = "secret"

    wb.security.lockStructure = True
    wb.security.lockWindows = True
    wb.security.lockRevision = True
    wb.security.workbookPassword = "wbpass"
    wb.security.revisionsPassword = "revpass"

    wb.calculation.calcMode = "auto"
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.refMode = "R1C1"
    wb.calculation.iterate = True
    wb.calculation.iterateCount = 100
    wb.calculation.iterateDelta = 0.001
    wb.calculation.fullPrecision = True
    wb.calculation.calcCompleted = True
    wb.calculation.calcOnSave = True
    wb.calculation.concurrentCalc = True
    wb.calculation.forceFullCalc = True

    view = wb.views[0]
    view.showHorizontalScroll = False
    view.showVerticalScroll = False
    view.showSheetTabs = False
    view.tabRatio = 500
    view.firstSheet = 0
    view.activeTab = 0
    view.autoFilterDateGrouping = False

    wb.custom_doc_props.append(StringProperty("TextProp", "value"))
    wb.custom_doc_props.append(IntProperty("IntProp", 42))
    wb.custom_doc_props.append(FloatProperty("FloatProp", 1.5))
    wb.custom_doc_props.append(BoolProperty("BoolProp", True))
    wb.custom_doc_props.append(DateTimeProperty("DateProp", dt.datetime(2024, 5, 6, 7, 8, 9)))
    return wb


def _objects_workbook() -> Workbook:
    wb = Workbook()
    ws = wb.active
    ws.title = "Objects"

    rows = [
        ["Product", "Qty", "Price", "Status"],
        ["Alpha", 10, 1.5, "New"],
        ["Beta", 20, 2.5, "Active"],
        ["Gamma", 30, 3.5, "Closed"],
        ["Delta", 40, 4.5, "New"],
    ]
    for row in rows:
        ws.append(row)
    for row in ws.iter_rows(min_row=2, min_col=2, max_col=3):
        for cell in row:
            cell.number_format = "#,##0.00"

    table = Table(displayName="ObjectsTable", ref="A1:D5")
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    table.tableColumns = [
        TableColumn(id=1, name="Product"),
        TableColumn(id=2, name="Qty"),
        TableColumn(id=3, name="Price"),
        TableColumn(id=4, name="Status"),
    ]
    ws.add_table(table)

    list_dv = DataValidation(
        type="list",
        formula1='"New,Active,Closed"',
        allow_blank=True,
        showDropDown=False,
        showErrorMessage=True,
        errorTitle="Bad status",
        error="Pick one of the list values",
        promptTitle="Status",
        prompt="Choose a status",
        showInputMessage=True,
    )
    list_dv.add("D2:D20")
    ws.add_data_validation(list_dv)

    whole_dv = DataValidation(
        type="whole",
        operator="between",
        formula1="0",
        formula2="1000",
        allow_blank=False,
        showErrorMessage=True,
        error="Quantity out of range",
        errorTitle="Quantity",
    )
    whole_dv.add("B2:B20")
    ws.add_data_validation(whole_dv)

    custom_dv = DataValidation(
        type="custom",
        formula1="LEN(D2)<=10",
        allow_blank=True,
        error="Too long",
        errorTitle="Status",
    )
    custom_dv.add("D2:D20")
    ws.add_data_validation(custom_dv)

    ws.conditional_formatting.add(
        "B2:B5",
        CellIsRule(
            operator="greaterThan",
            formula=["15"],
            fill=PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"),
            font=Font(color="9C0006"),
        ),
    )
    ws.conditional_formatting.add(
        "C2:C5",
        ColorScaleRule(
            start_type="min",
            start_color="FF0000",
            mid_type="percentile",
            mid_value=50,
            mid_color="FFFF00",
            end_type="max",
            end_color="00FF00",
        ),
    )
    ws.conditional_formatting.add(
        "B2:B5",
        DataBarRule(start_type="min", end_type="max", color="638EC6", showValue=True),
    )
    ws.conditional_formatting.add(
        "C2:C5",
        IconSetRule("3TrafficLights1", "percent", [0, 33, 67], showValue=True),
    )
    ws.conditional_formatting.add(
        "A2:A5",
        FormulaRule(
            formula=["LEN($A2)>3"],
            fill=PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid"),
        ),
    )
    ws.conditional_formatting.add(
        "B2:B5",
        Rule(
            type="top10",
            rank=3,
            dxf=DifferentialStyle(font=Font(bold=True, color="006100")),
        ),
    )
    ws.conditional_formatting.add(
        "C2:C5",
        Rule(
            type="aboveAverage",
            aboveAverage=True,
            dxf=DifferentialStyle(font=Font(italic=True, color="9C0006")),
        ),
    )
    ws.conditional_formatting.add(
        "A2:A5",
        Rule(
            type="containsText",
            operator="containsText",
            text="a",
            formula=['NOT(ISERROR(SEARCH("a",A2)))'],
            dxf=DifferentialStyle(font=Font(bold=True, color="0000FF")),
        ),
    )

    style = NamedStyle(name="Fancy")
    style.font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
    style.fill = PatternFill("solid", fgColor="4472C4")
    style.border = Border(
        left=Side(style="thin", color="000000"),
        right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"),
        bottom=Side(style="thin", color="000000"),
    )
    style.alignment = openpyxl.styles.Alignment(horizontal="center", vertical="center")
    style.number_format = "@"
    wb.add_named_style(style)
    ws["A1"].style = "Fancy"
    return wb


def _write_chart(wb: Workbook, chart, title: str) -> None:
    ws = wb.create_sheet(title)
    ws.add_chart(chart, "E2")


def _charts_workbook() -> Workbook:
    wb = Workbook()
    data = wb.active
    data.title = "ChartData"
    data.append(["Category", "Series 1", "Series 2", "Series 3", "Series 4"])
    for index, category in enumerate(["A", "B", "C", "D"]):
        data.append([category, 10 + index, 20 + index, 30 + index, 40 + index])

    categories = Reference(data, min_col=1, min_row=2, max_row=5)
    values = Reference(data, min_col=2, min_row=1, max_col=3, max_row=5)
    one_series = Reference(data, min_col=2, min_row=1, max_row=5)

    bar = BarChart()
    bar.type = "col"
    bar.style = 10
    bar.title = "Bar"
    bar.x_axis.title = "Category"
    bar.y_axis.title = "Value"
    bar.add_data(values, titles_from_data=True)
    bar.set_categories(categories)
    _write_chart(wb, bar, "Bar")

    bar3d = BarChart3D()
    bar3d.title = "Bar 3D"
    bar3d.add_data(values, titles_from_data=True)
    bar3d.set_categories(categories)
    _write_chart(wb, bar3d, "Bar3D")

    line = LineChart()
    line.title = "Line"
    line.style = 12
    line.add_data(values, titles_from_data=True)
    line.set_categories(categories)
    _write_chart(wb, line, "Line")

    line3d = LineChart3D()
    line3d.title = "Line 3D"
    line3d.add_data(values, titles_from_data=True)
    line3d.set_categories(categories)
    _write_chart(wb, line3d, "Line3D")

    pie = PieChart()
    pie.title = "Pie"
    pie.add_data(one_series, titles_from_data=True)
    pie.set_categories(categories)
    _write_chart(wb, pie, "Pie")

    pie3d = PieChart3D()
    pie3d.title = "Pie 3D"
    pie3d.add_data(one_series, titles_from_data=True)
    pie3d.set_categories(categories)
    _write_chart(wb, pie3d, "Pie3D")

    doughnut = DoughnutChart()
    doughnut.title = "Doughnut"
    doughnut.add_data(one_series, titles_from_data=True)
    doughnut.set_categories(categories)
    _write_chart(wb, doughnut, "Doughnut")

    area = AreaChart()
    area.title = "Area"
    area.add_data(values, titles_from_data=True)
    area.set_categories(categories)
    _write_chart(wb, area, "Area")

    area3d = AreaChart3D()
    area3d.title = "Area 3D"
    area3d.add_data(values, titles_from_data=True)
    area3d.set_categories(categories)
    _write_chart(wb, area3d, "Area3D")

    scatter = ScatterChart()
    scatter.title = "Scatter"
    scatter.x_axis.title = "X"
    scatter.y_axis.title = "Y"
    scatter.series.append(
        Series(
            xVal=AxDataSource(numRef=NumRef(f="ChartData!$B$2:$B$5")),
            yVal=NumDataSource(numRef=NumRef(f="ChartData!$C$2:$C$5")),
            tx=SeriesLabel(strRef=StrRef(f="ChartData!$C$1")),
        )
    )
    _write_chart(wb, scatter, "Scatter")

    radar = RadarChart()
    radar.title = "Radar"
    radar.add_data(values, titles_from_data=True)
    radar.set_categories(categories)
    _write_chart(wb, radar, "Radar")

    bubble = BubbleChart()
    bubble.title = "Bubble"
    bubble.series.append(
        Series(
            xVal=AxDataSource(numRef=NumRef(f="ChartData!$B$2:$B$5")),
            yVal=NumDataSource(numRef=NumRef(f="ChartData!$C$2:$C$5")),
            bubbleSize=NumDataSource(numRef=NumRef(f="ChartData!$D$2:$D$5")),
            tx=SeriesLabel(strRef=StrRef(f="ChartData!$C$1")),
        )
    )
    _write_chart(wb, bubble, "Bubble")

    stock = StockChart()
    stock.title = "Stock"
    stock.add_data(
        Reference(data, min_col=2, min_row=1, max_col=5, max_row=5),
        titles_from_data=True,
    )
    stock.set_categories(categories)
    _write_chart(wb, stock, "Stock")

    surface = SurfaceChart()
    surface.title = "Surface"
    surface.add_data(values, titles_from_data=True)
    surface.set_categories(categories)
    _write_chart(wb, surface, "Surface")

    surface3d = SurfaceChart3D()
    surface3d.title = "Surface 3D"
    surface3d.add_data(values, titles_from_data=True)
    surface3d.set_categories(categories)
    _write_chart(wb, surface3d, "Surface3D")

    return wb


def _images_workbook() -> Workbook:
    wb = Workbook()
    ws = wb.active
    ws.title = "Images"

    plain = XLImage(io.BytesIO(_PNG))
    plain.width = 32
    plain.height = 32
    ws.add_image(plain, "B2")

    one_cell = XLImage(io.BytesIO(_PNG))
    one_cell.anchor = OneCellAnchor(
        _from=AnchorMarker(col=3, row=3),
        ext=XDRPositiveSize2D(pixels_to_EMU(48), pixels_to_EMU(48)),
    )
    ws.add_image(one_cell)

    two_cell = XLImage(io.BytesIO(_PNG))
    two_cell.anchor = TwoCellAnchor(
        editAs="twoCell",
        _from=AnchorMarker(col=1, row=8),
        to=AnchorMarker(col=4, row=11),
    )
    ws.add_image(two_cell)

    absolute = XLImage(io.BytesIO(_PNG))
    absolute.anchor = AbsoluteAnchor(
        pos=XDRPoint2D(pixels_to_EMU(10), pixels_to_EMU(10)),
        ext=XDRPositiveSize2D(pixels_to_EMU(32), pixels_to_EMU(32)),
    )
    ws.add_image(absolute)
    return wb


def _chartsheet_workbook() -> Workbook:
    wb = Workbook()
    data = wb.active
    data.title = "Data"
    data.append(["Category", "Value"])
    for index, category in enumerate(["A", "B", "C", "D"]):
        data.append([category, index + 1])

    chart = BarChart()
    chart.title = "Chartsheet chart"
    chart.add_data(Reference(data, min_col=2, min_row=1, max_row=5), titles_from_data=True)
    chart.set_categories(Reference(data, min_col=1, min_row=2, max_row=5))

    sheet = wb.create_chartsheet("ChartSheet")
    sheet.add_chart(chart)
    return wb


def _long_tail_workbook() -> Workbook:
    wb = Workbook()
    ws = wb.active
    ws.title = "LongTail"
    ws["A1"] = "Comment"
    ws["A2"] = "https://example.com"
    ws["A3"] = "Internal"
    ws["A4"] = "Rich"

    ws["A1"].comment = Comment("A classic note", "Golden", width=200, height=100)
    ws["A2"].hyperlink = "https://example.com"
    ws["A2"].style = "Hyperlink"
    ws["A3"].hyperlink = "#LongTail!A1"
    ws["A3"].style = "Hyperlink"

    from openpyxl.cell.rich_text import CellRichText, TextBlock
    from openpyxl.cell.text import InlineFont

    ws["A4"].value = CellRichText(
        [
            TextBlock(InlineFont(b=True, color="FF0000"), "Bold red"),
            " plain ",
            TextBlock(InlineFont(i=True), "italic"),
        ]
    )

    ws.merge_cells("B2:C2")
    ws["B2"] = "Merged"
    ws.column_dimensions["E"].width = 20
    ws.row_dimensions[5].height = 25
    ws.page_setup.orientation = "portrait"
    ws.print_area = "A1:C10"
    return wb


def _save(wb: Workbook, path: Path, errors: list[dict]) -> None:
    try:
        wb.save(path)
    except Exception as exc:  # pragma: no cover - surfaced in the manifest
        errors.append(
            {
                "file": path.name,
                "error": f"{type(exc).__name__}: {exc}",
                "traceback": traceback.format_exc(),
            }
        )
        raise
    finally:
        wb.close()


def generate(out_dir: Path) -> dict:
    out_dir.mkdir(parents=True, exist_ok=True)
    errors: list[dict] = []
    builders = {
        "metadata.xlsx": _metadata_workbook,
        "objects.xlsx": _objects_workbook,
        "charts.xlsx": _charts_workbook,
        "images.xlsx": _images_workbook,
        "chartsheet.xlsx": _chartsheet_workbook,
        "long_tail.xlsx": _long_tail_workbook,
    }
    for name, builder in builders.items():
        wb = builder()
        _save(wb, out_dir / name, errors)

    manifest = {
        "openpyxl_version": openpyxl.__version__,
        "python_version": platform.python_version(),
        "generated_utc": dt.datetime.now(dt.UTC).isoformat(),
        "files": sorted(builders),
        "errors": errors,
    }
    (out_dir / "manifest.json").write_text(
        json.dumps(manifest, indent=2, ensure_ascii=False) + "\n", encoding="utf-8"
    )
    return manifest


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--out",
        type=Path,
        default=Path(__file__).resolve().parents[1] / "tests" / "data" / "openpyxl_parity",
        help="output directory (default: tests/data/openpyxl_parity)",
    )
    args = parser.parse_args(argv)
    manifest = generate(args.out)
    print(
        f"generated {len(manifest['files'])} goldens with openpyxl "
        f"{manifest['openpyxl_version']} in {args.out}"
    )
    if manifest["errors"]:
        for error in manifest["errors"]:
            print(f"FAILED {error['file']}: {error['error']}")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
