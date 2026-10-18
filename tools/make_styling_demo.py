#!/usr/bin/env python3
"""Generate the xlspy capability-demo workbooks.

The workbooks exercise the styling and formatting features of xlspy 0.7.0 so
they can be opened in Excel and verified by eye (or checked with
``tools/validate_excel_com.ps1``).

Usage:
    python3 tools/make_styling_demo.py [--out DIR]

Default output directory: ``tmp/xlspy-test-files`` (relative to the repo root).

Files:
    styling_demo.xlsx — ten sheets: flat styles, fonts, borders, alignment,
        fills (patterns and gradients), conditional formatting, tables,
        charts, layout options, number formats.
    styling_demo.xlsb — flat styles and tuple number formats (XLSB backend).
"""

from __future__ import annotations

import argparse
import datetime
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

from xlspy import (  # noqa: E402
    Alignment,
    Border,
    CellIsRule,
    Chart,
    ChartSeries,
    Color,
    ColorScaleRule,
    DataBarRule,
    DifferentialStyle,
    ExcelTable,
    F,
    Font,
    FormatObject,
    FormulaRule,
    GradientFill,
    IconSetRule,
    PatternFill,
    Protection,
    Side,
    Stop,
    Style,
    XlsbWriter,
    XlsxWriter,
)

HEADER = Style(font=Font(bold=True, color="FFFFFF"), fill="305496")
NOTE = Style(font=Font(italic=True, size=9, color="808080"))
SECTION = Style(bold=True)

BORDER_STYLES = [
    "thin", "medium", "thick", "dashed", "dotted", "double", "hair",
    "mediumDashed", "dashDot", "mediumDashDot", "dashDotDot",
    "mediumDashDotDot", "slantDashDot",
]

PATTERN_TYPES = [
    "solid", "gray125",
    "darkGray", "mediumGray", "lightGray", "gray0625",
    "darkHorizontal", "darkVertical", "darkDown", "darkUp",
    "darkGrid", "darkTrellis",
    "lightHorizontal", "lightVertical", "lightDown", "lightUp",
    "lightGrid", "lightTrellis",
]


def _serial(hours: float, minutes: float, seconds: float = 0.0,
            milliseconds: float = 0.0) -> float:
    """Excel time-of-day serial for a fraction of a day."""
    total = hours * 3600 + minutes * 60 + seconds + milliseconds / 1000.0
    return total / 86400.0


# ---------------------------------------------------------------------------
# Sheet builders
# ---------------------------------------------------------------------------

def _flat_sheet(writer: XlsxWriter | XlsbWriter) -> None:
    """Flat (0.6.0) style API plus tuple number formats."""
    long_text = "Text centered and wrapped inside the cell"
    data = [
        ["Element", "Demo", "Expected result"],
        ["fill = '#008000'", "Green background", "solid green fill"],
        ["fill + font_color + bold", "White text on navy", "bold white text"],
        ["bold + italic + underline", "Bold, italic, underlined", "three effects at once"],
        ["number_format", 1234.56, "1,234.57 (pl-PL locale: 1 234,57)"],
        ["horizontal + vertical + wrap", long_text, "text centered horizontally and vertically"],
        ["tuple (value, format)", (0.1234, F.PERCENTAGE), "12%"],
        ["tuple + Style together", (1234.5, F.TWO_DECIMALS), "format from the tuple wins; yellow fill from Style"],
    ]
    styles = {
        "B2": Style(fill="#008000"),
        "B3": Style(fill="#1F3864", font_color="#FFFFFF", bold=True),
        "B4": Style(bold=True, italic=True, underline=True),
        "B5": Style(number_format=F.TWO_DECIMALS),
        "B6": Style(horizontal="center", vertical="center", wrap=True),
        "B7": Style(fill="FFF2CC"),
        "B8": Style(fill="FFF2CC"),
    }
    writer.add_sheet("Styles")
    writer.write_sheet(
        data,
        header_style=HEADER,
        styles=styles,
        column_widths={"A": 38, "B": 34, "C": 54},
        row_heights={5: 30},
        auto_filter=False,
    )


def _font_sheet(writer: XlsxWriter | XlsbWriter) -> None:
    data = [
        ["Element", "Demo", "Expected result"],
        ["name='Segoe UI', size=14", "Segoe UI 14", "Excel shows Segoe UI, not Calibri"],
        ["charset=0, family=2", "charset/family", "styles.xml: <charset val=\"0\"/>, <family val=\"2\"/>"],
        ["scheme='major'", "Theme font (major)", "Excel uses the theme's major font (e.g. Calibri Light)"],
        ["name='Arial' without scheme", "Arial without scheme", "Excel shows Arial (no substitution for the theme font)"],
        ["outline=True", "Outline letters", "outline around the glyphs"],
        ["shadow=True", "Shadowed letters", "shadow below the glyphs"],
        ["condense=True", "Condensed letters", "tighter character spacing"],
        ["extend=True", "Expanded letters", "wider character spacing"],
        ["strike=True", "Strikethrough", "line through the text"],
        ["underline='single'", "Single underline", "one line"],
        ["underline='double'", "Double underline", "two lines"],
        ["underline='singleAccounting'", "Accounting 1", "line below the text baseline"],
        ["underline='doubleAccounting'", "Accounting 2", "double line below the baseline"],
        ["vertAlign='superscript'", "x2", "whole text raised (x²)"],
        ["vertAlign='subscript'", "H2O", "whole text lowered (H₂O)"],
        ["color=Color(auto=True)", "Automatic color", "styles.xml: auto=\"1\""],
        ["color=Color(theme=4, tint=0.4)", "Theme color + tint", "lighter shade of the theme accent color"],
        ["color=Color(indexed=3)", "Indexed color 3", "red from the indexed palette"],
        ["name='Arial' + rgb + bold + 18", "Red Arial 18", "bold red, 18 pt"],
    ]
    styles = {
        "B2": Style(font=Font(name="Segoe UI", size=14, charset=0, family=2)),
        "B3": Style(font=Font(name="Segoe UI", charset=0, family=2)),
        "B4": Style(font=Font(scheme="major", size=14)),
        "B5": Style(font=Font(name="Arial", size=12)),
        "B6": Style(font=Font(size=16, outline=True)),
        "B7": Style(font=Font(size=16, shadow=True)),
        "B8": Style(font=Font(size=14, condense=True)),
        "B9": Style(font=Font(size=14, extend=True)),
        "B10": Style(font=Font(strike=True)),
        "B11": Style(font=Font(underline="single")),
        "B12": Style(font=Font(underline="double")),
        "B13": Style(font=Font(underline="singleAccounting")),
        "B14": Style(font=Font(underline="doubleAccounting")),
        "B15": Style(font=Font(vertAlign="superscript")),
        "B16": Style(font=Font(vertAlign="subscript")),
        "B17": Style(font=Font(color=Color(auto=True), bold=True)),
        "B18": Style(font=Font(color=Color(theme=4, tint=0.4), bold=True, size=14)),
        "B19": Style(font=Font(color=Color(indexed=3), bold=True)),
        "B20": Style(font=Font(name="Arial", size=18, bold=True, color="FF0000")),
    }
    writer.add_sheet("Fonts")
    writer.write_sheet(
        data,
        header_style=HEADER,
        styles=styles,
        column_widths={"A": 38, "B": 30, "C": 56},
        row_heights={5: 28, 6: 28},
        auto_filter=False,
    )


def _border_sheet(writer: XlsxWriter) -> None:
    data = [["Element", "Demo", None, "Expected result"]]
    styles: dict[str, Style] = {}
    merges: list[str] = []
    tall_rows: list[int] = []
    row = 2  # 1-based row cursor

    for name in BORDER_STYLES:
        data.append([f"left={name}", None, None, f"styles.xml: <left style=\"{name}\"/>"])
        styles[f"B{row}"] = Style(border=Border(left=Side(name)))
        row += 1

    outer = {"left": Side("thin"), "right": Side("thin"),
             "top": Side("thin"), "bottom": Side("thin")}
    entries = [
        # (label, note, border, merge_rows)
        ("left=thin + bottom=double (red)",
         "thin left edge, red double bottom edge",
         Border(left=Side("thin"), bottom=Side("double", "C00000")), 1),
        ("start=thin, end=medium",
         "start/end = leading/trailing edge (LTR)",
         Border(start=Side("thin"), end=Side("medium")), 1),
        ("diagonal up (diagonal_direction='up')",
         "line from the bottom-left to the top-right corner",
         Border(diagonal=Side("hair", "C00000"), diagonal_direction="up"), 1),
        ("diagonal down (diagonal_down=True)",
         "line from the top-left to the bottom-right corner",
         Border(diagonal=Side("double", "0070C0"), diagonal_down=True), 1),
        ("diagonal in both directions",
         "two diagonal lines",
         Border(diagonal=Side("thin"), diagonal_up=True, diagonal_down=True), 1),
        ("vertical + horizontal inside a merged 2×2 block",
         "inner cross inside the merged block (vertical=red, horizontal=blue)",
         Border(**outer, vertical=Side("thin", "C00000"),
                horizontal=Side("thin", "0070C0")), 2),
        ("outline=False on a merged block",
         "no frame around the merged block",
         Border(**outer, outline=False), 2),
        ("default outline on a merged block",
         "frame around the whole merged block",
         Border(**outer), 2),
        ("no border (Border())",
         "cell without any frame",
         None, 1),
    ]
    data.append([None, None, None, None])
    row += 1
    for label, note, border, merge_rows in entries:
        data.append([label, None, None, note])
        if merge_rows == 2:
            data.append([None, None, None, None])
            for offset in range(merge_rows):
                for column in ("B", "C"):
                    if border is not None:
                        styles[f"{column}{row + offset}"] = Style(border=border)
                tall_rows.append(row + offset)
            merges.append(f"B{row}:C{row + merge_rows - 1}")
            row += merge_rows
        else:
            if border is not None:
                styles[f"B{row}"] = Style(border=border)
            row += 1

    writer.add_sheet("Borders")
    writer.write_sheet(
        data,
        header_style=HEADER,
        styles=styles,
        merges=merges,
        column_widths={"A": 42, "B": 16, "C": 16, "D": 52},
        row_heights={height - 1: 30 for height in tall_rows},
        auto_filter=False,
    )


def _alignment_sheet(writer: XlsxWriter) -> None:
    text = ("Allocation of the research and development budget by quarter, "
            "together with the management commentary")
    data = [
        ["Element", "Demo", "Expected result"],
        ["justify + justify_last_line + wrap", text, "every line stretched to the column width"],
        ["relative_indent=3 (justify + wrap)", text, "extra indent on the last line of the paragraph"],
        ["indent=3 (left)", "Indented text", "text pushed three characters from the left"],
        ["shrink_to_fit", "Very long text in a column that is 32 wide", "Excel shrinks the font so the text fits"],
        ["text_rotation=45", "45 degrees", "text rotated 45° upwards"],
        ["text_rotation=135", "135 degrees", "text rotated 135° (downwards)"],
        ["text_rotation=255", "Vertical", "vertical text, top to bottom"],
        ["reading_order=2 (RTL)", "Reading order 2 — RTL", "styles.xml: readingOrder=\"2\""],
        ["reading_order=1 (LTR)", "Reading order 1 — LTR", "styles.xml: readingOrder=\"1\""],
        ["wrap + vertical='top'", "Wrapped paragraph aligned to the top edge of the cell", "two lines aligned to the top"],
    ]
    styles = {
        "B2": Style(alignment=Alignment(horizontal="justify", wrap_text=True, justify_last_line=True)),
        "B3": Style(alignment=Alignment(horizontal="justify", wrap_text=True, relative_indent=3)),
        "B4": Style(alignment=Alignment(horizontal="left", indent=3)),
        "B5": Style(alignment=Alignment(shrink_to_fit=True)),
        "B6": Style(alignment=Alignment(text_rotation=45)),
        "B7": Style(alignment=Alignment(text_rotation=135)),
        "B8": Style(alignment=Alignment(text_rotation=255)),
        "B9": Style(alignment=Alignment(reading_order=2, horizontal="right")),
        "B10": Style(alignment=Alignment(reading_order=1, horizontal="left")),
        "B11": Style(alignment=Alignment(wrap_text=True, vertical="top")),
    }
    writer.add_sheet("Alignment")
    writer.write_sheet(
        data,
        header_style=HEADER,
        styles=styles,
        column_widths={"A": 40, "B": 32, "C": 54},
        row_heights={2: 30, 3: 30, 5: 45, 6: 60, 7: 80},
        auto_filter=False,
    )


def _fill_sheet(writer: XlsxWriter | XlsbWriter) -> None:
    data = [["Element", "Demo", None, None, "Expected result"]]
    styles: dict[str, Style] = {}
    row = 2  # 1-based row cursor

    for name in PATTERN_TYPES:
        if name == "solid":
            fill = PatternFill("solid", start_color="FFC000")
            note = "solid foreground color (fgColor)"
        else:
            fill = PatternFill(name, start_color="1F3864", end_color="FFF2CC")
            note = "pattern color (fgColor) + background (bgColor)"
        data.append([name, None, None, None, note])
        styles[f"B{row}"] = Style(fill=fill)
        row += 1

    data.append(["GradientFill", None, None, None, None])
    styles[f"A{row}"] = SECTION
    row += 1

    gradients = [
        ("linear, degree=90, two explicit Stop objects",
         GradientFill("linear", degree=90,
                      stop=[Stop("C00000", 0.0), Stop("0070C0", 1.0)]),
         "red → blue, top to bottom"),
        ("linear, degree=0, three bare colors",
         GradientFill("linear", degree=0, stop=["C00000", "FFFF00", "0070C0"]),
         "colors spread evenly at 0 / 0.5 / 1, left to right"),
        ("linear, degree=45, two colors",
         GradientFill("linear", degree=45, stop=["008000", "FFFFFF"]),
         "green → white at 45°"),
        ("path, extents 0.2",
         GradientFill("path", left=0.2, right=0.2, top=0.2, bottom=0.2,
                      stop=["FFF2CC", "BF8F00"]),
         "path (radial) gradient"),
    ]
    for label, fill, note in gradients:
        data.append([label, None, None, None, note])
        styles[f"B{row}"] = Style(fill=fill)
        row += 1

    block_label = f"gradient 90° repeated over B{row}:D{row + 2}"
    for extra in range(3):
        data.append([
            block_label if extra == 0 else None,
            None, None, None,
            "one GradientFill applied to a 3×3 range" if extra == 0 else None,
        ])
    styles[f"B{row}:D{row + 2}"] = Style(
        fill=GradientFill("linear", degree=90, stop=["FFF2CC", "BF8F00"])
    )

    writer.add_sheet("Fills")
    writer.write_sheet(
        data,
        header_style=HEADER,
        styles=styles,
        column_widths={"A": 40, "B": 14, "C": 14, "D": 14, "E": 52},
        auto_filter=False,
    )


def _cf_sheet(writer: XlsxWriter) -> None:
    headers = [
        "Data", "cellIs: full dxf", "formula: gradient dxf",
        "2-color scale", "3-color scale", "data bar",
        "5Arrows icons", "3TrafficLights1 icons",
        "shared dxf A", "shared dxf B",
    ]
    values_b = [3, 8, 1, 9, 4, 7, 2, 10, 5, 6]
    data = [headers]
    for index in range(10):
        data.append([
            index + 1,
            values_b[index],
            index + 1,
            (index + 1) * 2,
            (index + 1) * 5,
            index - 5,
            index + 1,
            index + 1,
            index + 1,
            index + 1,
        ])

    full_dxf = DifferentialStyle(
        font=Font(bold=True, color="9C0006"),
        number_format="0.00",
        fill=PatternFill("solid", start_color="FFC7CE"),
        alignment=Alignment(horizontal="center"),
        border=Border(left=Side("thin", "9C0006")),
        protection=Protection(locked=False),
    )
    warm_dxf = DifferentialStyle(
        font=Font(bold=True, color="9C6500"),
        fill=PatternFill("solid", start_color="FFEB9C"),
    )
    gradient_dxf = DifferentialStyle(
        font=Font(bold=True, color="9C0006"),
        fill=GradientFill("linear", degree=45, stop=["FFC7CE", "FFEB9C"]),
    )
    shared_dxf = DifferentialStyle(
        font=Font(italic=True, color="006100"),
        fill=PatternFill("solid", start_color="C6EFCE"),
    )

    rules = [
        CellIsRule("B2:B11", "lessThan", ["5"], dxf=full_dxf, priority=1),
        CellIsRule("B2:B11", "between", ["7", "8"], dxf=warm_dxf, priority=2),
        FormulaRule("C2:C11", "MOD($A2,2)=0", dxf=gradient_dxf, priority=3),
        ColorScaleRule(
            "D2:D11", FormatObject("min"), None, FormatObject("max"),
            ["F8696B", "63BE7B"],
        ),
        ColorScaleRule(
            "E2:E11", FormatObject("min"), FormatObject("percentile", 50),
            FormatObject("max"), ["F8696B", "FFEB84", "63BE7B"],
        ),
        DataBarRule("F2:F11", FormatObject("min"), FormatObject("max"), "638EC6"),
        IconSetRule("G2:G11", "5Arrows", [0, 20, 40, 60, 80], show_value=True),
        IconSetRule("H2:H11", "3TrafficLights1", [0, 35, 70],
                    percent=True, reverse=True),
        FormulaRule("I2:I11", "$A2>5", dxf=shared_dxf),
        FormulaRule("J2:J11", "$A2<=5", dxf=shared_dxf),
    ]

    styles: dict[str, Style] = {}
    notes = [
        "Priorities: rules 1–3 carry an explicit priority=1..3; the writer assigned 4.. "
        "to the remaining rules — the rules manager shows the order without duplicates.",
        "dxf dedup: columns I and J use the same DifferentialStyle object, so styles.xml "
        "contains one shared <dxf>.",
        "Column B: 'less than 5' (full dxf: red bold font, 0.00 format, pink fill, "
        "centered, left border, locked=False) and 'between 7 and 8' (yellow dxf).",
        "Column C: dxf with a gradient — Excel may render a dxf gradient fill as solid; "
        "the bold red font from the dxf must be visible.",
        "dxf protection (locked=False) shows up after enabling sheet protection: "
        "Review → Protect Sheet.",
        "Column H: reversed icons (reverse=True), three traffic lights.",
    ]
    note_row = len(data) + 1
    for offset, note in enumerate(notes):
        data.append([note])
        styles[f"A{note_row + offset}"] = NOTE

    writer.add_sheet("Conditional Formatting")
    writer.write_sheet(
        data,
        header_style=HEADER,
        styles=styles,
        conditional_formats=rules,
        column_widths={
            "A": 20, "B": 18, "C": 24, "D": 14, "E": 14,
            "F": 14, "G": 16, "H": 18, "I": 20, "J": 20,
        },
        freeze_panes="B2",
        auto_filter=False,
    )


def _table_sheet(writer: XlsxWriter) -> None:
    data = [
        ["Product", "Quantity", "Price", "Notes"],
        ["Apples", 1200, 12.5, "fresh"],
        ["Pears", 8, 7.25, "ripening"],
        ["Plums", 640, 9.9, "dried"],
        ["Cherries", 95, 23.0, "in season"],
        ["Apricots", 310, 15.75, "sweet"],
        [None, None, None, None],
        ["Table: ExcelTable display_name='VerifTable', ref='A1:D6', style TableStyleMedium9.", None, None, None],
        ["Check: Table Design tab, filter arrows in the header, banded rows, status bar sums.", None, None, None],
        ["openpyxl: ws.tables['VerifTable'].ref == 'A1:D6'.", None, None, None],
    ]
    writer.add_sheet("Table")
    writer.write_sheet(
        data,
        header_style=HEADER,
        tables=[ExcelTable(
            "VerifTable", "A1:D6",
            columns=["Product", "Quantity", "Price", "Notes"],
        )],
        styles={"C2:C6": Style(number_format=F.TWO_DECIMALS), "A8:A10": NOTE},
        column_widths={"A": 34, "B": 10, "C": 12, "D": 18},
        auto_filter=False,
    )


def _chart_sheet(writer: XlsxWriter | XlsbWriter) -> None:
    data = [
        ["Quarter", "Hardware", "Services"],
        ["Q1", 120, 80],
        ["Q2", 150, 95],
        ["Q3", 130, 110],
        ["Q4", 170, 125],
        [None, None, None],
        ["Charts: bar (E2, style 10), line (E18), pie (E34).", None, None],
        ["Check: three charts anchored at E2 / E18 / E34 with the titles below.", None, None],
    ]
    writer.add_sheet("Charts")
    writer.write_sheet(
        data,
        header_style=HEADER,
        charts=[
            Chart(
                "bar",
                [
                    ChartSeries("B2:B5", title_cell="B1", categories="A2:A5"),
                    ChartSeries("C2:C5", title_cell="C1", categories="A2:A5"),
                ],
                title="Revenue by quarter",
                anchor="E2",
                style=10,
            ),
            Chart(
                "line",
                [ChartSeries("B2:B5", title_cell="B1", categories="A2:A5")],
                title="Hardware trend",
                anchor="E18",
            ),
            Chart(
                "pie",
                [ChartSeries("C2:C5", title="Services")],
                title="Services share",
                anchor="E34",
            ),
        ],
        column_widths={"A": 40, "B": 12, "C": 12, "D": 3},
        auto_filter=False,
    )


def _layout_sheet(writer: XlsxWriter | XlsbWriter) -> None:
    data = [
        ["Column A", "Column B", "Column C", "Column D (hidden)", "Column E"],
        [None, None, None, None, None],
        [None, "Merged B3:C3 — one label spanning two columns", None, None, None],
        [1, 2, 3, 4, 5],
        ["Row 5 — height 30", None, None, None, None],
        [10, 20, 30, 40, 50],
        ["Row 7 — hidden", None, None, None, None],
        ["Row 8 — after the hidden row", None, None, None, None],
        [None, None, None, None, None],
        ["Column B is 28 wide.", None, None, None, None],
        ["Column D is hidden.", None, None, None, None],
        ["Row 5 is 30 pt high, row 7 is hidden.", None, None, None, None],
        ["Freeze B2: row 1 and column A stay visible while scrolling.", None, None, None, None],
        ["Auto-filter (auto_filter=True): filter arrows in row 1.", None, None, None, None],
    ]
    writer.add_sheet("Layout")
    writer.write_sheet(
        data,
        header_style=HEADER,
        styles={"B3:C3": Style(bold=True, fill="FFF2CC", horizontal="center")},
        merges=["B3:C3"],
        column_widths={"B": 28},
        hidden_columns=["D"],
        row_heights={4: 30},
        hidden_rows=[6],
        freeze_panes="B2",
        auto_filter=True,
    )


def _format_sheet(writer: XlsxWriter | XlsbWriter) -> None:
    data = [
        ["Format", "Value", "Expected result in the cell"],
        ["F.TWO_DECIMALS", (1234.5678, F.TWO_DECIMALS), "1,234.57 (pl-PL locale: 1 234,57)"],
        ["F.THOUSANDS_SEP", (12345, F.THOUSANDS_SEP), "12,345 (pl-PL locale: 12 345)"],
        ["F.PERCENTAGE", (0.1234, F.PERCENTAGE), "12%"],
        ["F.CURRENCY_PLN", (1234.5, F.CURRENCY_PLN), "1,234.50 zł"],
        ["F.CURRENCY_EUR", (1234.5, F.CURRENCY_EUR), "1,234.50 €"],
        ["F.SCIENTIFIC", (12345, F.SCIENTIFIC), "1.23E+04"],
        ["F.LEADING_ZEROS", (42, F.LEADING_ZEROS), "000000042"],
        ["F.TEXT", ("007", F.TEXT), "007 kept as text (no number conversion)"],
        ["F.DATE_ISO", (datetime.date(2026, 10, 18), F.DATE_ISO), "2026-10-18"],
        ["F.DATE_LONG", (datetime.date(2026, 10, 18), F.DATE_LONG), "18 October 2026 (pl-PL locale: 18 października 2026)"],
        ["F.DATETIME_ISO", (datetime.datetime(2026, 10, 18, 14, 30, 15), F.DATETIME_ISO), "2026-10-18T14:30:15"],
        ["F.TIME_HH_MM", (_serial(14, 30), F.TIME_HH_MM), "14:30"],
        ["F.TIME_MS", (_serial(14, 30, 15, 250), F.TIME_MS), "14:30:15.250"],
    ]
    writer.add_sheet("Formats")
    writer.write_sheet(
        data,
        header_style=HEADER,
        column_widths={"A": 22, "B": 26, "C": 56},
        auto_filter=False,
    )


def _write_xlsb(path: pathlib.Path) -> None:
    with XlsbWriter(path) as writer:
        _flat_sheet(writer)
        _font_sheet(writer)
        _fill_sheet(writer)
        _chart_sheet(writer)
        _layout_sheet(writer)
        _format_sheet(writer)
        writer.add_sheet("Flat")
        writer.write_sheet(
            [
                ["Product", "Quantity", "Price", "Notes"],
                ["Apples", 1200, (12.5, F.TWO_DECIMALS), "fresh"],
                ["Pears", 8, (7.25, F.TWO_DECIMALS), "ripening"],
                ["Plums", 640, (9.9, F.TWO_DECIMALS), "dried"],
            ],
            header_style=Style(fill="008000", font_color="FFFFFF", bold=True),
            styles={
                "B2:B4": Style(fill="FFF2CC"),
                "C3": Style(fill="FFC7CE", font_color="9C0006", underline=True),
                "D4": Style(italic=True),
            },
        )


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument(
        "--out",
        default=None,
        help="output directory (default: tmp/xlspy-test-files)",
    )
    args = parser.parse_args(argv)

    out_dir = pathlib.Path(args.out) if args.out else ROOT / "tmp" / "xlspy-test-files"
    out_dir.mkdir(parents=True, exist_ok=True)
    xlsx_path = out_dir / "styling_demo.xlsx"
    xlsb_path = out_dir / "styling_demo.xlsb"

    with XlsxWriter(xlsx_path) as writer:
        _flat_sheet(writer)
        _font_sheet(writer)
        _border_sheet(writer)
        _alignment_sheet(writer)
        _fill_sheet(writer)
        _cf_sheet(writer)
        _table_sheet(writer)
        _chart_sheet(writer)
        _layout_sheet(writer)
        _format_sheet(writer)
    _write_xlsb(xlsb_path)

    for path in (xlsx_path, xlsb_path):
        print(f"{path}  ({path.stat().st_size:,} B)".replace(",", " "))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
