"""Generate an XLSB workbook demonstrating the full styling capabilities.

Usage:
    python3 tools/make_xlsb_full_demo.py [--out DIR]

Writes ``xlsb_full_styling_demo.xlsb`` (default: ``tmp/xlspy-test-files/``).
Every feature below is supported by the XLSB backend; XLSX-only constructs
(conditional formatting, tables, relative indent, vertical/horizontal/start/end
border sides) are intentionally absent.
"""

from __future__ import annotations

import argparse
import pathlib

import sys

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent / "src"))

from xlspy import (  # noqa: E402
    Alignment,
    Border,
    Chart,
    ChartSeries,
    Font,
    GradientFill,
    PatternFill,
    Protection,
    Side,
    Stop,
    Style,
    XlsbWriter,
)

ROOT = pathlib.Path(__file__).resolve().parent.parent

HEADER = Style(font=Font(bold=True, color="FFFFFF"), fill="305496")


def _fonts_sheet(writer: XlsbWriter) -> None:
    data = [
        ["Element", "Demo", "Expected"],
        ["bold + italic + strike", "Bold italic struck", "all three at once"],
        ["size=18, color red", "Big red", "18 pt, red"],
        ["name='Arial'", "Arial typeface", "Arial, not Calibri"],
        ["underline single/double", "Single / Double", "one and two lines"],
        ["accounting underlines", "Acct1 / Acct2", "lines below baseline"],
        ["superscript / subscript", "x2 / H2O", "raised / lowered"],
        ["outline + shadow", "Outlined shadowed", "glyph effects"],
        ["condense + extend", "Condensed / Expanded", "spacing effects"],
        ["theme color + tint", "Themed", "lighter accent shade"],
        ["indexed color", "Indexed", "palette red"],
    ]
    styles = {
        "B2": Style(font=Font(bold=True, italic=True, strike=True)),
        "B3": Style(font=Font(size=18, color="FF0000")),
        "B4": Style(font=Font(name="Arial", size=12)),
        "B5": Style(font=Font(underline="single")),
        "B6": Style(font=Font(underline="singleAccounting")),
        "B7": Style(font=Font(vertAlign="superscript")),
        "B8": Style(font=Font(outline=True, shadow=True)),
        "B9": Style(font=Font(condense=True)),
        "B10": Style(font=Font(color="4472C4")),
        "B11": Style(font=Font(color="FF0000")),
    }
    writer.add_sheet("Fonts")
    writer.write_sheet(
        data, header_style=HEADER, styles=styles,
        column_widths={"A": 28, "B": 26, "C": 40}, auto_filter=False,
    )


def _fills_sheet(writer: XlsbWriter) -> None:
    data = [
        ["Element", "Demo", "Expected"],
        ["solid green", "Solid", "solid fill"],
        ["gray125 pattern", "Gray125", "Excel gray125"],
        ["darkGrid two colors", "Dark grid", "pattern fg + bg"],
        ["linear gradient 90°", "Red → blue", "top to bottom"],
        ["linear gradient 0°", "Green → white", "left to right"],
        ["path gradient", "Radial", "from center"],
    ]
    styles = {
        "B2": Style(fill=PatternFill("solid", start_color="00B050")),
        "B3": Style(fill=PatternFill("gray125", start_color="1F3864", end_color="FFF2CC")),
        "B4": Style(fill=PatternFill("darkGrid", start_color="C00000", end_color="FFFFFF")),
        "B5": Style(fill=GradientFill("linear", degree=90,
                                      stop=[Stop("C00000", 0.0), Stop("0070C0", 1.0)])),
        "B6": Style(fill=GradientFill("linear", degree=0,
                                      stop=["008000", "FFFFFF"])),
        "B7": Style(fill=GradientFill("path", left=0.2, right=0.2, top=0.2,
                                      bottom=0.2, stop=["FFF2CC", "BF8F00"])),
    }
    writer.add_sheet("Fills")
    writer.write_sheet(
        data, header_style=HEADER, styles=styles,
        column_widths={"A": 28, "B": 26, "C": 40}, auto_filter=False,
    )


def _borders_sheet(writer: XlsbWriter) -> None:
    thin = Side("thin")
    data = [
        ["Element", "Demo", "Expected"],
        ["thin frame", "Framed", "frame on all sides"],
        ["thick red bottom", "Bottom", "red thick bottom edge"],
        ["dashed top + dotted left", "Dashed/dotted", "mixed edges"],
        ["double + hair", "Double/hair", "mixed edges"],
        ["slantDashDot diagonal up", "Diagonal /", "bottom-left to top-right"],
        ["diagonal both ways", "Diagonal X", "two diagonals"],
    ]
    full = Border(left=thin, right=thin, top=thin, bottom=thin)
    styles = {
        "B2": Style(border=full),
        "B3": Style(border=Border(bottom=Side("thick", "C00000"))),
        "B4": Style(border=Border(top=Side("dashed"), left=Side("dotted"))),
        "B5": Style(border=Border(top=Side("double"), bottom=Side("hair"))),
        "B6": Style(border=Border(diagonal=Side("slantDashDot", "0070C0"),
                                  diagonal_up=True)),
        "B7": Style(border=Border(diagonal=Side("thin"),
                                  diagonal_up=True, diagonal_down=True)),
    }
    writer.add_sheet("Borders")
    writer.write_sheet(
        data, header_style=HEADER, styles=styles,
        column_widths={"A": 32, "B": 26, "C": 40}, auto_filter=False,
    )


def _alignment_sheet(writer: XlsbWriter) -> None:
    data = [
        ["Element", "Demo", "Expected"],
        ["center + wrap + middle", "Centered wrapped text in a narrow column", "centered, wrapped"],
        ["left + indent 3", "Indented", "pushed from the left"],
        ["right + shrink", "Shrunk to fit a narrow column", "font shrinks"],
        ["rotation 45°", "Rotated", "45° upwards"],
        ["rotation 255 (vertical)", "Vertical", "top to bottom"],
        ["justify + justify last", "Justified paragraph stretching across lines", "stretched lines"],
    ]
    styles = {
        "B2": Style(alignment=Alignment(horizontal="center", vertical="center",
                                        wrap_text=True)),
        "B3": Style(alignment=Alignment(horizontal="left", indent=3)),
        "B4": Style(alignment=Alignment(horizontal="right", shrink_to_fit=True)),
        "B5": Style(alignment=Alignment(text_rotation=45)),
        "B6": Style(alignment=Alignment(text_rotation=255)),
        "B7": Style(alignment=Alignment(horizontal="justify", wrap_text=True,
                                        justify_last_line=True)),
    }
    writer.add_sheet("Alignment")
    writer.write_sheet(
        data, header_style=HEADER, styles=styles,
        column_widths={"A": 32, "B": 26, "C": 40},
        row_heights={1: 30}, auto_filter=False,
    )


def _protection_sheet(writer: XlsbWriter) -> None:
    data = [
        ["Element", "Demo", "Expected (after sheet protection is enabled in Excel)"],
        ["locked=False", "Editable", "editable when the sheet is protected"],
        ["locked=True (default)", "Locked", "locked when the sheet is protected"],
        ["hidden=True", "Hidden formula", "formula hidden when the sheet is protected"],
    ]
    styles = {
        "B2": Style(protection=Protection(locked=False)),
        "B3": Style(protection=Protection(locked=True)),
        "B4": Style(protection=Protection(hidden=True)),
    }
    writer.add_sheet("Protection")
    writer.write_sheet(
        data, header_style=HEADER, styles=styles,
        column_widths={"A": 32, "B": 20, "C": 58}, auto_filter=False,
    )


def _layout_sheet(writer: XlsbWriter) -> None:
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
        ["Column B is 28 wide; column D is hidden.", None, None, None, None],
        ["Freeze B2: row 1 and column A stay visible.", None, None, None, None],
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


def _charts_sheet(writer: XlsbWriter) -> None:
    data = [
        ["Quarter", "Hardware", "Services"],
        ["Q1", 120, 80],
        ["Q2", 150, 95],
        ["Q3", 130, 110],
        ["Q4", 170, 125],
        [None, None, None],
        ["Charts: bar (E2, style 10) + line (E18) on this sheet.", None, None],
        ["The bar chart reads both series; the line chart reads Hardware.", None, None],
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
                title="Revenue by quarter (XLSB)",
                anchor="E2",
                style=10,
            ),
            Chart(
                "line",
                [ChartSeries("B2:B5", title_cell="B1", categories="A2:A5")],
                title="Hardware trend (XLSB)",
                anchor="E18",
            ),
        ],
        column_widths={"A": 40, "B": 12, "C": 12, "D": 3},
        auto_filter=False,
    )


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--out", default=None,
                        help="output directory (default: tmp/xlspy-test-files)")
    args = parser.parse_args(argv)
    out_dir = pathlib.Path(args.out) if args.out else ROOT / "tmp" / "xlspy-test-files"
    out_dir.mkdir(parents=True, exist_ok=True)
    path = out_dir / "xlsb_full_styling_demo.xlsb"
    with XlsbWriter(path) as writer:
        _fonts_sheet(writer)
        _fills_sheet(writer)
        _borders_sheet(writer)
        _alignment_sheet(writer)
        _protection_sheet(writer)
        _layout_sheet(writer)
        _charts_sheet(writer)
    print(f"{path}  ({path.stat().st_size:,} B)".replace(",", " "))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
