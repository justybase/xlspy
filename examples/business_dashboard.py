"""A one-sheet sales dashboard: KPI cards, a slicer, charts and sparklines.

The source table ``Sprzedaz`` sits at the bottom of the sheet. Every figure on
the dashboard is a ``SUMIFS`` over it that counts only the rows still visible,
so the slicer on ``Region`` updates the cards, the summaries and the charts:

* the ``Widoczny`` column holds ``SUBTOTAL(103, A<row>)``, which is 1 for a row
  the table filter shows and 0 for one it hides;
* the summaries use plain A1 ranges (XLSB cannot hold structured references).

Usage:
    python examples/business_dashboard.py [--out dashboard.xlsx]

An ``.xlsb`` output name writes XLSB. XLSB has no formula cells and no slicers
or sparklines yet, so it gets the computed figures, the cards, the charts and
the table, and the slicer and sparklines are left out.
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "../src")))

from xlspy import (  # noqa: E402
    Chart, ChartAxis, ChartLegend, ChartMarker, ChartSeries, ExcelTable, Font, Formula,
    Shape, Slicer, SparklineGroup, Style, XlsbWriter, XlsxWriter,
)

REGIONS = ["Północ", "Południe", "Zachód", "Wschód"]
PRODUCTS = ["Laptopy", "Monitory", "Drukarki"]
CHANNELS = ["Sklep", "Online", "Partnerzy"]
MONTHS = ["Sty", "Lut", "Mar", "Kwi", "Maj", "Cze"]
HEADERS = ["Region", "Produkt", "Kanał", "Miesiąc", "Sprzedaż", "Koszt", "Marża", "Widoczny"]

HEADER_ROW = 50
FIRST_DATA_ROW = HEADER_ROW + 1
LAST_DATA_ROW = FIRST_DATA_ROW + len(REGIONS) * len(PRODUCTS) * len(CHANNELS) * len(MONTHS) - 1
LAST_COLUMN = "T"

TITLE = Style(fill="#1F4E79", font=Font(size=18, bold=True, color="#FFFFFF"), vertical="center")
SECTION = Style(font_color="#1F4E79", bold=True)
LABEL = Style(font_color="#595959", bold=True)
HEADER = Style(fill="#1F4E79", font_color="#FFFFFF", bold=True, horizontal="center")
MONEY = Style(number_format='#,##0 "zł"')
PLAIN_MONEY = Style(number_format="#,##0")
PERCENT = Style(number_format="0.0%")
COUNT = Style(number_format="#,##0")

CARD_FILL = "#F2F6FB"
CARD_LINE = "#9DC3E6"
FRAME = {"fill": "#FFFFFF", "border": "#D9D9D9", "font_size": 9, "font_color": "#404040"}


def sample_rows():
    """Deterministic (region, product, channel, month, sales, cost) rows, one per combination."""
    rows = []
    for ri, region in enumerate(REGIONS):
        for pi, product in enumerate(PRODUCTS):
            for ci, channel in enumerate(CHANNELS):
                for mi, month in enumerate(MONTHS):
                    sales = (40 + (ri * 31 + pi * 17 + ci * 13 + mi * 7) % 60) * 100
                    cost = sales * (55 + 5 * ((ri + pi + ci) % 3)) // 100
                    rows.append((region, product, channel, month, sales, cost))
    return rows


def _money(value):
    return f"{value:,.0f} zł".replace(",", " ")


def _percent(value):
    return f"{value:.1%}"


def _totals(rows):
    """The figures the formulas compute, for cached values and for XLSB."""
    # Every row is visible in the file as saved, so the sums run over all rows.
    margins = [(row[0], row[1], row[2], row[3], row[4], row[4] - row[5]) for row in rows]
    total_sales = sum(row[4] for row in rows)
    return {
        "sales": total_sales,
        "margin": sum(row[5] for row in margins),
        "orders": len(rows),
        "online": sum(row[4] for row in rows if row[2] == "Online"),
        "product": {p: sum(row[4] for row in rows if row[1] == p) for p in PRODUCTS},
        "channel": {c: sum(row[4] for row in rows if row[2] == c) for c in CHANNELS},
        "month": {m: sum(row[4] for row in rows if row[3] == m) for m in MONTHS},
        "month_margin": {m: sum(row[4] - row[5] for row in rows if row[3] == m) for m in MONTHS},
        "region_month": {
            (r, m): sum(row[4] for row in rows if row[0] == r and row[3] == m)
            for r in REGIONS for m in MONTHS
        },
    }


def _place(grid, ref, value):
    grid[int(ref[1:]) - 1][ord(ref[0]) - ord("A")] = value


def build_grid(rows, totals, use_formulas):
    """The sheet as a list of rows; formula cells carry their computed result."""
    grid = [[None] * (ord(LAST_COLUMN) - ord("A") + 1) for _ in range(LAST_DATA_ROW)]

    def formula(text, value):
        return Formula(text, value) if use_formulas else value

    def rng(letter):
        return f"${letter}${FIRST_DATA_ROW}:${letter}${LAST_DATA_ROW}"

    visible = rng("H")
    _place(grid, "A1", "Dashboard sprzedaży 2026")
    _place(grid, "A2", "Wybierz region w fragmentatorze; karty, wykresy i podsumowania przeliczą się.")
    _place(grid, "A4", "Sprzedaż (widoczne)")
    _place(grid, "C4", "Marża")
    _place(grid, "E4", "Zamówienia")
    _place(grid, "G4", "Udział Online")

    # Helper cells the KPI cards link to.
    _place(grid, "S4", "Wskaźniki dla kart")
    _place(grid, "S5", "Sprzedaż")
    _place(grid, "T5", formula(f"SUMIFS({rng('E')},{visible},1)", totals["sales"]))
    _place(grid, "S6", "Marża %")
    _place(grid, "T6", formula(
        f"IF(T5=0,0,SUMIFS({rng('G')},{visible},1)/T5)",
        totals["margin"] / totals["sales"]))
    _place(grid, "S7", "Zamówienia")
    _place(grid, "T7", formula(f"COUNTIFS({visible},1)", totals["orders"]))
    _place(grid, "S8", "Udział Online")
    _place(grid, "T8", formula(
        f'IF(T5=0,0,SUMIFS({rng("E")},{rng("C")},"Online",{visible},1)/T5)',
        totals["online"] / totals["sales"]))

    # Product summary (chart 1).
    _place(grid, "K14", "Sprzedaż wg produktu")
    _place(grid, "K15", "Produkt")
    _place(grid, "L15", "Sprzedaż")
    for index, product in enumerate(PRODUCTS):
        row = 16 + index
        _place(grid, f"K{row}", product)
        _place(grid, f"L{row}", formula(
            f"SUMIFS({rng('E')},{rng('B')},K{row},{visible},1)", totals["product"][product]))

    # Channel summary (doughnut).
    _place(grid, "K20", "Udział kanałów")
    _place(grid, "K21", "Kanał")
    _place(grid, "L21", "Sprzedaż")
    for index, channel in enumerate(CHANNELS):
        row = 22 + index
        _place(grid, f"K{row}", channel)
        _place(grid, f"L{row}", formula(
            f"SUMIFS({rng('E')},{rng('C')},K{row},{visible},1)", totals["channel"][channel]))

    # Monthly summary (combined chart: sales bars, margin line on the second axis).
    _place(grid, "K26", "Sprzedaż i marża wg miesiąca")
    _place(grid, "K27", "Miesiąc")
    _place(grid, "L27", "Sprzedaż")
    _place(grid, "M27", "Marża %")
    for index, month in enumerate(MONTHS):
        row = 28 + index
        _place(grid, f"K{row}", month)
        _place(grid, f"L{row}", formula(
            f"SUMIFS({rng('E')},{rng('D')},K{row},{visible},1)", totals["month"][month]))
        margin = totals["month_margin"][month] / totals["month"][month]
        _place(grid, f"M{row}", formula(
            f"IF(L{row}=0,0,SUMIFS({rng('G')},{rng('D')},K{row},{visible},1)/L{row})", margin))

    # Region trend: one row per region, months across, a sparkline in column R.
    _place(grid, "K35", "Trend wg regionu")
    _place(grid, "K36", "Region")
    for index, month in enumerate(MONTHS):
        _place(grid, f"{chr(ord('L') + index)}36", month)
    _place(grid, "R36", "Trend")
    for region_index, region in enumerate(REGIONS):
        row = 37 + region_index
        _place(grid, f"K{row}", region)
        for index, month in enumerate(MONTHS):
            column = chr(ord("L") + index)
            _place(grid, f"{column}{row}", formula(
                f"SUMIFS({rng('E')},{rng('A')},$K{row},{rng('D')},{column}$36,{visible},1)",
                totals["region_month"][(region, month)]))

    # The source table, Widoczny holding the filter flag.
    _place(grid, f"A{HEADER_ROW - 1}", "Dane źródłowe (tabela Sprzedaz)")
    for index, header in enumerate(HEADERS):
        _place(grid, f"{chr(ord('A') + index)}{HEADER_ROW}", header)
    for offset, (region, product, channel, month, sales, cost) in enumerate(rows):
        row = FIRST_DATA_ROW + offset
        _place(grid, f"A{row}", region)
        _place(grid, f"B{row}", product)
        _place(grid, f"C{row}", channel)
        _place(grid, f"D{row}", month)
        _place(grid, f"E{row}", sales)
        _place(grid, f"F{row}", cost)
        _place(grid, f"G{row}", formula(f"E{row}-F{row}", sales - cost))
        _place(grid, f"H{row}", formula(f"SUBTOTAL(103,A{row})", 1))

    return grid


def build_dashboard(path):
    use_formulas = not str(path).lower().endswith(".xlsb")
    rows = sample_rows()
    totals = _totals(rows)
    grid = build_grid(rows, totals, use_formulas)

    cards = [
        Shape(anchor="A5", width=196, height=60, fill=CARD_FILL, line=CARD_LINE, font_size=20,
              font_color="#1F4E79", bold=True, align="center", valign="center",
              link_cell="T5", text=_money(totals["sales"]), name="Karta sprzedaży"),
        Shape(anchor="C5", width=196, height=60, fill=CARD_FILL, line=CARD_LINE, font_size=20,
              font_color="#1F4E79", bold=True, align="center", valign="center",
              link_cell="T6", text=_percent(totals["margin"] / totals["sales"]), name="Karta marży"),
        Shape(anchor="E5", width=196, height=60, fill=CARD_FILL, line=CARD_LINE, font_size=20,
              font_color="#1F4E79", bold=True, align="center", valign="center",
              link_cell="T7", text=f"{totals['orders']}", name="Karta zamówień"),
        Shape(anchor="G5", width=196, height=60, fill=CARD_FILL, line=CARD_LINE, font_size=20,
              font_color="#1F4E79", bold=True, align="center", valign="center",
              link_cell="T8", text=_percent(totals["online"] / totals["sales"]), name="Karta Online"),
    ]

    charts = [
        Chart("bar", [ChartSeries("L16:L18", categories="K16:K18", title="Sprzedaż",
                                  color="#2E75B6")],
              title="Sprzedaż wg produktu", anchor="A9", width=13, height=7, legend=False,
              show_data_labels=True, gap_width=60, y_axis=ChartAxis(number_format="#,##0"),
              **FRAME),
        Chart("doughnut", [ChartSeries("L22:L24", categories="K22:K24")],
              title="Udział kanałów", anchor="F9", width=8, height=7, hole_size=60,
              legend=ChartLegend(position="bottom"), show_data_labels=True, **FRAME),
        Chart("bar", [
            ChartSeries("L28:L33", categories="K28:K33", title="Sprzedaż", color="#9DC3E6"),
            ChartSeries("M28:M33", categories="K28:K33", title="Marża %", chart_type="line",
                        secondary_axis=True, color="#C55A11", marker=ChartMarker()),
        ], title="Sprzedaż i marża wg miesiąca", anchor="A25", width=21, height=8,
           gap_width=60, y2_axis=ChartAxis(number_format="0%"), legend=ChartLegend(position="bottom"),
           **FRAME),
    ]

    widths = {letter: 14 for letter in "ABCDEFGH"}
    widths.update({"I": 3, "J": 3, "K": 20, "L": 14, "M": 12, "N": 12, "O": 12, "P": 12,
                   "Q": 12, "R": 16, "S": 18, "T": 14})
    styles = {
        "A1:H1": TITLE,
        "A2": Style(font_color="#595959", italic=True),
        "A4": LABEL, "C4": LABEL, "E4": LABEL, "G4": LABEL,
        "S4": SECTION, "S5:S8": LABEL,
        "T5": MONEY, "T6": PERCENT, "T7": COUNT, "T8": PERCENT,
        "K14": SECTION, "K20": SECTION, "K26": SECTION, "K35": SECTION,
        "A49": SECTION,
        "K15:L15": HEADER, "K21:L21": HEADER, "K27:M27": HEADER, "K36:R36": HEADER,
        "A50:H50": HEADER,
        "L16:L18": MONEY, "L22:L24": MONEY, "L28:L33": MONEY, "M28:M33": PERCENT,
        "L37:Q40": MONEY,
        f"E{FIRST_DATA_ROW}:G{LAST_DATA_ROW}": PLAIN_MONEY,
    }
    table = ExcelTable("Sprzedaz", f"A{HEADER_ROW}:H{LAST_DATA_ROW}", columns=HEADERS)

    options = dict(
        merges=["A1:H1", "A2:H2"],
        column_widths=widths,
        styles=styles,
        tables=[table],
        charts=charts,
        shapes=cards,
        show_grid_lines=False,
        zoom=90,
        tab_color="#1F4E79",
    )
    if use_formulas:
        options["slicers"] = [Slicer(table="Sprzedaz", column="Region", anchor="K4",
                                     width=144, height=180, caption="Region")]
        options["sparklines"] = [SparklineGroup(
            data="L37:Q40", location="R37:R40", kind="line", color="#2E75B6", markers=True,
            last_point=True,
        )]

    writer_class = XlsxWriter if use_formulas else XlsbWriter
    with writer_class(path) as writer:
        writer.add_sheet("Dashboard")
        writer.write_sheet(grid, **options)
    return use_formulas


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__.split("\n", 1)[0])
    parser.add_argument("--out", default="business_dashboard.xlsx",
                        help="output path; a .xlsb name writes XLSB")
    args = parser.parse_args(argv)
    xlsx = build_dashboard(args.out)
    if xlsx:
        print(f"Wrote {args.out}: cards, slicer, charts, sparklines and table.")
    else:
        print(f"Wrote {args.out}: cards, charts and table (no slicer or sparklines in XLSB yet).")


if __name__ == "__main__":
    main()
