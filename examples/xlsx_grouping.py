"""Row and column groups (the Excel +/- outline buttons), two levels each.

Each region has two detail groups (level 2) with subtotal rows, inside a
region group (level 1) closed by a region total. The months are grouped under
their quarter columns (level 2), and both quarters under the half-year total
(level 1). Summary rows sit below their group and summary columns to the
right of it, which is Excel's default.

Usage:
    python examples/xlsx_grouping.py [--out grouping.xlsx]

An ``.xlsb`` output name writes XLSB. XLSB has no formula cells, so it gets
the computed totals; XLSX gets SUM formulas with cached results.
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "../src")))

from xlspy import Formula, Group, Style, XlsbWriter, XlsxWriter  # noqa: E402

LETTERS = "ABCDEFGHIJKL"
MONTHS = ["Sty", "Lut", "Mar", "Kwi", "Maj", "Cze"]
# Region -> (detail rows, wholesale rows), six months each.
SALES = {
    "Północ": (
        [[50, 55, 60, 62, 66, 70], [45, 50, 58, 60, 65, 68]],
        [[80, 85, 90, 92, 95, 100], [70, 75, 82, 85, 88, 92]],
    ),
    "Południe": (
        [[55, 60, 65, 68, 72, 75], [48, 54, 62, 65, 69, 74]],
        [[85, 90, 95, 98, 102, 108], [75, 80, 88, 90, 94, 99]],
    ),
    "Zachód": (
        [[42, 47, 52, 55, 59, 63], [40, 44, 49, 52, 56, 60]],
        [[72, 78, 84, 87, 91, 96], [68, 73, 79, 82, 86, 91]],
    ),
}
# Months in D:F and H:J, quarter totals in G and K, the half-year in L.
MONTH_COLUMNS = [3, 4, 5, 7, 8, 9]

TITLE = Style(fill="#17375E", font_color="#FFFFFF", bold=True)
HEADER = Style(fill="#1F4E79", font_color="#FFFFFF", bold=True, horizontal="center")
SUBTOTAL = Style(fill="#DDEBF7", bold=True)
REGION_TOTAL = Style(fill="#9BC2E6", bold=True)


def _with_totals(months):
    """Six monthly values -> the D..L cells: months, Q1, months, Q2, H1."""
    q1, q2 = sum(months[:3]), sum(months[3:])
    return months[:3] + [q1] + months[3:] + [q2, q1 + q2]


def _cell(formula, value, use_formulas):
    return Formula(formula, value) if use_formulas else value


def _row(cat1, cat2, values, row_number, use_formulas, sources=None):
    """One sheet row. ``sources`` are the 1-based rows a subtotal adds up."""
    cells = [None, cat1, cat2]
    for offset, value in enumerate(values):
        col = LETTERS[3 + offset]
        if col == "G":
            formula = f"SUM(D{row_number}:F{row_number})"
        elif col == "K":
            formula = f"SUM(H{row_number}:J{row_number})"
        elif col == "L":
            formula = f"G{row_number}+K{row_number}"
        elif sources:
            formula = "+".join(f"{col}{source}" for source in sources)
        else:
            cells.append(value)
            continue
        cells.append(_cell(formula, value, use_formulas))
    return cells


def build_rows(use_formulas):
    rows = [
        [None, "Przykład: podwójne grupowanie wierszy i kolumn"],
        [None, "Kategoria 1", "Kategoria 2", *MONTHS[:3], "Q1", *MONTHS[3:], "Q2", "Razem H1"],
    ]
    row_groups, subtotal_rows, region_rows = [], [], []
    for region, (detail, wholesale) in SALES.items():
        start = len(rows)  # 0-based index of the region's first row
        subtotals = []
        for label, block in (("Detal", detail), ("Hurt", wholesale)):
            first = len(rows)
            for months in block:
                rows.append(_row(region, label, _with_totals(months), len(rows) + 1, use_formulas))
            sums = [sum(column) for column in zip(*(_with_totals(m) for m in block))]
            rows.append(_row(region, f"Suma {label}", sums, len(rows) + 1, use_formulas,
                             sources=[first + 1, first + 2]))
            row_groups.append(Group(first, first + 1))
            subtotal_rows.append(len(rows) - 1)
            subtotals.append((len(rows), sums))
        total = [a + b for a, b in zip(subtotals[0][1], subtotals[1][1])]
        rows.append(_row(region, "SUMA REGIONU", total, len(rows) + 1, use_formulas,
                         sources=[subtotals[0][0], subtotals[1][0]]))
        row_groups.append(Group(start, len(rows) - 2))
        region_rows.append(len(rows) - 1)

    styles = {"B1:L1": TITLE, "B2:L2": HEADER}
    for index in subtotal_rows:
        styles[f"B{index + 1}:L{index + 1}"] = SUBTOTAL
    for index in region_rows:
        styles[f"B{index + 1}:L{index + 1}"] = REGION_TOTAL
    return rows, row_groups, styles


# Months D:F and H:J at level 2, inside D:K at level 1. G, K and L are the
# summary columns, to the right of their groups.
COLUMN_GROUPS = [Group("D", "K"), Group("D", "F"), Group("H", "J")]


def write(path):
    is_xlsb = path.lower().endswith(".xlsb")
    rows, row_groups, styles = build_rows(use_formulas=not is_xlsb)
    writer_cls = XlsbWriter if is_xlsb else XlsxWriter
    with writer_cls(path) as writer:
        writer.add_sheet("Grupowanie")
        writer.write_sheet(
            rows,
            styles=styles,
            column_widths={"A": 3, "B": 14, "C": 16},
            row_groups=row_groups,
            column_groups=COLUMN_GROUPS,
            freeze_panes="",
            auto_filter=False,
        )
    return path


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--out", default="grouping.xlsx", help="output .xlsx or .xlsb file")
    args = parser.parse_args()
    print(f"wrote {write(args.out)}")


if __name__ == "__main__":
    main()
