"""Stream a large SQL result to XLSX (or XLSB) with bounded memory.

The rows come straight from a DB-API cursor, one at a time. Nothing is
collected into a list, and with ``useSharedStrings=False`` the text stays in
the cells, so no shared-string table grows with the export.

This script uses the standard-library ``sqlite3`` module and builds its own
sample table so it runs anywhere. Replace ``connect()`` and the query with your
database driver (for example a pyodbc or nzpy cursor).

Usage:
    python examples/xlsx_streaming_export.py [--rows 200000] [--xlsb]
"""
import argparse
import os
import sqlite3
import sys
import tracemalloc

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "../src")))

from xlspy import XlsbWriter, XlsxWriter  # noqa: E402


def build_sample_database(rows):
    connection = sqlite3.connect(":memory:")
    connection.execute("CREATE TABLE sales (id INTEGER, customer TEXT, amount REAL, sold_on TEXT)")
    connection.executemany(
        "INSERT INTO sales VALUES (?, ?, ?, ?)",
        ((i, f"Customer {i % 5000}", (i * 37) % 1000 / 10, f"2026-{(i % 12) + 1:02d}-01") for i in range(rows)),
    )
    return connection


def cursor_rows(cursor):
    """Yield the header, then one row per fetched record, without buffering."""
    yield [column[0] for column in cursor.description]
    while True:
        row = cursor.fetchone()
        if row is None:
            break
        yield list(row)


def export(connection, path, use_xlsb=False):
    cursor = connection.execute("SELECT id, customer, amount, sold_on FROM sales ORDER BY id")
    writer_class = XlsbWriter if use_xlsb else XlsxWriter
    with writer_class(path, useSharedStrings=False) as writer:
        writer.add_sheet("Sales")
        writer.write_sheet(cursor_rows(cursor), freeze_panes="A2", auto_filter=True)
    return path


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--rows", type=int, default=200_000)
    parser.add_argument("--xlsb", action="store_true", help="write XLSB instead of XLSX")
    args = parser.parse_args()

    connection = build_sample_database(args.rows)
    path = "sales_export.xlsb" if args.xlsb else "sales_export.xlsx"

    tracemalloc.start()
    export(connection, path, use_xlsb=args.xlsb)
    _, peak = tracemalloc.get_traced_memory()
    tracemalloc.stop()
    print(f"wrote {path} ({args.rows:,} rows); peak Python allocations {peak / 2**20:.1f} MB")


if __name__ == "__main__":
    main()
