"""Mix DataFrames and plain row data across several sheets.

``write_dataframe`` behaves exactly like ``write_sheet``: it writes into the
last sheet added with ``add_sheet``, so you can freely combine DataFrame
tables with hand-written rows in one workbook.

Run:
    python examples/dataframe_writer_methods.py [output_dir]
"""
import datetime
import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "../src")))

try:
    import pandas as pd
except ImportError:  # pragma: no cover - depends on the environment
    sys.exit("This example requires pandas. Install it with: pip install xlspy[pandas]")

from xlspy import XlsbWriter, XlsxWriter


def build_workbook(writer_cls, path: str) -> None:
    df = pd.DataFrame(
        {
            "region": ["North", "South", "East"],
            "revenue": [120000.0, 85000.0, 64000.0],
        }
    )

    with writer_cls(path) as writer:
        # Sheet 1: a DataFrame table
        writer.add_sheet("Revenue")
        writer.write_dataframe(df)

        # Sheet 2: plain rows written by hand
        writer.add_sheet("Notes")
        writer.write_sheet(
            [
                ["Generated", datetime.datetime(2024, 1, 15, 9, 0)],
                ["Owner", "finance-team"],
            ]
        )

        # Sheet 3: hidden helper sheet
        writer.add_sheet("Hidden", hidden=True)
        writer.write_dataframe(pd.DataFrame({"v": [1, 2, 3]}))


def main(output_dir: str) -> None:
    xlsx_path = os.path.join(output_dir, "writer_methods.xlsx")
    xlsb_path = os.path.join(output_dir, "writer_methods.xlsb")

    build_workbook(XlsxWriter, xlsx_path)
    build_workbook(XlsbWriter, xlsb_path)

    print(f"Wrote {xlsx_path}")
    print(f"Wrote {xlsb_path}")


if __name__ == "__main__":
    directory = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "output"
    )
    os.makedirs(directory, exist_ok=True)
    main(directory)
