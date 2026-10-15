"""Write a polars DataFrame to .xlsx and .xlsb.

The exact same API works for polars -- ``to_xlsx`` / ``to_xlsb`` and the
``write_dataframe`` writer method accept either a pandas or a polars frame.
polars lacks an index, so ``include_index=True`` is rejected with a clear
``ValueError``. polars is optional: ``pip install xlspy[polars]``.

Run:
    python examples/polars_to_xlsx_xlsb.py [output_dir]
"""
import datetime
import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "../src")))

try:
    import polars as pl
except ImportError:  # pragma: no cover - depends on the environment
    sys.exit("This example requires polars. Install it with: pip install xlspy[polars]")

from xlspy import F, to_xlsb, to_xlsx


def main(output_dir: str) -> None:
    df = pl.DataFrame(
        {
            "region": ["North", "South", "East"],
            "revenue": [120000.0, 85000.0, None],
            "margin": [0.21, 0.18, 0.15],
            "sold_on": [
                datetime.date(2024, 1, 15),
                datetime.date(2024, 2, 1),
                datetime.date(2024, 3, 10),
            ],
            "lead_time": [
                datetime.timedelta(hours=36),
                datetime.timedelta(hours=12),
                None,
            ],
        }
    )

    formats = {
        "revenue": F.CURRENCY_PLN,
        "margin": F.PERCENTAGE,
        "sold_on": F.DATE_ISO,
        "lead_time": F.TIME_HH_MM,
    }

    xlsx_path = os.path.join(output_dir, "polars.xlsx")
    xlsb_path = os.path.join(output_dir, "polars.xlsb")
    # Nulls are written as empty cells by default; use na_rep to label them.
    to_xlsx(df, xlsx_path, formats=formats, na_rep="n/a")
    to_xlsb(df, xlsb_path, formats=formats, na_rep="n/a")

    print(f"Wrote {xlsx_path}")
    print(f"Wrote {xlsb_path}")


if __name__ == "__main__":
    directory = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "output"
    )
    os.makedirs(directory, exist_ok=True)
    main(directory)
