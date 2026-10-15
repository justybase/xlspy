"""Write a pandas DataFrame to an .xlsx file with a single call.

The one-liner ``to_xlsx(df, path)`` is the quickest way to persist a table.
pandas is optional: install it with ``pip install xlspy[pandas]``.

Run:
    python examples/dataframe_to_xlsx.py [output_dir]
"""
import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "../src")))

try:
    import pandas as pd
except ImportError:  # pragma: no cover - depends on the environment
    sys.exit("This example requires pandas. Install it with: pip install xlspy[pandas]")

from xlspy import to_xlsx


def main(output_dir: str) -> None:
    df = pd.DataFrame(
        {
            "name": ["Alice", "Bob", "Charlie"],
            "age": [25, 30, 35],
            "score": [1.5, 2.75, 3.0],
            "active": [True, False, True],
            "joined": pd.to_datetime(["2024-01-15", "2023-06-01", "2022-12-31"]),
        }
    )

    path = os.path.join(output_dir, "dataframe_to_xlsx.xlsx")
    to_xlsx(df, path)

    print(f"Wrote {len(df)} rows x {len(df.columns)} columns to {path}")


if __name__ == "__main__":
    directory = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "output"
    )
    os.makedirs(directory, exist_ok=True)
    main(directory)
