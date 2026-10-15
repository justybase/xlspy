"""Write a pandas DataFrame to a compact .xlsb file.

The API is identical to the .xlsx example -- only the shortcut (and the file
extension) change. XLSB files are binary and usually much smaller, and the
XLSB writer streams rows, which keeps memory usage low for big tables.

Run:
    python examples/dataframe_to_xlsb.py [output_dir]
"""
import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "../src")))

try:
    import pandas as pd
except ImportError:  # pragma: no cover - depends on the environment
    sys.exit("This example requires pandas. Install it with: pip install xlspy[pandas]")

from xlspy import to_xlsb


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

    path = os.path.join(output_dir, "dataframe_to_xlsb.xlsb")
    to_xlsb(df, path)

    size = os.path.getsize(path)
    print(f"Wrote {len(df)} rows to {path} ({size} bytes)")


if __name__ == "__main__":
    directory = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "output"
    )
    os.makedirs(directory, exist_ok=True)
    main(directory)
