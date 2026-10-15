"""Round-trip a DataFrame through xlspy: write it, read it back, rebuild it.

Demonstrates that data written with ``to_xlsx`` / ``to_xlsb`` comes back
cleanly through ``ExcelReader``, so you can use xlspy as a plain I/O layer
between pandas and spreadsheets.

Run:
    python examples/dataframe_roundtrip.py [output_dir]
"""
import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "../src")))

try:
    import pandas as pd
except ImportError:  # pragma: no cover - depends on the environment
    sys.exit("This example requires pandas. Install it with: pip install xlspy[pandas]")

from xlspy import ExcelReader, F, to_xlsb, to_xlsx


def read_back(path: str, sheet: str = "Sheet1"):
    with ExcelReader(path) as reader:
        rows = reader.read_all(sheet)
    if not rows:
        return None
    return pd.DataFrame(rows[1:], columns=rows[0])


def main(output_dir: str) -> None:
    original = pd.DataFrame(
        {
            "name": ["Alice", "Bob", "Charlie"],
            "age": [25, 30, 35],
            "joined": pd.to_datetime(["2024-01-15", "2023-06-01", "2022-12-31"]),
        }
    )

    for ext, shortcut in (("xlsx", to_xlsx), ("xlsb", to_xlsb)):
        path = os.path.join(output_dir, f"roundtrip.{ext}")
        shortcut(original, path, formats={"joined": F.DATE_ISO})

        restored = read_back(path)
        print(f"--- {ext} -> {path} ---")
        print(restored)
        print()


if __name__ == "__main__":
    directory = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "output"
    )
    os.makedirs(directory, exist_ok=True)
    main(directory)
