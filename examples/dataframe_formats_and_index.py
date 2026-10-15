"""Per-column formats, the index and missing values.

Shows the three ``write_dataframe`` knobs:

* ``include_index=True`` -- prepend the pandas index as leading column(s),
* ``formats={column: format}`` -- apply number/date formats (see ``xlspy.F``),
* ``na_rep`` -- the text used for NaN/NaT/None (default is an empty cell).

Run:
    python examples/dataframe_formats_and_index.py [output_dir]
"""
import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "../src")))

try:
    import pandas as pd
except ImportError:  # pragma: no cover - depends on the environment
    sys.exit("This example requires pandas. Install it with: pip install xlspy[pandas]")

from xlspy import F, to_xlsx


def main(output_dir: str) -> None:
    df = pd.DataFrame(
        {
            "product": ["Gadget", "Widget", "Gizmo"],
            "revenue": [100000.0, 85000.0, 64000.0],
            "margin": [0.21, 0.18, None],
            "sold_on": pd.to_datetime(["2024-01-15", "2024-02-01", "2024-03-10"]),
        },
        index=pd.Index([101, 102, 103], name="sku"),
    )

    path = os.path.join(output_dir, "formats_and_index.xlsx")
    to_xlsx(
        df,
        path,
        include_index=True,
        na_rep="n/a",
        formats={
            "revenue": F.CURRENCY_PLN,
            "margin": F.PERCENTAGE,
            "sold_on": F.DATE_ISO,
        },
    )

    print(f"Wrote {path}")


if __name__ == "__main__":
    directory = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "output"
    )
    os.makedirs(directory, exist_ok=True)
    main(directory)
