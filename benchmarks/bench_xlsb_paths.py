"""Which path each XLSB write scenario takes, and how long it takes.

    python benchmarks/bench_xlsb_paths.py [--rows 100000] [--only substring]

Prints, per scenario, the seconds, the rows the native encoder took, the rows Python
wrote cell by cell, and the per-cell calls into Python. The scenarios are those of
``tests/test_native_write_coverage.py`` (``tests/native_paths.py``).
"""

from __future__ import annotations

import argparse
import os
import sys
import tempfile
import time
import warnings

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, os.path.join(ROOT, "src"))
sys.path.insert(0, os.path.join(ROOT, "tests"))

from native_paths import SCENARIOS, write  # noqa: E402


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--rows", type=int, default=100_000)
    parser.add_argument("--only", help="run the scenarios whose name contains this text")
    args = parser.parse_args()
    warnings.simplefilter("ignore")
    print(f"{'scenario':34} {'seconds':>8} {'native':>9} {'python':>9} {'scalar':>9} {'formula':>9}")
    with tempfile.TemporaryDirectory(prefix="xlspy-paths-") as directory:
        for scenario in SCENARIOS:
            if args.only and args.only not in scenario.name:
                continue
            path = os.path.join(directory, "out.xlsb")
            start = time.perf_counter()
            counts = write(scenario, args.rows, __import__("pathlib").Path(path))
            seconds = time.perf_counter() - start
            note = f"  <- {scenario.gap}" if scenario.gap else ""
            print(
                f"{scenario.name:34} {seconds:8.2f} {counts['native']:>9} {counts['python_rows']:>9} "
                f"{counts['scalar_hook']:>9} {counts['formula_calls']:>9}{note}"
            )


if __name__ == "__main__":
    main()
