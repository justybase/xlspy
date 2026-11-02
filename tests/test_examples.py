"""Every script in examples/ runs and writes its workbooks into the directory it is
given, not into the repository."""
from __future__ import annotations

import importlib.util
import subprocess
import sys
from pathlib import Path

import pytest

from helpers import write_workbook

EXAMPLES = Path(__file__).resolve().parents[1] / "examples"

NOT_RUN = {
    "benchmark.py": "needs a Netezza database",
    "performance_test.py": "a benchmark run of several minutes",
    "write_benchmark.py": "a benchmark run of several minutes",
    "business_dashboard.py": "test_business_dashboard_example.py runs it",
}
# Scripts that take an output directory as their only argument.
TAKES_DIRECTORY = {
    "dataframe_formats_and_index.py", "dataframe_roundtrip.py", "dataframe_to_xlsb.py",
    "dataframe_to_xlsx.py", "dataframe_writer_methods.py", "format_demo.py",
    "polars_to_xlsx_xlsb.py",
}
EXTRA_ARGS = {"xlsx_streaming_export.py": ["--rows", "2000"]}
NEEDS = {
    "dataframe_formats_and_index.py": "pandas", "dataframe_roundtrip.py": "pandas",
    "dataframe_to_xlsb.py": "pandas", "dataframe_to_xlsx.py": "pandas",
    "dataframe_writer_methods.py": "pandas", "polars_to_xlsx_xlsb.py": "polars",
}


def _scripts() -> list[str]:
    return sorted(path.name for path in EXAMPLES.glob("*.py") if path.name not in NOT_RUN)


def test_every_example_is_run_or_has_a_reason() -> None:
    names = {path.name for path in EXAMPLES.glob("*.py")}
    assert set(NOT_RUN) <= names, "a skipped example no longer exists"
    assert names == set(_scripts()) | set(NOT_RUN)


@pytest.mark.parametrize("script", _scripts())
def test_example_runs(tmp_path: Path, script: str) -> None:
    module = NEEDS.get(script)
    if module is not None and importlib.util.find_spec(module) is None:
        pytest.skip(f"needs {module}")
    if script == "update_existing.py":
        rows = [["Name", "Amount"], ["old", 1]]
        write_workbook(tmp_path / "template.xlsx", rows, sheet="Data")
        write_workbook(tmp_path / "template.xlsb", rows, sheet="Data")
    before = sorted(path.name for path in EXAMPLES.iterdir())
    args = [str(tmp_path)] if script in TAKES_DIRECTORY else EXTRA_ARGS.get(script, [])
    result = subprocess.run(
        [sys.executable, str(EXAMPLES / script), *args],
        cwd=tmp_path, capture_output=True, text=True, timeout=240,
    )
    assert result.returncode == 0, result.stderr[-2000:]
    assert sorted(path.name for path in EXAMPLES.iterdir()) == before, "the example wrote into examples/"
    written = [path for path in tmp_path.rglob("*") if path.suffix in {".xlsx", ".xlsb", ".xlsm"}]
    assert written, "the example wrote no workbook"
