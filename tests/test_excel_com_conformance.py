"""Excel COM conformance suite driven by the shared spec in comparison/com/cases.json.

Writer tests: xlspy writes every case in XLSX and XLSB; Microsoft Excel (COM) opens each file
and checks sheet names, visibility, AutoFilter, every cell value, the used range and number formats.

Reader tests: Excel writes the same cases; xlspy reads them back and must return the spec values.

Requires Windows, Microsoft Excel with COM automation, and PowerShell 7+ (pwsh).
"""

from __future__ import annotations

import datetime as dt
import json
import math
import shutil
import subprocess
from pathlib import Path

import pytest

from xlspy import ExcelReader, XlsbWriter, XlsxWriter

REPO_ROOT = Path(__file__).resolve().parents[2]
COM_DIR = REPO_ROOT / "comparison" / "com"
SPEC_PATH = COM_DIR / "cases.json"
EXCEL_EPOCH = dt.datetime(1899, 12, 30)


def _excel_available() -> bool:
    if not __import__("sys").platform.startswith("win"):
        return False
    try:
        import winreg

        winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Excel.Application\CLSID")
        return True
    except OSError:
        return False


pytestmark = [
    pytest.mark.skipif(not _excel_available(), reason="Microsoft Excel COM is not available"),
    pytest.mark.skipif(shutil.which("pwsh") is None, reason="PowerShell 7 (pwsh) is required"),
]


# --- spec loading ---------------------------------------------------------------------------

def _load_cases() -> list[dict]:
    return json.loads(SPEC_PATH.read_text(encoding="utf-8"))["cases"]


CASES = _load_cases()
FORMATS = ("xlsx", "xlsb")


def _expand_sheet(sheet: dict) -> list[list]:
    """Mirror of Expand-SheetRows in comparison/com/ExcelComSpec.psm1."""
    gen = sheet.get("generator")
    if gen is None:
        return [list(row) for row in sheet["rows"]]
    count = int(gen["rows"])
    if "columns" in gen:
        cols = int(gen["columns"])
        rows = [[f"C{j:02d}" for j in range(1, cols + 1)]]
        rows += [[float(i * 1000 + j) for j in range(1, cols + 1)] for i in range(1, count + 1)]
        return rows
    rows = [list(gen["header"])]
    rows += [[float(i), f"row-{i}", i * 0.25] for i in range(1, count + 1)]
    return rows


def _cell_value(cell):
    """Spec cell -> Python value as the library should read it / write it."""
    if isinstance(cell, dict):
        if "date" in cell:
            return dt.datetime.strptime(cell["date"], "%Y-%m-%dT%H:%M:%S")
        if "value" in cell:
            return cell["value"]
    return cell


def _cell_number_format(cell):
    if isinstance(cell, dict) and "numberFormat" in cell:
        return cell["numberFormat"]
    return None


# --- writer side ----------------------------------------------------------------------------

def _write_case(case: dict, path: Path, fmt: str) -> None:
    writer_cls = XlsxWriter if fmt == "xlsx" else XlsbWriter
    with writer_cls(str(path)) as workbook:
        for sheet in case["sheets"]:
            workbook.add_sheet(sheet["name"], hidden=bool(sheet.get("hidden", False)))
            rows = _expand_sheet(sheet)
            data = []
            for row in rows:
                out_row = []
                for cell in row:
                    value = _cell_value(cell)
                    fmt_code = _cell_number_format(cell)
                    out_row.append((value, fmt_code) if fmt_code else value)
                data.append(out_row)
            workbook.write_sheet(data, auto_filter=bool(sheet.get("autoFilter", False)))


# --- Excel-side helpers ---------------------------------------------------------------------

def _run_pwsh(args: list[str]) -> subprocess.CompletedProcess:
    return subprocess.run(
        ["pwsh", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", *args],
        capture_output=True,
        text=True,
        encoding="utf-8",
        timeout=900,
    )


def _verify(targets: list[dict], out_dir: Path) -> list[dict]:
    targets_path = out_dir / "verify-targets.json"
    results_path = out_dir / "verify-results.json"
    targets_path.write_text(json.dumps(targets), encoding="utf-8")
    proc = _run_pwsh(
        [
            str(COM_DIR / "Invoke-ExcelVerify.ps1"),
            "-SpecPath", str(SPEC_PATH),
            "-TargetsPath", str(targets_path),
            "-OutPath", str(results_path),
        ]
    )
    assert results_path.exists(), f"Excel verifier produced no results.\n{proc.stdout}\n{proc.stderr}"
    return json.loads(results_path.read_text(encoding="utf-8"))


@pytest.fixture(scope="session")
def com_session(tmp_path_factory):
    base = tmp_path_factory.mktemp("excel_com")
    written = base / "written"
    written.mkdir()

    targets = []
    for case in CASES:
        for fmt in FORMATS:
            path = written / f"{case['id']}.{fmt}"
            _write_case(case, path, fmt)
            targets.append({"file": str(path), "caseId": case["id"], "format": fmt})
    writer_results = _verify(targets, base)

    fixtures = base / "excel-fixtures"
    proc = _run_pwsh(
        [
            str(COM_DIR / "New-ExcelFixture.ps1"),
            "-SpecPath", str(SPEC_PATH),
            "-OutDir", str(fixtures),
        ]
    )
    assert proc.returncode == 0, f"New-ExcelFixture failed:\n{proc.stdout}\n{proc.stderr}"
    fixture_targets = json.loads((fixtures / "targets.json").read_text(encoding="utf-8"))

    return {"writer": writer_results, "fixtures": fixture_targets, "fixture_dir": fixtures}


def _results_for(results: list[dict], case_id: str, fmt: str) -> list[dict]:
    return [r for r in results if r["caseId"] == case_id and r["format"] == fmt]


def _failures(results: list[dict]) -> str:
    return "\n".join(
        f"{r['check']} [{r['sheet']}]: expected {r['expected']!r}, got {r['actual']!r} {r['message']}"
        for r in results
        if not r["pass"]
    )


# --- tests: Excel reads what xlspy wrote -----------------------------------------------------

def _is_structure_check(check: str) -> bool:
    return not check.startswith(("cells:", "numberFormat:"))


@pytest.mark.parametrize("fmt", FORMATS)
@pytest.mark.parametrize("case_id", [c["id"] for c in CASES])
def test_excel_opens_xlspy_structure(com_session, case_id, fmt):
    """Excel opens the file; sheet names, order, visibility, AutoFilter and used range match the spec."""
    results = [r for r in _results_for(com_session["writer"], case_id, fmt) if _is_structure_check(r["check"])]
    assert results, "no Excel checks were produced for this file"
    failures = _failures(results)
    assert not failures, failures


@pytest.mark.parametrize("fmt", FORMATS)
@pytest.mark.parametrize("case_id", [c["id"] for c in CASES])
def test_excel_reads_xlspy_cell_values(com_session, case_id, fmt):
    """Every cell value and number format written by xlspy reads back in Excel as specified."""
    results = [r for r in _results_for(com_session["writer"], case_id, fmt) if not _is_structure_check(r["check"])]
    if not results:
        pytest.skip("no cell checks for this case")
    failures = _failures(results)
    assert not failures, failures


# --- tests: xlspy reads what Excel wrote -----------------------------------------------------

def _normalise_for_compare(value):
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return float(value)
    return value


def _same_cell(expected, actual) -> bool:
    expected = _normalise_for_compare(expected)
    actual = _normalise_for_compare(actual)
    if expected is None:
        return actual in (None, "")
    if isinstance(expected, dt.datetime):
        if not isinstance(actual, dt.datetime):
            return False
        return math.isclose((expected - EXCEL_EPOCH).total_seconds(), (actual - EXCEL_EPOCH).total_seconds(), abs_tol=1.0)
    if isinstance(expected, float):
        return isinstance(actual, float) and math.isclose(expected, actual, rel_tol=1e-9, abs_tol=1e-12)
    return expected == actual


@pytest.mark.parametrize("fmt", FORMATS)
@pytest.mark.parametrize("case_id", [c["id"] for c in CASES])
def test_xlspy_reads_excel_authored_file(com_session, case_id, fmt):
    case = next(c for c in CASES if c["id"] == case_id)
    fixture = next(
        t for t in com_session["fixtures"] if t["caseId"] == case_id and t["format"] == fmt
    )
    with ExcelReader(fixture["file"]) as reader:
        names = reader.get_sheet_names()
        expected_names = [s["name"] for s in case["sheets"]]
        assert names == expected_names

        for sheet in case["sheets"]:
            expected_rows = [[_cell_value(c) for c in row] for row in _expand_sheet(sheet)]
            if not expected_rows:
                continue
            width = max(len(r) for r in expected_rows)
            expected_rows = [r + [None] * (width - len(r)) for r in expected_rows]
            actual_rows = [list(r) for r in reader.read_all(sheet["name"])]
            actual_rows = [r + [None] * (width - len(r)) for r in actual_rows]

            # Rows that are entirely empty may be trimmed by the reader; compare the populated prefix.
            while actual_rows and all(v in (None, "") for v in actual_rows[-1]) and len(actual_rows) > len(expected_rows):
                actual_rows.pop()
            assert len(actual_rows) >= len(expected_rows), f"{sheet['name']}: too few rows"
            for r, (exp_row, act_row) in enumerate(zip(expected_rows, actual_rows), start=1):
                for c, (exp, act) in enumerate(zip(exp_row, act_row), start=1):
                    assert _same_cell(exp, act), (
                        f"{sheet['name']} row {r} col {c}: expected {exp!r}, got {act!r}"
                    )


# --- tests: in-place updaters on an Excel-authored template -----------------------------------

@pytest.mark.parametrize("fmt", FORMATS)
def test_xlspy_updater_replaces_sheet_and_excel_verifies(com_session, tmp_path, fmt):
    from xlspy import XlsbUpdater, XlsxUpdater

    template = next(
        t for t in com_session["fixtures"] if t["caseId"] == "update_template" and t["format"] == fmt
    )
    target = tmp_path / f"updated.{fmt}"
    shutil.copyfile(template["file"], target)

    expected_case = next(c for c in CASES if c["id"] == "update_result")
    data_rows = _expand_sheet(expected_case["sheets"][0])
    headers, body = data_rows[0], [[_cell_value(c) for c in row] for row in data_rows[1:]]

    updater_cls = XlsxUpdater if fmt == "xlsx" else XlsbUpdater
    updater = updater_cls(str(target))
    try:
        updater.replace_sheet_data("Data", body, headers=headers)
        updater.save(str(target))
    finally:
        updater.close()

    results = _verify(
        [{"file": str(target), "caseId": "update_result", "format": fmt}],
        tmp_path,
    )
    failures = _failures(results)
    assert not failures, failures
