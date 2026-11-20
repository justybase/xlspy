"""Shared helpers for the openpyxl compatibility harness."""

from __future__ import annotations

import datetime as dt
import math
import tomllib
import zipfile
from pathlib import Path

import pytest

IGNORED_PARTS = {"docProps/core.xml"}

# The manifest is the harness's source of truth: a test xfails only for a gap
# listed here. The documentation table is kept in sync by
# tests/openpyxl_compat/test_gap_manifest.py, not read by this harness.
_GAPS_FILE = Path(__file__).with_name("expected_gaps.toml")


def _open_gaps() -> dict[str, dict]:
    data = tomllib.loads(_GAPS_FILE.read_text(encoding="utf-8"))
    gaps = {}
    for entry in data["gap"]:
        gap = entry["id"]
        if gap in gaps:
            raise RuntimeError(f"duplicate gap {gap} in {_GAPS_FILE.name}")
        if not entry.get("reason"):
            raise RuntimeError(f"gap {gap} has no reason in {_GAPS_FILE.name}")
        gaps[gap] = entry
    return gaps


GAPS = _open_gaps()
OPEN_GAPS = frozenset(GAPS)


def gap_xfail(message: str) -> None:
    """Expected failure for an open gap; any other missing feature fails the test."""
    gap = message.split(":", 1)[0]
    entry = GAPS.get(gap)
    if entry is not None:
        pytest.xfail(f"{message} ({entry['area']})")
    pytest.fail(
        f"{message} (gap {gap!r} is not listed in tests/openpyxl_compat/expected_gaps.toml)",
        pytrace=False,
    )


def openpyxl_load(path, **kwargs):
    openpyxl = pytest.importorskip("openpyxl")
    return openpyxl.load_workbook(str(path), **kwargs)


def xlspy_load(path, **kwargs):
    from xlspy import load_workbook

    return load_workbook(str(path), **kwargs)


def close_quietly(workbook) -> None:
    try:
        workbook.close()
    except Exception:
        pass


def zip_parts(path) -> dict[str, bytes]:
    with zipfile.ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def diff_parts(before: dict[str, bytes], after: dict[str, bytes], ignored=IGNORED_PARTS):
    problems = []
    for name in sorted(set(before) | set(after)):
        if name in ignored:
            continue
        if name not in before:
            problems.append(f"+{name}")
        elif name not in after:
            problems.append(f"-{name}")
        elif before[name] != after[name]:
            problems.append(f"~{name}")
    return problems


def normalize(value):
    if isinstance(value, bool) or value is None:
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    if isinstance(value, (dt.datetime, dt.date, dt.time)):
        return value.isoformat()
    if isinstance(value, dt.timedelta):
        return value.total_seconds()
    return value


def require(obj, attr: str, gap: str):
    """Return ``obj.attr`` or xfail with the gap ID when it is missing."""
    if not hasattr(obj, attr):
        gap_xfail(f"{gap}: {type(obj).__name__}.{attr} is missing")
    return getattr(obj, attr)


def require_path(obj, path: str, gap: str):
    current = obj
    for part in path.split("."):
        current = require(current, part, gap)
    return current


def fail_mismatch(gap: str, label: str, actual, expected):
    """Fail with a useful diff; used when the API exists but the value is wrong."""
    if normalize(actual) != normalize(expected):
        pytest.fail(
            f"{gap}: {label}: {actual!r} != {expected!r}",
            pytrace=False,
        )


def xfail_mismatch(gap: str, label: str, actual, expected):
    if normalize(actual) != normalize(expected):
        gap_xfail(f"{gap}: {label}: {actual!r} != {expected!r}")


def assert_close(actual, expected, *, rel_tol=1e-9, abs_tol=1e-9):
    if isinstance(actual, float) and isinstance(expected, float):
        assert math.isclose(actual, expected, rel_tol=rel_tol, abs_tol=abs_tol), (
            f"{actual!r} != {expected!r}"
        )
    else:
        assert normalize(actual) == normalize(expected), f"{actual!r} != {expected!r}"


def cell_grid(worksheet, values_only: bool = True):
    return [
        list(row) if values_only else [cell.value for cell in row]
        for row in worksheet.iter_rows(values_only=values_only)
    ]


def golden_dir() -> Path:
    return Path(__file__).resolve().parents[1] / "data" / "openpyxl_parity"


def golden_path(name: str) -> Path:
    return golden_dir() / name
