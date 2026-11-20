"""The gap manifest and the documented parity matrix name the same open gaps."""

from __future__ import annotations

import re
from pathlib import Path

from ._util import GAPS

PARITY_DOC = Path(__file__).resolve().parents[2] / "docs" / "openpyxl-parity.md"
_ROW = re.compile(r"^\|\s*([A-Z]\d+)\s*\|(.*)$")
_HEADER = "| ID |"


def documented_open_gaps() -> set[str]:
    """The IDs the parity document lists without a **done** marker."""
    gaps: set[str] = set()
    in_matrix = False
    for line in PARITY_DOC.read_text(encoding="utf-8").splitlines():
        if line.startswith(_HEADER):
            in_matrix = True
            continue
        match = _ROW.match(line)
        if match is None or not in_matrix:
            continue
        gap, rest = match.groups()
        if not gap.startswith("L") and "**done**" not in rest:
            gaps.add(gap)
    return gaps


def test_manifest_lists_every_documented_open_gap():
    assert set(GAPS) == documented_open_gaps(), (
        "tests/openpyxl_compat/expected_gaps.toml and docs/openpyxl-parity.md "
        "disagree; update both"
    )


def test_every_manifest_gap_has_an_area_and_reason():
    for gap, entry in GAPS.items():
        assert entry["area"].strip(), f"gap {gap} has no area"
        assert entry["reason"].strip(), f"gap {gap} has no reason"
