"""The XLSB writer keeps typical workbooks in the native row encoder.

``count_paths`` (native_paths.py) counts the rows Python had to write cell by cell
and the per-cell calls into Python. The scenarios that are expected to be fast must
stay at zero; a first row for each new style variant is the one allowed exception
(Python registers the style, the following rows are native). The scenarios with a
known gap are marked ``xfail(strict=True)``: when a gap is closed the marker has to go.
"""

import pytest
from native_paths import SCENARIOS, Scenario, write

from xlspy import _accel

pytestmark = pytest.mark.skipif(not _accel.HAVE_NATIVE, reason="native extension not built")

ROWS = 600


def _id(scenario: Scenario) -> str:
    return scenario.name


@pytest.mark.parametrize("scenario", [s for s in SCENARIOS if s.gap is None], ids=_id)
def test_typical_workbooks_are_written_natively(tmp_path, scenario):
    counts = write(scenario, ROWS, tmp_path / "a.xlsb")
    assert counts["python_rows"] <= scenario.max_python_rows, dict(counts)
    assert counts["native"] >= ROWS, dict(counts)
    if scenario.max_scalar_hook is not None:
        assert counts["scalar_hook"] <= scenario.max_scalar_hook, dict(counts)


@pytest.mark.parametrize("scenario", [s for s in SCENARIOS if s.gap is not None], ids=_id)
@pytest.mark.xfail(strict=True, reason="a known gap: see Scenario.gap")
def test_known_gaps_are_still_gaps(tmp_path, scenario):
    """The wanted state; fails while the gap is open. Remove the marker when it closes."""
    counts = write(scenario, ROWS, tmp_path / "a.xlsb")
    assert counts["python_rows"] <= 1, dict(counts)
    assert counts["scalar_hook"] == 0, dict(counts)


def test_counters_see_the_python_rows(tmp_path):
    """The counters are not blind: rich text with inline strings does go through Python."""
    scenario = next(s for s in SCENARIOS if s.name == "rich text, inline strings")
    counts = write(scenario, 50, tmp_path / "a.xlsb")
    assert counts["python_rows"] >= 50
