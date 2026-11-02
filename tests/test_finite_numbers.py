"""NaN and infinity are refused by the writers; finite numbers are written."""
from __future__ import annotations

import math
from decimal import Decimal

import pytest

from xlspy import XlsbWriter, XlsxWriter, load_workbook


@pytest.mark.parametrize("use_shared_strings", [True, False])
@pytest.mark.parametrize("value", [math.nan, math.inf, -math.inf, Decimal("NaN"), Decimal("Infinity")])
def test_xlsx_writer_rejects_non_finite_numbers(tmp_path, value, use_shared_strings):
    with pytest.raises(ValueError, match="finite"):
        with XlsxWriter(tmp_path / "bad.xlsx", use_shared_strings=use_shared_strings) as writer:
            writer.add_sheet("S")
            writer.write_sheet([[1, value, "x"]])


@pytest.mark.parametrize("use_shared_strings", [True, False])
@pytest.mark.parametrize("value", [math.nan, math.inf, -math.inf, Decimal("NaN")])
def test_xlsb_writer_rejects_non_finite_numbers(tmp_path, value, use_shared_strings):
    with pytest.raises(ValueError, match="finite"):
        with XlsbWriter(tmp_path / "bad.xlsb", use_shared_strings=use_shared_strings) as writer:
            writer.add_sheet("S")
            writer.write_sheet([[1, value, "x"]])


def test_writers_still_write_finite_numbers(tmp_path):
    path = tmp_path / "finite.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("S")
        writer.write_sheet([[1.5, -2.0, Decimal("3.25")]])
    assert [cell.value for cell in load_workbook(path).active[1]] == [1.5, -2, 3.25]
