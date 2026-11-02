"""Formula cells: cached values, names Excel needs a prefix for, value checks."""
from __future__ import annotations

from decimal import Decimal

import pytest
from helpers import part_text

from xlspy import Formula, XlsxWriter
from xlspy.formulas import normalize_formula


@pytest.mark.parametrize(
    "value",
    [float("inf"), float("-inf"), float("nan"), Decimal("Infinity"), Decimal("-Infinity"), Decimal("NaN")],
)
def test_formula_rejects_non_finite_cached_values(value):
    with pytest.raises(ValueError, match="finite"):
        Formula("1/0", value=value)


def test_formula_keeps_finite_cached_values(tmp_path):
    path = tmp_path / "finite.xlsx"
    with XlsxWriter(path) as writer:
        writer.write_sheet([[Formula("1/3", value=Decimal("0.25")), Formula("2*2", value=1e308)]])

    xml = part_text(path, "xl/worksheets/sheet1.xml")
    assert "<v>0.25</v>" in xml
    assert "inf" not in xml.lower()


def test_formula_bool_value_message_says_bool_is_refused():
    with pytest.raises(TypeError, match="not a bool"):
        Formula("A1", value=True)


def test_formula_names_get_the_prefix_excel_needs():
    assert normalize_formula("=COT(1)") == "_xlfn.COT(1)"
    assert normalize_formula("=SORT(A1:A3)") == "_xlfn._xlws.SORT(A1:A3)"
    assert normalize_formula("=FILTER(A1:A3,B1:B3>1)") == "_xlfn._xlws.FILTER(A1:A3,B1:B3>1)"
    assert normalize_formula("=SUM(1,2)") == "SUM(1,2)"
