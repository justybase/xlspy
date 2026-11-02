"""Data validations: list sources, dynamic formulas, round trips."""
from __future__ import annotations

import pytest

from xlspy import DataValidation, XlsxWriter

openpyxl = pytest.importorskip("openpyxl")


@pytest.mark.parametrize(
    "formula",
    [
        "OFFSET($B$1,0,0,5,1)",
        "INDEX($B:$B,1):INDEX($B:$B,COUNTA($B:$B))",
        'INDIRECT("Lists!A1:A5")',
        "'Lists, 2024'!$A$1:$A$5",
        "$B$1:$B$5",
        "MyNamedRange",
    ],
)
def test_list_validation_accepts_ranges_and_dynamic_formulas(formula):
    rule = DataValidation("list", "A1:A5", formula1=formula)
    assert rule.formula1 == formula


@pytest.mark.parametrize("formula", ["Yes,No", "$B$1:$B$2,$C$1", "OFFSET($B$1,0,0),X"])
def test_list_validation_still_rejects_unquoted_literal_lists(formula):
    with pytest.raises(ValueError, match="double quotes"):
        DataValidation("list", "A1:A5", formula1=formula)


def test_dynamic_list_validation_round_trips(tmp_path):
    path = tmp_path / "dynamic.xlsx"
    with XlsxWriter(path) as writer:
        writer.write_sheet(
            [["pick", "options"], [None, "a"], [None, "b"]],
            data_validations=[DataValidation("list", "A2:A3", formula1="OFFSET($B$2,0,0,COUNTA($B:$B)-1,1)")],
        )

    sheet = openpyxl.load_workbook(path).active
    (rule,) = sheet.data_validations.dataValidation
    assert rule.type == "list"
    assert rule.formula1 == "OFFSET($B$2,0,0,COUNTA($B:$B)-1,1)"
