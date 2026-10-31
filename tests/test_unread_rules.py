"""Rules the model cannot read survive a save, or the save is refused."""

import pytest

openpyxl = pytest.importorskip("openpyxl")

from openpyxl.worksheet.datavalidation import DataValidation as OpenpyxlValidation  # noqa: E402

from xlspy import DataValidation, Workbook, load_workbook  # noqa: E402
from xlspy.workbook import _xlsb_reader  # noqa: E402


@pytest.fixture
def xlsx(tmp_path):
    """A prompt-only ("any value") validation next to a list validation."""
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    prompt = OpenpyxlValidation(prompt="Type a name", showInputMessage=True)
    prompt.add("A1:A5")
    choices = OpenpyxlValidation(type="list", formula1='"a,b"')
    choices.add("B1:B5")
    sheet.add_data_validation(prompt)
    sheet.add_data_validation(choices)
    path = tmp_path / "prompt.xlsx"
    workbook.save(path)
    return path


def _validations(path):
    return [(item.type, str(item.sqref)) for item in openpyxl.load_workbook(path).active.data_validations.dataValidation]


def test_xlsx_appending_keeps_unread_validations(xlsx, tmp_path):
    out = tmp_path / "out.xlsx"
    workbook = load_workbook(xlsx)
    try:
        workbook.active.add_data_validation(DataValidation(type="whole", sqref="C1:C5", operator="greater_than", formula1="0"))
        workbook.save(out)
    finally:
        workbook.close()
    assert (None, "A1:A5") in _validations(out)
    assert len(_validations(out)) == 3


def test_xlsx_rewrite_that_would_drop_unread_validations_is_refused(xlsx, tmp_path):
    workbook = load_workbook(xlsx)
    try:
        workbook.active.data_validations.clear()
        with pytest.raises(NotImplementedError, match="1 data validations that xlspy cannot read"):
            workbook.save(tmp_path / "out.xlsx")
    finally:
        workbook.close()


def test_xlsb_rewrite_that_would_drop_unread_validations_is_refused(tmp_path, monkeypatch):
    workbook = Workbook()
    sheet = workbook.active
    sheet.add_data_validation(DataValidation(type="list", sqref="A1:A5", formula1='"a,b"'))
    sheet.add_data_validation(DataValidation(type="list", sqref="B1:B5", formula1='"c,d"'))
    path = tmp_path / "in.xlsb"
    workbook.save(path)
    # Stand-in for a rule the model has no type for (Excel's "any value").
    calls = []
    read = _xlsb_reader.dval_rule

    def first_unreadable(payload, context):
        calls.append(payload)
        return None if len(calls) == 1 else read(payload, context)

    monkeypatch.setattr(_xlsb_reader, "dval_rule", first_unreadable)
    loaded = load_workbook(path)
    try:
        assert len(loaded.active.data_validations) == 1
        loaded.save(tmp_path / "untouched.xlsb")
        loaded.active.data_validations.clear()
        with pytest.raises(NotImplementedError, match="data validations that xlspy cannot read"):
            loaded.save(tmp_path / "out.xlsb")
    finally:
        loaded.close()
