"""A conditional-format edit in a loaded XLSB keeps the Excel 2010 extension of untouched rules."""

import zipfile
from pathlib import Path


from xlspy import FormulaRule, load_workbook
from xlspy.biff_utils import iter_record_payloads

FIXTURE = Path(__file__).parent / "data" / "x14_databars.xlsb"  # two data bars, made by Excel
BEGIN_CF14, EXT_COLOR, BEGIN_RULE = 0x416, 0x41F, 0x1CF


def _ids(path):
    with zipfile.ZipFile(path) as archive:
        return [record for record, _ in iter_record_payloads(archive.read("xl/worksheets/sheet1.bin"))]


def _edit(tmp_path, edit):
    out = tmp_path / "out.xlsb"
    workbook = load_workbook(FIXTURE)
    try:
        edit(workbook.active)
        workbook.save(out)
    finally:
        workbook.close()
    return out


def test_fixture_has_the_extension():
    ids = _ids(FIXTURE)
    assert ids.count(BEGIN_CF14) == 2 and ids.count(EXT_COLOR) == 3


def test_adding_a_rule_keeps_the_extension(tmp_path):
    out = _edit(tmp_path, lambda s: s.conditional_formatting.add("E1", FormulaRule("E1", ["TRUE"])))
    ids = _ids(out)
    assert ids.count(BEGIN_CF14) == 2 and ids.count(EXT_COLOR) == 3
    assert ids.count(BEGIN_RULE) == 3
    reloaded = load_workbook(out)
    try:
        assert len(reloaded.active.conditional_formatting.all_rules()) == 3
    finally:
        reloaded.close()


def test_removed_rule_takes_only_its_extension(tmp_path):
    def edit(sheet):
        sheet.conditional_formatting.remove(sheet.conditional_formatting.all_rules()[0])

    out = _edit(tmp_path, edit)
    ids = _ids(out)
    assert ids.count(BEGIN_RULE) == 1 and ids.count(BEGIN_CF14) == 1


def test_changed_rule_loses_only_its_extension(tmp_path):
    from dataclasses import replace

    def edit(sheet):
        item = sheet.conditional_formatting[0]
        item.rules[0] = replace(item.rules[0], min_length=20)

    ids = _ids(_edit(tmp_path, edit))
    assert ids.count(BEGIN_RULE) == 2 and ids.count(BEGIN_CF14) == 1


def test_extension_stays_after_the_base_rules_and_before_the_end(tmp_path):
    out = _edit(tmp_path, lambda s: s.conditional_formatting.add("E1", FormulaRule("E1", ["TRUE"])))
    ids = _ids(out)
    last_rule_end = max(i for i, record in enumerate(ids) if record == 0x1CE)
    assert ids.index(BEGIN_CF14) > last_rule_end
    assert ids[-1] == 0x82 and ids.index(0x470) < len(ids) - 1


def test_conditional_format_on_a_new_sheet_of_a_loaded_xlsb(tmp_path):
    from xlspy import Workbook

    first = tmp_path / "first.xlsb"
    book = Workbook()
    book.active["A1"] = 1
    book.save(first)
    book.close()
    loaded = load_workbook(first)
    sheet = loaded.create_sheet("Extra")
    sheet["A1"] = 5
    sheet.conditional_formatting.add("A1", FormulaRule("A1", ["A1>1"]))
    target = tmp_path / "second.xlsb"
    loaded.save(target)
    loaded.close()
    again = load_workbook(target)
    assert len(list(again["Extra"].conditional_formatting)) == 1
    again.close()
