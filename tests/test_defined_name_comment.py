"""A defined name keeps its comment in both formats, through the writers, the updaters and the model."""

import pytest

from xlspy import DefinedName, Workbook, XlsbUpdater, XlsbWriter, XlsxUpdater, XlsxWriter, load_workbook
from xlspy.reader import ExcelReader

FORMATS = pytest.mark.parametrize("suffix", ["xlsx", "xlsb"])
TEXT = 'what it is: <a "quoted" & odd> name'


def _names(path):
    with ExcelReader(path) as reader:
        return {item.name: (item.formula, item.comment, item.hidden) for item in reader.get_defined_names()}


@FORMATS
def test_the_writers_keep_the_comment(tmp_path, suffix):
    path = tmp_path / f"w.{suffix}"
    writer_class = XlsxWriter if suffix == "xlsx" else XlsbWriter
    with writer_class(path) as writer:
        writer.set_defined_name(DefinedName("Rate", "Data!$A$1", comment=TEXT))
        writer.set_defined_name(DefinedName("Plain", "Data!$A$2"))
        writer.set_defined_name(DefinedName("Hid", "Data!$A$3", hidden=True, comment="shy"))
        writer.add_sheet("Data")
        writer.write_sheet([[1], [2], [3]])
    names = _names(path)
    assert names["Rate"][1] == TEXT
    assert names["Plain"][1] is None
    assert names["Hid"][1:] == ("shy", True)


@FORMATS
def test_the_updaters_keep_the_comment(tmp_path, suffix):
    source = tmp_path / f"s.{suffix}"
    workbook = Workbook()
    workbook.active.title = "Data"
    workbook.active.append([1])
    workbook.save(source)
    out = tmp_path / f"o.{suffix}"
    updater_class = XlsxUpdater if suffix == "xlsx" else XlsbUpdater
    updater = updater_class(source)
    updater.set_defined_name(DefinedName("Rate", "Data!$A$1", comment=TEXT))
    updater.save(out)
    updater.close()
    assert _names(out)["Rate"][1] == TEXT
    # Replacing a name with one without a comment removes the comment.
    updater = updater_class(out)
    updater.set_defined_name(DefinedName("Rate", "Data!$A$1"))
    updater.save(tmp_path / f"p.{suffix}")
    updater.close()
    assert _names(tmp_path / f"p.{suffix}")["Rate"][1] is None


@FORMATS
def test_the_model_saves_a_changed_comment_of_a_loaded_workbook(tmp_path, suffix):
    source = tmp_path / f"s.{suffix}"
    workbook = Workbook()
    workbook.active.title = "Data"
    workbook.active.append([1])
    workbook.defined_names.add(DefinedName("Rate", "Data!$A$1", comment="old"))
    workbook.save(source)
    assert _names(source)["Rate"][1] == "old"

    loaded = load_workbook(source)
    assert loaded.defined_names["Rate"].comment == "old"
    loaded.defined_names.remove("Rate")
    loaded.defined_names.add(DefinedName("Rate", "Data!$A$1", comment="new"))
    out = tmp_path / f"o.{suffix}"
    loaded.save(out)
    loaded.close()
    assert _names(out)["Rate"][1] == "new"


@FORMATS
def test_a_name_without_a_comment_has_none(tmp_path, suffix):
    source = tmp_path / f"s.{suffix}"
    workbook = Workbook()
    workbook.active.append([1])
    workbook.defined_names.add(DefinedName("Rate", "Sheet!$A$1"))
    workbook.save(source)
    loaded = load_workbook(source)
    try:
        assert loaded.defined_names["Rate"].comment is None
    finally:
        loaded.close()
