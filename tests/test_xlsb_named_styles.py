"""Named styles can be added to a loaded XLSB workbook."""

import pytest

from xlspy import Font, NamedStyle, PatternFill, Workbook, load_workbook


@pytest.fixture
def source(tmp_path):
    workbook = Workbook()
    workbook.active["A1"] = "hello"
    path = tmp_path / "source.xlsb"
    workbook.save(path)
    workbook.close()
    return path


def test_named_style_added_and_applied(source, tmp_path):
    out = tmp_path / "out.xlsb"
    workbook = load_workbook(source)
    try:
        workbook.add_named_style(NamedStyle(
            "Highlight", font=Font(bold=True, color="FFFF0000"),
            fill=PatternFill("solid", start_color="FFFFFF00", end_color="FFFFFF00"), number_format="0.00",
        ))
        workbook.active["A1"].style = "Highlight"
        workbook.save(out)
    finally:
        workbook.close()
    reloaded = load_workbook(out)
    try:
        assert "Highlight" in [getattr(s, "name", s) for s in reloaded.named_styles]
        cell = reloaded.active["A1"]
        assert cell.font.bold and cell.number_format == "0.00"
        assert str(cell.fill.start_color.rgb).endswith("FFFF00")
    finally:
        reloaded.close()


def test_duplicate_named_style_is_refused(source):
    workbook = load_workbook(source)
    try:
        workbook.add_named_style(NamedStyle("Once", font=Font(italic=True)))
        with pytest.raises(ValueError):
            workbook.add_named_style(NamedStyle("Once", font=Font(italic=True)))
    finally:
        workbook.close()
