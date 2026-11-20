"""The writer and the loaded style table render style fragments identically.

They used to carry two copies of the color and border encoders; the copies had
already drifted on ``tint`` formatting. These tests pin the shared output.
"""

from __future__ import annotations

import pytest

from xlspy import chart_parts
from xlspy._style_xml import border_xml, color_attrs
from xlspy.styles import Border, Side
from xlspy.styles._keys import _border_key
from xlspy.workbook._styles import XlsxStyleTable
from xlspy.xlsx_writer import XlsxWriter


@pytest.mark.parametrize(
    "key",
    [
        ("rgb", "FF112233"),
        ("indexed", 4),
        ("auto",),
        ("theme", 2, None),
        ("theme", 2, 0.5),
        ("theme", 2, 1.0),
    ],
)
def test_writer_and_loaded_table_share_color_attributes(key):
    assert XlsxWriter._color_attrs(key) == XlsxStyleTable._color_attrs(key)
    assert color_attrs(key) == XlsxStyleTable._color_attrs(key)


def test_writer_and_loaded_table_share_border_xml():
    border = Border(
        left=Side("thin", "FF0000"),
        right=Side(color="00FF00"),
        top=Side("dotted"),
        diagonal=Side("double", "0000FF"),
        vertical=Side("medium", None),
        horizontal=Side(style=None, color=None),
        diagonal_up=True,
        diagonal_down=True,
        outline=False,
    )
    key = _border_key(border)
    writer = XlsxWriter._border_xml(None, key)
    loaded = XlsxStyleTable.empty()._border_xml_for(key)
    assert writer == loaded == border_xml(key)
    assert 'diagonalUp="1"' in writer
    assert 'diagonalDown="1"' in writer
    assert 'outline="0"' in writer
    assert '<left style="thin"><color rgb="FFFF0000"/></left>' in writer
    assert '<right><color rgb="FF00FF00"/></right>' in writer


def test_one_column_letter_table():
    letters = chart_parts.column_letters()
    assert XlsxWriter._generate_column_letters(None) == letters
    assert len(letters) == 16384
    assert letters[0] == "A" and letters[25] == "Z"
    assert letters[26] == "AA" and letters[-1] == "XFD"
