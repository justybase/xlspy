"""A named style whose fill is a colour string must be addable to an XLSB style table.

Before the fix, XlsbStyleTable.add_named_style passed the string straight to the
PatternFill/GradientFill key function and raised AttributeError ('str' object has
no attribute 'stop'). A colour string is a solid fill of that colour.
"""

from xlspy.styles import NamedStyle, fill_key_for
from xlspy._xlsb_styles import XlsbStyleTable


def test_named_style_with_colour_string_fill_is_added():
    table = XlsbStyleTable.from_bin(None)
    index = table.add_named_style(NamedStyle(name="Tinted", fill="00FF00"))
    assert isinstance(index, int)


def test_colour_string_fill_is_a_solid_rgb_fill():
    assert fill_key_for("00FF00") == ("pattern", "solid", ("rgb", "FF00FF00"), None)


def test_no_fill_has_no_key():
    assert fill_key_for(None) is None
