"""Formula cells in the native XLSX row encoder (C5): same bytes as the Python writer.

The native encoder writes a Formula cell itself, except for the cases it cannot
reproduce exactly (array formulas, values that need an escape, floats beyond
2**53, Decimal values, bool, text over the cell limit), which go to the Python
writer. Every case here must produce the same sheet XML, or the same error,
through both paths.
"""

import contextlib
import io
import zipfile
from decimal import Decimal
from unittest import mock

import pytest

import xlspy.xlsx_writer
from xlspy import Formula, XlsxWriter
from xlspy import _accel

needs_native = pytest.mark.skipif(not _accel.HAVE_NATIVE, reason="native extension not built")

ASTRAL = "\U0001D11E"  # outside the BMP: two UTF-16 units

CASES = [
    Formula("A1+B1"),
    Formula("SUM(A1:A3)", value=6),
    Formula('A1&"<x>"', value='a & <b> "c"'),
    Formula("A1*2", value=0.5),
    Formula("A1*2", value=-0.0),
    Formula("A1*2", value=1.0),
    Formula("A1*2", value=1e-07),
    Formula("A1*2", value=0.1 + 0.2),
    Formula("A1*2", value=2**60),
    Formula("A1*2", value=-5),
    Formula("A1*2", value=2**70),
    Formula("A1*2", value=1e20),
    Formula("A1*2", value=Decimal("1.50")),
    Formula("A1*2", value="_x0041_"),
    Formula("A1*2", value="tab\tand\nnewline"),
    Formula("A1*2", value="bell\x07"),
    Formula("A1*2", value="zolw " + ASTRAL),
    Formula("A1", value=3, array_ref="A1:A1"),
    Formula('IF(A1>0,"x","y")', value="x"),
    Formula("A1"),
    (Formula("A1+1"), "0.00"),
    (Formula("A1+1", value=2), "0.00"),
]


def _sheet_xml(rows, shared, native=True):
    """The worksheet part written by XlsxWriter, through the native or the Python row encoder."""
    patch = (
        contextlib.nullcontext()
        if native
        else mock.patch.object(xlspy.xlsx_writer, "accel_encode_xlsx_row", lambda *a, **k: None)
    )
    with patch:
        buffer = io.BytesIO()
        with XlsxWriter(buffer, useSharedStrings=shared) as writer:
            writer.add_sheet("T")
            writer.write_sheet(rows)
    with zipfile.ZipFile(io.BytesIO(buffer.getvalue())) as archive:
        return archive.read("xl/worksheets/sheet1.xml")


@needs_native
@pytest.mark.parametrize("shared", [True, False], ids=["shared", "inline"])
def test_formula_sheet_matches_python_writer(shared):
    rows = [["Header"]] + [[case] for case in CASES]
    rows.append([CASES[1], 1, "text", CASES[2], CASES[21]])
    assert _sheet_xml(rows, shared, native=True) == _sheet_xml(rows, shared, native=False)


@needs_native
@pytest.mark.parametrize(
    "formula, expected",
    [
        (Formula("A1+1"), b'<row r="1"><c><f>A1+1</f></c></row>'),
        (Formula("A1+1", value=2), b'<row r="1"><c><f>A1+1</f><v>2</v></c></row>'),
        (Formula("A1+1", value=2.5), b'<row r="1"><c><f>A1+1</f><v>2.5</v></c></row>'),
        (Formula("A1+1", value=3.0), b'<row r="1"><c><f>A1+1</f><v>3</v></c></row>'),
        (Formula("A1&B1", value="x&y"),
         b'<row r="1"><c t="str"><f>A1&amp;B1</f><v>x&amp;y</v></c></row>'),
        (Formula('A1&"<"', value="a"),
         b'<row r="1"><c t="str"><f>A1&amp;"&lt;"</f><v>a</v></c></row>'),
        (Formula('A1&"<"'),
         b'<row r="1"><c><f>A1&amp;"&lt;"</f></c></row>'),
    ],
)
def test_native_formula_row_bytes(formula, expected):
    result = _accel.encode_xlsx_row([formula], {}, [], 0, 0, 0, False)
    assert result is not None, "a plain formula cell must be encoded natively"
    assert result[0] == expected


@needs_native
@pytest.mark.parametrize(
    "formula",
    [
        Formula("A1*2", value=Decimal("1.5")),
        Formula("A1*2", value=2**70),
        Formula("A1*2", value=1e20),
        Formula("A1*2", value="_x0041_"),
        Formula("A1*2", value="bell\x07"),
        Formula("A1", value=3, array_ref="A1:A1"),
    ],
    ids=["decimal", "bigint", "huge-float", "x-escape", "control", "array"],
)
def test_native_encoder_declines_what_the_python_writer_must_write(formula):
    assert _accel.encode_xlsx_row([formula], {}, [], 0, 0, 0, False) is None


@needs_native
@pytest.mark.parametrize(
    "value, error",
    [("x" * 40000, ValueError), ("lone \ud800", UnicodeEncodeError)],
    ids=["over-limit", "lone-surrogate"],
)
def test_formula_text_errors_match_python_writer(value, error):
    rows = [[Formula("A1", value=value)]]
    for native in (True, False):
        with pytest.raises(error):
            _sheet_xml(rows, True, native=native)
