"""Hyperlink cells in the native XLSX row encoder (C5b): the same package as the Python writer.

The native encoder writes the display text of a link with no style of its own,
in the built-in link style, and returns the link so the sheet's <hyperlinks>
and relationships come out as before. Links with a style of their own, in a
styled column, with a format, or with the link style not yet registered go to
the Python writer, which registers the style on the first link.
"""

import contextlib
import io
import zipfile
from unittest import mock

import pytest

import xlspy.xlsx_writer
from xlspy import Hyperlink, Style, XlsxWriter
from xlspy import _accel

needs_native = pytest.mark.skipif(not _accel.HAVE_NATIVE, reason="native extension not built")

ROWS = [
    ["Name", "Link"],
    ["plain", Hyperlink("https://example.com")],
    ["a", Hyperlink("https://example.com/a&b", text="Open <this>")],
    [Hyperlink("#T!A1", text="Back"), 5],
    [Hyperlink("mailto:x@example.com", tooltip="mail us")],
    ["text", " leading space", "tab\tx"],
    [Hyperlink("https://example.com/c", text="bell\x07")],
    [(Hyperlink("https://example.com/d"), Style(bold=True))],
    [Hyperlink("https://example.com/e", text="x" * 300)],
    [None, Hyperlink("https://example.com/f")],
    [Hyperlink("https://example.com/g"), Hyperlink("https://example.com/h", text="two")],
]


def _package(rows, shared, native=True):
    """The package parts written by XlsxWriter, through the native or the Python row encoder.

    docProps carries the save time, so it is left out of the comparison.
    """
    patch = (
        contextlib.nullcontext()
        if native
        else mock.patch.object(xlspy._xlsx_writer_sheet, "accel_encode_xlsx_row", lambda *a, **k: None)
    )
    with patch:
        buffer = io.BytesIO()
        with XlsxWriter(buffer, use_shared_strings=shared) as writer:
            writer.add_sheet("T")
            writer.write_sheet(rows)
    with zipfile.ZipFile(io.BytesIO(buffer.getvalue())) as archive:
        return {
            name: archive.read(name)
            for name in archive.namelist()
            if not name.startswith("docProps/")
        }


@needs_native
@pytest.mark.parametrize("shared", [True, False], ids=["shared", "inline"])
def test_hyperlink_package_matches_python_writer(shared):
    assert _package(ROWS, shared, native=True) == _package(ROWS, shared, native=False)


@needs_native
def test_link_rows_after_the_first_are_written_natively():
    # The parity test above would also pass if every link went to the Python
    # writer, so count the rows the native encoder wrote with link records.
    original = xlspy._xlsx_writer_sheet.accel_encode_xlsx_row
    native_link_rows = []

    def spy(*args, **kwargs):
        result = original(*args, **kwargs)
        if result is not None and len(result[3]):
            native_link_rows.append(len(result[3]))
        return result

    with mock.patch.object(xlspy._xlsx_writer_sheet, "accel_encode_xlsx_row", spy):
        _package(ROWS, True, native=True)
    assert len(native_link_rows) >= 3


@needs_native
@pytest.mark.parametrize("shared", [True, False], ids=["shared", "inline"])
def test_native_link_row_bytes_and_records(shared):
    link = Hyperlink("https://example.com", text="Open")
    result = _accel.encode_xlsx_row([link], {}, [], 0, 0, 0, shared, None, None, None, 5)
    assert result is not None, "a plain link cell must be encoded natively"
    xml, _uniq, _all, links = result
    if shared:
        assert xml == b'<row r="1"><c t="s" s="5"><v>0</v></c></row>'
    else:
        assert xml == b'<row r="1"><c t="inlineStr" s="5"><is><t>Open</t></is></c></row>'
    assert links == [(0, 0, link)]


@needs_native
@pytest.mark.parametrize(
    "row, kwargs",
    [
        ([Hyperlink("https://example.com")], {}),  # the link style is not registered yet
        ([Hyperlink("https://example.com")], {"link_xf": 0}),
        ([Hyperlink("https://example.com")], {"link_xf": 5, "style_ids": ((0, 7),)}),  # styled column
        ([(Hyperlink("https://example.com"), Style(bold=True))], {"link_xf": 5}),  # own style
        ([(Hyperlink("https://example.com"), "0.00")], {"link_xf": 5}),  # a number format
    ],
    ids=["unregistered", "zero-xf", "styled-column", "own-style", "format"],
)
def test_native_encoder_declines_links_the_python_writer_must_write(row, kwargs):
    assert _accel.encode_xlsx_row(row, {}, [], 0, 0, 0, False, **kwargs) is None


@needs_native
@pytest.mark.parametrize(
    "error",
    [ValueError],
    ids=["over-limit"],
)
def test_link_text_errors_match_python_writer(error):
    rows = [[Hyperlink("https://example.com", text="x" * 40000)]]
    for native in (True, False):
        with pytest.raises(error):
            _package(rows, True, native=native)
