"""Differential tests: the regex and C fast paths for data_only reads must match
the ElementTree path."""

import datetime
import io

import pytest

from xlspy import ExcelReader, XlsxWriter
from xlspy import _accel, _xlsx_read
from test_comments_and_editing import replace_parts

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
STRICT_NS = "http://purl.oclc.org/ooxml/spreadsheetml/main"
SHEET = "xl/worksheets/sheet1.xml"


def _sheet(body, root=f'<worksheet xmlns="{MAIN_NS}">'):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + root
            + "<sheetData>" + body + "</sheetData></worksheet>").encode("utf-8")


def _base_file(tmp_path):
    """A workbook with a shared string "text" at index 0 and a date style in the styles part."""
    path = tmp_path / "base.xlsx"
    with XlsxWriter(path) as writer:
        writer.write_sheet([["text", 1, 2.5, True, datetime.datetime(2024, 1, 5)]])
    return path


def _both_paths(path, monkeypatch):
    with ExcelReader(str(path)) as reader:
        fast = list(_xlsx_read.iter_indexed_rows(reader, "Sheet1", data_only=True))
    monkeypatch.setattr(_xlsx_read, "_probe_fast", lambda stream: None)
    with ExcelReader(str(path)) as reader:
        slow = list(_xlsx_read.iter_indexed_rows(reader, "Sheet1", data_only=True))
    monkeypatch.undo()
    return fast, slow


def _fast_path_applies(path):
    with ExcelReader(str(path)) as reader:
        with reader._archive.open(SHEET) as stream:
            return _xlsx_read._probe_fast(stream) is not None


def _all_paths(path, monkeypatch):
    """The same sheet read by C, by the Python fast path, and by ElementTree."""
    with ExcelReader(str(path)) as reader:
        from_c = list(_xlsx_read.iter_indexed_rows(reader, "Sheet1", data_only=True))
    with monkeypatch.context() as patch:
        patch.setattr(_xlsx_read._accel, "read_xlsx_worksheet", lambda *args: None)
        with ExcelReader(str(path)) as reader:
            from_python = list(_xlsx_read.iter_indexed_rows(reader, "Sheet1", data_only=True))
    with monkeypatch.context() as patch:
        patch.setattr(_xlsx_read, "_probe_fast", lambda stream: None)
        with ExcelReader(str(path)) as reader:
            from_tree = list(_xlsx_read.iter_indexed_rows(reader, "Sheet1", data_only=True))
    return from_c, from_python, from_tree


def _same_typed_rows(left, right):
    assert len(left) == len(right)
    for (left_number, left_cells), (right_number, right_cells) in zip(left, right):
        assert left_number == right_number
        assert len(left_cells) == len(right_cells)
        for (left_col, left_value), (right_col, right_value) in zip(left_cells, right_cells):
            assert left_col == right_col
            assert type(left_value) is type(right_value)
            assert left_value == right_value


def _sheet_bytes(path):
    with ExcelReader(str(path)) as reader:
        with reader._archive.open(SHEET) as stream:
            return stream.read()


def _c_items(path):
    """Raw C items, before the caller resolves rows that need ElementTree."""
    with ExcelReader(str(path)) as reader:
        with reader._archive.open(SHEET) as stream:
            probe = _xlsx_read._probe_fast(stream)
            assert probe is not None
            rows = _accel.read_xlsx_worksheet(
                stream, probe[4], reader._shared_string_array,
                _xlsx_read._date_style_flags(reader), reader._date_1904)
            assert rows is not None
            return list(rows)


class _Trickle:
    """A stream whose read() returns at most ``size`` bytes per call."""

    def __init__(self, data, size):
        self._data = data
        self._size = size
        self._pos = 0

    def read(self, size):
        end = min(len(self._data), self._pos + min(size, self._size))
        chunk = self._data[self._pos:end]
        self._pos = end
        return chunk


SUPPORTED_BODIES = {
    "shared_strings_attribute_order": '<row r="1"><c t="s" s="0" r="A1"><v>0</v></c></row>',
    "style_then_type": '<row r="1"><c r="A1" s="1"><v>45292</v></c><c r="B1" s="0" t="n"><v>7</v></c></row>',
    "number_formats": '<row r="1"><c r="A1"><v>1E+3</v></c><c r="B1"><v>-0.5</v></c>'
                      '<c r="C1"><v>12345678901234567890</v></c><c r="D1"><v>1.5e-7</v></c></row>',
    "booleans_and_errors": '<row r="1"><c r="A1" t="b"><v>1</v></c><c r="B1" t="b"><v>0</v></c>'
                           '<c r="C1" t="e"><v>#DIV/0!</v></c></row>',
    "formula_with_cached_value": '<row r="1"><c r="A1"><f>1+1</f><v>2</v></c>'
                                 '<c r="B1" t="str"><f>"a"</f><v>a</v></c></row>',
    "shared_formula_cells": '<row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">B1*2</f><v>4</v></c></row>'
                            '<row r="2"><c r="A2"><f t="shared" si="0"/><v>6</v></c></row>',
    "self_closing_cells_and_rows": '<row r="1"><c r="A1"/><c r="B1" s="1"/><c r="C1"><v>3</v></c></row>'
                                   '<row r="3"/><row r="5"><c r="A5"><v>9</v></c></row>',
    "missing_refs_and_row_numbers": '<row><c><v>1</v></c><c><v>2</v></c></row>'
                                    '<row><c r="D2"><v>3</v></c><c><v>4</v></c></row>',
    "gap_in_columns": '<row r="1"><c r="A1"><v>1</v></c><c r="E1"><v>5</v></c><c><v>6</v></c></row>',
    "empty_value_element": '<row r="1"><c r="A1" t="n"></c><c r="B1"><v></v></c><c r="C1" t="s"><v></v></c></row>',
    "simple_inline_string": '<row r="1"><c r="A1" t="inlineStr"><is><t>plain</t></is></c>'
                            '<c r="B1" t="inlineStr"><is><t xml:space="preserve"> spaced </t></is></c></row>',
    "xstring_escape_in_inline": '<row r="1"><c r="A1" t="inlineStr"><is><t>x_x000D_y</t></is></c></row>',
    "rich_inline_string_falls_back": '<row r="1"><c r="A1" t="inlineStr"><is><r><t>a</t></r><r><t>b</t></r></is></c></row>',
    "entities_in_inline_fall_back": '<row r="1"><c r="A1" t="inlineStr"><is><t>a &amp; b</t></is></c></row>',
    "entities_in_str_fall_back": '<row r="1"><c r="A1" t="str"><v>x &lt; y</v></c></row>',
    "carriage_return_in_str_falls_back": '<row r="1"><c r="A1" t="str"><v>a\r\nb</v></c></row>',
    "unknown_child_falls_back": '<row r="1"><c r="A1"><v>1</v><extLst/></c><c r="B1"><v>2</v></c></row>',
    "lowercase_reference_falls_back": '<row r="1"><c r="a1"><v>1</v></c></row>',
    "whitespace_before_cell_end": '<row r="1"><c /><c r="B1" ><v>1</v></c>'
                                  '<c r="C1" /><c r="D1" ><v>2</v></c></row>',
    "non_ascii_attribute_value": '<row note="é" r="1"><c r="A1"><v>1</v></c></row>',
    "prefixed_cell_falls_back": '<row r="1"><c r="A1"><v>1</v></c><x:c r="B1" xmlns:x="' + MAIN_NS + '"><x:v>2</x:v></x:c></row>',
}

PREFIXED_SHEETS = {
    "prefixed_worksheet": (
        '<?xml version="1.0" encoding="UTF-8"?>'
        f'<x:worksheet xmlns:x="{MAIN_NS}"><x:sheetData>'
        '<x:row r="1"><x:c r="A1"><x:v>1</x:v></x:c></x:row>'
        '</x:sheetData></x:worksheet>'
    ).encode("utf-8"),
}


@pytest.mark.parametrize("name", sorted(SUPPORTED_BODIES))
def test_fast_path_matches_element_tree(tmp_path, monkeypatch, name):
    path = _base_file(tmp_path)
    replace_parts(path, {SHEET: _sheet(SUPPORTED_BODIES[name])})
    fast, slow = _both_paths(path, monkeypatch)
    assert fast == slow


@pytest.mark.parametrize("name", sorted(PREFIXED_SHEETS))
def test_prefixed_sheets_use_element_tree_and_still_read(tmp_path, monkeypatch, name):
    path = _base_file(tmp_path)
    replace_parts(path, {SHEET: PREFIXED_SHEETS[name]})
    assert not _fast_path_applies(path)
    fast, slow = _both_paths(path, monkeypatch)
    assert fast == slow == [(1, [(0, 1)])]


def test_fast_path_applies_to_plain_sheets(tmp_path):
    path = _base_file(tmp_path)
    replace_parts(path, {SHEET: _sheet(SUPPORTED_BODIES["style_then_type"])})
    assert _fast_path_applies(path)


def test_strict_namespace_sheet_matches(tmp_path, monkeypatch):
    path = _base_file(tmp_path)
    body = '<row r="1"><c r="A1"><v>5</v></c><c r="B1" t="s"><v>0</v></c></row>'
    replace_parts(path, {SHEET: _sheet(body, root=f'<worksheet xmlns="{STRICT_NS}">')})
    fast, slow = _both_paths(path, monkeypatch)
    assert fast == slow


@pytest.mark.parametrize("name", sorted(SUPPORTED_BODIES))
def test_c_reader_matches_python_and_element_tree(tmp_path, monkeypatch, name):
    path = _base_file(tmp_path)
    replace_parts(path, {SHEET: _sheet(SUPPORTED_BODIES[name])})
    from_c, from_python, from_tree = _all_paths(path, monkeypatch)
    _same_typed_rows(from_c, from_python)
    _same_typed_rows(from_c, from_tree)


@pytest.mark.parametrize("name", sorted(PREFIXED_SHEETS))
def test_c_reader_declines_prefixed_sheets(tmp_path, monkeypatch, name):
    path = _base_file(tmp_path)
    replace_parts(path, {SHEET: PREFIXED_SHEETS[name]})
    from_c, from_python, from_tree = _all_paths(path, monkeypatch)
    _same_typed_rows(from_c, from_python)
    _same_typed_rows(from_c, from_tree)
    assert from_c == [(1, [(0, 1)])]


# Bodies the C scanner must hand to ElementTree instead of guessing.
FALLBACK_BODIES = {
    "rich_inline_string_falls_back",
    "entities_in_inline_fall_back",
    "entities_in_str_fall_back",
    "carriage_return_in_str_falls_back",
    "unknown_child_falls_back",
    "lowercase_reference_falls_back",
    "prefixed_cell_falls_back",
}


@pytest.mark.skipif(not _accel.HAVE_C_EXT, reason="C extension not built")
@pytest.mark.parametrize("name", sorted(SUPPORTED_BODIES))
def test_c_reader_signals_fallback_instead_of_guessing(tmp_path, name):
    path = _base_file(tmp_path)
    replace_parts(path, {SHEET: _sheet(SUPPORTED_BODIES[name])})
    items = _c_items(path)
    assert any(isinstance(item, bytes) for item in items) == (name in FALLBACK_BODIES)


@pytest.mark.skipif(not _accel.HAVE_C_EXT, reason="C extension not built")
def test_c_reader_handles_tiny_stream_chunks(tmp_path):
    path = tmp_path / "chunks.xlsx"
    rows = [[index, f"text-{index}", index * 0.25, index % 2 == 0] for index in range(400)]
    with XlsxWriter(path) as writer:
        writer.write_sheet(rows)

    raw = _sheet_bytes(path)
    marker = raw.index(b"<sheetData>") + len(b"<sheetData>")
    cut = marker + (len(raw) - marker) // 3
    probe = _xlsx_read._probe_fast(io.BytesIO(raw[:cut]))
    assert probe is not None
    rest = raw[cut:]
    with ExcelReader(str(path)) as reader:
        expected = list(_xlsx_read.iter_indexed_rows(reader, "Sheet1"))
        for size in (1, 3, 7, 64):
            rows = _accel.read_xlsx_worksheet(
                _Trickle(rest, size), probe[4], reader._shared_string_array,
                _xlsx_read._date_style_flags(reader), reader._date_1904)
            got = list(_xlsx_read._c_rows(reader, rows, probe[2], probe[3]))
            _same_typed_rows(got, expected)


def test_c_reader_resumes_row_numbers_after_a_fallback(tmp_path, monkeypatch):
    path = _base_file(tmp_path)
    body = ('<row r="12 "><c r="A12"><v>1</v></c></row>'
            '<row><c r="B13"><v>2</v></c><c><v>3</v></c></row>')
    replace_parts(path, {SHEET: _sheet(body)})
    from_c, from_python, from_tree = _all_paths(path, monkeypatch)
    _same_typed_rows(from_c, from_python)
    _same_typed_rows(from_c, from_tree)
    assert from_c == [(12, [(0, 1)]), (13, [(1, 2), (2, 3)])]


def test_c_reader_honors_the_1904_date_system(tmp_path, monkeypatch):
    path = _base_file(tmp_path)
    body = ('<row r="1"><c r="A1" s="1"><v>0</v></c><c r="B1" s="1"><v>59</v></c>'
            '<c r="C1" s="1"><v>60</v></c><c r="D1" s="1"><v>43000.5</v></c></row>')
    replace_parts(path, {SHEET: _sheet(body)})

    with ExcelReader(str(path)) as reader:
        reader._date_1904 = True
        from_c = list(_xlsx_read.iter_indexed_rows(reader, "Sheet1"))
    with monkeypatch.context() as patch:
        patch.setattr(_xlsx_read._accel, "read_xlsx_worksheet", lambda *args: None)
        with ExcelReader(str(path)) as reader:
            reader._date_1904 = True
            from_python = list(_xlsx_read.iter_indexed_rows(reader, "Sheet1"))
    _same_typed_rows(from_c, from_python)
    assert [cell for _number, cells in from_c for _col, cell in cells] == [
        datetime.datetime(1904, 1, 1),
        datetime.datetime(1904, 2, 29),
        datetime.datetime(1904, 3, 1),
        datetime.datetime(2021, 9, 23, 12, 0),
    ]


def test_c_reader_reads_iso_dates_and_empty_date_cells(tmp_path, monkeypatch):
    path = _base_file(tmp_path)
    body = ('<row r="1"><c r="A1" t="d"><v>2024-02-29T12:30:00</v></c>'
            '<c r="B1" t="d"><v></v></c></row>')
    replace_parts(path, {SHEET: _sheet(body)})
    from_c, from_python, from_tree = _all_paths(path, monkeypatch)
    assert from_c == [(1, [(0, datetime.datetime(2024, 2, 29, 12, 30)), (1, None)])]
    _same_typed_rows(from_c, from_python)
    _same_typed_rows(from_c, from_tree)


def test_c_reader_propagates_number_errors(tmp_path, monkeypatch):
    path = _base_file(tmp_path)
    replace_parts(path, {SHEET: _sheet('<row r="1"><c r="A1"><v>abc</v></c></row>')})
    with pytest.raises(ValueError):
        with ExcelReader(str(path)) as reader:
            list(_xlsx_read.iter_indexed_rows(reader, "Sheet1"))
    with monkeypatch.context() as patch:
        patch.setattr(_xlsx_read._accel, "read_xlsx_worksheet", lambda *args: None)
        with pytest.raises(ValueError):
            with ExcelReader(str(path)) as reader:
                list(_xlsx_read.iter_indexed_rows(reader, "Sheet1"))


def test_c_reader_propagates_invalid_utf8(tmp_path):
    path = _base_file(tmp_path)
    xml = (b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
           b'<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>\xff</t></is></c>'
           b'</row></sheetData></worksheet>')
    replace_parts(path, {SHEET: xml})
    with pytest.raises(UnicodeDecodeError):
        with ExcelReader(str(path)) as reader:
            list(_xlsx_read.iter_indexed_rows(reader, "Sheet1"))


def test_c_reader_propagates_truncated_utf8(tmp_path):
    path = _base_file(tmp_path)
    xml = (b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
           b'<sheetData><row r="1"><c r="A1"><v>1</v></c></row>'
           b'</sheetData></worksheet>\xe2\x82')
    replace_parts(path, {SHEET: xml})
    with pytest.raises(UnicodeDecodeError):
        with ExcelReader(str(path)) as reader:
            list(_xlsx_read.iter_indexed_rows(reader, "Sheet1"))


def test_values_match_across_chunk_boundaries(tmp_path, monkeypatch):
    # Rows far larger than the chunk size, so rows straddle chunk boundaries.
    path = tmp_path / "chunks.xlsx"
    rows = [[index, f"text-{index}-" + "x" * 300, index * 0.25, index % 2 == 0] for index in range(3000)]
    with XlsxWriter(path) as writer:
        writer.write_sheet(rows)
    from_c, from_python, from_tree = _all_paths(path, monkeypatch)
    _same_typed_rows(from_c, from_python)
    _same_typed_rows(from_c, from_tree)
    assert len(from_c) == 3000


def test_writer_output_matches_element_tree(tmp_path, monkeypatch):
    path = tmp_path / "mixed.xlsx"
    with XlsxWriter(path) as writer:
        writer.write_sheet([
            ["h1", "h2", "h3", "h4"],
            [None, "a & b < c", "line\nbreak", "_x0041_"],
            [datetime.datetime(2024, 2, 29, 12, 30), -1, 1.25e20, False],
            [],
            [None, None, None, "end"],
        ])
    fast, slow = _both_paths(path, monkeypatch)
    assert fast == slow
