"""Regression tests for incremental rows and formula text reading."""

import gc
import tracemalloc
import zipfile

import pytest

from xlspy import ExcelReader, Formula, XlsxWriter
from xlspy._xlsx_read import translate_formula
from test_comments_and_editing import replace_parts


def test_ragged_rows_do_not_lose_cells(tmp_path):
    path = tmp_path / "ragged.xlsx"
    with XlsxWriter(path) as writer:
        writer.write_sheet([[], [1, 2, 3], [4]])
    with ExcelReader(str(path)) as reader:
        assert reader.read_all("Sheet1") == [[], [1, 2, 3], [4]]


def test_short_rows_keep_historical_padding(tmp_path):
    path = tmp_path / "padding.xlsx"
    with XlsxWriter(path) as writer:
        writer.write_sheet([[1, 2], [3], [4, 5, 6]])
    with ExcelReader(str(path)) as reader:
        assert reader.read_all("Sheet1") == [[1, 2], [3, None], [4, 5, 6]]


def test_inline_rich_text_keeps_all_runs_and_spaces(tmp_path):
    path = tmp_path / "rich.xlsx"
    with XlsxWriter(path) as writer:
        writer.write_sheet([[1]])
    xml = b'''<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><r><t xml:space="preserve"> first </t></r><r><rPr><b/></rPr><t>second</t></r><t xml:space="preserve"> last </t></is></c></row></sheetData></worksheet>'''
    replace_parts(path, {"xl/worksheets/sheet1.xml": xml})
    with ExcelReader(str(path)) as reader:
        assert reader.read_all("Sheet1") == [[" first second last "]]


def test_first_row_is_yielded_before_sheet_eof(tmp_path, monkeypatch):
    path = tmp_path / "incremental.xlsx"
    with XlsxWriter(path, useSharedStrings=False) as writer:
        writer.write_sheet(([index, "x" * 200] for index in range(10000)))
    with ExcelReader(str(path)) as reader:
        original = reader._archive.open
        reads = []
        class CountingStream:
            def __init__(self, stream):
                self.stream = stream
            def read(self, size=-1):
                assert size > 0, "worksheet must not be read in one unbounded call"
                chunk = self.stream.read(size)
                reads.append(len(chunk))
                return chunk
            def __enter__(self):
                return self
            def __exit__(self, *args):
                self.stream.close()
        def counted(name, *args, **kwargs):
            stream = original(name, *args, **kwargs)
            return CountingStream(stream) if name == "xl/worksheets/sheet1.xml" else stream
        monkeypatch.setattr(reader._archive, "open", counted)
        rows = reader.get_rows("Sheet1")
        assert next(rows) == [0, "x" * 200]
        assert sum(reads) < reader._archive.getinfo("xl/worksheets/sheet1.xml").file_size
        rows.close()


def test_reader_row_memory_is_bounded(tmp_path):
    path = tmp_path / "memory.xlsx"
    with XlsxWriter(path, useSharedStrings=False) as writer:
        writer.write_sheet(([index, "x" * 120] for index in range(30000)))
    gc.collect()
    tracemalloc.start()
    try:
        with ExcelReader(str(path)) as reader:
            assert sum(1 for _ in reader.get_rows("Sheet1")) == 30000
        _, peak = tracemalloc.get_traced_memory()
    finally:
        tracemalloc.stop()
    assert peak < 4 * 1024 * 1024


def test_formula_text_shared_array_and_cached_values(tmp_path):
    path = tmp_path / "shared.xlsx"
    with XlsxWriter(path) as writer:
        writer.write_sheet([[1]])
    xml = b'''<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
<row r="1"><c r="A1"><v>5</v></c><c r="B1"><f t="shared" si="0" ref="B1:B2">A1+$A$1</f><v>10</v></c><c r="C1"><f t="array" ref="C1:C2">A1:A2*2</f><v>10</v></c></row>
<row r="2"><c r="A2"><v>6</v></c><c r="B2"><f t="shared" si="0"/><v>11</v></c><c r="C2"><v>12</v></c></row>
</sheetData></worksheet>'''
    replace_parts(path, {"xl/worksheets/sheet1.xml": xml})
    with ExcelReader(str(path)) as reader:
        assert reader.read_all("Sheet1") == [[5, 10, 10], [6, 11, 12]]
        assert reader.read_all("Sheet1", data_only=False) == [[5, "=A1+$A$1", "=A1:A2*2"], [6, "=A2+$A$1", 12]]


@pytest.mark.parametrize(("text", "expected"), [
    ('A1+$B1+C$1+$D$1', 'B2+$B2+D$1+$D$1'),
    ('SUM(A1:B3)+"A1"', 'SUM(B2:C4)+"A1"'),
    ("'A1 sheet'!A1+Sheet1!B2", "'A1 sheet'!B2+Sheet1!C3"),
    ('SUM(A:A)+SUM(1:3)+Table1[Column1]', 'SUM(B:B)+SUM(2:4)+Table1[Column1]'),
])
def test_shared_formula_translation(text, expected):
    assert translate_formula(text, "B2", "C3") == expected


def test_shared_formula_negative_reference():
    assert translate_formula("A1", "B2", "A1") == "#REF!"


def test_strict_open_xml_is_readable_and_refused_for_editing(tmp_path):
    import zipfile
    from xlspy import Comment, DefinedName, XlsxUpdater, XlsxWriter, load_workbook
    source, strict = tmp_path / "source.xlsx", tmp_path / "strict.xlsx"
    with XlsxWriter(source) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([["a", 1], ["b", 2]], comments={"A1": Comment("note", "Ann")})
        writer.set_defined_name(DefinedName("Total", "'Data'!$B$2"))
    replacements = [
        (b"http://schemas.openxmlformats.org/spreadsheetml/2006/main", b"http://purl.oclc.org/ooxml/spreadsheetml/main"),
        (b"http://schemas.openxmlformats.org/officeDocument/2006/relationships", b"http://purl.oclc.org/ooxml/officeDocument/relationships"),
    ]
    with zipfile.ZipFile(source) as original, zipfile.ZipFile(strict, "w") as output:
        for name in original.namelist():
            data = original.read(name)
            for old, new in replacements:
                data = data.replace(old, new)
            output.writestr(name, data)
    with ExcelReader(str(strict)) as reader:
        assert reader.get_sheet_names() == ["Data"]
        assert reader.read_all("Data") == [["a", 1], ["b", 2]]
        assert reader.get_comments("Data") == {"A1": Comment("note", "Ann")}
        assert reader.get_defined_names() == [DefinedName("Total", "'Data'!$B$2")]
    with pytest.raises(NotImplementedError, match="Strict"):
        XlsxUpdater(strict)
    with pytest.raises(NotImplementedError, match="Strict"):
        load_workbook(strict)
