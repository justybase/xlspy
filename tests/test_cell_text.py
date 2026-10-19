"""Cell text: Excel's 32,767-unit limit and XML-safe text, for every writer path.

Every case runs through each format (XLSX, XLSB), each string storage (shared,
inline) and each row encoder (C, and the Python encoder forced in-process).
Excel 16.0 rejects workbooks whose cells exceed the limit (verified over COM),
and it counts the limit in UTF-16 code units.
"""

import io
import struct
import zipfile
import xml.etree.ElementTree as ElementTree

import pytest

import xlspy.writer
import xlspy.xlsx_writer
from xlspy import ExcelReader, Formula, Hyperlink, Style, XlsbWriter, XlsxWriter
from xlspy import _accel

LIMIT = 32767
ASTRAL = "\U0001D11E"  # outside the BMP: two UTF-16 units

FORMATS = [
    pytest.param(XlsxWriter, "xlsx", True, id="xlsx-shared"),
    pytest.param(XlsxWriter, "xlsx", False, id="xlsx-inline"),
    pytest.param(XlsbWriter, "xlsb", True, id="xlsb-shared"),
    pytest.param(XlsbWriter, "xlsb", False, id="xlsb-inline"),
]


@pytest.fixture(params=["c", "python"])
def encoder(request, monkeypatch):
    """Run with the C row encoders, or with them replaced by the Python ones."""
    if request.param == "c":
        if not _accel.HAVE_C_EXT:
            pytest.skip("C extension not built")
    else:
        def no_c(*args, **kwargs):
            return None

        monkeypatch.setattr(xlspy.xlsx_writer, "accel_encode_xlsx_row", no_c)
        monkeypatch.setattr(xlspy.writer, "encode_xlsb_row", no_c)
        monkeypatch.setattr(xlspy.writer, "encode_xlsb_row_inline", no_c)
    return request.param


def _write(writer_cls, shared, rows, **options):
    buffer = io.BytesIO()
    with writer_cls(buffer, useSharedStrings=shared) as writer:
        writer.add_sheet("T")
        writer.write_sheet(rows, **options)
    return buffer.getvalue()


def _read(data, ext, tmp_path):
    path = tmp_path / f"read.{ext}"
    path.write_bytes(data)
    with ExcelReader(str(path)) as reader:
        return list(reader.get_rows("T"))


def _assert_rejected(writer_cls, shared, rows, error=ValueError, match="32767", **options):
    buffer = io.BytesIO()
    with pytest.raises(error, match=match):
        with writer_cls(buffer, useSharedStrings=shared) as writer:
            writer.add_sheet("T")
            writer.write_sheet(rows, **options)
    assert buffer.getvalue() == b"", "a rejected workbook must not be published"


# ------------------------------------------------------------- the limit


@pytest.mark.parametrize("writer_cls,ext,shared", FORMATS)
@pytest.mark.parametrize("text", [
    pytest.param("", id="empty"),
    pytest.param("a" * (LIMIT - 1), id="32766"),
    pytest.param("a" * LIMIT, id="32767"),
    pytest.param(ASTRAL * 16383, id="astral-32766-units"),
    pytest.param("a" + ASTRAL * 16383, id="astral-32767-units"),
    pytest.param("ż" * LIMIT, id="non-ascii-65534-utf8-bytes"),
    pytest.param("é" * 16383, id="combining-32766"),
    pytest.param("&<>" * 10922 + "&", id="xml-special-32767"),
])
def test_text_within_the_limit_round_trips(tmp_path, encoder, writer_cls, ext, shared, text):
    data = _write(writer_cls, shared, [["head", "tail"], [text, "after"]])
    rows = _read(data, ext, tmp_path)
    assert rows[1][1] == "after"
    assert rows[1][0] in ((text,) if text else ("", None))


@pytest.mark.parametrize("writer_cls,ext,shared", FORMATS)
@pytest.mark.parametrize("text", [
    pytest.param("a" * (LIMIT + 1), id="32768"),
    pytest.param("x" * 40_000, id="40000"),
    pytest.param(ASTRAL * 16384, id="astral-32768-units"),
    # 32,767 code points, but the astral character makes it 32,768 units.
    pytest.param("a" * (LIMIT - 1) + ASTRAL, id="32767-code-points-32768-units"),
])
def test_text_over_the_limit_is_rejected(encoder, writer_cls, ext, shared, text):
    _assert_rejected(writer_cls, shared, [["head"], [text]])


@pytest.mark.parametrize("writer_cls,ext,shared", FORMATS)
def test_oversized_value_deep_in_a_stream_publishes_nothing(encoder, writer_cls, ext, shared):
    def rows():
        for i in range(5000):
            yield [f"row {i}", i]
        yield ["x" * (LIMIT + 1), 0]
        yield ["never written", 1]

    _assert_rejected(writer_cls, shared, rows())


@pytest.mark.parametrize("writer_cls,ext,shared", FORMATS)
def test_styled_and_header_cells_are_checked(encoder, writer_cls, ext, shared):
    long_text = "x" * (LIMIT + 1)
    _assert_rejected(writer_cls, shared, [["head"], [long_text]],
                     styles={"A2:A2": Style(bold=True)})
    _assert_rejected(writer_cls, shared, [[long_text], ["body"]],
                     header_style=Style(bold=True))


@pytest.mark.parametrize("writer_cls,ext,shared", FORMATS)
def test_str_of_other_objects_is_checked(encoder, writer_cls, ext, shared):
    class Long:
        def __str__(self):
            return "y" * (LIMIT + 1)

    _assert_rejected(writer_cls, shared, [["head"], [Long()]])


@pytest.mark.parametrize("writer_cls,ext,shared", FORMATS)
def test_dataframe_export_is_checked(writer_cls, ext, shared):
    pandas = pytest.importorskip("pandas")
    frame = pandas.DataFrame({"name": ["ok", "z" * (LIMIT + 1)], "qty": [1, 2]})
    buffer = io.BytesIO()
    with pytest.raises(ValueError, match="32767"):
        with writer_cls(buffer, useSharedStrings=shared) as writer:
            writer.write_dataframe(frame)
    assert buffer.getvalue() == b""


@pytest.mark.parametrize("shared", [True, False], ids=["shared", "inline"])
def test_xlsx_hyperlink_text_and_formula_results_are_checked(encoder, shared):
    long_text = "x" * (LIMIT + 1)
    _assert_rejected(XlsxWriter, shared, [[Hyperlink("https://example.com", text=long_text)]])
    _assert_rejected(XlsxWriter, shared, [[Formula('REPT("x",32768)', value=long_text)]])


def test_xlsb_shared_string_record_keeps_the_full_text():
    # Shared XLSB text used to be cut to 32,767 units when the table was
    # written; at the limit the record must hold every unit.
    text = "a" + ASTRAL * 16383
    data = _write(XlsbWriter, True, [[text]])
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        sst = archive.read("xl/sharedStrings.bin")
    pos = sst.index(b"\x13")  # BrtSSTItem
    size_byte_count = 1 + sum(1 for b in sst[pos + 1:pos + 4] if b & 0x80)
    payload = sst[pos + 1 + size_byte_count:]
    flags, cch = payload[0], struct.unpack_from("<I", payload, 1)[0]
    assert (flags, cch) == (0, LIMIT)
    assert payload[5:5 + 2 * cch].decode("utf-16-le") == text


def _sst_writer(strings):
    """An XlsbWriter whose shared-string table holds ``strings``, ready to encode."""
    writer = XlsbWriter(io.BytesIO())
    writer._shared_strings = list(strings)
    writer._sst_unique_count = writer._sst_all_count = len(strings)
    return writer


@pytest.mark.parametrize("text", [
    pytest.param("a" * (LIMIT + 1), id="32768"),
    pytest.param(ASTRAL * 16384, id="astral-32768-units"),
    pytest.param("a" * (LIMIT - 1) + ASTRAL, id="32767-code-points-32768-units"),
])
@pytest.mark.parametrize("encoder", ["c", "python"])
def test_shared_string_table_refuses_text_over_the_limit(encoder, text, monkeypatch):
    # The writers reject these cells first; the table must not cut them to fit.
    if encoder == "c":
        if not _accel.HAVE_C_EXT:
            pytest.skip("C extension not built")
    else:
        monkeypatch.setattr(xlspy.writer, "accel_encode_xlsb_sst", lambda *args: None)
    with pytest.raises(ValueError, match="32767"):
        _sst_writer(["ok", text])._write_shared_strings_bin(io.BytesIO())


@pytest.mark.skipif(not _accel.HAVE_C_EXT, reason="C extension not built")
@pytest.mark.parametrize("text", [
    pytest.param("a" * LIMIT, id="limit"),
    pytest.param("a" + ASTRAL * 16383, id="astral-at-limit"),
    pytest.param("é" * 100, id="accented"),
])
def test_c_and_python_shared_string_tables_agree(text, monkeypatch):
    strings = ["short", text]
    c_bytes = io.BytesIO()
    _sst_writer(strings)._write_shared_strings_bin(c_bytes)
    monkeypatch.setattr(xlspy.writer, "accel_encode_xlsb_sst", lambda *args: None)
    python_bytes = io.BytesIO()
    _sst_writer(strings)._write_shared_strings_bin(python_bytes)
    assert c_bytes.getvalue() == python_bytes.getvalue()


# --------------------------------------------------- invalid Unicode


@pytest.mark.parametrize("writer_cls,ext,shared", FORMATS)
def test_lone_surrogate_is_refused(encoder, writer_cls, ext, shared):
    _assert_rejected(writer_cls, shared, [["head"], ["bad\ud800text"]],
                     error=UnicodeEncodeError, match="surrogate")


# ------------------------------------------- XML-unsafe and special text


SPECIAL = [
    "line1\nline2",
    "crlf\r\nend",
    "tab\tx",
    "bell\x07",
    "nul\x00",
    "ctl\x01\x1f",
    "lit _x0041_ text",
    "_x005F_",
    "_x00e9_ lower-case hex",
    "_x12_ not an escape",
    "ff￾￿",
    "a&b<c>d\"e'f",
    " leading space",
]


@pytest.mark.parametrize("writer_cls,ext,shared", FORMATS)
@pytest.mark.parametrize("styled", [False, True], ids=["plain", "styled"])
def test_special_text_round_trips(tmp_path, encoder, writer_cls, ext, shared, styled):
    options = {"styles": {"A2:M2": Style(italic=True)}} if styled else {}
    data = _write(writer_cls, shared, [["head"] * len(SPECIAL), SPECIAL], **options)
    assert _read(data, ext, tmp_path)[1] == SPECIAL


@pytest.mark.parametrize("shared", [True, False], ids=["shared", "inline"])
@pytest.mark.parametrize("styled", [False, True], ids=["plain", "styled"])
def test_xlsx_parts_stay_well_formed_xml(encoder, shared, styled):
    options = {"styles": {"A1:M1": Style(italic=True)}} if styled else {}
    data = _write(XlsxWriter, shared, [SPECIAL, [Formula('"x"', value="bell\x07")]], **options)
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        for name in archive.namelist():
            if name.endswith(".xml"):
                ElementTree.fromstring(archive.read(name))  # raises if ill-formed


def test_xlsx_escapes_follow_ooxml():
    data = _write(XlsxWriter, False, [["bell\x07", "crlf\r\n", "lit _x0041_"]])
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        sheet = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")
    assert "<t>bell_x0007_</t>" in sheet
    assert "<t>crlf_x000D_\n</t>" in sheet
    assert "<t>lit _x005F_x0041_</t>" in sheet


@pytest.mark.skipif(not _accel.HAVE_C_EXT, reason="C extension not built")
@pytest.mark.parametrize("text", ["x" * (LIMIT + 1), "a" * (LIMIT - 1) + ASTRAL],
                         ids=["32768", "32767-code-points-32768-units"])
def test_c_shared_row_encoders_hand_long_text_back_untouched(text):
    # A row the C encoder refuses goes to the Python encoder, which raises.
    # Strings the C encoder interned earlier in that row must be removed again,
    # or the shared-string indexes would drift from the caller's counters.
    from xlspy import _c_core

    for encode in (
        lambda d, s: _c_core.encode_xlsb_row(["new one", "new two", text], d, s, 1, 1, 1),
        lambda d, s: _c_core.encode_xlsx_row(["new one", "new two", text], d, s, 1, 1, 1, True),
    ):
        table, strings = {"old": 0}, ["old"]
        assert encode(table, strings) is None
        assert (table, strings) == ({"old": 0}, ["old"])


@pytest.mark.skipif(not _accel.HAVE_C_EXT, reason="C extension not built")
def test_c_encoders_hand_unusual_cells_back():
    from xlspy import _c_core

    class BrokenStr:
        def __str__(self):
            raise RuntimeError("str failed")

    class Plain:
        def __str__(self):
            return "plain"

    # Objects of other types (Formula, Hyperlink, anything with a __str__) go
    # to the Python encoder, which writes or refuses them; str() raising used
    # to drop the XLSB cell silently.
    for cell in (BrokenStr(), Plain(), Formula("A1"), Hyperlink("https://example.com")):
        assert _c_core.encode_xlsb_row_inline(["a", cell], 1) is None
        assert _c_core.encode_xlsb_row(["a", cell], {}, [], 0, 0, 1) is None
        assert _c_core.encode_xlsx_row(["a", cell], {}, [], 0, 0, 1, False) is None
    for text in ("bell\x07", "cr\r", "lit _x0041_", "ff￾"):
        assert _c_core.encode_xlsx_row([text], {}, [], 0, 0, 1, False) is None
    # Shared XLSX text is escaped when the table is written, so C keeps it.
    assert _c_core.encode_xlsx_row(["bell\x07"], {}, [], 0, 0, 1, True) is not None


@pytest.mark.skipif(not _accel.HAVE_C_EXT, reason="C extension not built")
def test_c_shared_encoders_keep_errors_raised_by_string_hashing():
    from xlspy import _c_core

    class BadHash(str):
        def __hash__(self):
            raise RuntimeError("hash failed")

    with pytest.raises(RuntimeError, match="hash failed"):
        _c_core.encode_xlsb_row(["ok", BadHash("x")], {}, [], 0, 0, 1)
    with pytest.raises(RuntimeError, match="hash failed"):
        _c_core.encode_xlsx_row(["ok", BadHash("x")], {}, [], 0, 0, 1, True)


@pytest.mark.parametrize("writer_cls,ext,shared", FORMATS)
def test_failing_str_is_reported_not_dropped(encoder, writer_cls, ext, shared):
    class BrokenStr:
        def __str__(self):
            raise RuntimeError("str failed")

    _assert_rejected(writer_cls, shared, [["head"], ["a", BrokenStr()]],
                     error=RuntimeError, match="str failed")


def test_c_and_python_xlsx_encoders_agree_on_ordinary_text(monkeypatch):
    if not _accel.HAVE_C_EXT:
        pytest.skip("C extension not built")
    rows = [["Name", "Note"], ["plain", "a&b"], ["żółw", ASTRAL + "x"],
            ["_x", "x_"], ["_x00G0_", "tab\tand\nnewline"]]
    outputs = []
    for use_c in (True, False):
        if not use_c:
            monkeypatch.setattr(xlspy.xlsx_writer, "accel_encode_xlsx_row", lambda *a, **k: None)
        for shared in (True, False):
            data = _write(XlsxWriter, shared, rows)
            with zipfile.ZipFile(io.BytesIO(data)) as archive:
                outputs.append([archive.read(n) for n in sorted(archive.namelist()) if n.endswith(".xml")])
    assert outputs[0] == outputs[2]  # shared: C == Python
    assert outputs[1] == outputs[3]  # inline: C == Python
