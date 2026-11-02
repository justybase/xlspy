"""Writer constructor keywords are snake_case; the old camelCase names say what they became."""

import io
import zipfile

import pytest

from xlspy import XlsbWriter, XlsxWriter

WRITERS = [XlsxWriter, XlsbWriter]


@pytest.mark.parametrize("writer", WRITERS)
@pytest.mark.parametrize("old,new", [("compressionLevel", "compression_level"),
                                     ("useSharedStrings", "use_shared_strings")])
def test_camel_case_keywords_name_their_replacement(writer, old, new):
    with pytest.raises(TypeError, match=f"argument '{old}' was renamed to '{new}'"):
        writer(io.BytesIO(), **{old: 1})


@pytest.mark.parametrize("writer", WRITERS)
def test_an_unknown_keyword_is_a_type_error(writer):
    with pytest.raises(TypeError, match="unexpected keyword argument 'colour'"):
        writer(io.BytesIO(), colour=1)


def _write(writer, **options) -> bytes:
    buffer = io.BytesIO()
    with writer(buffer, **options) as book:
        book.add_sheet("S")
        book.write_sheet([["repeated text"] * 20 for _ in range(200)])
    return buffer.getvalue()


@pytest.mark.parametrize("writer", WRITERS)
def test_compression_level_is_used(writer):
    assert len(_write(writer, compression_level=0)) > len(_write(writer, compression_level=9))


@pytest.mark.parametrize("writer,suffix", [(XlsxWriter, "xml"), (XlsbWriter, "bin")])
def test_use_shared_strings_false_keeps_text_out_of_the_table(writer, suffix):
    for shared in (True, False):
        with zipfile.ZipFile(io.BytesIO(_write(writer, use_shared_strings=shared))) as archive:
            names = archive.namelist()
            table = archive.read(f"xl/sharedStrings.{suffix}") if f"xl/sharedStrings.{suffix}" in names else b""
        text = "repeated text".encode("utf-16-le" if suffix == "bin" else "utf-8")
        assert (text in table) is shared
