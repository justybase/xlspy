"""Package XML is parsed through xlspy._safe_xml, which refuses DTDs."""

import io
import xml.etree.ElementTree as ET
from xml.parsers import expat

import pytest

from xlspy import _safe_xml

BILLION_LAUGHS = b"""<?xml version="1.0"?>
<!DOCTYPE lolz [
  <!ENTITY lol "lol">
  <!ENTITY lol2 "&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;">
]>
<root>&lol2;</root>"""

EXTERNAL = b"""<?xml version="1.0"?>
<!DOCTYPE root [<!ENTITY ext SYSTEM "file:///etc/passwd">]>
<root>&ext;</root>"""


def test_plain_part_parses_with_and_without_declaration():
    assert _safe_xml.fromstring(b'<?xml version="1.0"?><a x="1"/>').get("x") == "1"
    assert _safe_xml.fromstring('<a><b/></a>').find("b") is not None


def test_byte_order_mark_and_comments_before_root_are_allowed():
    data = b"\xef\xbb\xbf<?xml version='1.0'?><!-- note --><a/>"
    assert _safe_xml.fromstring(data).tag == "a"


def test_doctype_is_refused_in_bytes_and_str():
    with pytest.raises(ET.ParseError):
        _safe_xml.fromstring(BILLION_LAUGHS)
    with pytest.raises(ET.ParseError):
        _safe_xml.fromstring(EXTERNAL.decode())


def test_doctype_after_a_comment_is_still_refused():
    data = b"<?xml version='1.0'?><!-- one --><!-- two --><!DOCTYPE a><a/>"
    with pytest.raises(ET.ParseError):
        _safe_xml.fromstring(data)


def test_doctype_text_inside_the_root_is_not_a_declaration():
    data = b"<a><![CDATA[<!DOCTYPE not-a-dtd>]]></a>"
    assert _safe_xml.fromstring(data).text == "<!DOCTYPE not-a-dtd>"


def test_iterparse_refuses_a_doctype_in_the_stream():
    with pytest.raises(ET.ParseError):
        for _ in _safe_xml.iterparse(io.BytesIO(BILLION_LAUGHS), events=("end",)):
            pass


def test_iterparse_reads_a_plain_stream_in_small_chunks():
    data = b"<?xml version='1.0'?>" + b"<!-- pad -->" * 500 + b"<root><item>1</item><item>2</item></root>"

    class Trickle:
        """A stream that returns at most seven bytes per read."""

        def __init__(self, payload):
            self._stream = io.BytesIO(payload)

        def read(self, size=-1):
            return self._stream.read(7)

    tags = [node.tag for _, node in _safe_xml.iterparse(Trickle(data), events=("end",))]
    assert tags.count("item") == 2


def test_harden_expat_refuses_entity_declarations():
    parser = _safe_xml.harden_expat(expat.ParserCreate())
    # The handler's exception propagates out of pyexpat unchanged.
    with pytest.raises(ET.ParseError, match="document type declarations"):
        parser.Parse(BILLION_LAUGHS, True)


def test_harden_expat_accepts_a_plain_document():
    parser = _safe_xml.harden_expat(expat.ParserCreate())
    parser.Parse(b"<a><b/></a>", True)
