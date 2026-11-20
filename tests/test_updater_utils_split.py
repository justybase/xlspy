"""The updater helpers moved to focused modules, and their old names still work.

`updater_utils` is a facade for the import paths that predate the split: it
must expose the same objects, not copies, so identity checks in callers keep
working.
"""

from __future__ import annotations

import xlspy._xml_text as xml_text
import xlspy._zip_package as zip_package
import xlspy.updater_utils as facade


def test_package_helpers_are_the_same_objects():
    assert facade.ZipPackage is zip_package.ZipPackage
    assert facade.DiskStringIndex is zip_package.DiskStringIndex
    assert facade.STREAM_BUFFER_SIZE == zip_package.STREAM_BUFFER_SIZE
    assert facade.copy_stream is zip_package.copy_stream
    assert facade.copy_file_range is zip_package.copy_file_range


def test_xml_helpers_are_the_same_objects():
    for name in (
        "xml_attribute",
        "escape_xml_text",
        "unescape_xml_text",
        "column_index_to_letter",
        "column_letter_to_index",
        "local_xml_name",
        "replace_or_add_xml_attribute",
        "parse_relationships_xml",
        "resolve_zip_target",
    ):
        assert getattr(facade, name) is getattr(xml_text, name), name
