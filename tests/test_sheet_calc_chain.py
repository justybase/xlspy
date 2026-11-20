"""Reordering sheets must not leave a calcChain pointing at old sheet indexes.

calcChain entries name formula cells by sheet index, so any structural change to
the sheet list has to drop the part. Excel rebuilds it on the next calculation.
"""

from __future__ import annotations

from xml.etree import ElementTree as ET

from xlspy import XlsbUpdater, XlsbWriter, XlsxUpdater, XlsxWriter

REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"


def _workbook(path, writer_class):
    with writer_class(path) as writer:
        for name in ("One", "Two"):
            writer.add_sheet(name)
            writer.write_sheet([[name]])
    return path


def _inject_calc_chain(updater, part, rels_path, target, content_type):
    package = updater._package
    package.set(part, b"<calcChain/>")
    rels = ET.fromstring(package.get(rels_path))
    ET.SubElement(
        rels,
        f"{{{PKG_REL}}}Relationship",
        {"Id": "rIdCalcChain", "Type": REL + "/calcChain", "Target": target},
    )
    package.set(rels_path, ET.tostring(rels, encoding="utf-8"))
    types = ET.fromstring(package.get("[Content_Types].xml"))
    ET.SubElement(
        types,
        f"{{{CT}}}Override",
        {"PartName": "/" + part, "ContentType": content_type},
    )
    package.set("[Content_Types].xml", ET.tostring(types, encoding="utf-8"))


def _assert_dropped(updater, part, rels_path):
    package = updater._package
    assert not package.has(part)
    rels = ET.fromstring(package.get(rels_path))
    assert not any(
        node.get("Type", "").endswith("/calcChain") for node in rels
    )
    types = ET.fromstring(package.get("[Content_Types].xml"))
    assert all(node.get("PartName") != "/" + part for node in types)


def test_xlsx_reorder_drops_calc_chain(tmp_path):
    path = _workbook(tmp_path / "reorder.xlsx", XlsxWriter)
    updater = XlsxUpdater(path)
    try:
        _inject_calc_chain(
            updater,
            "xl/calcChain.xml",
            "xl/_rels/workbook.xml.rels",
            "calcChain.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml",
        )
        updater.reorder_sheets(["Two", "One"])
        _assert_dropped(updater, "xl/calcChain.xml", "xl/_rels/workbook.xml.rels")
        assert b"fullCalcOnLoad=\"1\"" in updater._package.get("xl/workbook.xml")
        updater.save(path)
    finally:
        updater.close()

    reopened = XlsxUpdater(path)
    try:
        assert reopened.get_sheet_names() == ["Two", "One"]
        _assert_dropped(reopened, "xl/calcChain.xml", "xl/_rels/workbook.xml.rels")
    finally:
        reopened.close()


def test_xlsb_reorder_drops_calc_chain(tmp_path):
    path = _workbook(tmp_path / "reorder.xlsb", XlsbWriter)
    updater = XlsbUpdater(path)
    try:
        _inject_calc_chain(
            updater,
            "xl/calcChain.bin",
            "xl/_rels/workbook.bin.rels",
            "calcChain.bin",
            "application/vnd.ms-excel.calcChain",
        )
        updater.reorder_sheets(["Two", "One"])
        _assert_dropped(updater, "xl/calcChain.bin", "xl/_rels/workbook.bin.rels")
        updater.save(path)
    finally:
        updater.close()

    reopened = XlsbUpdater(path)
    try:
        assert reopened.get_sheet_names() == ["Two", "One"]
        _assert_dropped(reopened, "xl/calcChain.bin", "xl/_rels/workbook.bin.rels")
    finally:
        reopened.close()
