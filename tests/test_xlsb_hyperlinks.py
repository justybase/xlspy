"""Hyperlinks in XLSB output: BrtHLink records and their sheet relationships."""

import re
import struct
import zipfile
from collections import Counter
from typing import NamedTuple

import pytest

from xlspy import Chart, ChartSeries, Comment, ExcelReader, Hyperlink, Style, XlsbWriter
from xlspy.biff_utils import iter_records
from xlspy._xlsb_metadata import read_wide

BRT_HLINK = 0x01EE


class _Link(NamedTuple):
    row: int
    col: int
    rel_id: str
    location: str
    tooltip: str
    display: str


def _sheet(path, number=1):
    with zipfile.ZipFile(path) as archive:
        return archive.read(f"xl/worksheets/sheet{number}.bin")


def _rels(path, number=1):
    with zipfile.ZipFile(path) as archive:
        return archive.read(f"xl/worksheets/_rels/sheet{number}.bin.rels").decode()


def _links(path):
    """Decode every BrtHLink record of sheet 1."""
    data = _sheet(path)
    result = []
    for record in iter_records(data):
        if record.record_id != BRT_HLINK:
            continue
        payload = data[record.data_start:record.data_end]
        first_row, last_row, first_col, _last_col = struct.unpack_from("<4I", payload, 0)
        assert first_row == last_row
        offset = 16
        strings = []
        for _ in range(4):
            text, offset = read_wide(payload, offset)
            strings.append(text)
        assert offset == len(payload)
        result.append(_Link(first_row, first_col, *strings))
    return result


def _rel_entries(rels):
    """(Id, relationship type suffix, Target) for every Relationship element."""
    return re.findall(r'Id="(rId\d+)" Type="[^"]*/([^"/]+)" Target="([^"]*)"', rels)


def test_external_link_has_relationship_and_tooltip(tmp_path):
    path = tmp_path / "links.xlsb"
    with XlsbWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([
            ["Page"],
            [Hyperlink("https://example.com/report", text="Report", tooltip="Tooltip")],
        ])

    (link,) = _links(path)
    assert (link.row, link.col) == (1, 0)
    assert link.location == ""
    assert link.tooltip == "Tooltip"
    assert link.display == "https://example.com/report"
    rels = _rels(path)
    assert (
        f'Id="{link.rel_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" '
        f'Target="https://example.com/report" TargetMode="External"'
    ) in rels


def test_internal_link_has_location_and_no_relationship(tmp_path):
    path = tmp_path / "internal.xlsb"
    with XlsbWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([["x"], [Hyperlink("#'Data'!B5", text="Above")]])

    (link,) = _links(path)
    assert (link.row, link.col, link.rel_id) == (1, 0, "")
    assert link.location == link.display == "'Data'!B5"
    assert link.tooltip == ""
    assert "/hyperlink" not in _rels(path)


def test_one_relationship_per_target_and_ids_are_unique(tmp_path):
    path = tmp_path / "parts.xlsb"
    with XlsbWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(
            [["a", "b"],
             [Hyperlink("https://example.com/x"), Hyperlink("https://example.com/x")],
             [Hyperlink("mailto:someone@example.com"), Hyperlink("#'Data'!A1")]],
            charts=[Chart("bar", [ChartSeries("A2:A3")], anchor="E2")],
            comments={"B1": Comment("note", "Author")},
        )

    rels = _rels(path)
    entries = _rel_entries(rels)
    ids = [rel_id for rel_id, _kind, _target in entries]
    assert len(ids) == len(set(ids))
    kinds = {rel_id: kind for rel_id, kind, _target in entries}
    assert "drawing" in kinds.values() and "comments" in kinds.values()

    link_targets = {rel_id: target for rel_id, kind, target in entries if kind == "hyperlink"}
    assert sorted(link_targets.values()) == ["https://example.com/x", "mailto:someone@example.com"]

    links = _links(path)
    assert len(links) == 4
    uses = Counter(link.rel_id for link in links if link.rel_id)
    assert set(uses) == set(link_targets)
    assert sorted(uses.values()) == [1, 2]


def test_link_text_reads_back(tmp_path):
    path = tmp_path / "styled.xlsb"
    with XlsbWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([
            [Hyperlink("https://example.com", text="Plain")],
            [(Hyperlink("https://example.com/own", text="Own"), Style(bold=True))],
        ])

    with ExcelReader(str(path)) as reader:
        values = [row[0] for row in reader.get_rows("Data")]
    assert values == ["Plain", "Own"]


def test_hyperlink_cell_cannot_carry_a_number_format(tmp_path):
    path = tmp_path / "format.xlsb"
    with pytest.raises(TypeError, match="cannot carry a number format"):
        with XlsbWriter(path) as writer:
            writer.add_sheet("Data")
            writer.write_sheet([[(Hyperlink("https://example.com"), "0.00")]])
