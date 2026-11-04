"""Edits take time in proportion to their number, and the range index behind them finds
what going through every range finds.

The limits are far above what these take now; the loops they guard took minutes at these
sizes, and nothing else would notice: each operation works, only slowly.
"""
from __future__ import annotations

import random
import struct
import time
import tracemalloc

import pytest

from xlspy import Comment, Hyperlink, Workbook, XlsbWriter, XlsxWriter, load_workbook
from xlspy._xlsb_metadata import check_merged_refs
from xlspy.biff_utils import build_record
from xlspy.refs import RangeIndex
from xlspy.workbook._coords import MultiCellRange
from xlspy.xlsb_updater import XlsbUpdater
from xlspy.xlsx_updater import XlsxUpdater

FORMATS = {"xlsx": (XlsxWriter, XlsxUpdater), "xlsb": (XlsbWriter, XlsbUpdater)}
_LIMIT = 15.0


def _source(tmp_path, ext, rows=40000, **options):
    path = tmp_path / f"source.{ext}"
    with FORMATS[ext][0](str(path)) as writer:
        writer.add_sheet("S")
        writer.write_sheet(([row, f"t{row}"] for row in range(rows)), auto_filter=False, **options)
    return path


def _timed_update(tmp_path, ext, change, **options):
    updater = FORMATS[ext][1](str(_source(tmp_path, ext, **options)))
    try:
        start = time.perf_counter()
        change(updater)
        updater.save(str(tmp_path / f"out.{ext}"))
        assert time.perf_counter() - start < _LIMIT
    finally:
        updater.close()
    return tmp_path / f"out.{ext}"


@pytest.mark.parametrize("ext", FORMATS)
def test_many_notes_at_once(tmp_path, ext):
    out = _timed_update(tmp_path, ext, lambda updater: updater.set_comments(
        "S", {f"C{row}": Comment(f"n{row}", f"author {row % 500}") for row in range(1, 20000)}))
    updater = FORMATS[ext][1](str(out))
    try:
        notes = updater.get_comments("S")
    finally:
        updater.close()
    assert len(notes) == 19999
    assert notes["C777"] == Comment("n777", "author 277")


@pytest.mark.parametrize("ext", FORMATS)
def test_many_links_at_once(tmp_path, ext):
    out = _timed_update(tmp_path, ext, lambda updater: updater.set_hyperlinks(
        "S", {f"C{row}": Hyperlink(f"https://x{row}.example/") for row in range(1, 20000)}))
    sheet = load_workbook(str(out)).active
    assert sheet["C777"].hyperlink.target == "https://x777.example/"


@pytest.mark.parametrize("ext", FORMATS)
def test_many_edits_on_a_sheet_with_many_merges(tmp_path, ext):
    merges = [f"G{row}:H{row}" for row in range(1, 20000)]
    _timed_update(tmp_path, ext, lambda updater: updater.update_cells(
        "S", {f"G{row}": row for row in range(1, 20000)}), merges=merges)
    with pytest.raises(ValueError, match="top-left"):
        _timed_update(tmp_path, ext, lambda updater: updater.update_cells("S", {"H19999": 1}), merges=merges)


@pytest.mark.parametrize("ext", FORMATS)
def test_many_merges_in_the_model(tmp_path, ext):
    loaded = load_workbook(str(_source(tmp_path, ext)))
    sheet = loaded.active
    start = time.perf_counter()
    for row in range(1, 20000):
        sheet.merge_cells(f"G{row}:H{row}")
    with pytest.raises(ValueError, match="overlaps 'G19999:H19999'"):
        sheet.merge_cells("H19999:I19999")
    for row in range(1, 20000):
        assert sheet.cell(row, 7)._sheet is sheet
    for row in range(1, 10000):
        sheet.unmerge_cells(f"G{row}:H{row}")
    loaded.save(str(tmp_path / f"out.{ext}"))
    assert time.perf_counter() - start < _LIMIT
    assert len(load_workbook(str(tmp_path / f"out.{ext}")).active.merged_cells) == 10000


def _brute(ranges, query):
    c1, r1, c2, r2 = query
    return [bounds for bounds in ranges
            if bounds[0] <= c2 and c1 <= bounds[2] and bounds[1] <= r2 and r1 <= bounds[3]]


def _range(rng):
    # Near each other, far apart, and wider than the index keeps in its buckets.
    row = rng.choice([rng.randrange(200), rng.randrange(1048576)])
    height = rng.choice([0, 0, 1, 5, 63, 64, 65, 400, 511, 512, 513, 5000, 1048575 - row])
    column = rng.randrange(30)
    return column, row, column + rng.choice([0, 1, 4]), min(row + height, 1048575)


@pytest.mark.parametrize("seed", range(300))
def test_the_index_finds_what_going_through_every_range_finds(seed):
    rng = random.Random(seed)
    index, present = RangeIndex(), []
    for _ in range(rng.randint(0, 60)):
        bounds = _range(rng)
        if bounds in present:
            continue
        if present and rng.random() < 0.2:
            gone = present.pop(rng.randrange(len(present)))
            index.remove(gone, str(gone))
        index.add(bounds, str(bounds))
        present.append(bounds)
    for _ in range(60):
        query = _range(rng)
        assert index.overlapping(query) == [(*bounds, str(bounds)) for bounds in _brute(present, query)]


def test_a_range_claiming_four_billion_rows_takes_no_memory_to_index():
    index = RangeIndex()
    tracemalloc.start()
    try:
        for column in range(100):
            index.add((column, 0, column, 0xFFFFFFFF))
        assert tracemalloc.get_traced_memory()[1] < 1_000_000
    finally:
        tracemalloc.stop()
    assert len(index.overlapping((5, 4_000_000_000, 5, 4_000_000_000))) == 1


def test_merge_records_of_a_damaged_sheet_are_checked_without_a_table_of_their_rows():
    record = build_record(0xB0, struct.pack("<IIII", 0, 0xFFFFFFFF, 0, 3))
    with pytest.raises(ValueError, match="top-left"):
        check_merged_refs(record * 2000, ["B9"])
    check_merged_refs(record * 2000, ["A1", "E9"])


def test_merged_ranges_follow_changes_made_directly_to_the_public_list():
    merged = MultiCellRange(["A1:B2", "D4:E5"])
    assert "A1" in merged and merged.first_overlapping((0, 0, 0, 0)) == "A1:B2"
    merged.ranges.remove("A1:B2")
    assert "A1" not in merged and "A1:B2" not in merged
    merged.ranges.append(type(merged.ranges[0])("G7:H8"))
    assert "H8" in merged and "G7:H8" in merged
    merged.add("G7:H8")
    assert len(merged) == 2
    merged.remove("G7:H8")
    assert list(merged) == ["D4:E5"]
    with pytest.raises(ValueError, match="not merged"):
        merged.remove("G7:H8")
    merged.ranges.clear()
    assert not merged and "D4" not in merged


def test_a_new_workbook_with_many_merges():
    workbook = Workbook()
    sheet = workbook.active
    start = time.perf_counter()
    for row in range(1, 20000):
        sheet.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
    assert time.perf_counter() - start < _LIMIT
    assert sheet.cell(5, 2).__class__.__name__ == "MergedCell"
    assert sheet.cell(5, 1).__class__.__name__ == "Cell"
    assert sheet.cell(5, 3).__class__.__name__ == "Cell"
