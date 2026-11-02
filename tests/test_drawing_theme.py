"""The theme part a workbook with drawings refers to (images and charts)."""
from __future__ import annotations

import re
import struct
import zipfile
import zlib

import pytest
from helpers import part_text

from xlspy import Chart, ChartSeries, Image, XlsxWriter

openpyxl = pytest.importorskip("openpyxl")


DATA = [["h"], [1], [2], [3]]


def _png_bytes(width=2, height=2):
    raw = b"".join(b"\x00" + b"\xff\x00\x00" * width for _ in range(height))

    def chunk(kind, data):
        return (
            struct.pack(">I", len(data)) + kind + data
            + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF)
        )

    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
        + chunk(b"IDAT", zlib.compress(raw))
        + chunk(b"IEND", b"")
    )


def _png_file(tmp_path):
    path = tmp_path / "pixel.png"
    path.write_bytes(_png_bytes())
    return path


def _check_package_consistency(path):
    """Every content-type override names a part, and every part a rels file targets exists.

    In particular a theme part must be referenced from the workbook rels.
    """
    with zipfile.ZipFile(path) as archive:
        names = set(archive.namelist())
        content_types = archive.read("[Content_Types].xml").decode()
        for part in re.findall(r'PartName="/([^"]+)"', content_types):
            assert part in names, f"content type for missing part {part}"
        targeted = set()
        for rels_name in (n for n in names if n.endswith(".rels")):
            base = rels_name.replace("_rels/", "").removesuffix(".rels")
            folder = base.rpartition("/")[0]
            for target, mode in re.findall(
                r'Target="([^"]+)"(\s+TargetMode="External")?', archive.read(rels_name).decode()
            ):
                if mode:
                    continue
                parts = (folder.split("/") if folder else []) + target.split("/")
                resolved = []
                for piece in parts:
                    if piece == "..":
                        resolved.pop()
                    elif piece:
                        resolved.append(piece)
                full = "/".join(resolved)
                assert full in names, f"{rels_name} targets missing part {full}"
                targeted.add(full)
        for name in names:
            if name.startswith("xl/theme/"):
                assert name in targeted, f"{name} is in the package but no rels file targets it"


def test_image_only_workbook_references_its_theme(tmp_path):
    path = tmp_path / "image.xlsx"
    with XlsxWriter(path) as writer:
        writer.write_sheet([["x"]], images=[Image(_png_file(tmp_path), anchor="B2")])

    rels = part_text(path, "xl/_rels/workbook.xml.rels")
    assert 'Target="theme/theme1.xml"' in rels
    _check_package_consistency(path)
    openpyxl.load_workbook(path)


def test_chart_and_image_on_different_sheets_share_one_theme(tmp_path):
    path = tmp_path / "mixed.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Pictures")
        writer.write_sheet([["x"]], images=[Image(_png_file(tmp_path), anchor="B2")])
        writer.add_sheet("Chart")
        writer.write_sheet(
            [["k", "v"], ["a", 1], ["b", 2]],
            charts=[Chart("bar", [ChartSeries("B2:B3", categories="A2:A3")])],
        )

    assert part_text(path, "xl/_rels/workbook.xml.rels").count("theme/theme1.xml") == 1
    _check_package_consistency(path)


def test_workbook_without_drawings_has_no_theme(tmp_path):
    path = tmp_path / "plain.xlsx"
    with XlsxWriter(path) as writer:
        writer.write_sheet(DATA)

    with zipfile.ZipFile(path) as archive:
        assert "xl/theme/theme1.xml" not in archive.namelist()
    assert "theme" not in part_text(path, "xl/_rels/workbook.xml.rels")
    _check_package_consistency(path)
