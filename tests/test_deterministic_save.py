"""The same edits to a loaded workbook save to the same bytes in every process.

Pending edits are kept in sets of A1 references; string hashing differs between
processes (PYTHONHASHSEED), so saving must not depend on set order.
"""
from __future__ import annotations

import os
import pathlib
import subprocess
import sys
import zipfile

import pytest

SCRIPT = r'''
import sys
from xlspy import Comment, Font, Hyperlink, Style, XlsbWriter, XlsxWriter, load_workbook

source, target = sys.argv[1], sys.argv[2]
writer_class = XlsbWriter if source.endswith(".xlsb") else XlsxWriter
with writer_class(source) as writer:
    writer.add_sheet("Data")
    writer.write_sheet([[f"r{row}", row] for row in range(1, 30)])
workbook = load_workbook(source)
ws = workbook["Data"]
bold = Style(font=Font(bold=True))
for row in range(1, 30, 3):
    ws.cell(row, 2).value = row * 10
    ws.cell(row, 1).style = bold
    if source.endswith(".xlsx"):  # an XLSB comment record carries a fresh GUID
        ws.cell(row, 3).comment = Comment(f"note {row}", "me")
        ws.cell(row, 4).hyperlink = Hyperlink(f"https://example.com/{row}")
workbook.save(target)
workbook.close()
'''


def _save_in_process(tmp_path: pathlib.Path, suffix: str, seed: int) -> dict[str, bytes]:
    script = tmp_path / "edit.py"
    script.write_text(SCRIPT, encoding="utf-8")
    source = tmp_path / f"source{seed}{suffix}"
    target = tmp_path / f"saved{seed}{suffix}"
    env = dict(os.environ, PYTHONHASHSEED=str(seed))
    subprocess.run([sys.executable, str(script), str(source), str(target)], check=True, env=env)
    with zipfile.ZipFile(target) as archive:
        return {name: archive.read(name) for name in archive.namelist() if name != "docProps/core.xml"}


@pytest.mark.parametrize("suffix", [".xlsx", ".xlsb"])
def test_saving_edits_does_not_depend_on_hash_seed(tmp_path: pathlib.Path, suffix: str) -> None:
    first = _save_in_process(tmp_path, suffix, 1)
    for seed in (2, 3):
        assert _save_in_process(tmp_path, suffix, seed) == first
