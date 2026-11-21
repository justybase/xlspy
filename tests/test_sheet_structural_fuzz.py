"""Random structural edits keep both packages readable and self-consistent.

The fuzz is deterministic: a bad sequence can be replayed from the seed. After
every operation the updater is reopened and the file is loaded by the object
model, so a dangling sheet part, relationship or workbook entry fails here.
"""

from __future__ import annotations

import random

import pytest

from xlspy import XlsbUpdater, XlsbWriter, XlsxUpdater, XlsxWriter, load_workbook

FORMATS = {
    "xlsx": (XlsxWriter, XlsxUpdater),
    "xlsb": (XlsbWriter, XlsbUpdater),
}
SEED = 20261120


def _create(path, writer_class):
    with writer_class(path) as writer:
        for name in ("Alpha", "Beta", "Gamma"):
            writer.add_sheet(name)
            writer.write_sheet([[name, 1], [name, 2]])


@pytest.mark.parametrize("ext", sorted(FORMATS))
def test_random_sheet_structure_edits_stay_readable(tmp_path, ext):
    writer_class, updater_class = FORMATS[ext]
    path = tmp_path / f"structure.{ext}"
    _create(path, writer_class)
    rng = random.Random(SEED)

    for step in range(30):
        updater = updater_class(path)
        try:
            names = updater.get_sheet_names()
            action = rng.randrange(6)
            if action == 0:
                name = f"Added {step}"
                updater.add_sheet(name, rng.choice([None, 0, len(names)]))
            elif action == 1 and len(names) > 1:
                updater.remove_sheet(rng.choice(names))
            elif action == 2 and len(names) > 1:
                order = list(names)
                rng.shuffle(order)
                updater.reorder_sheets(order)
            elif action == 3:
                old = rng.choice(names)
                new = f"Renamed {step}"
                if new not in names:
                    updater.rename_sheet(old, new)
            elif action == 4:
                states = updater.get_sheet_states()
                visible = [name for name in names if states[name] == "visible"]
                candidate = rng.choice(names)
                if len(visible) > 1 or states[candidate] != "visible":
                    updater.set_sheet_state(
                        candidate, rng.choice(["hidden", "veryHidden", "visible"])
                    )
            elif action == 5:
                updater.set_active_sheet(rng.randrange(len(names)))
            updater.save(path)
        finally:
            updater.close()

        reopened = updater_class(path)
        try:
            current = reopened.get_sheet_names()
            assert len(current) == len(set(current))
            assert set(reopened.get_sheet_states()) == set(current)
            for part in reopened._sheets.values():
                assert reopened._package.has(part)
            assert reopened.get_sheet_names() == current
        finally:
            reopened.close()

        book = load_workbook(path)
        try:
            assert book.get_sheet_names() == current
        finally:
            book.close()
