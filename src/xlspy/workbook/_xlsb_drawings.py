"""Add charts and images to a sheet of a loaded XLSB package.

DrawingML parts are the same XML in XLSB and XLSX, so the parts come from the
XLSX writer's scratch workbook and are grafted with the XLSX code; only the
sheet itself differs: it points at its drawing with a ``BrtDrawing`` record.
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING, cast

from .._xlsb_metadata import replace_records, wide_string
from ..biff_utils import build_record
from ._drawings import new_drawing_objects
from .._xlsx_edit import _bytes
from ._objects import _Sheet, _add_drawing, _Package, _rels_path, _scratch, _sheet_rels
from ._xlsb_layout import _AFTER_PRINT

if TYPE_CHECKING:
    from ..xlsb_updater import XlsbUpdater
    from ._worksheet import Worksheet

BRT_DRAWING = 0x226


class _DrawingHook:
    """Stands in for the XLSX sheet: records the relationship id of a new drawing."""

    rid: str | None = None

    def insert(self, _name: str, fragment: str) -> None:
        match = re.search(r'r:id="([^"]+)"', fragment)
        self.rid = match[1] if match else None


def add_drawing_objects(updater: XlsbUpdater, sheet: Worksheet) -> None:
    added = {name: items for name, items in new_drawing_objects(sheet).items() if items}
    if not added:
        return
    objects = {"charts": [], "images": [], "tables": [], "data_validations": [],
               "conditional_formatting": [], **added}
    scratch = _scratch(sheet, objects, None)
    package = _Package(updater)
    path = updater._sheets[sheet.title]
    rels = package.rels(path)
    hook = _DrawingHook()
    host = cast(_Sheet, hook)  # only ``insert`` is used for a new drawing
    _add_drawing(package, host, path, rels, scratch, _sheet_rels(scratch))
    package.set(_rels_path(path), _bytes(rels))
    if hook.rid is not None:
        data = package.get(path)
        package.set(path, replace_records(
            data, {BRT_DRAWING}, build_record(BRT_DRAWING, wide_string(hook.rid)),
            before=tuple(item for item in _AFTER_PRINT if item != BRT_DRAWING),
        ))
    package.commit()
