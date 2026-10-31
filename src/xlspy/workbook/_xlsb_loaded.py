"""Save edits back into a loaded XLSB package."""

from __future__ import annotations

import os
import shutil
import tempfile

from ..refs import parse_cell_range
from . import _props
from ._loaded import _value_for_save
from ._drawings import check_removals, new_drawing_objects, remember_drawing_objects, remove_drawing_objects, save_chart_edits
from ._objects import _object_values, check_unread, object_signatures
from ._xlsb_layout import changed_groups, layout_state, patch_layout, save_builtin_names


def _merge_signature(sheet):
    return frozenset(str(value) for value in sheet.merged_cells.ranges)


def remember_saved_state(sheet) -> None:
    """Record what the file holds, so that save can tell which parts were edited."""
    sheet._saved_merges = _merge_signature(sheet)
    sheet._saved_object_counts = {
        name: len(_object_values(sheet, name))
        for name in ("charts", "images", "tables", "data_validations", "conditional_formatting")
    }
    sheet._saved_object_signatures = object_signatures(sheet)
    sheet._saved_layout = layout_state(sheet)
    remember_drawing_objects(sheet)


def remember_workbook_state(workbook) -> None:
    _props.remember(workbook)


def _unsupported(what: str):
    return NotImplementedError(
        f"{what} cannot be saved into a loaded XLSB workbook yet; "
        "build the workbook with Workbook() and save it as .xlsb instead"
    )


def _is_new(sheet) -> bool:
    return sheet._path is None


def _reject_unsupported_sheet_edits(sheet) -> None:
    if _is_new(sheet):
        for name in ("charts", "images", "tables"):
            if _object_values(sheet, name):
                raise _unsupported(f"{name.replace('_', ' ')} on a new sheet")
        return
    for name, added in new_drawing_objects(sheet).items():
        if added:
            raise _unsupported(f"adding {name}")
    check_removals(sheet)
    saved = sheet._saved_object_signatures
    current = object_signatures(sheet)
    if current["tables"] != saved.get("tables", ()):
        raise _unsupported("table changes")
    check_unread(sheet, [
        name for name in ("data_validations", "conditional_formatting")
        if current[name] != saved.get(name, ())
    ])


def _reject_unsupported_edits(workbook) -> None:
    """Fail loudly, before anything is written, on edits that cannot be saved yet."""
    styles = workbook._styles
    if workbook._styles_dirty and not hasattr(styles, "to_bin"):
        raise NotImplementedError(
            "style changes are not supported when saving a loaded XLSB workbook"
        )
    if styles.is_modified() and not getattr(workbook, "_xlsb_styles_path", None):
        raise _unsupported("style changes (the file has no style sheet)")
    for sheet in workbook._sheets:
        _reject_unsupported_sheet_edits(sheet)
    if any(sheet._part is None for sheet in workbook._chartsheets):
        raise _unsupported("a new chartsheet")


def _sync_sheet_list(workbook, updater) -> None:
    """Remove, rename, add and reorder the package's sheets to match the model."""
    sheets = workbook._sheets
    loaded = [sheet for sheet in sheets if not _is_new(sheet)]
    kept = {sheet._original_title for sheet in loaded}
    for name in updater.get_sheet_names():
        if name in workbook._loaded_worksheet_names and name not in kept:
            updater.remove_sheet(name)
    renamed = [sheet for sheet in loaded if sheet.title != sheet._original_title]
    # Through temporary names, so that sheets can swap titles.
    taken = set(updater.get_sheet_names()) | {sheet.title for sheet in sheets}
    temporary = {}
    for sheet in renamed:
        number = len(temporary)
        while f"~xlspy{number}" in taken:
            number += 1
        name = f"~xlspy{number}"
        taken.add(name)
        temporary[sheet] = name
        updater.rename_sheet(sheet._original_title, name)
    for sheet, name in temporary.items():
        updater.rename_sheet(name, sheet.title)
    for sheet in sheets:
        if _is_new(sheet):
            sheet._path = updater.add_sheet(sheet.title)
    titles = [sheet.title for sheet in sheets]
    current = updater.get_sheet_names()
    # Sheets the model does not hold (if any) keep their positions.
    order = list(titles)
    for index, name in enumerate(current):
        if name not in titles:
            order.insert(min(index, len(order)), name)
    updater.reorder_sheets(order)
    states = updater.get_sheet_states()
    for sheet in sheets:
        if states.get(sheet.title) != sheet.sheet_state:
            updater.set_sheet_state(sheet.title, sheet.sheet_state)
    active = order.index(titles[min(workbook._active_index, len(titles) - 1)])
    if active != updater.get_active_sheet():
        updater.set_active_sheet(active)


def _save_defined_names(workbook, updater, renamed) -> None:
    names = {(item.name, item.formula, item.scope) for item in workbook.defined_names}
    previous = getattr(workbook, "_loaded_defined_names", names)
    if names == previous:
        return
    for name, _formula, scope in previous - names:
        scope = renamed.get(scope, scope)
        # Names local to a removed sheet went with it.
        if scope is None or scope in updater._sheets:
            updater.remove_defined_name(name, scope)
    for item in workbook.defined_names:
        if item.scope is None or item.scope in updater._sheets:
            updater.set_defined_name(item)
    workbook._loaded_defined_names = names


def _save_cells(sheet, updater, everything: bool) -> None:
    updates = {}
    style_ids = {}
    if everything:
        cells = [cell for row in sheet._cells.values() for cell in row.values()]
    else:
        cells = [
            cell for cell in (sheet._existing_cell(row, column) for row, column in list(sheet._dirty_cells))
            if cell is not None
        ]
    for cell in cells:
        value = _value_for_save(cell, updater._date_1904)
        if everything and value is None and not cell._style_id:
            continue
        updates[cell.coordinate] = value
        style_ids[cell.coordinate] = cell._style_id
    if not everything:
        for ref in sheet._dirty_styles:
            col, row, _, _ = parse_cell_range(ref)
            cell = sheet._existing_cell(row + 1, col + 1)
            if cell is not None:
                style_ids[ref] = cell._style_id
    if updates or style_ids:
        updater.update_cells(sheet.title, updates, styles=style_ids)


def _save_sheet(sheet, updater) -> None:
    new = getattr(sheet, "_saved_layout", None) is None
    _save_cells(sheet, updater, new)
    groups = changed_groups(sheet)
    if groups - {"print_names"}:
        path, data = updater._sheet_bin(sheet.title)
        patched = patch_layout(data, sheet, groups)
        if patched != data:
            updater._apply_parts({path: patched})
    save_builtin_names(updater, sheet, groups)
    merges = sheet._merged_cells.ranges
    if merges if new else _merge_signature(sheet) != sheet._saved_merges:
        updater.set_merged_cells(sheet.title, sorted(str(item) for item in merges))
    validations_changed = (
        bool(sheet.data_validations) if new else
        object_signatures(sheet)["data_validations"]
        != sheet._saved_object_signatures.get("data_validations", ())
    )
    if validations_changed:
        updater.set_data_validations(sheet.title, sheet.data_validations)
    formats_changed = (
        bool(sheet.conditional_formatting) if new else
        object_signatures(sheet)["conditional_formatting"]
        != sheet._saved_object_signatures.get("conditional_formatting", ())
    )
    if formats_changed:
        updater.set_conditional_formats(
            sheet.title,
            [(item.sqref, item.pivot, item.rules) for item in sheet.conditional_formatting],
            sheet._workbook._styles.dxf_id,
        )
    if new or sheet._dirty_hyperlinks:
        links = {
            cell.coordinate: cell._hyperlink
            for row in sheet._cells.values() for cell in row.values()
            if cell._hyperlink is not None
        }
        if links or not new:
            updater.set_hyperlinks(sheet.title, links)
    if new:
        comments = {
            cell.coordinate: cell._comment
            for row in sheet._cells.values() for cell in row.values()
            if cell._comment is not None
        }
    else:
        comments = {}
        for ref in sheet._dirty_comments:
            col, row, _, _ = parse_cell_range(ref)
            cell = sheet._existing_cell(row + 1, col + 1)
            comments[ref] = cell.comment if cell is not None else None
    if comments:
        updater.set_comments(sheet.title, comments)


def save_loaded_xlsb(workbook, filename) -> None:
    updater = workbook._updater
    if updater is None:
        raise ValueError("this workbook was not loaded from an XLSB package")

    _reject_unsupported_edits(workbook)

    renamed = {sheet._original_title: sheet.title for sheet in workbook._sheets if not _is_new(sheet)}
    _sync_sheet_list(workbook, updater)
    _save_defined_names(workbook, updater, renamed)
    for sheet in workbook._sheets:
        _save_sheet(sheet, updater)
        remove_drawing_objects(updater, sheet)
        save_chart_edits(updater, sheet)
    _props.save_properties(updater, workbook)
    # After the sheets: their conditional formats may add differential styles.
    styles = workbook._styles
    if styles.is_modified():
        updater._package.set(workbook._xlsb_styles_path, styles.to_bin())

    if isinstance(filename, (str, os.PathLike)):
        updater.save(os.fspath(filename))
    else:
        with tempfile.TemporaryDirectory(prefix="xlspy-workbook-") as directory:
            temporary = os.path.join(directory, "workbook.xlsb")
            updater.save(temporary)
            with open(temporary, "rb") as source:
                shutil.copyfileobj(source, filename)

    styles.mark_saved()
    workbook._styles_dirty = False
    for sheet in workbook._sheets:
        sheet._original_title = sheet.title
        sheet._clear_dirty()
        remember_saved_state(sheet)
    workbook._loaded_worksheet_names = set(workbook.sheetnames)
    workbook._loaded_active_index = workbook._active_index
