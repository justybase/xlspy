"""Save edits back into a loaded XLSB package."""

from __future__ import annotations

import os
import shutil
import tempfile

from ..refs import parse_cell_range
from ._loaded import _meta_signature, _value_for_save
from ._objects import _object_values, object_signatures
from ._print import print_settings_for

# Index of the merged ranges in _meta_signature(): merges are saved on their own.
_MERGES = 4


def _layout_signature(sheet):
    """Sheet metadata other than merges: views, dimensions, print, protection."""
    signature = _meta_signature(sheet)
    # Reading row_dimensions[n] or column_dimensions["X"] creates an entry with
    # default values; only entries that set something count as edits.
    columns, rows = (
        tuple(entry for entry in entries if any(entry[1:]))
        for entries in signature[:2]
    )
    return (columns, rows) + signature[2:_MERGES] + signature[_MERGES + 1:]


def _merge_signature(sheet):
    return frozenset(str(value) for value in sheet.merged_cells.ranges)


def remember_saved_state(sheet) -> None:
    """Record what the file holds, so that save can tell which parts were edited."""
    sheet._meta_signature = _layout_signature(sheet)
    sheet._saved_merges = _merge_signature(sheet)
    sheet._saved_object_counts = {
        name: len(_object_values(sheet, name))
        for name in ("charts", "images", "tables", "data_validations", "conditional_formatting")
    }
    sheet._saved_object_signatures = object_signatures(sheet)
    sheet._saved_print_settings = print_settings_for(sheet)


def _properties_signature(workbook):
    return repr(workbook.properties), repr(workbook.custom_doc_props)


def remember_workbook_state(workbook) -> None:
    workbook._saved_properties = _properties_signature(workbook)


def _unsupported(what: str):
    return NotImplementedError(
        f"{what} cannot be saved into a loaded XLSB workbook yet; "
        "build the workbook with Workbook() and save it as .xlsb instead"
    )


def _reject_unsupported_sheet_edits(sheet) -> None:
    if _layout_signature(sheet) != sheet._meta_signature:
        raise _unsupported(
            "changes to column/row dimensions, sheet views, freeze panes, "
            "auto-filter, print settings or protection"
        )
    if print_settings_for(sheet) != sheet._saved_print_settings:
        raise _unsupported("print setting changes")
    counts = sheet._saved_object_counts
    for name in ("charts", "images"):
        if len(_object_values(sheet, name)) != counts.get(name, 0):
            raise _unsupported(f"adding or removing {name}")
    saved = sheet._saved_object_signatures
    current = object_signatures(sheet)
    for name, label in (
        ("tables", "table"), ("conditional_formatting", "conditional formatting"),
    ):
        if current[name] != saved.get(name, ()):
            raise _unsupported(f"{label} changes")


def _reject_unsupported_edits(workbook) -> None:
    """Fail loudly on edits the XLSB updater cannot persist yet."""
    sheets = workbook._sheets
    original = [sheet._original_title for sheet in sheets]
    if (
        any(sheet.title != sheet._original_title for sheet in sheets)
        or set(workbook.sheetnames) != workbook._loaded_worksheet_names
        or len(set(original)) != len(original)
        or original != list(workbook._updater.get_sheet_names())
    ):
        raise NotImplementedError(
            "adding, removing, renaming or reordering sheets is not supported "
            "when saving a loaded XLSB workbook"
        )
    styles = workbook._styles
    if workbook._styles_dirty and not hasattr(styles, "to_bin"):
        raise NotImplementedError(
            "style changes are not supported when saving a loaded XLSB workbook"
        )
    if styles.is_modified() and not getattr(workbook, "_xlsb_styles_path", None):
        raise _unsupported("style changes (the file has no style sheet)")
    if _properties_signature(workbook) != getattr(workbook, "_saved_properties", None):
        raise _unsupported("document property changes")
    for sheet in sheets:
        _reject_unsupported_sheet_edits(sheet)


def save_loaded_xlsb(workbook, filename) -> None:
    updater = workbook._updater
    if updater is None:
        raise ValueError("this workbook was not loaded from an XLSB package")

    _reject_unsupported_edits(workbook)

    names = {
        (item.name, item.formula, item.scope) for item in workbook.defined_names
    }
    previous = getattr(workbook, "_loaded_defined_names", names)
    if names != previous:
        for name, _formula, scope in previous - names:
            updater.remove_defined_name(name, scope)
        for item in workbook.defined_names:
            updater.set_defined_name(item)
        workbook._loaded_defined_names = names

    styles = workbook._styles
    if styles.is_modified():
        updater._package.set(workbook._xlsb_styles_path, styles.to_bin())

    for sheet in workbook._sheets:
        updates = {}
        style_ids = {}
        for row, column in list(sheet._dirty_cells):
            cell = sheet._existing_cell(row, column)
            if cell is not None:
                updates[cell.coordinate] = _value_for_save(
                    cell, updater._date_1904
                )
                style_ids[cell.coordinate] = cell._style_id
        for ref in sheet._dirty_styles:
            col, row, _, _ = parse_cell_range(ref)
            cell = sheet._existing_cell(row + 1, col + 1)
            if cell is not None:
                style_ids[ref] = cell._style_id
        if _merge_signature(sheet) != sheet._saved_merges:
            updater.set_merged_cells(sheet.title, sorted(str(item) for item in sheet._merged_cells.ranges))
        if updates or style_ids:
            updater.update_cells(sheet.title, updates, styles=style_ids)
        if (object_signatures(sheet)["data_validations"]
                != sheet._saved_object_signatures.get("data_validations", ())):
            updater.set_data_validations(sheet.title, sheet.data_validations)
        if sheet._dirty_hyperlinks:
            updater.set_hyperlinks(sheet.title, {
                cell.coordinate: cell._hyperlink
                for row in sheet._cells.values() for cell in row.values()
                if cell._hyperlink is not None
            })
        if sheet._dirty_comments:
            comments = {}
            for ref in sheet._dirty_comments:
                col, row, _, _ = parse_cell_range(ref)
                cell = sheet._existing_cell(row + 1, col + 1)
                comments[ref] = cell.comment if cell is not None else None
            updater.set_comments(sheet.title, comments)

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
