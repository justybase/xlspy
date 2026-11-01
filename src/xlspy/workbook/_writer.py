"""Save the workbook model through the existing fast writers."""

from __future__ import annotations

from ..exceptions import UnsupportedFeatureError
import datetime as _datetime
import os
import re
import shutil
import tempfile

from ..utils.datetime import time_to_days, timedelta_to_days
from .._metadata import SheetProtection as MetadataProtection
from ..formulas import Formula
from ..rich_text import CellRichText
from ..styles import Style
from ..writer import XlsbWriter
from ..xlsx_writer import XlsxWriter
from ._loaded import save_loaded
from ._objects import _as_excel_table
from ._print import print_settings_for
from ._model import Workbook, Worksheet, WriteOnlyWorksheet



def _value_for_new(cell):
    if isinstance(cell._formula, Formula):
        return cell._formula
    if isinstance(cell._formula, str):
        return Formula(cell._formula)
    link = cell._hyperlink
    if link is not None:
        from ..hyperlinks import Hyperlink

        text = cell._value if isinstance(cell._value, str) else link.text
        return Hyperlink(link.target, text=text, tooltip=link.tooltip)
    value = cell._value
    if isinstance(value, _datetime.time):
        return time_to_days(value)
    if isinstance(value, _datetime.timedelta):
        return timedelta_to_days(value)
    return value


# An empty style: written as the default look instead of the bold header.
_PLAIN = Style()
# Values that _value_for_new returns unchanged (no formula or hyperlink).
_PLAIN_TYPES = frozenset({int, float, bool, str, _datetime.date, _datetime.datetime, type(None)})


def _cell_values(sheet: Worksheet):
    """Rows for the writer. A styled cell with a value is written as
    ``(value, Style)``; styled empty cells go to ``_styles``."""
    if isinstance(sheet, WriteOnlyWorksheet):
        for row in sheet._rows:
            yield list(row)
        return
    max_row = max(sheet.max_row, 1)
    # A row past the last cell still goes out when it has a height, a flag or an outline
    # level (openpyxl writes such rows too); its <row> carries the attributes.
    for number, dimension in sheet.row_dimensions.items():
        if number > max_row and (dimension.height or dimension.hidden or dimension.outlineLevel or dimension.collapsed):
            max_row = number
    style_at = sheet._workbook._styles.style_at
    styles: dict[int, Style] = {}
    plain_types = _PLAIN_TYPES
    cells = sheet._cells
    for row_number in range(1, max_row + 1):
        row_cells = cells.get(row_number)
        if not row_cells:
            # An empty row: the writers count rows by position, so it is still yielded.
            yield []
            continue
        # Each row stops at its last cell (styled empty cells are in
        # ``_cells`` too, so their columns are kept). A missing trailing
        # value is an empty cell, so the row never builds a list the width of
        # the whole sheet.
        row = [None] * max(row_cells)
        for column, cell in row_cells.items():
            if cell._formula is None and cell._hyperlink is None:
                value = cell._value
                if type(value) not in plain_types:
                    value = _value_for_new(cell)
            else:
                value = _value_for_new(cell)
            style_id = cell._style_id
            if style_id and value is not None:
                style = styles.get(style_id)
                if style is None:
                    style = styles[style_id] = style_at(style_id)
                value = (value, style)
            elif row_number == 1 and isinstance(value, (str, CellRichText)):
                # The writers make first-row text bold by default; a model
                # cell keeps the look it has, as in openpyxl.
                value = (value, _PLAIN)
            row[column - 1] = value
        yield row


def _styles(sheet: Worksheet):
    """Styles of empty cells, which ``_cell_values`` cannot carry."""
    if isinstance(sheet, WriteOnlyWorksheet):
        return None
    result = {}
    for row in sheet._cells.values():
        for cell in row.values():
            if cell._style_id and _value_for_new(cell) is None:
                result[cell.coordinate] = sheet._workbook._styles.style_at(cell._style_id)
    return result or None


def _comments(sheet: Worksheet):
    if isinstance(sheet, WriteOnlyWorksheet):
        return None
    result = {}
    for row in sheet._cells.values():
        for cell in row.values():
            if cell.comment is not None:
                result[cell.coordinate] = cell.comment
    return result or None


def _sheet_protection(sheet: Worksheet):
    protection = sheet.protection
    if not protection.sheet:
        return None
    password = protection.password
    if password is not None and not re.fullmatch(r"[0-9A-Fa-f]{1,4}", password):
        password = None
    # The model follows OOXML (True = protected); the writer takes "allowed".
    return MetadataProtection(
        enabled=True,
        select_locked_cells=not protection.selectLockedCells,
        select_unlocked_cells=not protection.selectUnlockedCells,
        hashed_password=password,
    )


def _common_options(sheet: Worksheet, *, xlsb: bool = False):
    widths = {}
    hidden_columns = []
    for letter in sheet.column_dimensions:
        dimension = sheet.column_dimensions[letter]
        index = 0
        for char in letter.upper():
            index = index * 26 + ord(char) - 64
        index -= 1
        if dimension.width is not None:
            widths[index] = dimension.width
        if dimension.hidden:
            hidden_columns.append(index)
    heights = {}
    hidden_rows = []
    for row, dimension in sheet.row_dimensions.items():
        if dimension.height is not None:
            heights[row - 1] = dimension.height
        if dimension.hidden:
            hidden_rows.append(row - 1)
    options = {
        "styles": _styles(sheet),
        "merges": [str(value) for value in sheet.merged_cells.ranges] or None,
        "column_widths": widths or None,
        "hidden_columns": hidden_columns or None,
        "row_heights": heights or None,
        "hidden_rows": hidden_rows or None,
        "freeze_panes": sheet.freeze_panes or "",
        "comments": _comments(sheet),
        "sheet_protection": _sheet_protection(sheet),
        "show_grid_lines": sheet.sheet_view.showGridLines is not False,
        "show_row_col_headers": sheet.sheet_view.showRowColHeaders is not False,
        "show_zero_values": sheet.sheet_view.showZeros is not False,
        "right_to_left": sheet.sheet_view.rightToLeft is True,
        "tab_color": sheet.sheet_properties.tabColor,
        "zoom": sheet.sheet_view.zoomScale,
        "charts": sheet.charts or None,
        "tables": [_as_excel_table(table) for table in sheet.tables] or None,
        "conditional_formats": list(sheet.conditional_formatting.all_rules()) or None,
    }
    print_settings = print_settings_for(sheet)
    for name, items in (
        ("images", sheet.images), ("data_validations", sheet.data_validations),
        ("print_settings", print_settings),
    ):
        if items:
            options[name] = items
    if sheet.ignored_errors:
        if xlsb:
            raise UnsupportedFeatureError("ignored errors can be saved to XLSX only")
        options["ignored_errors"] = list(sheet.ignored_errors)
    options["auto_filter"] = bool(sheet.auto_filter.ref)
    options["auto_filter_ref"] = sheet.auto_filter.ref or None
    return options


def _set_properties(writer, props) -> None:
    created = props.created
    if isinstance(created, _datetime.datetime) and created.tzinfo is None:
        # The model's default is datetime.now(): local time.
        created = created.astimezone(_datetime.timezone.utc)
    writer.set_properties(
        title=props.title, subject=props.subject, author=props.creator,
        keywords=props.keywords, comments=props.description,
        category=props.category, status=props.contentStatus,
        created=created if isinstance(created, _datetime.datetime) else None,
    )


def _save_new(workbook: Workbook, filename) -> None:
    extension = ""
    if isinstance(filename, (str, os.PathLike)):
        extension = os.fspath(filename).lower().rsplit(".", 1)[-1]
    from ._chartsheets import add_chartsheets, check_chartsheets, pending_chartsheets
    from ._package_kind import set_package_kind

    chartsheets = pending_chartsheets(workbook)
    if chartsheets:
        if extension == "xlsb":
            raise UnsupportedFeatureError("chartsheets can be saved to XLSX only")
        check_chartsheets(workbook)
    if chartsheets or extension in {"xlsm", "xltx", "xltm"}:
        # Written as XLSX, then the package is finished: chartsheets, the
        # macro-enabled or template content type and the VBA project.
        with tempfile.TemporaryDirectory(prefix="xlspy-workbook-") as directory:
            temporary = os.path.join(directory, "workbook.xlsx")
            _write_new(workbook, temporary, XlsxWriter, "xlsx")
            from ..xlsx_updater import XlsxUpdater

            updater = XlsxUpdater(temporary)
            try:
                add_chartsheets(updater, workbook)
                set_package_kind(updater, extension, workbook.vba_archive)
                if isinstance(filename, (str, os.PathLike)):
                    updater.save(os.fspath(filename))
                else:
                    filename.write(updater.to_bytes())
            finally:
                updater.close()
        return
    writer_class = XlsbWriter if extension == "xlsb" else XlsxWriter
    _write_new(workbook, filename, writer_class, extension)


def _write_new(workbook: Workbook, filename, writer_class, extension: str) -> None:
    with writer_class(filename) as writer:
        _set_properties(writer, workbook.properties)
        for definition in workbook.defined_names:
            writer.set_defined_name(definition)
        for sheet in workbook._sheets:
            writer.add_sheet(sheet.title, hidden=sheet.sheet_state != "visible")
            writer.write_sheet(
                _cell_values(sheet),
                **_common_options(sheet, xlsb=extension == "xlsb"),
            )
        writer.save()


def _save_loaded_to_path(workbook: Workbook, filename) -> None:
    if isinstance(filename, (str, os.PathLike)):
        save_loaded(workbook, os.fspath(filename))
        return
    with tempfile.TemporaryDirectory(prefix="xlspy-workbook-") as directory:
        temporary = os.path.join(directory, "workbook.xlsx")
        save_loaded(workbook, temporary)
        with open(temporary, "rb") as source:
            shutil.copyfileobj(source, filename)


def save_workbook(workbook: Workbook, filename) -> None:
    if not workbook._sheets:
        workbook.create_sheet("Sheet")
    if workbook._loaded:
        if getattr(workbook, "_format", "xlsx") == "xlsb":
            from ._xlsb_loaded import save_loaded_xlsb

            save_loaded_xlsb(workbook, filename)
        else:
            _save_loaded_to_path(workbook, filename)
    else:
        added = [name for name in workbook._styles.named_style_defs if name != "Normal"]
        if added:
            # The new-file writers build their own styles.xml and do not write
            # these styles; refuse instead of dropping them silently.
            raise UnsupportedFeatureError(
                f"named styles {added!r} can only be written to a loaded workbook: save the "
                "new workbook first, load it, call add_named_style() and save again"
            )
        _save_new(workbook, filename)
