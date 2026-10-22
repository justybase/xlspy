"""Save the workbook model through the existing fast writers."""

from __future__ import annotations

import datetime as _datetime
import io
import os
import shutil
import tempfile

from ..formulas import Formula
from ..rich_text import CellRichText
from ..styles import Style
from ..refs import column_label
from ..writer import XlsbWriter
from ..xlsx_writer import XlsxWriter
from ._loaded import save_loaded
from ._print import print_settings_for
from ._model import Workbook, Worksheet, WriteOnlyWorksheet, _coordinate


def _serial_time(value: _datetime.time) -> float:
    return (
        value.hour * 3600 + value.minute * 60 + value.second
        + value.microsecond / 1_000_000
    ) / 86400.0


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
        return _serial_time(value)
    if isinstance(value, _datetime.timedelta):
        return value.total_seconds() / 86400.0
    return value


# An empty style: written as the default look instead of the bold header.
_PLAIN = Style()


def _cell_values(sheet: Worksheet):
    """Rows for the writer. A styled cell with a value is written as
    ``(value, Style)``; styled empty cells go to ``_styles``."""
    if isinstance(sheet, WriteOnlyWorksheet):
        for row in sheet._rows:
            yield list(row)
        return
    max_row = max(sheet.max_row, 1)
    max_column = max(sheet.max_column, 1)
    style_at = sheet._workbook._styles.style_at
    for row_number in range(1, max_row + 1):
        row = []
        for column in range(1, max_column + 1):
            cell = sheet._existing_cell(row_number, column)
            if cell is None:
                row.append(None)
                continue
            value = _value_for_new(cell)
            if cell._style_id and value is not None:
                value = (value, style_at(cell._style_id))
            elif row_number == 1 and isinstance(value, (str, CellRichText)):
                # The writers make first-row text bold by default; a model
                # cell keeps the look it has, as in openpyxl.
                value = (value, _PLAIN)
            row.append(value)
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
        "sheet_protection": sheet._sheet_protection,
        "show_grid_lines": sheet.sheet_view.showGridLines is not False,
        "tab_color": sheet.sheet_properties.tabColor,
        "zoom": sheet.sheet_view.zoomScale,
        "charts": sheet.charts or None,
        "tables": sheet.tables or None,
        "conditional_formats": list(sheet.conditional_formatting) or None,
    }
    print_settings = print_settings_for(sheet)
    for name, items in (
        ("images", sheet.images), ("data_validations", sheet.data_validations),
        ("print_settings", print_settings),
    ):
        if items:
            options[name] = items
    if xlsb:
        options["auto_filter"] = bool(sheet.auto_filter.ref)
    else:
        options["auto_filter"] = bool(sheet.auto_filter.ref)
        options["auto_filter_ref"] = sheet.auto_filter.ref
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
    if extension == "xlsm":
        raise ValueError(
            "saving a new macro-enabled workbook is not supported; "
            "load an existing .xlsm to keep its VBA project"
        )
    if extension in {"xltx", "xltm"}:
        raise ValueError("template formats are not supported by the new workbook API")
    writer_class = XlsbWriter if extension == "xlsb" else XlsxWriter
    with writer_class(filename) as writer:
        _set_properties(writer, workbook.properties)
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
        _save_loaded_to_path(workbook, filename)
    else:
        _save_new(workbook, filename)
