"""The Workbook model."""

from __future__ import annotations
from ..exceptions import UnsupportedFeatureError
import copy
import zipfile
from dataclasses import replace
import datetime as _datetime
from typing import TYPE_CHECKING, Any
from ..formulas import Formula
from .._metadata import DefinedName
from ..refs import check_sheet_name
from ..styles import (
    NamedStyle,
    NamedStyleList,
)
from ._styles import XlsxStyleTable
from ._coords import MultiCellRange
from ._properties import BookView, CalcProperties, CustomPropertyList, DefinedNameList, DocumentProperties, WorkbookSecurity
from ._worksheet import Chartsheet, Worksheet, WriteOnlyWorksheet

if TYPE_CHECKING:
    from .._output import WriteTarget


class Workbook:
    # Set by the loaders and the XLSB saver; a workbook built in memory has none of them.
    _updater: Any
    _xlsb_styles_path: Any
    _loaded_defined_names: Any

    def __init__(self, write_only: bool = False, iso_dates: bool = False) -> None:
        self.write_only = bool(write_only)
        self.iso_dates = bool(iso_dates)
        self._sheets: list[Worksheet] = []
        self._chartsheets: list[Chartsheet] = []
        self._active_index = 0
        self._styles = XlsxStyleTable.empty()
        self._styles_dirty = False
        self._properties = DocumentProperties()
        self._defined_names = DefinedNameList(self)
        self._defined_names_dirty = False
        self.template = False
        self.epoch = _datetime.datetime(1899, 12, 30)
        self.vba_archive: zipfile.ZipFile | None = None
        self.security = WorkbookSecurity()
        self.calculation = CalcProperties()
        self.views = [BookView()]
        self.custom_doc_props = CustomPropertyList()
        self.external_links: list = []
        self.named_styles = NamedStyleList(["Normal"])
        self._format = "xlsx"
        self._loaded = False
        self._loaded_active_index = 0
        self._loaded_workbook_meta = None
        self._read_only = False
        self._updater = None
        self._source_path: str | None = None
        self._sheet_visibility: dict[str, str] = {}
        self._loaded_worksheet_names: set[str] = set()
        self._mark_loaded = False
        if not self.write_only:
            self.create_sheet("Sheet")

    def __repr__(self) -> str:
        return f"<Workbook {self.sheetnames!r}>"

    @property
    def sheetnames(self) -> list[str]:
        return [sheet.title for sheet in self._sheets]

    def _check_title(self, title: str, sheet: Worksheet | Chartsheet | None = None) -> None:
        """Raise unless ``title`` can name ``sheet`` (default: a new sheet).

        Excel's naming rules apply, and no other worksheet or chartsheet may
        have the name in any letter case.
        """
        check_sheet_name(title)
        folded = title.casefold()
        for other in self._all_sheets():
            if other is not sheet and other.title.casefold() == folded:
                raise ValueError(f"a sheet named {other.title!r} already exists")

    @property
    def read_only(self) -> bool:
        return self._read_only

    @read_only.setter
    def read_only(self, value: bool) -> None:
        self._read_only = bool(value)

    @property
    def worksheets(self) -> list[Worksheet]:
        return list(self._sheets)

    @property
    def chartsheets(self) -> list[Chartsheet]:
        return list(self._chartsheets)

    @property
    def active(self) -> Worksheet:
        if not self._sheets:
            self.create_sheet()
        return self._sheets[self._active_index]

    @active.setter
    def active(self, value: int | Worksheet) -> None:
        if isinstance(value, int) and not isinstance(value, bool):
            self._active_index = value
        elif value in self._sheets:
            self._active_index = self._sheets.index(value)
        else:
            raise ValueError("active must be a sheet index or Worksheet")

    def __getitem__(self, key: int | str) -> Worksheet:
        if isinstance(key, int) and not isinstance(key, bool):
            return self._sheets[key]
        for sheet in self._sheets:
            if sheet.title == key:
                return sheet
        raise KeyError(f"Worksheet {key!r} does not exist")

    def get_sheet_by_name(self, name: str) -> Worksheet:
        return self[name]

    def index(self, worksheet: Worksheet) -> int:
        return self._sheets.index(worksheet)

    def create_sheet(self, title: str | None = None, index: int | None = None) -> Worksheet:
        if len(self._sheets) >= 1 and title is None:
            title = f"Sheet{len(self._sheets) + 1}"
        elif title is None:
            title = "Sheet"
        self._check_title(title)
        cls =WriteOnlyWorksheet if self.write_only else Worksheet
        sheet = cls(self, title, index=index)
        if index is None:
            self._sheets.append(sheet)
        else:
            self._sheets.insert(index, sheet)
            for position, item in enumerate(self._sheets):
                item._index = position
        sheet._loaded = self._loaded
        return sheet

    def _all_sheets(self) -> tuple[Worksheet | Chartsheet, ...]:
        """Worksheets, then chartsheets: every sheet name in the workbook."""
        return (*self._sheets, *self._chartsheets)

    def create_chartsheet(self, title: str | None = None, index: int | None = None) -> Chartsheet:
        """Add a chartsheet after the other sheets (XLSX only), as openpyxl does."""
        if self.write_only:
            raise TypeError("create_chartsheet is not available in write-only mode")
        if title is None:
            taken = {sheet.title.casefold() for sheet in self._all_sheets()}
            number = 1
            while f"chart{number}" in taken:
                number += 1
            title = f"Chart{number}"
        self._check_title(title)
        count = len(self._sheets) + len(self._chartsheets)
        if index is not None and index < count:
            raise UnsupportedFeatureError("new chartsheets are added after the other sheets")
        sheet = Chartsheet(self, title, index=count)
        self._chartsheets.append(sheet)
        return sheet

    def move_sheet(self, sheet: Worksheet | str, offset: int = 0) -> None:
        """Move a sheet ``offset`` positions (negative: towards the front), as openpyxl does."""
        if isinstance(sheet, str):
            sheet = self[sheet]
        if sheet not in self._sheets:
            raise ValueError("worksheet is not part of this workbook")
        active = self._sheets[self._active_index] if self._sheets else None
        index = self._sheets.index(sheet)
        self._sheets.remove(sheet)
        self._sheets.insert(max(0, min(index + offset, len(self._sheets))), sheet)
        for position, item in enumerate(self._sheets):
            item._index = position
        if active is not None:
            self._active_index = self._sheets.index(active)

    def remove(self, worksheet: Worksheet) -> None:
        if worksheet not in self._sheets:
            raise ValueError("worksheet is not part of this workbook")
        self._sheets.remove(worksheet)
        if self._active_index >= len(self._sheets):
            self._active_index = max(0, len(self._sheets) - 1)
        if self._loaded:
            self._mark_loaded = True

    def copy_worksheet(self, from_worksheet: Worksheet) -> Worksheet:
        if self.write_only:
            raise TypeError("copy_worksheet is not available in write-only mode")
        title = from_worksheet.title
        base = title
        counter = 1
        taken = {sheet.title.casefold() for sheet in self._all_sheets()}
        while title.casefold() in taken:
            title = f"{base}{counter}"
            counter += 1
        target = self.create_sheet(title)
        for row_number, row in from_worksheet._cells.items():
            for column, cell in row.items():
                new_cell = target.cell(row_number, column, cell.value)
                new_cell._style_id = cell._style_id
                # Own copies, as openpyxl makes: editing one sheet leaves the other alone.
                new_cell._comment = copy.copy(cell.comment)
                new_cell._hyperlink = copy.copy(cell.hyperlink)
                new_cell._named_style = cell._named_style
        target._merged_cells = MultiCellRange(from_worksheet._merged_cells.ranges)
        for name in ("_column_dimensions", "_row_dimensions"):
            getattr(target, name).update(copy.deepcopy(dict(getattr(from_worksheet, name))))
        target._row_heights = dict(from_worksheet._row_heights)
        target._hidden_rows = set(from_worksheet._hidden_rows)
        target._freeze_panes = from_worksheet.freeze_panes
        target._auto_filter = copy.deepcopy(from_worksheet.auto_filter)
        target._sheet_state = from_worksheet.sheet_state
        target._sheet_properties = copy.deepcopy(from_worksheet.sheet_properties)
        target._sheet_view = copy.deepcopy(from_worksheet.sheet_view)
        for name in ("page_setup", "page_margins", "print_options", "HeaderFooter",
                     "sheet_format", "row_breaks", "col_breaks", "ignored_errors"):
            setattr(target, name, copy.deepcopy(getattr(from_worksheet, name)))
        target._mark_dirty("copy")
        return target

    @property
    def properties(self) -> DocumentProperties:
        return self._properties

    @properties.setter
    def properties(self, value: DocumentProperties) -> None:
        if not isinstance(value, DocumentProperties):
            raise TypeError("properties must be DocumentProperties")
        self._properties = value

    def add_named_style(self, style: NamedStyle) -> None:
        if not isinstance(style, NamedStyle):
            raise TypeError("style must be a NamedStyle")
        self._styles.add_named_style(style)
        self.named_styles.add(style)
        self._styles_dirty = True

    @property
    def defined_names(self) -> DefinedNameList:
        if not isinstance(self._defined_names, DefinedNameList):
            self._defined_names = DefinedNameList(self, self._defined_names)
        return self._defined_names

    def _rename_sheet(self, sheet: Worksheet, title: str) -> None:
        """Point formulas, validations, conditional formats, internal links and
        defined names at the new title, as Excel does (loaded charts: on save)."""
        from ..formula.sheets import rename_sheet_references

        old = sheet.title

        def renamed(text: Any) -> Any:
            if not isinstance(text, str) or "!" not in text:
                return text
            return rename_sheet_references(text, old, title)

        for other in self._sheets:
            for row in other._cells.values():
                for cell in row.values():
                    link = cell._hyperlink
                    if link is not None and link.is_internal:
                        target = "#" + renamed(link.location)
                        if target != link.target:
                            cell.hyperlink = replace(link, target=target)
                    formula = cell._formula
                    if formula is None:
                        continue
                    text = formula.text if isinstance(formula, Formula) else formula
                    updated = renamed(text)
                    if updated != text:
                        cell.value = (
                            replace(formula, text=updated) if isinstance(formula, Formula)
                            else "=" + updated
                        )
            for index, rule in enumerate(other._data_validations):
                values = {name: renamed(getattr(rule, name)) for name in ("formula1", "formula2")}
                if any(values[name] != getattr(rule, name) for name in values):
                    other._data_validations[index] = replace(rule, **values)
            for item in other.conditional_formatting:
                for index, rule in enumerate(item.rules):
                    formulas = getattr(rule, "formula", None)
                    if formulas:
                        updated = tuple(renamed(text) for text in formulas)
                        if updated != tuple(formulas):
                            item.rules[index] = replace(rule, formula=updated)
        names = self.defined_names
        names._items = [
            DefinedName(
                item.name, renamed(item.formula), title if item.scope == old else item.scope,
                item.localSheetId, item.hidden, item.comment,
            ) if renamed(item.formula) != item.formula or item.scope == old else item
            for item in names._items
        ]
        if self._loaded:
            self._mark_loaded = True

    def save(self, filename: WriteTarget) -> None:
        if self._read_only:
            raise ValueError("a read-only workbook cannot be saved")
        from ._writer import save_workbook
        save_workbook(self, filename)

    def close(self) -> None:
        if self._updater is not None:
            self._updater.close()
            self._updater = None
        for sheet in self._sheets:
            reader = getattr(sheet, "_reader", None)
            if reader is not None:
                reader.close()
                sheet._reader = None
