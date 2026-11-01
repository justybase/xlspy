"""Workbook and sheet property objects: views, calculation, protection, filters, breaks."""

from __future__ import annotations
from dataclasses import dataclass, field
import datetime as _datetime
from ..printing import PageMargins as PublicPageMargins
from .._metadata import DefinedName
from .._metadata import SheetProtection as PublicSheetProtection
from ..styles import (
    Color,
)
from ._print import (
    PageMargins,
    PageSetupProperties,
)


@dataclass
class WorkbookSecurity:
    workbookPassword: str | None = None
    workbookPasswordCharacterSet: str | None = None
    revisionsPassword: str | None = None
    revisionsPasswordCharacterSet: str | None = None
    lockStructure: bool | None = None
    lockWindows: bool | None = None
    lockRevision: bool | None = None
    revisionsAlgorithmName: str | None = None
    revisionsHashValue: str | None = None
    revisionsSaltValue: str | None = None
    revisionsSpinCount: int | None = None
    workbookAlgorithmName: str | None = None
    workbookHashValue: str | None = None
    workbookSaltValue: str | None = None
    workbookSpinCount: int | None = None


@dataclass
class CalcProperties:
    calcId: int = 124519
    calcMode: str | None = None
    fullCalcOnLoad: bool = True
    refMode: str | None = None
    iterate: bool | None = None
    iterateCount: int | None = None
    iterateDelta: float | None = None
    fullPrecision: bool | None = None
    calcCompleted: bool | None = None
    calcOnSave: bool | None = None
    concurrentCalc: bool | None = None
    concurrentManualCount: int | None = None
    forceFullCalc: bool | None = None


@dataclass
class BookView:
    visibility: str = "visible"
    minimized: bool = False
    showHorizontalScroll: bool = True
    showVerticalScroll: bool = True
    showSheetTabs: bool = True
    xWindow: int | None = None
    yWindow: int | None = None
    windowWidth: int | None = None
    windowHeight: int | None = None
    tabRatio: int = 600
    firstSheet: int = 0
    activeTab: int = 0
    autoFilterDateGrouping: bool = True


class CustomPropertyList(list):
    """A list of custom document properties, keyed by ``name`` for convenience."""

    @property
    def names(self):
        return [item.name for item in self]


class DefinedNameList:
    """A list/dict hybrid of :class:`DefinedName` entries."""

    def __init__(self, workbook=None, definitions=None):
        self._workbook = workbook
        self._items: list[DefinedName] = []
        for definition in definitions or []:
            self.add(definition, stage=False)

    def __iter__(self):
        return iter(self._items)

    def __len__(self):
        return len(self._items)

    def __bool__(self):
        return bool(self._items)

    def __repr__(self):
        return f"DefinedNameList({self._items!r})"

    def __getitem__(self, key):
        if isinstance(key, int) and not isinstance(key, bool):
            return self._items[key]
        if isinstance(key, str):
            folded = key.casefold()
            for definition in self._items:
                if definition.name.casefold() == folded and definition.scope is None:
                    return definition
            for definition in self._items:
                if definition.name.casefold() == folded:
                    return definition
            raise KeyError(key)
        raise TypeError("defined name keys must be a name or an index")

    def __contains__(self, key):
        try:
            self[key]
        except (KeyError, TypeError):
            return False
        return True

    def __setitem__(self, key, value):
        self.add(value)

    @property
    def names(self):
        return [definition.name for definition in self._items]

    def add(self, definition, *, stage: bool = True):
        if not isinstance(definition, DefinedName):
            raise TypeError("defined names must be DefinedName instances")
        self._items = [
            item for item in self._items
            if not (item.name.casefold() == definition.name.casefold()
                    and item.scope == definition.scope)
        ]
        self._items.append(definition)
        if stage and self._workbook is not None:
            if self._workbook._updater is not None:
                self._workbook._updater.set_defined_name(definition)
            else:
                self._workbook._defined_names_dirty = True

    append = add

    def remove(self, name, scope=None):
        if not isinstance(name, str):
            raise TypeError("defined name must be a string")
        folded = name.casefold()
        target = None
        for definition in self._items:
            if definition.name.casefold() == folded and (
                scope is None or definition.scope == scope
            ):
                target = definition
                break
        if target is None:
            raise KeyError(name)
        self._items.remove(target)
        if self._workbook is not None:
            if self._workbook._updater is not None:
                self._workbook._updater.remove_defined_name(target.name, target.scope)
            else:
                self._workbook._defined_names_dirty = True
        return target


@dataclass
class DocumentProperties:
    creator: str = "xlspy"
    title: str | None = None
    subject: str | None = None
    description: str | None = None
    keywords: str | None = None
    category: str | None = None
    lastModifiedBy: str | None = None
    created: _datetime.datetime = field(default_factory=_datetime.datetime.now)
    modified: _datetime.datetime = field(default_factory=_datetime.datetime.now)
    revision: str | None = None
    contentStatus: str | None = None
    identifier: str | None = None
    language: str | None = None
    version: str | None = None
    lastPrinted: _datetime.datetime | None = None


@dataclass
class Outline:
    applyStyles: bool | None = None
    summaryBelow: bool | None = None
    summaryRight: bool | None = None
    showOutlineSymbols: bool | None = None


@dataclass
class SheetProperties:
    tabColor: Color | str | None = None
    codeName: str | None = None
    pageSetUpPr: PageSetupProperties = field(default_factory=PageSetupProperties)
    outlinePr: Outline = field(default_factory=Outline)
    enableFormatConditionsCalculation: bool | None = None
    filterMode: bool | None = None
    published: bool | None = None
    syncHorizontal: bool | None = None
    syncRef: str | None = None
    syncVertical: bool | None = None
    transitionEvaluation: bool | None = None
    transitionEntry: bool | None = None


@dataclass
class SheetFormatProperties:
    baseColWidth: int = 8
    defaultColWidth: float | None = None
    defaultRowHeight: float = 15
    customHeight: bool | None = None
    zeroHeight: bool | None = None
    thickTop: bool | None = None
    thickBottom: bool | None = None
    outlineLevelRow: int | None = None
    outlineLevelCol: int | None = None


@dataclass
class Pane:
    xSplit: float | int | None = None
    ySplit: float | int | None = None
    topLeftCell: str | None = None
    activePane: str = "topLeft"
    state: str = "split"


@dataclass
class Selection:
    pane: str | None = None
    activeCell: str = "A1"
    activeCellId: int | None = None
    sqref: str = "A1"


@dataclass
class SheetView:
    windowProtection: bool | None = None
    showFormulas: bool | None = None
    showGridLines: bool | None = None
    showRowColHeaders: bool | None = None
    showZeros: bool | None = None
    rightToLeft: bool | None = None
    tabSelected: bool | None = None
    showRuler: bool | None = None
    showOutlineSymbols: bool | None = None
    defaultGridColor: bool | None = None
    showWhiteSpace: bool | None = None
    view: str | None = None
    topLeftCell: str | None = None
    colorId: int | None = None
    zoomScale: int | None = None
    zoomScaleNormal: int | None = None
    zoomScaleSheetLayoutView: int | None = None
    zoomScalePageLayoutView: int | None = None
    zoomToFit: bool | None = None
    workbookViewId: int = 0
    selection: list[Selection] = field(default_factory=list)
    pane: Pane | None = None


@dataclass
class FilterColumn:
    colId: int = 0
    hiddenButton: bool = False
    showButton: bool = True
    vals: list[str] = field(default_factory=list)
    blank: bool = False
    # Criteria other than a value list (customFilters, top10, dynamicFilter,
    # colorFilter, iconFilter), kept as the XML they were read from. A column
    # holds either a value list or one of these, so the writer emits them alone.
    criteria: list[str] = field(default_factory=list)


@dataclass
class SortCondition:
    ref: str | None = None
    descending: bool | None = None
    sortBy: str | None = None
    customList: str | None = None
    dxfId: int | None = None
    iconSet: str | None = None
    iconId: int | None = None


@dataclass
class SortState:
    ref: str | None = None
    columnSort: bool | None = None
    caseSensitive: bool | None = None
    sortMethod: str | None = None
    sortCondition: list[SortCondition] = field(default_factory=list)


@dataclass
class AutoFilter:
    ref: str | None = None
    filterColumn: list[FilterColumn] = field(default_factory=list)
    sortState: SortState | None = None

    def add_filter_column(self, col_id, vals, blank=False):
        self.filterColumn.append(
            FilterColumn(colId=col_id, vals=list(vals or []), blank=bool(blank))
        )

    def add_sort_condition(self, ref, descending=False):
        if self.sortState is None:
            self.sortState = SortState(ref=self.ref)
        self.sortState.sortCondition.append(
            SortCondition(ref=ref, descending=bool(descending))
        )


@dataclass
class Break:
    id: int = 0
    min: int = 0
    max: int = 16383
    man: bool = True
    pt: int | None = None


class RowBreak:
    def __init__(self, brk=()):
        self.brk = list(brk)

    @property
    def count(self):
        return len(self.brk)

    @property
    def manualBreakCount(self):
        return len(self.brk)

    def append(self, brk=None):
        if not isinstance(brk, Break):
            brk = Break(id=self.count + 1)
        self.brk.append(brk)

    def __len__(self):
        return len(self.brk)

    def __bool__(self):
        return bool(self.brk)


class ColBreak(RowBreak):
    pass


@dataclass
class SheetProtection:
    sheet: bool = False
    objects: bool = False
    scenarios: bool = False
    formatCells: bool = True
    formatRows: bool = True
    formatColumns: bool = True
    insertColumns: bool = True
    insertRows: bool = True
    insertHyperlinks: bool = True
    deleteColumns: bool = True
    deleteRows: bool = True
    selectLockedCells: bool = False
    selectUnlockedCells: bool = False
    sort: bool = True
    autoFilter: bool = True
    pivotTables: bool = True
    password: str | None = None
    algorithmName: str | None = None
    saltValue: str | None = None
    spinCount: int | None = None
    hashValue: str | None = None

    def __setattr__(self, name, value):
        # As in openpyxl, a password set directly is stored as its legacy hash.
        if name == "password" and value is not None:
            from .._metadata import password_hash

            value = password_hash(value)
        object.__setattr__(self, name, value)

    def set_password(self, value: str | None = "", already_hashed: bool = False) -> None:
        if value is not None and not already_hashed:
            from .._metadata import password_hash

            value = password_hash(value)
        object.__setattr__(self, "password", value)

    def enable(self) -> None:
        self.sheet = True

    def disable(self) -> None:
        self.sheet = False

    def __bool__(self) -> bool:
        return self.sheet


def _as_model_protection(value) -> "SheetProtection":
    """Return the model SheetProtection for ``ws.protection`` assignments.

    Accepts the model class (openpyxl-style fields) or the public
    ``xlspy.SheetProtection`` (``enabled``, ``select_*_cells`` meaning "allowed").
    The public class is converted here so the model always holds one type.
    """
    if value is None:
        return SheetProtection()
    if isinstance(value, SheetProtection):
        return value
    if isinstance(value, PublicSheetProtection):
        model = SheetProtection(sheet=value.enabled, objects=value.enabled, scenarios=value.enabled,
                                selectLockedCells=not value.select_locked_cells,
                                selectUnlockedCells=not value.select_unlocked_cells)
        if value.password_hash is not None:
            model.set_password(value.password_hash, already_hashed=True)
        return model
    raise TypeError("protection must be a xlspy SheetProtection, a SheetProtection or None")


def _as_model_page_margins(value) -> "PageMargins":
    """Return the model PageMargins for ``ws.page_margins`` (mutable holder)."""
    if isinstance(value, PageMargins):
        return value
    if isinstance(value, PublicPageMargins):
        return PageMargins(**{name: getattr(value, name) for name in
                              ("left", "right", "top", "bottom", "header", "footer")})
    raise TypeError("page_margins must be a xlspy PageMargins or PageMargins")


@dataclass
class CustomProperty:
    name: str
    value: object


@dataclass
class DocumentSecurity:
    lockStructure: bool = False
    lockWindows: bool = False
    lockRevision: bool = False
