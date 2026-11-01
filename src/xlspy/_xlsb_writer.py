"""Optional metadata methods for XlsbWriter, outside its numeric row path."""

from typing import TYPE_CHECKING, Any

from ._metadata import DefinedName

if TYPE_CHECKING:
    from ._metadata import WorkbookProtection
    from ._xlsb_formulas import FormulaContext


class XlsbWriterMetadata:
    if TYPE_CHECKING:
        # Provided by WriterFrontEnd / StagedWriter, which the writer mixes in first.
        def _require_writable(self) -> None: ...
        # Attributes of XlsbWriter (writer.py) that these methods read and write.
        _xlsb_formula_context: FormulaContext | None
        _worksheet_data: list[tuple[str, Any, bool]]
        _workbook_protection: WorkbookProtection | None
        _filtered_data_list: list[Any]
        _slicer_cache_names: Any
        _sheet_print: dict[int, Any]

    def _formula_context(self):
        if self._xlsb_formula_context is None:
            from ._xlsb_formulas import FormulaContext
            self._xlsb_formula_context = FormulaContext(lambda: [name for name, _, _ in self._worksheet_data])
        return self._xlsb_formula_context

    def set_defined_name(self, definition: DefinedName) -> None:
        self._require_writable()
        self._formula_context().set_name(definition)

    def get_defined_names(self) -> list[DefinedName]:
        return self._xlsb_formula_context.defined_names() if self._xlsb_formula_context is not None else []

    def remove_defined_name(self, name: str, scope: str | None = None) -> None:
        self._require_writable()
        DefinedName(name, '0', scope)
        if self._xlsb_formula_context is not None:
            self._xlsb_formula_context.remove_name(name, scope)

    def _metadata_workbook(self, data):
        if self._xlsb_formula_context is not None:
            from ._xlsb_edit import request_recalculation
            data = self._xlsb_formula_context.write_workbook(data)
            data = request_recalculation(data)
        if self._workbook_protection is not None:
            from ._xlsb_metadata import replace_records, workbook_protection_bin
            data = replace_records(data, {0x216, 0x2A5}, workbook_protection_bin(self._workbook_protection), before=(0x87, 0x8F))
        return data

    def _register_filter_names(self):
        # Name indexes already written into cell formulas must stay stable.
        # Append filter names after user/future names and use the same XTI table.
        from ._xlsb_formulas import _Name
        from .refs import quote_sheet_name, column_label
        context = self._formula_context()  # created by _write_workbook_parts before these run
        for item in self._filtered_data_list:
            title = self._worksheet_data[item.sheet_index][0]
            first = f'${column_label(item.start_column)}${item.start_row + 1}'
            last = f'${column_label(item.end_column)}${item.end_row + 1}'
            definition = DefinedName('_FilterDatabase', quote_sheet_name(title) + '!' + first + ':' + last, title)
            context.names.append(_Name(definition.name, title, 0x21, definition=definition))

    def _register_slicer_names(self):
        # Excel links each slicer cache to a workbook name holding #N/A; the
        # name record carries flag 0x10 as Excel writes it.
        from ._xlsb_formulas import _Name
        context = self._formula_context()  # created by _write_workbook_parts before these run
        for cache_name in self._slicer_cache_names:
            context.names.append(_Name(cache_name, None, 0x10, definition=DefinedName(cache_name, '#N/A')))

    def _register_print_names(self):
        # Built-in (fBuiltin) sheet names, as Excel stores _xlnm.Print_Area
        # and _xlnm.Print_Titles in XLSB.
        from ._xlsb_formulas import _Name
        from .printing import print_area_name, print_titles_name
        context = self._formula_context()  # created by _write_workbook_parts before these run
        for sheet_index in sorted(self._sheet_print):
            settings = self._sheet_print[sheet_index]
            title = self._worksheet_data[sheet_index][0]
            values = []
            if settings.print_area is not None:
                values.append(('Print_Area', print_area_name(title, settings.print_area)))
            titles = print_titles_name(title, settings)
            if titles is not None:
                values.append(('Print_Titles', titles))
            for name, formula in values:
                definition = DefinedName(name, formula, title)
                context.names.append(_Name(name, title, 0x20, definition=definition))
