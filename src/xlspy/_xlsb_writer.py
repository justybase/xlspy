"""Optional metadata methods for XlsbWriter, outside its numeric row path."""

from ._metadata import DefinedName, WorkbookProtection


class XlsbWriterMetadata:
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

    def set_workbook_protection(self, protection: WorkbookProtection) -> None:
        self._require_writable()
        if not isinstance(protection, WorkbookProtection):
            raise TypeError("protection must be WorkbookProtection")
        self._workbook_protection = protection

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
        context = self._xlsb_formula_context
        for item in self._filtered_data_list:
            title = self._worksheet_data[item.sheet_index][0]
            first = f'${column_label(item.start_column)}${item.start_row + 1}'
            last = f'${column_label(item.end_column)}${item.end_row + 1}'
            definition = DefinedName('_FilterDatabase', quote_sheet_name(title) + '!' + first + ':' + last, title)
            context.names.append(_Name(definition.name, title, 0x21, definition=definition))
