"""PivotTable.autofit_columns: Excel's "Autofit column widths on update" (XLSX and XLSB).

By default a pivot does not change the column widths of the sheet it sits on when
Excel refreshes it on open. ``autofit_columns=True`` turns the setting on.
"""

import io
import zipfile

import pytest

from xlspy import ExcelTable, PivotTable, PivotValue, XlsbWriter, XlsxWriter
from xlspy._xlsb_pivots import _VIEW_PREFIX, _view_prefix

ROWS = [
    ["Region", "Produkt", "Sprzedaż"],
    ["Północ", "Kawa", 10],
    ["Południe", "Herbata", 20],
    ["Północ", "Woda", 30],
]


def _book(writer_cls, **pivot_kwargs):
    table = ExcelTable("TabelaTest", "A1:C4")
    buffer = io.BytesIO()
    with writer_cls(buffer) as writer:
        writer.add_sheet("Dane")
        writer.write_sheet(
            ROWS,
            tables=[table],
            pivots=[PivotTable(
                source=table, rows="Region",
                values=[PivotValue("Sprzedaż", "sum")],
                location="A8", **pivot_kwargs,
            )],
        )
    return buffer.getvalue()


def _part(data, name):
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return archive.read(name)


def test_xlsx_autofit_is_off_by_default():
    part = _part(_book(XlsxWriter), "xl/pivotTables/pivotTable1.xml").decode()
    assert 'useAutoFormatting="0"' in part
    assert 'useAutoFormatting="1"' not in part


def test_xlsx_autofit_on_when_requested():
    part = _part(_book(XlsxWriter, autofit_columns=True), "xl/pivotTables/pivotTable1.xml").decode()
    assert 'useAutoFormatting="1"' in part
    assert 'useAutoFormatting="0"' not in part


def test_xlsb_autofit_bit_is_off_by_default():
    part = _part(_book(XlsbWriter), "xl/pivotTables/pivotTable1.bin")
    assert _view_prefix(False) in part
    assert _view_prefix(True) not in part


def test_xlsb_autofit_bit_on_when_requested():
    part = _part(_book(XlsbWriter, autofit_columns=True), "xl/pivotTables/pivotTable1.bin")
    assert _view_prefix(True) in part
    assert _view_prefix(False) not in part


def test_xlsb_prefixes_differ_only_in_the_autofit_bit():
    on, off = _view_prefix(True), _view_prefix(False)
    assert on == _VIEW_PREFIX
    assert len(on) == len(off)
    diff = [i for i, (a, b) in enumerate(zip(on, off)) if a != b]
    assert diff == [5]
    assert on[5] ^ off[5] == 0x01


@pytest.mark.parametrize("value", [1, "yes", None])
def test_autofit_columns_must_be_a_bool(value):
    with pytest.raises(TypeError, match="autofit_columns"):
        PivotTable(source="TabelaTest", rows="Region",
                   values=[PivotValue("Sprzedaż", "sum")], autofit_columns=value)
