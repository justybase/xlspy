"""Native Excel checks for hyperlinks and chart layout in XLSB output."""

import os

import pytest

from xlspy import Chart, ChartSeries, Hyperlink, Style, XlsbWriter
from test_excel_com_metadata import excel, open_workbook  # noqa: F401

pytestmark = [pytest.mark.com, pytest.mark.skipif(os.environ.get('XLSPY_RUN_EXCEL_COM') != '1', reason='Set XLSPY_RUN_EXCEL_COM=1')]

XL_UNDERLINE_SINGLE = 2
HYPERLINK_BLUE_BGR = 0xC16305  # #0563C1 as Excel's BGR integer


def test_hyperlinks_in_excel(tmp_path, open_workbook):
    path = tmp_path / 'links.xlsb'
    with XlsbWriter(path) as writer:
        writer.add_sheet('Data')
        writer.write_sheet([
            ['Target', 'Description'],
            [Hyperlink('https://example.com/report', text='Report', tooltip='Tooltip'), 'external'],
            [Hyperlink('mailto:someone@example.com', text='Mail'), 'mail'],
            [Hyperlink("#'Data'!A1", text='Start'), 'internal'],
            [(Hyperlink('https://example.com/report', text='Second'), Style(bold=True)), 'repeated address'],
        ])
    book = open_workbook(path)
    sheet = book.Worksheets('Data')

    links = sheet.Hyperlinks
    assert links.Count == 4
    by_cell = {link.Range.Address: link for link in links}
    assert set(by_cell) == {'$A$2', '$A$3', '$A$4', '$A$5'}

    raport = by_cell['$A$2']
    assert raport.Address == 'https://example.com/report'
    assert raport.SubAddress == ''
    assert raport.TextToDisplay == 'Report'
    assert raport.ScreenTip == 'Tooltip'

    assert by_cell['$A$3'].Address == 'mailto:someone@example.com'

    internal = by_cell['$A$4']
    assert internal.SubAddress == "'Data'!A1"
    assert internal.Address == ''
    assert internal.TextToDisplay == 'Start'

    assert by_cell['$A$5'].Address == 'https://example.com/report'
    assert by_cell['$A$5'].TextToDisplay == 'Second'
    assert sheet.Range('A5').Font.Bold

    styled = sheet.Range('A2')
    assert styled.Value == 'Report'
    assert styled.Font.Underline == XL_UNDERLINE_SINGLE
    assert int(styled.Font.Color) == HYPERLINK_BLUE_BGR


def test_xlsb_chart_legend_does_not_overlap_plot_in_excel(tmp_path, open_workbook):
    path = tmp_path / 'chart_layout.xlsb'
    with XlsbWriter(path) as writer:
        writer.add_sheet('Data')
        writer.write_sheet(
            [['Region', 'Sales'], ['North', 10], ['South', 20], ['East', 30]],
            charts=[Chart('bar', [ChartSeries('B2:B4', title_cell='B1', categories='A2:A4')],
                          title='Sales', anchor='D2')],
        )
    book = open_workbook(path)
    chart = book.Worksheets('Data').ChartObjects(1).Chart
    assert chart.HasLegend
    # IncludeInLayout is False when the legend is drawn over the plot area.
    assert chart.Legend.IncludeInLayout
    assert chart.ChartTitle.IncludeInLayout
