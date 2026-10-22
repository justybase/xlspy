"""Native Excel checks for XLSB features that match the XLSX writer."""

import datetime
import io
import os

import pytest

from xlspy import (
    CellRichText, Chart, ChartSeries, InlineFont, Style, TextBlock, DataValidation, Formula, Image, PageMargins, PrintSettings, XlsbWriter,
    XlsxWriter,
)
from test_excel_com_metadata import excel, open_workbook  # noqa: F401
from test_xlsb_charts_extended import CHARTS, ROWS
from test_excel_com_properties_charts import GIF_1X1, _bmp, _png

XL_CATEGORY, XL_VALUE = 1, 2
# xlLine, xlLineMarkers, xlXYScatter, xlXYScatterLines, xlXYScatterLinesNoMarkers
_LINE_TYPES = {4, 65, -4169, 74, 75}

pytestmark = [pytest.mark.com, pytest.mark.skipif(os.environ.get('XLSPY_RUN_EXCEL_COM') != '1', reason='Set XLSPY_RUN_EXCEL_COM=1')]


def test_chart_over_formula_cells_in_excel(tmp_path, open_workbook):
    path = tmp_path / 'formula_chart.xlsb'
    with XlsbWriter(path) as writer:
        writer.add_sheet('Data')
        writer.write_sheet(
            [['Region', 'Share'], ['North', Formula('1+1', 2)], ['South', Formula('2+2')],
             ['East', (Formula('3+3', 6), '0.0')]],
            charts=[Chart('pie', [ChartSeries('B2:B4', categories='A2:A4')], anchor='D2')],
        )
    book = open_workbook(path)
    book.Application.Calculate()
    series = book.Worksheets('Data').ChartObjects(1).Chart.SeriesCollection(1)
    assert tuple(series.Values) == (2.0, 4.0, 6.0)
    assert tuple(series.XValues) == ('North', 'South', 'East')



def test_default_date_formats_match_in_excel(tmp_path, open_workbook):
    rows = [['Data', 'Date and time'], [datetime.date(2026, 10, 22), datetime.datetime(2026, 10, 22, 13, 45, 30)]]
    facts = {}
    for writer_class, suffix in ((XlsxWriter, 'xlsx'), (XlsbWriter, 'xlsb')):
        path = tmp_path / f'dates.{suffix}'
        with writer_class(path) as writer:
            writer.add_sheet('Dates')
            writer.write_sheet(rows)
        sheet = open_workbook(path).Worksheets('Dates')
        facts[suffix] = [(sheet.Range(a).NumberFormat, sheet.Range(a).Text) for a in ('A2', 'B2')]
    # NumberFormat comes back in Excel's language (rrrr in Polish); the text does not.
    assert facts['xlsx'] == facts['xlsb']
    assert [text for _, text in facts['xlsb']] == ['2026-10-22', '2026-10-22 13:45']


RICH_ROWS = [
    ['Header'],
    [CellRichText('Total: ', TextBlock(InlineFont(b=True, color='C00000', sz=14), '1 234'), ' EUR')],
    [CellRichText(TextBlock(InlineFont(i=True, u='double', rFont='Arial'), 'italic'),
                  TextBlock(InlineFont(strike=True, vertAlign='superscript'), 'x2'), ' end')],
    [(CellRichText('cell ', TextBlock(InlineFont(color='0070C0'), 'style')), Style(fill='#FFF2CC'))],
    [CellRichText(TextBlock(InlineFont(b=True), 'whole'))],
    [CellRichText('no format')],
]


def _rich_facts(sheet):
    facts = []
    for row in range(1, len(RICH_ROWS) + 1):
        cell = sheet.Cells(row, 1)
        text = cell.Value
        fonts = []
        for index in range(1, len(text) + 1):
            font = cell.GetCharacters(index, 1).Font
            fonts.append((font.Bold, font.Italic, font.Color, font.Size, font.Name,
                          font.Underline, font.Strikethrough, font.Superscript))
        facts.append((text, cell.Interior.Color, fonts))
    return facts


def test_rich_text_matches_xlsx_in_excel(tmp_path, open_workbook):
    facts = {}
    for writer_class, suffix, shared in (
        (XlsxWriter, 'xlsx', True), (XlsbWriter, 'xlsb', True), (XlsbWriter, 'xlsb', False),
    ):
        path = tmp_path / f'rich_{int(shared)}.{suffix}'
        with writer_class(path, useSharedStrings=shared) as writer:
            writer.add_sheet('Text')
            writer.write_sheet(RICH_ROWS)
        facts[(suffix, shared)] = _rich_facts(open_workbook(path).Worksheets('Text'))
    assert facts[('xlsb', True)] == facts[('xlsx', True)]
    assert facts[('xlsb', False)] == facts[('xlsx', True)]
    total = facts[('xlsb', True)][1][2]
    assert total[0][:4] == (False, False, 0, 11) and total[7][:4] == (True, False, 0x0000C0, 14)

def _get(read, default=None):
    try:
        return read()
    except Exception:  # the property does not apply to this chart type
        return default


def _chart_facts(chart):
    facts = {
        'type': chart.ChartType,
        'series': chart.SeriesCollection().Count,
        'legend': chart.HasLegend,
        'legend_position': _get(lambda: chart.Legend.Position) if chart.HasLegend else None,
        'legend_in_layout': _get(lambda: chart.Legend.IncludeInLayout) if chart.HasLegend else None,
        'title': _get(lambda: chart.ChartTitle.Text) if chart.HasTitle else None,
        'title_size': _get(lambda: chart.ChartTitle.Format.TextFrame2.TextRange.Font.Size) if chart.HasTitle else None,
    }
    for name, axis_type in (('x', XL_CATEGORY), ('y', XL_VALUE)):
        if _get(lambda: chart.HasAxis(axis_type)):
            axis = chart.Axes(axis_type)
            facts[f'{name}_axis'] = (
                _get(lambda: axis.MinimumScale), _get(lambda: axis.MaximumScale),
                _get(lambda: axis.MajorUnit), _get(lambda: axis.TickLabels.NumberFormat),
                axis.HasTitle and axis.AxisTitle.Text,
            )
    for index in range(1, facts['series'] + 1):
        series = chart.SeriesCollection(index)
        facts[f'series{index}'] = (
            tuple(series.Values), _get(lambda: tuple(series.XValues)), series.Name,
            series.HasDataLabels, _get(lambda: series.MarkerStyle), _get(lambda: series.MarkerSize),
            # Smoothing applies to lines only; area charts report noise here.
            series.Smooth if facts['type'] in _LINE_TYPES else None,
        )
    return facts


@pytest.mark.parametrize('name', sorted(CHARTS))
def test_xlsb_chart_matches_xlsx_in_excel(tmp_path, open_workbook, name):
    chart = CHARTS[name]
    facts = {}
    for writer_class, suffix in ((XlsxWriter, 'xlsx'), (XlsbWriter, 'xlsb')):
        path = tmp_path / f'{name}.{suffix}'
        with writer_class(path) as writer:
            writer.add_sheet('Data')
            writer.write_sheet(ROWS, charts=[chart])
        book = open_workbook(path)
        excel_chart = book.Worksheets('Data').ChartObjects(1).Chart
        facts[suffix] = _chart_facts(excel_chart)
        explicit = [item.color for item in chart.series]
        if any(explicit):
            facts[suffix]['colours'] = [
                int(excel_chart.SeriesCollection(index + 1).Format.Fill.ForeColor.RGB
                    if chart.chart_type != 'scatter'
                    else excel_chart.SeriesCollection(index + 1).Format.Line.ForeColor.RGB)
                for index in range(len(explicit))
            ]
    # XLSB charts deliberately use 18 pt titles by default (see chart_parts).
    if chart.title_size is None:
        del facts['xlsx']['title_size']
        del facts['xlsb']['title_size']
    assert facts['xlsb'] == facts['xlsx']


MSO_PICTURE = 13


def test_pictures_in_excel(tmp_path, open_workbook):
    path = tmp_path / 'pictures.xlsb'
    with XlsbWriter(path) as writer:
        writer.add_sheet('Pictures')
        writer.write_sheet(
            [['k', 'v'], ['a', 1], ['b', 2]],
            images=[
                Image(_png(40, 20), anchor='B2'),
                Image(io.BytesIO(_png(40, 20)), anchor='B8', width=80),
                Image(GIF_1X1, anchor='F2', width=24),
                Image(_bmp(30, 12), anchor='F8'),
            ],
            charts=[Chart('bar', [ChartSeries('B2:B3', categories='A2:A3')], anchor='J2')],
        )
    shapes = open_workbook(path).Worksheets('Pictures').Shapes
    assert shapes.Count == 5
    pictures = [shapes(i) for i in range(1, 6) if shapes(i).Type == MSO_PICTURE]
    assert len(pictures) == 4
    # Shape sizes are in points: 0.75 point per pixel.
    sizes = sorted((round(shape.Width / 0.75), round(shape.Height / 0.75)) for shape in pictures)
    assert sizes == sorted([(40, 20), (80, 40), (24, 24), (30, 12)])
    corners = sorted(shape.TopLeftCell.Address.replace('$', '') for shape in pictures)
    assert corners == ['B2', 'B8', 'F2', 'F8']


XL_LANDSCAPE = 2
PRINT_SETTINGS = PrintSettings(
    orientation='landscape', paper_size='A4', margins=PageMargins(0.5, 0.6, 0.7, 0.8, 0.25, 0.35),
    print_area='A1:D20', repeat_rows='1:2', repeat_columns='A:A', fit_to_width=1, fit_to_height=0,
    center_horizontally=True, center_vertically=True, header='&LReport&RPage &P of &N',
    footer='&CFooter', row_breaks=[10], col_breaks=[2],
)


def _page_facts(sheet):
    setup = sheet.PageSetup
    return {
        'orientation': setup.Orientation, 'paper': setup.PaperSize, 'zoom': setup.Zoom,
        'fit': (setup.FitToPagesWide, setup.FitToPagesTall), 'area': setup.PrintArea,
        'title_rows': setup.PrintTitleRows, 'title_columns': setup.PrintTitleColumns,
        'centre': (setup.CenterHorizontally, setup.CenterVertically),
        'header': (setup.LeftHeader, setup.CenterHeader, setup.RightHeader),
        'footer': (setup.LeftFooter, setup.CenterFooter, setup.RightFooter),
        # Margins are in points: 72 per inch.
        'margins': tuple(round(value / 72, 2) for value in (
            setup.LeftMargin, setup.RightMargin, setup.TopMargin, setup.BottomMargin,
            setup.HeaderMargin, setup.FooterMargin)),
        'row_breaks': [sheet.HPageBreaks(i).Location.Row for i in range(1, sheet.HPageBreaks.Count + 1)
                       if sheet.HPageBreaks(i).Type == -4135],
        'col_breaks': [sheet.VPageBreaks(i).Location.Column for i in range(1, sheet.VPageBreaks.Count + 1)
                       if sheet.VPageBreaks(i).Type == -4135],
    }


def _print_facts(tmp_path, open_workbook, settings):
    """Page setup of the same sheet written as XLSX and as XLSB."""
    rows = [['h1', 'h2', 'h3', 'h4']] + [[i, i * 2, i * 3, i * 4] for i in range(30)]
    facts = {}
    for writer_class, suffix in ((XlsxWriter, 'xlsx'), (XlsbWriter, 'xlsb')):
        path = tmp_path / f'print.{suffix}'
        with writer_class(path) as writer:
            writer.add_sheet('Print')
            writer.write_sheet(rows, print_settings=settings)
            writer.add_sheet('Other')
            writer.write_sheet([['x']])
        sheet = open_workbook(path).Worksheets('Print')
        if suffix == 'xlsx':
            try:
                sheet.PageSetup.Orientation
            except Exception:  # pywintypes.com_error
                # Excel reads the page setup through the default printer.
                pytest.skip('Excel cannot reach the default printer to read the page setup')
        facts[suffix] = _page_facts(sheet)
    return facts


@pytest.mark.parametrize('settings', [
    PRINT_SETTINGS,
    PrintSettings(scale=80),
    PrintSettings(orientation='portrait', header='&COnly'),
    # Without any page option Excel asks the default printer for one; an
    # orientation lets it read the sheet's setup without a reachable printer.
    PrintSettings(repeat_columns='B:C', orientation='portrait'),
], ids=['everything', 'scale', 'portrait_header', 'title_columns'])
@pytest.mark.printer
def test_print_settings_match_xlsx_in_excel(tmp_path, open_workbook, settings):
    facts = _print_facts(tmp_path, open_workbook, settings)
    assert facts['xlsb'] == facts['xlsx']


@pytest.mark.printer
def test_print_settings_in_excel(tmp_path, open_workbook):
    facts = _print_facts(tmp_path, open_workbook, PRINT_SETTINGS)['xlsb']
    # PaperSize depends on what the default printer supports; the parity test
    # above compares it with the XLSX file instead.
    assert facts['orientation'] == XL_LANDSCAPE
    assert facts['zoom'] is False and facts['fit'] == (1, False)
    assert facts['area'] == '$A$1:$D$20'
    assert facts['title_rows'] == '$1:$2' and facts['title_columns'] == '$A:$A'
    assert facts['centre'] == (True, True)
    assert facts['header'] == ('Report', '', 'Page &P of &N')
    assert facts['footer'] == ('', 'Footer', '')
    assert facts['margins'] == (0.5, 0.6, 0.7, 0.8, 0.25, 0.35)
    assert facts['row_breaks'] == [11] and facts['col_breaks'] == [3]


VALIDATIONS = [
    DataValidation('list', 'A2:A10', '"New,Active,Closed"', prompt_title='Status', prompt='Choose',
                   show_input_message=True, error_title='Error', error='Only from the list'),
    DataValidation('list', 'B2:B10', '$H$2:$H$4', show_dropdown=False, allow_blank=False),
    DataValidation('list', 'C2', 'Lists!$A$1:$A$3'),
    DataValidation('whole', 'D2:D10', 1, 100, error_style='warning'),
    DataValidation('decimal', 'E2:E10', 0.5, operator='greater_than', error_style='information'),
    DataValidation('date', 'F2:F10', datetime.date(2024, 1, 1), datetime.date(2024, 12, 31),
                   operator='not_between'),
    DataValidation('time', 'G2', datetime.time(8, 0), operator='less_than_or_equal'),
    DataValidation('text_length', 'I2:I10 K2:K3', 10, operator='less_than', show_error_message=False),
    DataValidation('custom', 'J2:J10', 'MOD(J2,2)=0'),
    DataValidation('whole', 'L2:L5', '$M$1', '$M$2'),
]


def _validation_facts(sheet):
    facts = {}
    for cell in ('A2', 'A10', 'B3', 'C2', 'D2', 'E2', 'F2', 'G2', 'I2', 'K3', 'J5', 'L3'):
        rule = sheet.Range(cell).Validation
        facts[cell] = (
            rule.Type, rule.Operator if rule.Type not in (3, 7) else None, rule.Formula1,
            rule.Formula2 if rule.Type not in (3, 7) else None, rule.AlertStyle, rule.IgnoreBlank,
            rule.InCellDropdown if rule.Type == 3 else None, rule.ShowInput, rule.ShowError,
            rule.InputTitle, rule.InputMessage, rule.ErrorTitle, rule.ErrorMessage,
        )
    return facts


def test_data_validation_matches_xlsx_in_excel(tmp_path, open_workbook):
    rows = [['a'] * 13] + [[None] * 7 + [f'h{i}'] for i in range(3)]
    facts = {}
    for writer_class, suffix in ((XlsxWriter, 'xlsx'), (XlsbWriter, 'xlsb')):
        path = tmp_path / f'validation.{suffix}'
        with writer_class(path) as writer:
            writer.add_sheet('Data')
            writer.write_sheet(rows, data_validations=VALIDATIONS)
            writer.add_sheet('Lists')
            writer.write_sheet([['x'], ['y'], ['z']])
        facts[suffix] = _validation_facts(open_workbook(path).Worksheets('Data'))
    assert facts['xlsb'] == facts['xlsx']
    # The relative custom formula moves with the cell. Formula1 uses the
    # list separator of Excel's locale (";" in Polish, for example).
    assert facts['xlsb']['J5'][2].replace(';', ',') == '=MOD(J5,2)=0'
    assert facts['xlsb']['C2'][2] == '=Lists!$A$1:$A$3'


def _cf_rules():
    from xlspy import (
        AboveAverageRule, Alignment, BlanksRule, Border, CellIsRule, DifferentialStyle, DuplicateRule,
        Font, FormulaRule, PatternFill, Protection, Side, TextRule, TimePeriodRule, TopBottomRule,
    )
    red = DifferentialStyle(font=Font(bold=True, color='9C0006'), fill=PatternFill('solid', start_color='FFC7CE'))
    green = DifferentialStyle(fill=PatternFill('solid', end_color='C6EFCE'))
    percent = DifferentialStyle(number_format='0.0%', font=Font(italic=True))
    box = DifferentialStyle(border=Border(left=Side('thin', 'FF0000'), bottom=Side('double', '0000FF')))
    fancy = DifferentialStyle(font=Font(name='Arial', size=14, underline='single', strike=True, color='00B050'),
                              fill=PatternFill('gray125', start_color='FF0000', end_color='0000FF'))
    aligned = DifferentialStyle(alignment=Alignment(horizontal='center', wrap_text=True),
                                protection=Protection(locked=False, hidden=True))
    return [
        CellIsRule('A2:A20', 'between', ['5', '10'], dxf=red),
        CellIsRule('A2:A20', 'greaterThan', ['$K$1'], dxf=green, stop_if_true=True),
        CellIsRule('B2:B20', 'equal', ['"yes"'], dxf=percent),
        FormulaRule('C2:C20', 'MOD(C2,2)=0', dxf=box),
        FormulaRule('C2:C20', 'COUNTIF(C$2:C$20,C2)>1', dxf=fancy),
        TopBottomRule('D2:D20', rank=3, dxf=red),
        TopBottomRule('D2:D20', rank=10, percent=True, bottom=True, dxf=green),
        AboveAverageRule('E2:E20', dxf=red),
        AboveAverageRule('E2:E20', below=True, equal_average=True, dxf=green),
        AboveAverageRule('E2:E20', std_dev=1, dxf=aligned),
        TextRule('F2:F20', 'contains', 'ab', dxf=red),
        TextRule('F2:F20', 'notContains', 'x"y', dxf=green),
        TextRule('F2:F20', 'beginsWith', 'Jo', dxf=percent),
        TextRule('F2:F20', 'endsWith', 'son', dxf=box),
        DuplicateRule('G2:G20', dxf=red),
        DuplicateRule('G2:G20', unique=True, dxf=green),
        *[TimePeriodRule('I2:I20', period, dxf=fancy) for period in (
            'today', 'yesterday', 'tomorrow', 'last7Days', 'thisWeek', 'lastWeek', 'nextWeek',
            'thisMonth', 'lastMonth', 'nextMonth')],
        BlanksRule('J2:J20', dxf=red),
        BlanksRule('J2:J20', 'no_blanks', dxf=green),
        BlanksRule('J2:J20', 'errors', dxf=percent),
        BlanksRule('J2:J20', 'no_errors', dxf=box),
    ]


def _cf_rows():
    today = datetime.date.today()
    names = ['abc', 'Johnson', 'xyz', 'x"y', 'Brown', 'abc']
    rows = [['h'] * 10 + [None, 7]]
    for i in range(19):
        day = today + datetime.timedelta(days=i * 3 - 30)
        rows.append([i, 'yes' if i % 3 else 'no', i % 5, i * i, i % 7, names[i % len(names)], i % 4, None,
                     day,
                     None if i % 4 == 0 else (Formula('1/0') if i % 4 == 1 else i)])
    return rows


def _cf_facts(sheet):
    facts = {}
    conditions = sheet.Range('A2:J20').FormatConditions
    for index in range(1, conditions.Count + 1):
        condition = conditions(index)
        facts[index] = tuple(_get(read) for read in (
            lambda: condition.Type, lambda: condition.Operator, lambda: condition.Formula1,
            lambda: condition.Formula2, lambda: condition.Priority, lambda: condition.StopIfTrue,
            lambda: condition.AppliesTo.Address, lambda: condition.Font.Bold, lambda: condition.Font.Italic,
            lambda: condition.Font.Color, lambda: condition.Font.Name, lambda: condition.Font.Size,
            lambda: condition.Font.Underline, lambda: condition.Font.Strikethrough,
            lambda: condition.Interior.Color, lambda: condition.Interior.PatternColor,
            lambda: condition.Interior.Pattern, lambda: condition.NumberFormat,
            lambda: condition.Borders(1).LineStyle, lambda: condition.Borders(4).LineStyle,
            lambda: condition.TopBottom, lambda: condition.Rank, lambda: condition.Percent,
            lambda: condition.AboveBelow, lambda: condition.NumStdDev, lambda: condition.Text,
            lambda: condition.TextOperator, lambda: condition.DateOperator, lambda: condition.DupeUnique,
        ))
    for row in range(2, 21):
        for column in 'ABCDEFGIJ':
            shown = sheet.Range(f'{column}{row}').DisplayFormat
            facts[f'{column}{row}'] = (shown.Interior.Color, shown.Font.Bold, shown.Font.Color, shown.NumberFormat)
    return facts


def test_conditional_formats_match_xlsx_in_excel(tmp_path, open_workbook):
    facts = {}
    for writer_class, suffix in ((XlsxWriter, 'xlsx'), (XlsbWriter, 'xlsb')):
        path = tmp_path / f'rules.{suffix}'
        with writer_class(path) as writer:
            writer.add_sheet('Data')
            writer.write_sheet(_cf_rows(), conditional_formats=_cf_rules())
        book = open_workbook(path)
        book.Application.Calculate()
        facts[suffix] = _cf_facts(book.Worksheets('Data'))
    assert len([key for key in facts['xlsb'] if isinstance(key, int)]) == len(_cf_rules())
    differences = {key: (facts['xlsx'][key], facts['xlsb'].get(key)) for key in facts['xlsx']
                   if facts['xlsx'][key] != facts['xlsb'].get(key)}
    assert not differences
