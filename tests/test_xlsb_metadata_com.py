"""Native Excel checks for XLSB metadata and tokenized formulas."""

import os

import pytest

from xlspy import Comment, DefinedName, Formula, Protection, SheetProtection, Style, WorkbookProtection, XlsbWriter, XlsbUpdater, ExcelReader, ExcelTable, Chart, ChartSeries
from test_excel_com_metadata import excel, open_workbook  # noqa: F401

pytestmark = [pytest.mark.com, pytest.mark.skipif(os.environ.get('XLSPY_RUN_EXCEL_COM') != '1', reason='Set XLSPY_RUN_EXCEL_COM=1')]


@pytest.mark.parametrize(('text','expected'), [
    ('A2*2', 20), ('SUM(A2:A4)', 60), ('IF(A2>0,"yes","no")', 'yes'),
    ('ROUND(A2/3,2)', 3.33), ('Other!A1+$A$2', 15), ('A2*Scale', 20),
    ('CONCAT("ok","!")', 'ok!'), ('IFERROR(1/0,99)', 99),
    ('CHOOSE(2,10,20,30)',20), ('SUMIFS(A2:A4,A2:A4,">10")',50),
    ('XLOOKUP(20,A2:A4,A2:A4)',20), ('IF(A2<0,"yes","no")','no'),
    ('other!A1',5), ('IFERROR(42,99)',42), ('COUNTIFS(A2:A4,">10",A2:A4,"<30")',1),
])
def test_binary_formula_in_excel(tmp_path, open_workbook, text, expected):
    path = tmp_path / 'formula.xlsb'
    with XlsbWriter(path) as writer:
        writer.set_defined_name(DefinedName('Scale', '2'))
        writer.add_sheet('Data')
        # An intentionally stale cache proves Excel calculated the tokens.
        writer.write_sheet([['Amount','Result'], [10,Formula(text,-777)], [20], [30]])
        writer.add_sheet('Other')
        writer.write_sheet([[5]])
    book = open_workbook(path)
    assert book.Worksheets('Data').Range('B2').Value == expected


def test_notes_names_protection_in_excel(tmp_path, open_workbook):
    path = tmp_path / 'metadata.xlsb'
    with XlsbWriter(path) as writer:
        writer.add_sheet('Data')
        writer.write_sheet([['Amount'],[10]], comments={'A2':Comment('Zażółć <&>\nline','Jan',300,120),'D9':Comment('outside','Ola')}, sheet_protection=SheetProtection(password='abc'), styles={'A2':Style(protection=Protection(locked=False))})
        writer.set_defined_name(DefinedName('Amount',"'Data'!$A$2"))
        writer.set_workbook_protection(WorkbookProtection(password='abc'))
    book = open_workbook(path)
    sheet = book.Worksheets('Data')
    assert sheet.Range('A2').Comment.Text() == 'Zażółć <&>\nline'
    assert sheet.Range('A2').Comment.Author == 'Jan'
    assert sheet.Range('A2').Comment.Shape.Width == pytest.approx(225,abs=1)
    assert sheet.Range('A2').Comment.Shape.Height == pytest.approx(90,abs=1)
    assert sheet.Range('D9').Comment.Text() == 'outside'
    assert sheet.ProtectContents
    assert not sheet.Range('A2').Locked
    assert book.ProtectStructure
    assert book.Names('Amount').RefersToRange.Value == 10


@pytest.mark.parametrize('array_ref', ['B2:B4', 'B2:C4'])
def test_array_formula_creation_in_excel(tmp_path, open_workbook, array_ref):
    path = tmp_path / 'array.xlsb'
    with XlsbWriter(path) as writer:
        writer.add_sheet('Data')
        writer.write_sheet([['Input','Array','Array2'], [10,Formula('A2:A4*3',30,array_ref)], [20], [30]])
    book = open_workbook(path)
    sheet = book.Worksheets('Data')
    assert sheet.Range('B2:B4').Value == ((30.0,), (60.0,), (90.0,))
    assert sheet.Range('B2').HasArray
    if array_ref == 'B2:C4':
        assert sheet.Range('C2:C4').Value == ((30.0,), (60.0,), (90.0,))


def test_array_extends_past_input_rows(tmp_path, open_workbook):
    path = tmp_path / 'array_tail.xlsb'
    with XlsbWriter(path) as writer:
        writer.add_sheet('Data')
        writer.write_sheet([['Array'], [Formula('ROW(A2:A5)',2,'A2:A5')]])
    book = open_workbook(path)
    assert book.Worksheets('Data').Range('A2:A5').Value == ((2.0,), (3.0,), (4.0,), (5.0,))


@pytest.fixture
def excel_binary_template(tmp_path, excel):
    path = tmp_path / 'excel_template.xlsb'
    book = excel.Workbooks.Add()
    try:
        while book.Worksheets.Count > 1:
            book.Worksheets(book.Worksheets.Count).Delete()
        sheet = book.Worksheets(1)
        sheet.Name = 'Data'
        other = book.Worksheets.Add(After=sheet)
        other.Name = 'Other'
        other.Range('A1').Value = 5
        sheet.Range('A1:G1').Value = (('Input','Double','Calculated','Note','Array','Reserved','New'),)
        sheet.Range('A2:A101').Value = tuple((i,) for i in range(1,101))
        sheet.Range('B2:B101').FormulaR1C1 = '=RC[-1]*2'
        book.Names.Add(Name='Scale',RefersTo='=2')
        sheet.Range('C2').Formula = '=A2*Scale'
        sheet.Range('E2:E4').FormulaArray = '=A2:A4*3'
        sheet.Range('A1').AddComment('original note')
        sheet.Range('D9').AddComment('keep outside data')
        button = sheet.Buttons().Add(20,20,80,24)
        button.Name = 'PreservedButton'
        button.Caption = 'Keep me'
        book.Names.Add(Name='KeepName',RefersTo='=Other!$A$1')
        sheet.Names.Add(Name='LocalAmount',RefersTo='=Data!$A$3')
        sheet.Protect(Password='abc')
        book.Protect(Password='abc',Structure=True)
        book.SaveAs(str(path),FileFormat=50)
    finally:
        book.Close(False)
        button = sheet = other = book = None
        import gc
        gc.collect()
    return path


def test_read_excel_generated_binary_formulas_and_metadata(excel_binary_template):
    with ExcelReader(str(excel_binary_template)) as reader:
        comments=reader.get_comments('Data')
        assert comments['A1'].text == 'original note'
        names={name.name:name for name in reader.get_defined_names()}
        assert names['Scale'].formula == '2'
        assert names['LocalAmount'].scope == 'Data'
        rows=reader.read_all('Data',data_only=False)
        assert rows[1][1] == '=A2*2'
        assert rows[2][1] == '=A3*2'
        assert rows[1][4] == '=A2:A4*3'
        assert rows[2][4] == 6
        assert rows[1][2] == '=A2*Scale'


def test_update_excel_generated_binary_preserves_button_and_formulas(excel_binary_template,open_workbook):
    path=excel_binary_template
    updater=XlsbUpdater(path)
    try:
        updater.set_comments('Data',{'A1':None,'C8':Comment('new note','Jan',240,100)})
        updater.update_cells('Data',{'A2':15,'G2':Formula('SUM(A2:A4)',20)})
        updater.set_defined_name(DefinedName('Scale','3'))
        updater.remove_defined_name('LocalAmount','Data')
        updater.set_sheet_protection('Data',SheetProtection(enabled=False))
        updater.set_workbook_protection(WorkbookProtection(password='next'))
        updater.save()
    finally:
        updater.close()
    book=open_workbook(path)
    sheet=book.Worksheets('Data')
    assert sheet.Range('A2').Value == 15
    assert sheet.Range('B2').Value == 30
    assert sheet.Range('C2').Value == 45
    assert sheet.Range('G2').Value == 20
    assert sheet.Range('A1').Comment is None
    assert sheet.Range('D9').Comment.Text() == 'keep outside data'
    assert sheet.Range('C8').Comment.Text() == 'new note'
    assert sheet.Buttons('PreservedButton').Caption == 'Keep me'
    assert sheet.Range('E2:E4').Value == ((45.0,), (6.0,), (9.0,))
    assert book.Names('KeepName').RefersToRange.Value == 5
    assert not sheet.ProtectContents
    assert book.ProtectStructure
    book.Unprotect('next')
    assert not book.ProtectStructure
    with ExcelReader(str(path)) as reader:
        assert 'LocalAmount' not in {name.name for name in reader.get_defined_names()}


def test_comments_coexist_with_binary_tables_and_charts(tmp_path,open_workbook):
    path=tmp_path/'report.xlsb'
    with XlsbWriter(path) as writer:
        writer.add_sheet('Data')
        writer.write_sheet([['Product','Units'],['Tea',10],['Coffee',20]],
                           tables=[ExcelTable('Sales','A1:B3')],
                           charts=[Chart('bar',[ChartSeries('B2:B3',categories='A2:A3')])],
                           comments={'B2':Comment('check','Jan')})
    book=open_workbook(path)
    sheet=book.Worksheets('Data')
    assert sheet.ListObjects.Count==1
    assert sheet.ChartObjects().Count==1
    assert sheet.Range('B2').Comment.Text()=='check'


def test_shared_binary_formulas_open_in_excel(tmp_path,open_workbook):
    from test_xlsb_metadata import make_shared_binary
    path=tmp_path/'shared.xlsb'
    make_shared_binary(path)
    book=open_workbook(path)
    sheet=book.Worksheets('Data')
    assert sheet.Range('B2:B3').Value==((20.0,),(40.0,))
    assert sheet.Range('B2:B3').Formula==(('=A2*2',),('=A3*2',))


def test_delete_all_binary_notes_preserves_excel_button(excel_binary_template,open_workbook):
    updater=XlsbUpdater(excel_binary_template)
    try:
        updater.set_comments('Data',{'A1':None,'D9':None})
        updater.save()
    finally:
        updater.close()
    book=open_workbook(excel_binary_template)
    sheet=book.Worksheets('Data')
    assert sheet.Comments.Count==0
    assert sheet.Buttons('PreservedButton').Caption=='Keep me'


def test_deleted_binary_name_does_not_shift_formula_references(tmp_path, open_workbook):
    path = tmp_path / 'names.xlsb'
    with XlsbWriter(path) as writer:
        writer.set_defined_name(DefinedName('FirstValue', '2'))
        writer.set_defined_name(DefinedName('SecondValue', '3'))
        writer.add_sheet('Data')
        writer.write_sheet([[Formula('FirstValue', 2), Formula('SecondValue', 3)]])
    updater = XlsbUpdater(path)
    try:
        updater.remove_defined_name('FirstValue')
        updater.set_defined_name(DefinedName('ThirdValue', '4'))
        updater.save()
    finally:
        updater.close()
    book = open_workbook(path)
    sheet = book.Worksheets('Data')
    assert sheet.Range('A1').Value != 2  # Excel localizes the displayed #NAME? text.
    assert sheet.Range('B1').Value == 3
