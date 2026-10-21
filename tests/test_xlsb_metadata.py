"""Binary metadata round trips, preservation, transactions and formula tokens."""

import io
import struct
import tracemalloc
import zipfile

import pytest

from xlspy import (
    Comment, DefinedName, ExcelReader, Formula, SheetProtection, Style,
    WorkbookProtection, XlsbWriter, XlsbUpdater,
)
from xlspy._xlsb_metadata import records, parse_comments, comments_bin
from xlspy.biff_utils import build_record
from xlspy._xlsb_formulas import FormulaContext, compile_formula, decode_formula, parsed_formula
from xlspy.comments import find_comment_parts
from xlspy import reader as reader_module
from test_comments_and_editing import payloads, replace_parts


@pytest.mark.parametrize('expression', ['1e999', 'UnknownFunction(1)', 'Table1[Amount]', '[other.xlsx]Data!A1', '{1,2}'])
def test_binary_formula_rejects_unsupported_or_invalid_expression(expression):
    with pytest.raises((ValueError, NotImplementedError)):
        compile_formula(expression, FormulaContext(['Data']))


def make_binary(path, **options):
    with XlsbWriter(path) as writer:
        writer.add_sheet('Data')
        writer.write_sheet([['Amount','Double'], [10,20], [20,40], [30,60]], **options)
        writer.add_sheet('Other')
        writer.write_sheet([['unchanged']])


@pytest.mark.parametrize('shared',[True,False])
def test_notes_binary_roundtrip(tmp_path,shared):
    path=tmp_path/'notes.xlsb'
    with XlsbWriter(path,useSharedStrings=shared) as writer:
        writer.add_sheet('Data')
        writer.write_sheet([[1]],comments={'A1':Comment('Zażółć <&>\n_x0041_\rline','Jan'), 'D9':Comment('empty','Ola')})
    with ExcelReader(str(path)) as reader:
        assert reader.get_comments('Data') == {'A1':Comment('Zażółć <&>\n_x0041_\rline','Jan'), 'D9':Comment('empty','Ola')}
    updater=XlsbUpdater(path)
    try:
        updater.set_comments('Data',{'A1':None,'B3':Comment('new','Jan')})
        updater.save()
    finally:
        updater.close()
    with ExcelReader(str(path)) as reader:
        assert reader.get_comments('Data') == {'D9':Comment('empty','Ola'),'B3':Comment('new','Jan')}


def test_binary_note_cleanup(tmp_path):
    path=tmp_path/'cleanup.xlsb'
    make_binary(path,comments={'A1':Comment('last','Jan')})
    with zipfile.ZipFile(path) as archive:
        original_parts=find_comment_parts(archive.read,archive.namelist(),'xl/worksheets/sheet1.bin')[2]
    updater=XlsbUpdater(path)
    try:
        updater.set_comments('Data',{'A1':None})
        updater.save()
    finally:
        updater.close()
    parts=payloads(path)
    assert all(part[0] not in parts for part in original_parts.values())
    assert not any(r.record_id==0x227 for r,p in records(parts['xl/worksheets/sheet1.bin']))


@pytest.mark.parametrize('stream',[False,True])
def test_binary_notes_survive_data_replacement(tmp_path,stream):
    path=tmp_path/'replace.xlsb'
    make_binary(path,comments={'D20':Comment('keep','Jan')})
    updater=XlsbUpdater(path)
    try:
        method=updater.replace_sheet_data_stream if stream else updater.replace_sheet_data
        method('Data',[[1]])
        updater.save()
    finally:
        updater.close()
    with ExcelReader(str(path)) as reader:
        assert reader.get_comments('Data') == {'D20':Comment('keep','Jan')}


@pytest.mark.parametrize('force_python',[False,True])
def test_binary_formula_caches_and_text(tmp_path,force_python,monkeypatch):
    if force_python:
        monkeypatch.setattr(reader_module,'HAVE_C_EXT',False)
    path=tmp_path/'formulas.xlsb'
    with XlsbWriter(path) as writer:
        writer.set_defined_name(DefinedName('Scale','2'))
        writer.add_sheet('Data')
        writer.write_sheet([['Amount','Result','Label'],[10,Formula('A2*Scale',20),Formula('IF(A2>0,"yes","no")','yes')]])
    with ExcelReader(str(path)) as reader:
        assert reader.read_all('Data')[1] == [10,20,'yes']
        assert reader.read_all('Data',data_only=False)[1] == [10,'=A2*Scale','=IF(A2>0,"yes","no")']


def test_sparse_binary_edits_preserve_styles_notes_and_other_parts(tmp_path):
    path=tmp_path/'sparse.xlsb'
    make_binary(path,comments={'A2':Comment('keep','Jan')},styles={'A2':Style(bold=True)})
    before=payloads(path)
    updater=XlsbUpdater(path)
    try:
        updater.update_cells('Data',{'A2':15,'C2':Formula('A2*2',30),'D8':'Zażółć\x01\r\n','B3':None})
        updater.save()
    finally:
        updater.close()
    after=payloads(path)
    for part in ('xl/styles.bin','xl/worksheets/sheet2.bin','xl/comments/comment1.bin','xl/drawings/commentsDrawing1.vml'):
        assert before[part] == after[part]
    with ExcelReader(str(path)) as reader:
        assert reader.get_comments('Data') == {'A2':Comment('keep','Jan')}
        rows=reader.read_all('Data',data_only=False)
        assert rows[1][:3] == [15,20,'=A2*2']
        assert rows[-1][3] == 'Zażółć\x01\r\n'
    cells=[p for r,p in records(after['xl/worksheets/sheet1.bin']) if r.record_id==5 and struct.unpack_from('<I',p)[0]==0]
    assert any(struct.unpack_from('<I',p,4)[0]>0 for p in cells)


def test_sparse_binary_clear_keeps_style(tmp_path):
    path=tmp_path/'clear.xlsb'
    make_binary(path,styles={'A2':Style(bold=True)},comments={'A2':Comment('keep','Jan')})
    updater=XlsbUpdater(path)
    try:
        updater.update_cells('Data',{'A2':None})
        updater.save()
    finally:
        updater.close()
    assert any(r.record_id==1 and struct.unpack_from('<I',p,4)[0]>0 for r,p in records(payloads(path)['xl/worksheets/sheet1.bin']))


def test_binary_names_and_protection(tmp_path):
    path=tmp_path/'names.xlsb'
    with XlsbWriter(path) as writer:
        writer.add_sheet('Data')
        writer.write_sheet([[1]],sheet_protection=SheetProtection(password='abc'))
        writer.set_defined_name(DefinedName('Amount',"'Data'!$A$1"))
        writer.set_defined_name(DefinedName('Amount','42','Data'))
        writer.set_workbook_protection(WorkbookProtection(password='abc'))
    updater=XlsbUpdater(path)
    try:
        updater.set_defined_name(DefinedName('amount','99'))
        updater.remove_defined_name('Amount','Data')
        assert updater.get_defined_names()==[DefinedName('amount','99')]
        updater.set_sheet_protection('Data',SheetProtection(enabled=False))
        updater.set_workbook_protection(WorkbookProtection(lock_structure=False))
        updater.save()
    finally:
        updater.close()
    with ExcelReader(str(path)) as reader:
        assert reader.get_defined_names()==[DefinedName('amount','99')]
    assert not any(r.record_id==0x217 for r,p in records(payloads(path)['xl/worksheets/sheet1.bin']))
    assert not any(r.record_id==0x216 for r,p in records(payloads(path)['xl/workbook.bin']))


@pytest.mark.parametrize('changes',[{'A2':object()},{'A2':float('nan')},{'A1:B2':1},{'A2':Formula('1',array_ref='A2:A3')}])
def test_binary_point_edit_failures_are_atomic(tmp_path,changes):
    path=tmp_path/'atomic.xlsb'
    make_binary(path)
    updater=XlsbUpdater(path)
    try:
        before=updater.to_bytes()
        with pytest.raises((ValueError,TypeError,NotImplementedError)):
            updater.update_cells('Data',changes)
        assert updater.to_bytes()==before
    finally:
        updater.close()


def test_binary_staging_failure_rolls_back(tmp_path,monkeypatch):
    path=tmp_path/'rollback.xlsb'
    make_binary(path)
    updater=XlsbUpdater(path)
    try:
        before=updater.to_bytes()
        original=updater._package.set
        def failing(name,data):
            if name=='[Content_Types].xml':
                raise OSError('simulated failure')
            original(name,data)
        monkeypatch.setattr(updater._package,'set',failing)
        with pytest.raises(OSError,match='simulated'):
            updater.set_comments('Data',{'A1':Comment('new','Jan')})
        assert updater.to_bytes()==before
    finally:
        updater.close()


@pytest.mark.parametrize('formula',[
    'A1+$B2+C$3+$D$4','SUM(A1:B3)','IF(A1>0,"yes","no")','ROUND(A1/3,2)',
    'IFERROR(1/0,99)','-A1^2','(A1+B1)*2','A1&"x"','SUM(A:A)','SUM(1:3)',
    'CONCAT("a","b")','COUNTIF(A1:A3,">0")','TRUE','1.25E3','SUM(A1,,B1)',
])
def test_formula_token_roundtrips(formula):
    context=FormulaContext(['Data'])
    encoded=compile_formula(formula,context,sheet='Data')
    decoded=decode_formula(encoded,context,sheet='Data')
    assert decoded.removeprefix('_xlfn.') == formula.replace('1.25E3','1250')


def test_array_binary_text_and_group_edit_guard(tmp_path):
    path=tmp_path/'array.xlsb'
    with XlsbWriter(path) as writer:
        writer.add_sheet('Data')
        writer.write_sheet([['Input','Array'],[10,Formula('A2:A4*3',30,'B2:B4')],[20,60],[30,90]])
    with ExcelReader(str(path)) as reader:
        assert reader.read_all('Data',data_only=False) == [['Input','Array'],[10,'=A2:A4*3'],[20,60],[30,90]]
    updater=XlsbUpdater(path)
    try:
        before=updater.to_bytes()
        with pytest.raises(ValueError,match='array/shared'):
            updater.update_cells('Data',{'B3':5})
        assert updater.to_bytes()==before
    finally:
        updater.close()


@pytest.mark.parametrize('force_python',[False,True])
def test_later_wider_binary_rows_not_truncated(tmp_path,force_python,monkeypatch):
    if force_python:
        monkeypatch.setattr(reader_module,'HAVE_C_EXT',False)
    path=tmp_path/'wide.xlsb'
    with XlsbWriter(path) as writer:
        writer.write_sheet([[1],[2,3,4],[5]])
    with ExcelReader(str(path)) as reader:
        assert reader.read_all('Sheet1')==[[1],[2,3,4],[5,None,None]]


def test_streamed_binary_rows_are_incremental(tmp_path,monkeypatch):
    path=tmp_path/'stream.xlsb'
    with XlsbWriter(path,useSharedStrings=False) as writer:
        writer.write_sheet(([i,'x'*100] for i in range(10000)))
    with ExcelReader(str(path)) as reader:
        original=reader._archive.open
        reads=[]
        class Stream:
            def __init__(self,source): self.source=source
            def read(self,size=-1):
                assert size>=0
                data=self.source.read(size)
                reads.append(len(data))
                return data
            def __enter__(self): return self
            def __exit__(self,*args): self.source.close()
        def wrapped(name,*args,**kwargs):
            source=original(name,*args,**kwargs)
            return Stream(source) if name=='xl/worksheets/sheet1.bin' else source
        monkeypatch.setattr(reader._archive,'open',wrapped)
        rows=reader.get_rows('Sheet1',streaming=True)
        assert next(rows)==[0,'x'*100]
        assert sum(reads)<reader._archive.getinfo('xl/worksheets/sheet1.bin').file_size
        rows.close()


def make_shared_binary(path):
    make_binary(path)
    parts=payloads(path)
    original=parts['xl/worksheets/sheet1.bin']
    output=[]
    row=-1
    expression=parsed_formula(b'\x01'+struct.pack('<I',1),struct.pack('<I',1))
    shared=parsed_formula(b'\x4c'+struct.pack('<IH',0,0xFFFF)+b'\x1e\x02\0\x05')
    for record,payload in records(original):
        if record.record_id==0:
            row=struct.unpack_from('<I',payload)[0]
        if row in (1,2) and record.record_id in range(1,12) and struct.unpack_from('<I',payload)[0]==1:
            output.append(build_record(9,payload[:8]+struct.pack('<d',row*20)+b'\x01\0'+expression))
            if row==1:
                output.append(build_record(0x1AB,struct.pack('<IIII',1,2,1,1)+shared))
        else:
            output.append(original[record.header_start:record.data_end])
    replace_parts(path,{'xl/worksheets/sheet1.bin':b''.join(output)})


def test_shared_binary_formula_reading_and_edit_guard(tmp_path):
    path=tmp_path/'shared.xlsb'
    make_shared_binary(path)
    with ExcelReader(str(path)) as reader:
        assert reader.read_all('Data',data_only=False)[1:3]==[[10,'=A2*2'],[20,'=A3*2']]
        assert reader.read_all('Data')[1:3]==[[10,20],[20,40]]
    updater=XlsbUpdater(path)
    try:
        with pytest.raises(ValueError,match='array/shared'):
            updater.update_cells('Data',{'B3':1})
    finally:
        updater.close()


@pytest.mark.parametrize('streamed',[False,True])
def test_binary_metadata_interleaves_with_existing_bulk_updates(tmp_path,streamed):
    path=tmp_path/'interleave.xlsb'
    make_binary(path)
    updater=XlsbUpdater(path)
    try:
        method=updater.replace_sheet_data_stream if streamed else updater.replace_sheet_data
        method('Data',[['first'],['second']])
        updater.update_cells('Data',{'A1':'inline'})
        method('Other',[['third']])
        updater.set_comments('Data',{'D9':Comment('keep','Jan')})
        updater.save()
    finally:
        updater.close()
    with ExcelReader(str(path)) as reader:
        assert reader.read_all('Data')==[['inline'],['second']]
        assert reader.read_all('Other')==[['third']]
    parts=payloads(path)
    count=next(struct.unpack('<II',payload)[0] for record,payload in records(parts['xl/sharedStrings.bin']) if record.record_id==0x9F)
    actual=sum(sum(record.record_id==7 for record,payload in records(parts[name])) for name in ('xl/worksheets/sheet1.bin','xl/worksheets/sheet2.bin'))
    assert count==actual


@pytest.mark.parametrize('c_ext', [True, False])
def test_streaming_and_formula_rows_match_bulk_shape(tmp_path, monkeypatch, c_ext):
    if not c_ext:
        monkeypatch.setattr(reader_module, 'HAVE_C_EXT', False)
    path = tmp_path / 'gaps.xlsb'
    with XlsbWriter(path) as writer:
        writer.add_sheet('Data')
        writer.write_sheet([[1, 2], [1, 2, 3], [1]])
    updated = tmp_path / 'updated.xlsb'
    updater = XlsbUpdater(path)
    try:
        updater.update_cells('Data', {'B7': 5, 'D9': None})
        updater.save(updated)
    finally:
        updater.close()
    with ExcelReader(str(updated)) as reader:
        bulk = reader.read_all('Data')
        assert reader.read_all('Data', streaming=True) == bulk
        assert reader.read_all('Data', data_only=False) == bulk
    assert bulk[6] == [None, 5, None] and len(bulk) == 9


def test_sheet_qualified_local_name_may_be_defined_after_use(tmp_path):
    path = tmp_path / 'names.xlsb'
    with XlsbWriter(path) as writer:
        writer.add_sheet('Data')
        writer.write_sheet([[Formula('Data!Rate*2'), 3]])
        writer.set_defined_name(DefinedName('Rate', 'Data!$B$1', 'Data'))
    with ExcelReader(str(path)) as reader:
        assert reader.get_defined_names() == [DefinedName('Rate', "'Data'!$B$1", 'Data')]
        assert reader.read_all('Data', data_only=False) == [['=Rate*2', 3]]


def test_array_formula_member_with_date_value(tmp_path):
    import datetime
    path = tmp_path / 'array.xlsb'
    with XlsbWriter(path) as writer:
        writer.add_sheet('Data')
        writer.write_sheet([[Formula('C1:D1', array_ref='A1:B1'), datetime.date(2020, 1, 1)]])
    with ExcelReader(str(path)) as reader:
        assert reader.read_all('Data') == [[0, 43831]]
