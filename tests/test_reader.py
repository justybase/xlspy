import io
import os
import shutil
import sys
import tempfile
import zipfile
import datetime

import pytest

# ``pythonpath = ["src"]`` in pyproject.toml makes this redundant under pytest;
# the insert below keeps ``python tests/test_reader.py`` working as well.
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '../src')))

from xlspy import XlsbWriter, XlsxWriter, ExcelReader, F
from xlspy import reader as xlspy_reader


def test_read_xlsb():
    """Write data with XlsbWriter, then read it back with ExcelReader."""
    tmpfile = tempfile.mktemp(suffix='.xlsb')
    try:
        data = [
            ['Name', 'Age', 'Active', 'Score', 'Birth'],
            ['Alice', 30, True, 95.5, datetime.date(1990, 5, 15)],
            ['Bob', 25, False, 88.0, datetime.date(1995, 8, 20)],
            ['Charlie', 35, True, 100.0, datetime.date(1985, 3, 10)],
            [None, None, None, None, None],
            ['Eve', 28, True, 72.5, datetime.date(1992, 12, 1)],
        ]

        with XlsbWriter(tmpfile) as writer:
            writer.add_sheet("TestSheet1")
            writer.write_sheet(data)

        with ExcelReader() as reader:
            reader.open(tmpfile)
            names = reader.get_sheet_names()
            assert 'TestSheet1' in names, f"Sheet not found: {names}"
            print(f"  Sheets: {names}")

            rows = reader.read_all('TestSheet1')
            assert len(rows) == len(data), f"Expected {len(data)} rows, got {len(rows)}"
            for i, (expected, actual) in enumerate(zip(data, rows)):
                for j, (e, a) in enumerate(zip(expected, actual)):
                    if e is None:
                        assert a is None or a == '', f"Row {i} col {j}: expected None, got {a!r}"
                    elif isinstance(e, datetime.date) and not isinstance(e, datetime.datetime):
                        assert isinstance(a, (datetime.date, datetime.datetime)), f"Row {i} col {j}: expected date, got {type(a)} {a!r}"
                    elif isinstance(e, bool):
                        assert a == e, f"Row {i} col {j}: expected {e}, got {a!r}"
                    elif isinstance(e, int):
                        assert a == e, f"Row {i} col {j}: expected {e}, got {a!r} (type={type(a)})"
                    elif isinstance(e, float):
                        assert abs(a - e) < 0.001, f"Row {i} col {j}: expected {e}, got {a!r}"
                    else:
                        assert a == e, f"Row {i} col {j}: expected {e!r}, got {a!r}"
            print(f"  Read {len(rows)} rows, all values match!")
    finally:
        os.remove(tmpfile)


def test_read_xlsx():
    """Write data with XlsxWriter, then read it back with ExcelReader."""
    tmpfile = tempfile.mktemp(suffix='.xlsx')
    try:
        data = [
            ['Name', 'Age', 'Active', 'Score', 'Birth'],
            ['Alice', 30, True, 95.5, datetime.date(1990, 5, 15)],
            ['Bob', 25, False, 88.0, datetime.date(1995, 8, 20)],
            ['Charlie', 35, True, 100.0, datetime.date(1985, 3, 10)],
            ['Eve', 28, True, 72.5, datetime.date(1992, 12, 1)],
        ]

        with XlsxWriter(tmpfile) as writer:
            writer.add_sheet("TestSheet1")
            writer.write_sheet(data)

        with ExcelReader() as reader:
            reader.open(tmpfile)
            names = reader.get_sheet_names()
            assert 'TestSheet1' in names, f"Sheet not found: {names}"
            print(f"  Sheets: {names}")

            rows = reader.read_all('TestSheet1')
            assert len(rows) == len(data), f"Expected {len(data)} rows, got {len(rows)}"
            for i, (expected, actual) in enumerate(zip(data, rows)):
                for j, (e, a) in enumerate(zip(expected, actual)):
                    if e is None:
                        assert a is None or a == '', f"Row {i} col {j}: expected None, got {a!r}"
                    elif isinstance(e, datetime.date) and not isinstance(e, datetime.datetime):
                        assert isinstance(a, (datetime.date, datetime.datetime)), f"Row {i} col {j}: expected date, got {type(a)} {a!r}"
                    elif isinstance(e, bool):
                        assert a == e, f"Row {i} col {j}: expected {e}, got {a!r}"
                    elif isinstance(e, int):
                        assert a == e, f"Row {i} col {j}: expected {e}, got {a!r} (type={type(a)})"
                    elif isinstance(e, float):
                        assert abs(a - e) < 0.001, f"Row {i} col {j}: expected {e}, got {a!r}"
                    else:
                        assert a == e, f"Row {i} col {j}: expected {e!r}, got {a!r}"
            print(f"  Read {len(rows)} rows, all values match!")
    finally:
        os.remove(tmpfile)


def test_multi_sheet_xlsb():
    """Test reading multiple sheets from XLSB."""
    tmpfile = tempfile.mktemp(suffix='.xlsb')
    try:
        with XlsbWriter(tmpfile) as writer:
            writer.add_sheet("Sheet1")
            writer.write_sheet([['A', 'B'], [1, 2]])
            writer.add_sheet("Sheet2")
            writer.write_sheet([['X', 'Y', 'Z'], ['a', 'b', 'c']])

        with ExcelReader() as reader:
            reader.open(tmpfile)
            names = reader.get_sheet_names()
            print(f"  Sheets: {names}")
            assert len(names) == 2
            assert 'Sheet1' in names
            assert 'Sheet2' in names

            rows1 = reader.read_all('Sheet1')
            assert len(rows1) == 2
            assert rows1[0] == ['A', 'B']

            rows2 = reader.read_all('Sheet2')
            assert len(rows2) == 2
            assert rows2[1] == ['a', 'b', 'c']
            print(f"  Multi-sheet XLSB OK!")
    finally:
        os.remove(tmpfile)


def test_multi_sheet_xlsx():
    """Test reading multiple sheets from XLSX."""
    tmpfile = tempfile.mktemp(suffix='.xlsx')
    try:
        with XlsxWriter(tmpfile) as writer:
            writer.add_sheet("Sheet1")
            writer.write_sheet([['A', 'B'], [1, 2]])
            writer.add_sheet("Sheet2")
            writer.write_sheet([['X', 'Y', 'Z'], ['a', 'b', 'c']])

        with ExcelReader() as reader:
            reader.open(tmpfile)
            names = reader.get_sheet_names()
            print(f"  Sheets: {names}")
            assert len(names) == 2
            assert 'Sheet1' in names
            assert 'Sheet2' in names

            rows1 = reader.read_all('Sheet1')
            assert len(rows1) == 2
            assert rows1[0] == ['A', 'B']

            rows2 = reader.read_all('Sheet2')
            assert len(rows2) == 2
            assert rows2[1] == ['a', 'b', 'c']
            print(f"  Multi-sheet XLSX OK!")
    finally:
        os.remove(tmpfile)


def test_row_generator_xlsb():
    """Test reading XLSB using generator."""
    tmpfile = tempfile.mktemp(suffix='.xlsb')
    try:
        with XlsbWriter(tmpfile) as writer:
            writer.add_sheet("Data")
            writer.write_sheet([['h1', 'h2'], ['v1', 'v2'], ['v3', 'v4']])

        with ExcelReader() as reader:
            reader.open(tmpfile)
            count = 0
            for row in reader.get_rows("Data"):
                count += 1
                assert len(row) == 2
            assert count == 3, f"Expected 3 rows (header + 2 data), got {count}"
            print(f"  Generator XLSB OK!")
    finally:
        os.remove(tmpfile)


def test_sheet_not_found():
    """Test error on reading non-existent sheet."""
    tmpfile = tempfile.mktemp(suffix='.xlsb')
    try:
        with XlsbWriter(tmpfile) as writer:
            writer.add_sheet("OnlySheet")
            writer.write_sheet([['a']])

        with ExcelReader() as reader:
            reader.open(tmpfile)
            try:
                reader.read_all("NonExistent")
                assert False, "Should have raised an exception"
            except (KeyError, ValueError):
                print(f"  Sheet not found error correctly raised!")
    finally:
        os.remove(tmpfile)


def test_read_complex_xlsx_values():
    """Test reading various data types from XLSX."""
    tmpfile = tempfile.mktemp(suffix='.xlsx')
    try:
        data = [
            ['Text', 'Integer', 'Float', 'Boolean', 'Date'],
            ['Hello', 42, 3.14, True, '2024-01-15'],
            ['World', -7, -0.001, False, '2024-06-30'],
            ['', 0, 1e10, True, ''],
            ['Spaces  ', 1000000, 0.5, False, None],
        ]

        with XlsxWriter(tmpfile) as writer:
            writer.add_sheet("Types")
            writer.write_sheet(data)

        with ExcelReader() as reader:
            reader.open(tmpfile)
            rows = reader.read_all('Types')
            assert len(rows) == len(data)
            for i, (expected, actual) in enumerate(zip(data, rows)):
                for j, (e, a) in enumerate(zip(expected, actual)):
                    if e is None:
                        assert a is None or a == '', f"Row {i} col {j}: expected None, got {a!r}"
                    elif isinstance(e, str):
                        assert a == e, f"Row {i} col {j}: expected {e!r}, got {a!r}"
                    elif isinstance(e, bool):
                        assert a == e, f"Row {i} col {j}: expected {e}, got {a!r}"
                    elif isinstance(e, int):
                        assert a == e, f"Row {i} col {j}: expected {e}, got {a!r} (type={type(a)})"
                    elif isinstance(e, float):
                        assert abs(a - e) < 0.001, f"Row {i} col {j}: expected {e}, got {a!r}"
            print(f"  Complex XLSX values OK!")
    finally:
        os.remove(tmpfile)


def test_xlsx_writer_with_bytesio():
    """Write XLSX to BytesIO, then read back with ExcelReader from file."""
    buf = io.BytesIO()
    data = [['Name', 'Value'], ['foo', 42]]
    with XlsxWriter(buf) as writer:
        writer.add_sheet("Sheet1")
        writer.write_sheet(data)

    tmpfile = tempfile.mktemp(suffix='.xlsx')
    try:
        with open(tmpfile, 'wb') as f:
            f.write(buf.getvalue())
        with ExcelReader(tmpfile) as reader:
            names = reader.get_sheet_names()
            assert 'Sheet1' in names, f"XLSX BytesIO: sheets={names}"
            rows = reader.read_all('Sheet1')
            assert rows == [['Name', 'Value'], ['foo', 42]]
    finally:
        os.remove(tmpfile)


def test_xlsb_writer_with_bytesio():
    """Write XLSB to BytesIO, then read back with ExcelReader from file."""
    buf = io.BytesIO()
    data = [['Name', 'Value'], ['foo', 42]]
    with XlsbWriter(buf) as writer:
        writer.add_sheet("Sheet1")
        writer.write_sheet(data)

    tmpfile = tempfile.mktemp(suffix='.xlsb')
    try:
        with open(tmpfile, 'wb') as f:
            f.write(buf.getvalue())
        with ExcelReader(tmpfile) as reader:
            names = reader.get_sheet_names()
            assert 'Sheet1' in names, f"XLSB BytesIO: sheets={names}"
            rows = reader.read_all('Sheet1')
            assert rows == [['Name', 'Value'], ['foo', 42]]
    finally:
        os.remove(tmpfile)


def test_writer_bytesio_valid_zip():
    """Verify BytesIO output is a valid ZIP for both writers."""
    data = [['A', 'B'], [1, 2]]

    for suffix, writer_cls in [('.xlsx', XlsxWriter), ('.xlsb', XlsbWriter)]:
        buf = io.BytesIO()
        with writer_cls(buf) as writer:
            writer.add_sheet("Data")
            writer.write_sheet(data)

        buf.seek(0)
        with zipfile.ZipFile(buf) as zf:
            assert len(zf.namelist()) > 0, f"{suffix}: empty zip"


def test_sheet_names_roundtrip_xlsx_and_xlsb():
    """Create XLSX and XLSB with two sheets, verify get_sheet_names() matches."""
    sheet_names = ["FirstSheet", "SecondSheet"]
    data = [["A", "B"], [1, 2]]

    for suffix, writer_cls in [('.xlsx', XlsxWriter), ('.xlsb', XlsbWriter)]:
        tmpfile = tempfile.mktemp(suffix=suffix)
        try:
            with writer_cls(tmpfile) as writer:
                for name in sheet_names:
                    writer.add_sheet(name)
                    writer.write_sheet(data)

            with ExcelReader(tmpfile) as reader:
                names = reader.get_sheet_names()
                assert names == sheet_names, \
                    f"{suffix}: expected {sheet_names}, got {names}"
        finally:
            os.remove(tmpfile)


def test_reader_with_path_in_constructor():
    """Regression: ``with ExcelReader(path)`` must auto-open the file."""
    tmpfile = tempfile.mktemp(suffix='.xlsb')
    try:
        data = [["Name", "Value"], ["foo", 42]]
        with XlsbWriter(tmpfile) as writer:
            writer.add_sheet("Sheet1")
            writer.write_sheet(data)

        # This pattern is documented on PyPI and must work:
        with ExcelReader(tmpfile) as reader:
            names = reader.get_sheet_names()
            assert 'Sheet1' in names, \
                f"ExcelReader(path) with-statement: sheets={names!r} (empty = bug!)"

            rows = reader.read_all('Sheet1')
            assert len(rows) == 2
            assert rows[0] == ['Name', 'Value']
            assert rows[1] == ['foo', 42]

        # Also test ExcelReader() + .open(path) still works (backward compat)
        with ExcelReader() as reader:
            reader.open(tmpfile)
            assert 'Sheet1' in reader.get_sheet_names()

        print("  Reader with path in constructor: OK")
    finally:
        os.remove(tmpfile)


def _write_workbook(path, data):
    writer_cls = XlsbWriter if str(path).endswith('.xlsb') else XlsxWriter
    with writer_cls(str(path)) as writer:
        writer.add_sheet('Sheet1')
        writer.write_sheet(data)
    return str(path)


def _read_workbook(path):
    with ExcelReader(path) as reader:
        return reader.read_all('Sheet1')


@pytest.mark.parametrize('ext', ['.xlsx', '.xlsb'])
@pytest.mark.parametrize('read_dates_first', [True, False])
def test_custom_number_formats_do_not_leak_between_workbooks(tmp_path, ext, read_dates_first):
    """Regression: custom number formats are scoped to a single workbook.

    Both workbooks below assign the same custom numFmtId to different formats
    (a number format in one, a date format in the other).  The format mapping
    used to be process-global with first-wins semantics, so reading one file
    after the other in the same process misread the second file (dates as raw
    numbers, or numbers as datetimes).
    """
    numbers = _write_workbook(tmp_path / f"numbers{ext}", [['num'], [(100000.0, F.TWO_DECIMALS)]])
    dates = _write_workbook(tmp_path / f"dates{ext}", [['dt'], [(datetime.datetime(2024, 1, 15), F.DATE_ISO)]])

    def read_numbers():
        value = _read_workbook(numbers)[1][0]
        assert not isinstance(value, datetime.datetime), f"number read as date: {value!r}"
        assert value == 100000.0, f"expected 100000.0, got {value!r}"

    def read_dates():
        value = _read_workbook(dates)[1][0]
        assert isinstance(value, datetime.datetime), f"date read as {type(value)}: {value!r}"
        assert value == datetime.datetime(2024, 1, 15), f"expected 2024-01-15, got {value!r}"

    if read_dates_first:
        read_dates()
        read_numbers()
    else:
        read_numbers()
        read_dates()


_EMPTY_ROW_CASES = [
    ('mid_one', [['v', 't'], [1, 'a'], [None, None], [2, 'b']]),
    ('mid_two', [['v', 't'], [1, 'a'], [None, None], [None, None], [2, 'b']]),
    ('mid_three', [['v', 't'], [1, 'a'], [None, None], [None, None], [None, None], [2, 'b']]),
    ('trailing_two', [['v', 't'], [1, 'a'], [None, None], [None, None]]),
    ('leading_two', [[None, None], [None, None], [1, 'a']]),
    ('leading_and_trailing', [[None, None], [1, 'a'], [None, None], [None, None], [2, 'b'], [None, None]]),
    ('single_column_trailing', [['v'], [1.0], [None], [None]]),
]


@pytest.mark.parametrize('use_c_ext', [False, True], ids=['python', 'c_ext'])
@pytest.mark.parametrize('data', [_case[1] for _case in _EMPTY_ROW_CASES],
                         ids=[_case[0] for _case in _EMPTY_ROW_CASES])
def test_xlsb_empty_rows_roundtrip(tmp_path, monkeypatch, use_c_ext, data):
    """Regression: all-empty XLSB rows must round-trip exactly.

    Consecutive all-empty rows used to be duplicated (2 empty rows read back
    as 3, 3 as 6 -- the spurious count grew quadratically), while trailing and
    leading all-empty rows were dropped or mis-aligned.  The python fallback
    reader and the C extension must behave identically.
    """
    if use_c_ext:
        if not xlspy_reader.HAVE_C_EXT:
            pytest.skip('C extension not available')
    else:
        monkeypatch.setattr(xlspy_reader, 'HAVE_C_EXT', False)

    path = _write_workbook(tmp_path / 'empty_rows.xlsb', data)
    assert _read_workbook(path) == data


@pytest.mark.parametrize('use_c_ext', [False, True], ids=['python', 'c_ext'])
def test_xlsb_sheet_with_only_empty_rows_keeps_rows(tmp_path, monkeypatch, use_c_ext):
    """An all-empty sheet keeps one row per written row.

    Cell widths fall back to a single column for the python reader: XLSB row
    records carry no width information and the writer emits no cell records
    for all-empty rows.
    """
    if use_c_ext:
        if not xlspy_reader.HAVE_C_EXT:
            pytest.skip('C extension not available')
    else:
        monkeypatch.setattr(xlspy_reader, 'HAVE_C_EXT', False)

    path = _write_workbook(tmp_path / 'only_empty.xlsb', [[None, None], [None, None]])
    assert _read_workbook(path) == [[None], [None]]


def test_xlsx_empty_rows_roundtrip(tmp_path):
    """XLSX is the reference behaviour for empty rows (python and C readers
    must match it wherever the XLSB metadata allows)."""
    data = [[None, None], [1, 'a'], [None, None], [None, None], [2, 'b'], [None, None]]
    path = _write_workbook(tmp_path / 'empty_rows.xlsx', data)
    assert _read_workbook(path) == data


def test_reader_instance_reuse_does_not_leak_state(tmp_path):
    """One reader instance may open several files in turn.

    Sheets, shared strings and number formats must not leak between files, and
    re-opening an earlier file must give the same result as the first read.
    """
    first = _write_workbook(tmp_path / 'first.xlsx', [['num'], [(100000.0, F.TWO_DECIMALS)]])
    second = _write_workbook(tmp_path / 'second.xlsb', [['dt'], [(datetime.datetime(2024, 1, 15), F.DATE_ISO)]])

    reader = ExcelReader()
    try:
        reader.open(first)
        assert reader.get_sheet_names() == ['Sheet1']
        assert reader.read_all('Sheet1') == [['num'], [100000.0]]

        reader.close()
        reader.open(second)
        assert reader.get_sheet_names() == ['Sheet1']
        assert reader.read_all('Sheet1') == [['dt'], [datetime.datetime(2024, 1, 15)]]

        reader.close()
        reader.open(first)
        assert reader.read_all('Sheet1') == [['num'], [100000.0]]
    finally:
        reader.close()


def test_reader_open_replaces_previous_archive(tmp_path):
    """Calling ``open`` twice without ``close`` must drop the previous archive.

    Leaking it would keep the file handle open until garbage collection and
    mix state from both workbooks.
    """
    first = _write_workbook(tmp_path / 'first.xlsx', [['a'], [1]])
    second = _write_workbook(tmp_path / 'second.xlsx', [['b'], [2]])

    reader = ExcelReader()
    try:
        reader.open(first)
        stale_archive = reader._archive
        reader.open(second)
        assert stale_archive.fp is None, "the previous archive was not closed"
        assert reader.get_sheet_names() == ['Sheet1']
        assert reader.read_all('Sheet1') == [['b'], [2]]
    finally:
        reader.close()


@pytest.mark.filterwarnings('ignore::UserWarning')  # zipfile warns about duplicates
@pytest.mark.parametrize('ext', ['.xlsx', '.xlsb'])
def test_reader_uses_last_duplicate_zip_member(tmp_path, ext):
    """Files with duplicate ZIP members (seen in the wild) read the last copy,
    matching ``zipfile``'s name resolution and preventing entry ambiguity."""
    first = _write_workbook(tmp_path / f'first{ext}', [['value'], [111]])
    second = _write_workbook(tmp_path / f'second{ext}', [['value'], [222]])
    target = tmp_path / f'duplicate{ext}'
    shutil.copyfile(second, target)

    member = 'xl/worksheets/sheet1.bin' if ext == '.xlsb' else 'xl/worksheets/sheet1.xml'
    with zipfile.ZipFile(first) as archive:
        duplicate_payload = archive.read(member)
    with zipfile.ZipFile(target, 'a') as archive:
        archive.writestr(member, duplicate_payload)

    assert _read_workbook(str(target)) == [['value'], [111]]


if __name__ == '__main__':
    print("=== ExcelReader Tests ===")
    print("\n1. XLSB read/write roundtrip...")
    test_read_xlsb()
    print("\n2. XLSX read/write roundtrip...")
    test_read_xlsx()
    print("\n3. Multi-sheet XLSB...")
    test_multi_sheet_xlsb()
    print("\n4. Multi-sheet XLSX...")
    test_multi_sheet_xlsx()
    print("\n5. Row generator XLSB...")
    test_row_generator_xlsb()
    print("\n6. Sheet not found error handling...")
    test_sheet_not_found()
    print("\n7. Complex XLSX values...")
    test_read_complex_xlsx_values()
    print("\n8. Reader with path in constructor...")
    test_reader_with_path_in_constructor()
    print("\n=== All tests passed! ===")
