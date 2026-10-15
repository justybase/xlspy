"""Date handling: Excel serial numbers, the fictitious 1900-02-29 and 1904.

openpyxl is the reference implementation here: its ``to_excel``/``from_excel``
mirror Excel's behaviour, so both the raw conversions and the writer/reader
round trips are compared against it.
"""

import datetime as dt
import struct
import zipfile

import pytest

from xlspy import ExcelReader, F, XlsbWriter, XlsxWriter
from xlspy import reader as xlspy_reader
from xlspy.biff_utils import iter_records
from xlspy.excel_dates import datetime_to_serial, serial_to_datetime

from zip_utils import rewrite_workbook_members


def _write_workbook(path, data):
    writer_cls = XlsbWriter if str(path).endswith(".xlsb") else XlsxWriter
    with writer_cls(str(path)) as writer:
        writer.add_sheet("Sheet1")
        writer.write_sheet(data)
    return path


def _read_workbook(path):
    with ExcelReader(str(path)) as reader:
        return reader.read_all("Sheet1")


def _xlsb_sheet_doubles(path):
    with zipfile.ZipFile(path) as archive:
        sheet = archive.read("xl/worksheets/sheet1.bin")
    return [
        struct.unpack_from("<d", sheet, record.data_start + 8)[0]
        for record in iter_records(sheet)
        if record.record_id == 0x05
    ]


def _set_xlsb_date_system(path):
    """Flip the ``f1904`` bit of the BrtWbProp record in ``workbook.bin``."""

    with zipfile.ZipFile(path) as archive:
        workbook = archive.read("xl/workbook.bin")
    for record in iter_records(workbook):
        if record.record_id == 0x99 and record.length >= 4:
            payload = bytearray(workbook[record.data_start : record.data_end])
            payload[0] |= 0x01
            updated = workbook[: record.data_start] + bytes(payload) + workbook[record.data_end :]
            rewrite_workbook_members(path, {"xl/workbook.bin": updated})
            return
    raise AssertionError("BrtWbProp record not found in workbook.bin")


EARLY_DATES = [dt.date(1900, 1, 1), dt.date(1900, 2, 28), dt.date(1900, 3, 1), dt.date(2024, 1, 15)]
REPRESENTATIVE_MAC_DATES = [
    dt.date(1904, 1, 2),
    dt.date(1969, 12, 31),
    dt.date(2024, 1, 15),
    dt.datetime(2024, 1, 15, 10, 30),
]


def _importorskip_openpyxl_datetime():
    pytest.importorskip("openpyxl")
    from openpyxl.utils.datetime import CALENDAR_MAC_1904, from_excel, to_excel

    return CALENDAR_MAC_1904, from_excel, to_excel


@pytest.mark.parametrize(
    "serial", [1, 1.5, 2, 59, 59.5, 60, 60.5, 61, 61.5, 45306, 45306.4375]
)
def test_serial_to_datetime_matches_openpyxl(serial):
    _, from_excel, _ = _importorskip_openpyxl_datetime()
    assert serial_to_datetime(serial) == from_excel(serial)


def test_serial_below_one_keeps_datetime_convention():
    """openpyxl returns ``datetime.time`` for serials below 1; xlspy keeps its
    documented behaviour of returning datetimes for date-formatted cells."""
    assert serial_to_datetime(0.5) == dt.datetime(1899, 12, 30, 12, 0)
    assert serial_to_datetime(0.5, date_1904=True) == dt.datetime(1904, 1, 1, 12, 0)


@pytest.mark.parametrize("serial", [1, 1.5, 59.5, 60, 61, 43843, 43844.4375])
def test_serial_to_datetime_matches_openpyxl_1904(serial):
    mac_epoch, from_excel, _ = _importorskip_openpyxl_datetime()
    assert serial_to_datetime(serial, date_1904=True) == from_excel(serial, epoch=mac_epoch)


@pytest.mark.parametrize(
    "value",
    [
        dt.datetime(1900, 1, 1),
        dt.datetime(1900, 2, 28),
        dt.datetime(1900, 2, 28, 23, 59),
        dt.datetime(1900, 3, 1),
        dt.datetime(2024, 1, 15, 10, 30),
    ],
)
def test_datetime_to_serial_matches_openpyxl(value):
    _, _, to_excel = _importorskip_openpyxl_datetime()
    assert datetime_to_serial(value) == pytest.approx(to_excel(value), rel=1e-12)


@pytest.mark.parametrize("value", [dt.datetime(1904, 1, 1), dt.date(2024, 1, 15)])
def test_datetime_to_serial_matches_openpyxl_1904(value):
    mac_epoch, _, to_excel = _importorskip_openpyxl_datetime()
    assert datetime_to_serial(value, date_1904=True) == pytest.approx(
        to_excel(value, epoch=mac_epoch), rel=1e-12
    )


@pytest.mark.parametrize("ext", [".xlsx", ".xlsb"])
def test_writers_store_excel_compatible_serials(tmp_path, ext):
    """1900-01-01 is serial 1: Excel counts the fictitious 1900-02-29."""
    path = _write_workbook(
        tmp_path / f"early{ext}", [[(value, F.DATE_SHORT)] for value in EARLY_DATES]
    )
    expected_values = [dt.datetime(v.year, v.month, v.day) for v in EARLY_DATES]

    if ext == ".xlsx":
        openpyxl = pytest.importorskip("openpyxl")
        workbook = openpyxl.load_workbook(path, data_only=True)
        read_by_openpyxl = [row[0] for row in workbook.active.values]
        workbook.close()
        assert read_by_openpyxl == expected_values
    else:
        assert _xlsb_sheet_doubles(path) == [datetime_to_serial(v) for v in EARLY_DATES]

    assert [row[0] for row in _read_workbook(path)] == expected_values


def test_reader_matches_openpyxl_for_early_1900_dates(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Sheet1"
    for index, value in enumerate(EARLY_DATES, 1):
        cell = sheet.cell(row=index, column=1, value=value)
        cell.number_format = "yyyy-mm-dd"
    path = tmp_path / "early_openpyxl.xlsx"
    workbook.save(path)
    workbook.close()

    expected = [dt.datetime(v.year, v.month, v.day) for v in EARLY_DATES]
    reference = openpyxl.load_workbook(path, data_only=True)
    assert [row[0] for row in reference.active.values] == expected  # sanity
    reference.close()
    assert [row[0] for row in _read_workbook(path)] == expected


def test_reader_maps_fictitious_1900_02_29_like_openpyxl(tmp_path):
    _, from_excel, _ = _importorskip_openpyxl_datetime()
    xlsxwriter = pytest.importorskip("xlsxwriter")

    path = tmp_path / "serial60.xlsx"
    workbook = xlsxwriter.Workbook(str(path))
    sheet = workbook.add_worksheet("Sheet1")
    sheet.write_number(0, 0, 60, workbook.add_format({"num_format": "yyyy-mm-dd"}))
    workbook.close()

    assert from_excel(60) == dt.datetime(1900, 2, 28)
    assert _read_workbook(path) == [[from_excel(60)]]


def test_reader_honors_1904_xlsx_like_openpyxl(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")
    mac_epoch, _, _ = _importorskip_openpyxl_datetime()

    workbook = openpyxl.Workbook()
    workbook.epoch = mac_epoch
    sheet = workbook.active
    sheet.title = "Sheet1"
    for index, value in enumerate(REPRESENTATIVE_MAC_DATES, 1):
        cell = sheet.cell(row=index, column=1, value=value)
        cell.number_format = "yyyy-mm-dd"
    path = tmp_path / "mac.xlsx"
    workbook.save(path)
    workbook.close()

    reference = openpyxl.load_workbook(path, data_only=True)
    expected = [row[0] for row in reference.active.values]
    reference.close()
    assert expected == [
        value if isinstance(value, dt.datetime) else dt.datetime(value.year, value.month, value.day)
        for value in REPRESENTATIVE_MAC_DATES
    ]
    assert [row[0] for row in _read_workbook(path)] == expected


@pytest.mark.parametrize("use_c_ext", [False, True], ids=["python", "c_ext"])
def test_reader_honors_1904_xlsb_flag(tmp_path, monkeypatch, use_c_ext):
    """The 1904 flag changes what every stored serial means.

    A workbook written under the 1900 system stores 2020-01-14 as serial 43843;
    after flipping the workbook flag, the same serial means 2024-01-15.
    """
    if use_c_ext:
        if not xlspy_reader.HAVE_C_EXT:
            pytest.skip("C extension not available")
    else:
        monkeypatch.setattr(xlspy_reader, "HAVE_C_EXT", False)

    path = _write_workbook(tmp_path / "mac.xlsb", [[(dt.date(2020, 1, 14), F.DATE_SHORT)]])
    _set_xlsb_date_system(path)

    assert _read_workbook(path) == [[dt.datetime(2024, 1, 15)]]
