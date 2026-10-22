"""Print settings in XLSB output, checked against the records Excel writes.

The expected payloads come from an XLSX with the same settings that Excel 16
saved as XLSB.
"""

import struct
import zipfile

import pytest

from xlspy import PageMargins, PrintSettings, XlsbWriter
from xlspy.biff_utils import iter_records
from xlspy._xlsb_metadata import read_wide

SETTINGS = PrintSettings(
    orientation="landscape", paper_size="A4", margins=PageMargins(0.5, 0.6, 0.7, 0.8, 0.25, 0.35),
    print_area="A1:D20", repeat_rows="1:2", repeat_columns="A:A", fit_to_width=1, fit_to_height=0,
    center_horizontally=True, center_vertically=True, header="&LReport&RPage &P of &N",
    footer="&CFooter", row_breaks=[10], col_breaks=[2],
)

# Payloads Excel wrote for SETTINGS.
EXCEL = {
    0x1DD: "13 00",
    0x1DC: ("00 00 00 00 00 00 e0 3f 33 33 33 33 33 33 e3 3f 66 66 66 66 66 66 e6 3f "
            "9a 99 99 99 99 99 e9 3f 00 00 00 00 00 00 d0 3f 66 66 66 66 66 66 d6 3f"),
    0x1DE: ("09 00 00 00 64 00 00 00 58 02 00 00 58 02 00 00 01 00 00 00 01 00 00 00 "
            "01 00 00 00 00 00 00 00 02 00 ff ff ff ff"),
    0x188: "01 00 00 00 01 00 00 00",
    0x18A: "01 00 00 00 01 00 00 00",
}
EXCEL_BREAKS = [
    "0a 00 00 00 00 00 00 00 ff 3f 00 00 01 00 00 00 00 00 00 00",
    "02 00 00 00 00 00 00 00 ff ff 0f 00 01 00 00 00 00 00 00 00",
]


def _write(tmp_path, settings, name="print.xlsb"):
    path = tmp_path / name
    with XlsbWriter(path) as writer:
        writer.add_sheet("Print")
        writer.write_sheet([["h1", "h2"]] + [[i, i] for i in range(30)], print_settings=settings)
    return path


def _records(path, part="xl/worksheets/sheet1.bin"):
    data = zipfile.ZipFile(path).read(part)
    return [(record.record_id, data[record.data_start:record.data_end]) for record in iter_records(data)]


def test_records_match_excel(tmp_path):
    records = _records(_write(tmp_path, SETTINGS))
    payloads = {kind: payload for kind, payload in records}
    for kind, expected in EXCEL.items():
        assert payloads[kind] == bytes.fromhex(expected), hex(kind)
    assert [payload for kind, payload in records if kind == 0x18C] == [bytes.fromhex(h) for h in EXCEL_BREAKS]

    flags, = struct.unpack_from("<H", payloads[0x1DF])
    assert flags == 0x0C
    header, offset = read_wide(payloads[0x1DF], 2)
    footer, offset = read_wide(payloads[0x1DF], offset)
    assert (header, footer) == ("&LReport&RPage &P of &N", "&CFooter")
    assert payloads[0x1DF][offset:] == b"\xff" * 16
    assert 0x1E0 in payloads

    kinds = [kind for kind, _ in records]
    order = [0x1DD, 0x1DC, 0x1DE, 0x1DF, 0x1E0, 0x188, 0x189, 0x18A, 0x18B]
    assert [kind for kind in kinds if kind in order] == order
    # fFitToPage in BrtWsProp.
    assert payloads[0x93][1] & 0x01


def test_print_area_and_titles_are_builtin_names(tmp_path):
    names = [payload for kind, payload in _records(_write(tmp_path, SETTINGS), "xl/workbook.bin") if kind == 0x27]
    by_name = {read_wide(payload, 9)[0]: payload for payload in names}
    area, titles = by_name["Print_Area"], by_name["Print_Titles"]
    for payload in (area, titles):
        flags, _key, sheet = struct.unpack_from("<IBI", payload)
        assert (flags, sheet) == (0x20, 0)  # fBuiltin, scoped to the first sheet
    # PtgArea3d over A1:D20, exactly as Excel stores it.
    assert bytes.fromhex("0f 00 00 00 3b 00 00 00 00 00 00 13 00 00 00 00 00 03 00") in area


def test_page_setup_defaults_follow_excel(tmp_path):
    def page_setup(settings):
        return dict(_records(_write(tmp_path, settings))).get(0x1DE)

    # Excel's defaults: Letter, 100%, no orientation (fNoOrient).
    assert page_setup(PrintSettings(scale=80)) == bytes.fromhex(
        "01 00 00 00 50 00 00 00 58 02 00 00 58 02 00 00 01 00 00 00 01 00 00 00 "
        "01 00 00 00 01 00 00 00 40 00 ff ff ff ff")
    assert page_setup(PrintSettings(orientation="portrait"))[32:34] == b"\x00\x00"
    # A header alone, like XLSX, has no page setup record.
    assert page_setup(PrintSettings(header="&COnly")) is None


def test_no_print_settings_keep_the_template(tmp_path):
    plain = tmp_path / "plain.xlsb"
    with XlsbWriter(plain) as writer:
        writer.add_sheet("Print")
        writer.write_sheet([["a"]])
    default = tmp_path / "default.xlsb"
    with XlsbWriter(default) as writer:
        writer.add_sheet("Print")
        writer.write_sheet([["a"]], print_settings=PrintSettings())
    assert _records(plain) == _records(default)


def test_print_settings_must_be_print_settings(tmp_path):
    with XlsbWriter(tmp_path / "bad.xlsb") as writer:
        writer.add_sheet("Print")
        with pytest.raises(TypeError, match="PrintSettings"):
            writer.write_sheet([["a"]], print_settings={"orientation": "landscape"})
        writer.write_sheet([["a"]])
