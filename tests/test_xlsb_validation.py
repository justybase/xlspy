"""Data validation in XLSB output, checked against the records Excel writes.

The expected payloads come from an XLSX with the same rules that Excel 16
saved as XLSB.
"""

import datetime
import zipfile

import pytest

from xlspy import DataValidation, XlsbWriter
from xlspy.biff_utils import iter_records

BEGIN, DVAL, END = 0x23D, 0x40, 0x23E

# Rule -> the BrtDVal payload Excel wrote for it.
EXCEL = [
    (DataValidation("list", "A2:A10", '"New,Active,Closed"', prompt_title="Status", prompt="Choose",
                    show_input_message=True, error_title="Error", error="Only from the list"),
     "83 01 0c 00 01 00 00 00 01 00 00 00 09 00 00 00 00 00 00 00 00 00 00 00 05 00 00 00 45 00 72 00 72 00 "
     "6f 00 72 00 12 00 00 00 4f 00 6e 00 6c 00 79 00 20 00 66 00 72 00 6f 00 6d 00 20 00 74 00 68 00 65 00 "
     "20 00 6c 00 69 00 73 00 74 00 06 00 00 00 53 00 74 00 61 00 74 00 75 00 73 00 06 00 00 00 43 00 68 00 "
     "6f 00 6f 00 73 00 65 00 25 00 00 00 17 11 00 4e 00 65 00 77 00 2c 00 41 00 63 00 74 00 69 00 76 00 65 "
     "00 2c 00 43 00 6c 00 6f 00 73 00 65 00 64 00 00 00 00 00 00 00 00 00 00 00 00 00"),
    (DataValidation("list", "B2:B10", "$H$2:$H$4", show_dropdown=False, allow_blank=False),
     "03 02 08 00 01 00 00 00 01 00 00 00 09 00 00 00 01 00 00 00 01 00 00 00 ff ff ff ff ff ff ff ff ff ff "
     "ff ff ff ff ff ff 0d 00 00 00 25 01 00 00 00 03 00 00 00 07 00 07 00 00 00 00 00 00 00 00 00 00 00 00 00"),
    (DataValidation("whole", "D2:D10", 1, 100, error_style="warning"),
     "11 01 08 00 01 00 00 00 01 00 00 00 09 00 00 00 03 00 00 00 03 00 00 00 ff ff ff ff ff ff ff ff ff ff "
     "ff ff ff ff ff ff 03 00 00 00 1e 01 00 00 00 00 00 03 00 00 00 1e 64 00 00 00 00 00"),
    (DataValidation("decimal", "E2:E10", 0.5, operator="greater_than", error_style="information"),
     "22 01 48 00 01 00 00 00 01 00 00 00 09 00 00 00 04 00 00 00 04 00 00 00 ff ff ff ff ff ff ff ff ff ff "
     "ff ff ff ff ff ff 09 00 00 00 1f 00 00 00 00 00 00 e0 3f 00 00 00 00 00 00 00 00 00 00 00 00"),
    (DataValidation("date", "F2:F10", datetime.date(2024, 1, 1), datetime.date(2024, 12, 31),
                    operator="not_between"),
     "04 01 18 00 01 00 00 00 01 00 00 00 09 00 00 00 05 00 00 00 05 00 00 00 ff ff ff ff ff ff ff ff ff ff "
     "ff ff ff ff ff ff 03 00 00 00 1e ec b0 00 00 00 00 03 00 00 00 1e 59 b2 00 00 00 00"),
    (DataValidation("text_length", "I2:I10 K2:K3", 10, operator="less_than", show_error_message=False),
     "06 01 50 00 02 00 00 00 01 00 00 00 09 00 00 00 08 00 00 00 08 00 00 00 01 00 00 00 02 00 00 00 0a 00 "
     "00 00 0a 00 00 00 ff ff ff ff ff ff ff ff ff ff ff ff ff ff ff ff 03 00 00 00 1e 0a 00 00 00 00 00 00 "
     "00 00 00 00 00 00 00"),
    # The relative reference is stored as PtgRefN, an offset from J2.
    (DataValidation("custom", "J2:J10", "MOD(J2,2)=0"),
     "07 01 08 00 01 00 00 00 01 00 00 00 09 00 00 00 09 00 00 00 09 00 00 00 ff ff ff ff ff ff ff ff ff ff "
     "ff ff ff ff ff ff 11 00 00 00 4c 00 00 00 00 00 c0 1e 02 00 41 27 00 1e 00 00 0b 00 00 00 00 00 00 00 "
     "00 00 00 00 00"),
]


def _records(path):
    data = zipfile.ZipFile(path).read("xl/worksheets/sheet1.bin")
    return [(record.record_id, data[record.data_start:record.data_end]) for record in iter_records(data)]


def _write(tmp_path, rules):
    path = tmp_path / "validation.xlsb"
    with XlsbWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([["a"] * 11], data_validations=rules)
        writer.add_sheet("Lists")
        writer.write_sheet([["x"], ["y"], ["z"]])
    return path


def test_records_match_excel(tmp_path):
    records = _records(_write(tmp_path, [rule for rule, _ in EXCEL]))
    kinds = [kind for kind, _ in records]
    assert kinds.count(BEGIN) == kinds.count(END) == 1
    begin = dict(records)[BEGIN]
    assert begin == bytes(14) + len(EXCEL).to_bytes(4, "little")
    payloads = [payload for kind, payload in records if kind == DVAL]
    for (rule, expected), payload in zip(EXCEL, payloads, strict=True):
        assert payload == bytes.fromhex(expected), rule
    # DVALS sit after the cell table and before the print options.
    assert kinds.index(0x92) < kinds.index(BEGIN) < kinds.index(0x1DD)


def test_cross_sheet_list_uses_a_3d_reference(tmp_path):
    records = _records(_write(tmp_path, [DataValidation("list", "C2", "Lists!$A$1:$A$3")]))
    (payload,) = [payload for kind, payload in records if kind == DVAL]
    # PtgArea3d (reference class) with the sheet's XTI index, rows 0-2 of column A.
    assert bytes.fromhex("3b 00 00 00 00 00 00 02 00 00 00 00 00 00 00") in payload


def test_no_validations_write_no_records(tmp_path):
    kinds = [kind for kind, _ in _records(_write(tmp_path, None))]
    assert BEGIN not in kinds and DVAL not in kinds


def test_data_validations_must_be_data_validations(tmp_path):
    with XlsbWriter(tmp_path / "bad.xlsb") as writer:
        writer.add_sheet("Data")
        with pytest.raises(TypeError, match="DataValidation instances"):
            writer.write_sheet([["a"]], data_validations=[{"type": "list"}])
        writer.write_sheet([["a"]])
