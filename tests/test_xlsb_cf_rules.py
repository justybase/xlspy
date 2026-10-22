"""Conditional format rules with a differential style in XLSB.

The expected field values are the ones Excel 16 writes when it saves the
same rules from an XLSX file as XLSB.
"""

import struct
import zipfile

import pytest

from xlspy import (
    AboveAverageRule, Alignment, BlanksRule, Border, CellIsRule, DifferentialStyle, DuplicateRule,
    Font, FormulaRule, GradientFill, PatternFill, Side, TextRule, TimePeriodRule, TopBottomRule,
    XlsbWriter, XlsxWriter,
)
from xlspy._xlsb_cf import dxf_payload
from xlspy._xlsb_formulas import FormulaContext, compile_formula
from xlspy.biff_utils import iter_records


def _payloads(data, record_id):
    return [data[r.data_start:r.data_end] for r in iter_records(data) if r.record_id == record_id]

RED = DifferentialStyle(font=Font(bold=True, color="9C0006"), fill=PatternFill("solid", start_color="FFC7CE"))


def _write(tmp_path, rules, rows=None):
    path = tmp_path / "cf.xlsb"
    with XlsbWriter(str(path)) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(rows or [["h", "t"]] + [[i, "abc"] for i in range(10)], conditional_formats=rules)
        writer.save()
    with zipfile.ZipFile(path) as archive:
        return archive.read("xl/worksheets/sheet1.bin"), archive.read("xl/styles.bin")


def _rules(sheet):
    """``(iType, iTemplate, dxfId, priority, iParam, flags, cbFmla1..3, text)`` per rule."""
    result = []
    for data in _payloads(sheet, 0x1CF):
        kind, template, dxf_id, priority, param = struct.unpack_from("<IIiII", data)
        flags, = struct.unpack_from("<H", data, 28)
        sizes = struct.unpack_from("<III", data, 30)
        count, = struct.unpack_from("<I", data, 42)
        text = None if count == 0xFFFFFFFF else data[46:46 + 2 * count].decode("utf-16-le")
        result.append((kind, template, dxf_id, priority, param, flags, *sizes, text))
    return result


def test_rule_fields_match_excel(tmp_path):
    rules = [
        CellIsRule("A2:A11", "between", ["5", "10"], dxf=RED),
        CellIsRule("A2:A11", "greaterThan", ["$K$1"], dxf=RED, stop_if_true=True),
        FormulaRule("A2:A11", "MOD(A2,2)=0", dxf=RED),
        TopBottomRule("A2:A11", rank=3, dxf=RED),
        TopBottomRule("A2:A11", rank=10, percent=True, bottom=True, dxf=RED),
        AboveAverageRule("A2:A11", dxf=RED),
        AboveAverageRule("A2:A11", below=True, equal_average=True, dxf=RED),
        AboveAverageRule("A2:A11", std_dev=1, dxf=RED),
        DuplicateRule("A2:A11", dxf=RED),
        DuplicateRule("A2:A11", unique=True, dxf=RED),
        TimePeriodRule("A2:A11", "lastWeek", dxf=RED),
        BlanksRule("A2:A11", "no_errors", dxf=RED),
        # Rules are written per range, so another range goes last.
        TextRule("B2:B11", "contains", "ab", dxf=RED),
    ]
    sheet, styles = _write(tmp_path, rules)
    fields = _rules(sheet)
    # One DXF shared by all rules; priorities follow the rule order.
    assert [field[2] for field in fields] == [0] * len(rules)
    assert [field[3] for field in fields] == list(range(1, len(rules) + 1))
    assert [record.record_id for record in iter_records(styles)].count(0x1FB) == 1
    summary = [(kind, template, param, flags, sizes, text)
               for kind, template, _, _, param, flags, *sizes, text in fields]
    assert summary == [
        # The cell-is operands are numbers and need no reference bytes.
        (1, 0, 1, 0x00, [3, 3, 0], None),
        # An absolute reference: 7 token bytes plus 2.
        (1, 0, 5, 0x02, [9, 0, 0], None),
        # MOD(A2,2)=0: a relative reference (PtgRefN) adds 2 as well.
        (2, 1, 0, 0x00, [19, 0, 0], None),
        (5, 5, 3, 0x00, [0, 0, 0], None),
        (5, 5, 10, 0x18, [0, 0, 0], None),
        (2, 0x19, 0, 0x04, [0, 0, 0], None),
        (2, 0x1E, 0, 0x00, [0, 0, 0], None),
        (2, 0x19, 1, 0x04, [0, 0, 0], None),
        (2, 0x1B, 0, 0x00, [0, 0, 0], None),
        (2, 0x07, 0, 0x00, [0, 0, 0], None),
        (2, 0x17, 4, 0x00, [68, 0, 0], None),
        (2, 0x0C, 0, 0x00, [15, 0, 0], None),
        (2, 0x08, 0, 0x00, [26, 0, 0], "ab"),
    ]


def test_cb_fmla_counts_two_bytes_per_reference(tmp_path):
    sheet, _ = _write(tmp_path, [FormulaRule("A2:A11", "AND(A2>B2,A2<$C$1)", dxf=RED)])
    data = _payloads(sheet, 0x1CF)[0]
    tokens, = struct.unpack_from("<I", data, 46)
    assert struct.unpack_from("<I", data, 30)[0] == tokens + 2 * 4


def test_time_period_formula_is_volatile(tmp_path):
    sheet, _ = _write(tmp_path, [TimePeriodRule("A2:A11", "today", dxf=RED)])
    data = _payloads(sheet, 0x1CF)[0]
    # The formula starts with PtgAttrSemi, as TODAY() is volatile.
    assert data[50:52] == b"\x19\x01"


@pytest.mark.parametrize("formula", ["TODAY()", "NOW()+1", "RAND()", "OFFSET(A1,1,1)"])
def test_volatile_functions_get_the_attr_semi_marker(formula):
    compiled = compile_formula(formula, FormulaContext(['Data']))
    assert compiled[4:6] == b"\x19\x01"


def test_edate_is_known():
    compiled = compile_formula("EDATE(A1,1)", FormulaContext(['Data']))
    # A fixed argument count: PtgFunc with function number 449.
    assert struct.pack("<H", 449) in compiled


def test_text_operators_and_text_parameter(tmp_path):
    rules = [TextRule("B2:B11", operator, 'x"y', dxf=RED)
             for operator in ("contains", "notContains", "beginsWith", "endsWith")]
    fields = _rules(_write(tmp_path, rules)[0])
    assert [(f[1], f[4], f[-1]) for f in fields] == [(8, op, 'x"y') for op in range(4)]


def test_text_cell_is_operators_are_refused(tmp_path):
    with pytest.raises(TypeError, match="XLSX-only"):
        _write(tmp_path, [CellIsRule("A2:A11", "containsText", ["a"], dxf=RED)])


def test_dxf_payload_matches_excel():
    payload = dxf_payload(RED, lambda _fmt: 0)
    assert payload.hex() == (
        "008000000500"
        # fill pattern solid, foreground and background FFC7CE
        "000005000101000c0005ff0000ffc7ceff" "02000c0005ff0000ffc7ceff"
        # font color 9C0006, bold 700
        "05000c0005ff00009c0006ff" "19000600bc02"
    )


def test_dxf_number_format_comes_before_its_string():
    payload = dxf_payload(DifferentialStyle(number_format="0.0%"), lambda _fmt: 167)
    assert payload[4:6] == b"\x02\x00"
    assert payload[6:12] == struct.pack("<HHH", 0x29, 6, 167)
    assert payload[12:14] == struct.pack("<H", 0x26)
    assert payload[16:].decode("utf-16-le").endswith("0.0%")


def test_dxf_border_writes_all_four_sides():
    payload = dxf_payload(DifferentialStyle(border=Border(left=Side("thin", "FF0000"))), lambda _fmt: 0)
    count, = struct.unpack_from("<H", payload, 4)
    kinds = [struct.unpack_from("<H", payload, 6 + 14 * i)[0] for i in range(count)]
    assert kinds == [6, 7, 8, 9]


@pytest.mark.parametrize("dxf", [
    DifferentialStyle(fill=GradientFill(stop=("FF0000", "0000FF"))),
    DifferentialStyle(font=Font(vertAlign="superscript")),
    DifferentialStyle(border=Border(diagonal=Side("thin"))),
    DifferentialStyle(alignment=Alignment(vertical="center")),
])
def test_unsupported_dxf_options_raise(tmp_path, dxf):
    with pytest.raises(TypeError, match="not supported in XLSB conditional formats"):
        _write(tmp_path, [FormulaRule("A2:A11", "A2>1", dxf=dxf)])


def test_xlsx_solid_dxf_fill_with_start_color_sets_background(tmp_path):
    path = tmp_path / "cf.xlsx"
    with XlsxWriter(str(path)) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([["h"], [1]], conditional_formats=[CellIsRule("A2", "equal", ["1"], dxf=RED)])
        writer.save()
    styles = zipfile.ZipFile(path).read("xl/styles.xml").decode()
    dxfs = styles[styles.index("<dxfs"):styles.index("</dxfs>")]
    assert 'bgColor rgb="FFFFC7CE"' in dxfs
