"""Shared data for the golden workbooks and the cross-validation tests.

Both ``tools/make_goldens.py`` and ``tests/test_cross_validation.py`` import
this module so the golden files, their JSON sidecars and the oracle tests all
describe the same workbook content.

The dates deliberately cover the format-classification cases that used to
cross-contaminate files (built-in short date, custom ISO date and datetime
formats, plus a serial number that is *not* a date).
"""

import datetime as dt

from xlspy import F

HEADER = ["text", "int", "float", "bool", "date_iso", "date_short", "datetime", "hole", "tail"]

ROWS = [
    ["plain", 0, 0.0, True, dt.date(2024, 1, 15), dt.date(2023, 6, 1), dt.datetime(2024, 1, 15, 10, 30, 0), None, "end"],
    ["  padded  ", -1, -0.001, False, dt.date(1900, 1, 1), dt.date(1999, 9, 9), dt.datetime(1999, 9, 9, 9, 9, 9), None, "end"],
    ["<tag> & \"quotes\" 'x'", 2147483647, 1e10, True, dt.date(2173, 10, 14), dt.date(2100, 12, 31), dt.datetime(2024, 2, 29, 23, 59, 59), "mid", "end"],
    ["zażółć gęślą jaźń ⚡", -2147483648, 1.23456789e-8, False, dt.date(2000, 2, 29), dt.date(2020, 12, 31), dt.datetime(2020, 12, 31, 0, 0, 0), None, "end"],
    ["line\nbreak\ttab", 123456789, 0.30000000000000004, False, dt.date(1970, 1, 1), dt.date(2024, 2, 29), dt.datetime(1970, 1, 1, 0, 0, 0), None, "end"],
]

SHEETS = {
    "Main": [HEADER, *ROWS],
    "Numbers": [["n"], [1.5], [2], [3]],
}

# Number formats used when the goldens are written (openpyxl / xlsxwriter) and
# by the xlspy writer in the oracle tests.
DATE_FORMAT = "yyyy-mm-dd"
SHORT_DATE_FORMAT = "mm-dd-yy"  # built-in numFmt 14
DATETIME_FORMAT = "yyyy-mm-dd hh:mm:ss"

# The 1904 date system cannot represent dates before 1904-01-01, so the 1904
# golden substitutes the single pre-1904 value from ROWS.
_MAC_SUBSTITUTIONS = {dt.date(1900, 1, 1): dt.date(1904, 1, 2)}


def sheets_1904():
    """``SHEETS`` adjusted for a workbook using the 1904 date system."""

    return {
        name: [[_MAC_SUBSTITUTIONS.get(cell, cell) for cell in row] for row in rows]
        for name, rows in SHEETS.items()
    }

COLUMN_FORMATS = {
    "date_iso": F.DATE_ISO,
    "date_short": F.DATE_SHORT,
    "datetime": F.DATETIME_ISO,
}

# openpyxl and xlsxwriter round floats to 15 significant digits (Excel's
# precision), so this 17-digit value is stored - and read back - as 0.3.
EXPECTED_OVERRIDES = {
    ("Main", 5, 2): 0.3,
}


def encode_cell(value):
    """Encode one expected cell value for a JSON sidecar.

    Dates and datetimes become ``{"datetime": <ISO>}``; xlspy normalises both
    to ``datetime`` when the cell carries a date number format.
    """

    if isinstance(value, dt.datetime):
        return {"datetime": value.isoformat()}
    if isinstance(value, dt.date):
        return {"datetime": dt.datetime(value.year, value.month, value.day).isoformat()}
    return value


def encoded_sheets(sheets=None, *, overrides=True):
    """The expected xlspy reading of ``SHEETS`` (or ``sheets``) as JSON-ready data.

    ``overrides`` applies ``EXPECTED_OVERRIDES`` for writers that round floats
    to Excel's 15 significant digits; pass ``overrides=False`` for the xlspy
    writer, which stores the exact IEEE double.
    """

    encoded = {}
    for name, rows in (SHEETS if sheets is None else sheets).items():
        encoded[name] = []
        for row_index, row in enumerate(rows):
            encoded[name].append(
                [
                    encode_cell(
                        EXPECTED_OVERRIDES.get((name, row_index, column_index), cell)
                        if overrides
                        else cell
                    )
                    for column_index, cell in enumerate(row)
                ]
            )
    return encoded


def writer_rows():
    """``SHEETS["Main"]`` prepared for the xlspy writers.

    Date columns carry an explicit ``(value, format)`` tuple so the writer
    emits the same number formats as the golden writers.
    """

    header, *rows = SHEETS["Main"]
    converted = [list(header)]
    for row in rows:
        converted.append(
            [
                (value, COLUMN_FORMATS[column])
                if column in COLUMN_FORMATS and value is not None
                else value
                for column, value in zip(header, row)
            ]
        )
    return converted
