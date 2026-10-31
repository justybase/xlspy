"""Number-format constants and helpers compatible with openpyxl.

The built-in format table and the date/time classification mirror
``openpyxl.styles.numbers``.  xlspy's writers only need the codes, but the
public helpers are useful to callers that inspect or normalise formats.
"""

from __future__ import annotations

import re

BUILTIN_FORMATS_MAX_SIZE = 164

BUILTIN_FORMATS = {
    0: "General",
    1: "0",
    2: "0.00",
    3: "#,##0",
    4: "#,##0.00",
    5: '"$"#,##0_);("$"#,##0)',
    6: '"$"#,##0_);[Red]("$"#,##0)',
    7: '"$"#,##0.00_);("$"#,##0.00)',
    8: '"$"#,##0.00_);[Red]("$"#,##0.00)',
    9: "0%",
    10: "0.00%",
    11: "0.00E+00",
    12: "# ?/?",
    13: "# ??/??",
    14: "mm-dd-yy",
    15: "d-mmm-yy",
    16: "d-mmm",
    17: "mmm-yy",
    18: "h:mm AM/PM",
    19: "h:mm:ss AM/PM",
    20: "h:mm",
    21: "h:mm:ss",
    22: "m/d/yy h:mm",
    37: "#,##0_);(#,##0)",
    38: "#,##0_);[Red](#,##0)",
    39: "#,##0.00_);(#,##0.00)",
    40: "#,##0.00_);[Red](#,##0.00)",
    41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)',
    42: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)',
    43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
    44: '_("$"* #,##0.00_)_("$"* \\(#,##0.00\\)_("$"* "-"??_)_(@_)',
    45: "mm:ss",
    46: "[h]:mm:ss",
    47: "mmss.0",
    48: "##0.0E+0",
    49: "@",
}

BUILTIN_FORMATS_REVERSE = {value: key for key, value in BUILTIN_FORMATS.items()}

FORMAT_GENERAL = "General"
FORMAT_TEXT = "@"
FORMAT_NUMBER = "0"
FORMAT_NUMBER_00 = "0.00"
FORMAT_NUMBER_COMMA_SEPARATED1 = "#,##0.00"
FORMAT_NUMBER_COMMA_SEPARATED2 = "#,##0.00_-"
FORMAT_PERCENTAGE = "0%"
FORMAT_PERCENTAGE_00 = "0.00%"
FORMAT_CURRENCY_USD = "$#,##0_-"
FORMAT_CURRENCY_USD_SIMPLE = '"$"#,##0.00_-'
FORMAT_CURRENCY_EUR_SIMPLE = "[$EUR ]#,##0.00_-"
FORMAT_DATE_YYYYMMDD2 = "yyyy-mm-dd"
FORMAT_DATE_YYMMDD = "yy-mm-dd"
FORMAT_DATE_DDMMYY = "dd/mm/yy"
FORMAT_DATE_DMYSLASH = "d/m/y"
FORMAT_DATE_DMYMINUS = "d-m-y"
FORMAT_DATE_DMMINUS = "d-m"
FORMAT_DATE_MYMINUS = "m-y"
FORMAT_DATE_XLSX14 = "mm-dd-yy"
FORMAT_DATE_XLSX15 = "d-mmm-yy"
FORMAT_DATE_XLSX16 = "d-mmm"
FORMAT_DATE_XLSX17 = "mmm-yy"
FORMAT_DATE_XLSX22 = "m/d/yy h:mm"
FORMAT_DATE_DATETIME = "yyyy-mm-dd h:mm:ss"
FORMAT_DATE_TIME1 = "h:mm AM/PM"
FORMAT_DATE_TIME2 = "h:mm:ss AM/PM"
FORMAT_DATE_TIME3 = "h:mm"
FORMAT_DATE_TIME4 = "h:mm:ss"
FORMAT_DATE_TIME5 = "mm:ss"
FORMAT_DATE_TIME6 = "h:mm:ss"
FORMAT_DATE_TIME7 = "i:s.S"
FORMAT_DATE_TIME8 = "h:mm:ss@"
FORMAT_DATE_TIMEDELTA = "[hh]:mm:ss"
FORMAT_DATE_YYMMDDSLASH = "yy/mm/dd@"

STRIP_RE = re.compile(r'".*?"|\[(?!hh?\]|mm?\]|ss?\])[^\]]*\]')
TIMEDELTA_RE = re.compile(r"\[hh?\](:mm(:ss(\.0{0,3})?)?)?|\[mm?\](:ss(\.0{0,3})?)?|\[ss?\](\.0{0,3})?")


def builtin_format_code(index: int) -> str | None:
    """Return one of the standard format codes by index, or None."""
    return BUILTIN_FORMATS.get(index)


def builtin_format_id(fmt: str | None) -> int | None:
    """Return the id of a standard format code, or None."""
    return BUILTIN_FORMATS_REVERSE.get(fmt)


def is_builtin(fmt: str | None) -> bool:
    return fmt in BUILTIN_FORMATS.values()


def is_date_format(fmt: str | None) -> bool:
    if fmt is None:
        return False
    fmt = fmt.split(";")[0]
    fmt = STRIP_RE.sub("", fmt)
    return re.search(r"(?<![_\\])[dmhysDMHYS]", fmt) is not None


def is_datetime(fmt: str | None) -> str | None:
    """Return ``"date"``, ``"time"`` or ``"datetime"`` for a date format."""
    if not is_date_format(fmt):
        return None
    fmt = fmt.split(";")[0]
    date = any(char in fmt for char in "dy")
    time = any(char in fmt for char in "hs")
    if date and time:
        return "datetime"
    if date:
        return "date"
    return "time"


def is_timedelta_format(fmt: str | None) -> bool:
    if fmt is None:
        return False
    fmt = fmt.split(";")[0]
    return TIMEDELTA_RE.search(fmt) is not None
