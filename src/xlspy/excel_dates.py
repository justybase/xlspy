"""Conversions between Excel serial numbers and Python datetimes.

Excel stores dates as serial numbers.  The Windows/1900 system counts days
from ``1899-12-30`` but *also* contains the famous fictitious leap day
1900-02-29: serial 1 is 1900-01-01 and serial 60 is the non-existent
1900-02-29.  The Mac/1904 system counts days from ``1904-01-01`` and has no
such quirk.

The functions below mirror ``openpyxl``'s ``to_excel``/``from_excel`` (used as
the reference implementation): serials ``1..59`` are shifted by one day when
read, serial 60 and 59 both read as 1900-02-28, and dates before 1900-03-01
are written one serial lower.
"""

import datetime

WINDOWS_EPOCH = datetime.datetime(1899, 12, 30)
MAC_EPOCH = datetime.datetime(1904, 1, 1)
_FIRST_MARCH_1900 = datetime.datetime(1900, 3, 1)


def serial_to_datetime(serial: float, date_1904: bool = False) -> datetime.datetime:
    """Convert an Excel serial number to a ``datetime``."""

    if date_1904:
        return MAC_EPOCH + datetime.timedelta(days=serial)
    if 1 <= serial < 60:
        serial += 1
    return WINDOWS_EPOCH + datetime.timedelta(days=serial)


def datetime_to_serial(value, date_1904: bool = False) -> float:
    """Convert a ``date``/``datetime`` to an Excel serial number.

    Timezone-aware datetimes keep their offset: the serial describes the
    wall-clock moment, exactly as the writers did before this helper existed.
    """

    moment = (
        value
        if isinstance(value, datetime.datetime)
        else datetime.datetime.combine(value, datetime.time())
    )
    naive = moment.replace(tzinfo=None)
    if date_1904:
        epoch = MAC_EPOCH
    elif naive < _FIRST_MARCH_1900:
        # 1900-01-01 is serial 1, not 2: Excel counts the fictitious
        # 1900-02-29, so serials before 1900-03-01 are one day lower.
        epoch = datetime.datetime(1899, 12, 31)
    else:
        epoch = WINDOWS_EPOCH
    if moment.tzinfo is not None:
        epoch = epoch.replace(tzinfo=moment.tzinfo)
    return (moment - epoch).total_seconds() / 86400.0
