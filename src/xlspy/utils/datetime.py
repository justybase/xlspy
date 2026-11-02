"""Excel/Python date conversions compatible with openpyxl."""

from __future__ import annotations

import datetime
import re
from ..excel_dates import MAC_EPOCH as MAC_EPOCH
from ..excel_dates import WINDOWS_EPOCH as WINDOWS_EPOCH

SECS_PER_DAY = 24 * 60 * 60

_DURATION_RE = re.compile(
    r"^P"
    r"(?:(?P<days>\d+(?:\.\d+)?)D)?"
    r"(?:T"
    r"(?:(?P<hours>\d+(?:\.\d+)?)H)?"
    r"(?:(?P<minutes>\d+(?:\.\d+)?)M)?"
    r"(?:(?P<seconds>\d+(?:\.\d+)?)S)?"
    r")?$"
)


def days_to_time(value: datetime.timedelta) -> datetime.time:
    minutes, seconds = divmod(value.seconds, 60)
    hours, minutes = divmod(minutes, 60)
    return datetime.time(hours, minutes, seconds, value.microseconds)


def time_to_days(value: datetime.time) -> float:
    return (
        value.hour * 3600
        + value.minute * 60
        + value.second
        + value.microsecond / 10**6
    ) / SECS_PER_DAY


def timedelta_to_days(value: datetime.timedelta) -> float:
    return value.total_seconds() / SECS_PER_DAY


def from_excel(value, epoch=WINDOWS_EPOCH, timedelta=False):
    """Convert an Excel serial number to a datetime, date, time or timedelta."""
    if value is None:
        return None

    if timedelta:
        result = datetime.timedelta(days=value)
        if result.microseconds:
            result = datetime.timedelta(
                seconds=result.total_seconds() // 1,
                microseconds=round(result.microseconds, -3),
            )
        return result

    day, fraction = divmod(value, 1)
    diff = datetime.timedelta(milliseconds=round(fraction * SECS_PER_DAY * 1000))
    if 0 <= value < 1 and diff.days == 0:
        return days_to_time(diff)
    if 0 < value < 60 and epoch == WINDOWS_EPOCH:
        day += 1
    return epoch + datetime.timedelta(days=day) + diff


def to_excel(value, epoch=WINDOWS_EPOCH):
    """Convert a date/datetime/time/timedelta to an Excel serial number."""
    if value is None:
        return None
    if isinstance(value, datetime.time):
        return time_to_days(value)
    if isinstance(value, datetime.timedelta):
        return timedelta_to_days(value)
    if not hasattr(value, "date"):
        value = datetime.datetime.combine(value, datetime.time())

    days = (value - epoch).days
    if 0 < days <= 60 and epoch == WINDOWS_EPOCH:
        days -= 1
    return days + time_to_days(value)


def from_ISO8601(formatted_string):
    """Parse the ISO-8601 date/time/duration forms Excel stores."""
    if not formatted_string:
        return None

    duration = _DURATION_RE.match(formatted_string)
    if duration is not None and any(duration.groups()):
        parts = {
            key: float(value)
            for key, value in duration.groupdict().items()
            if value is not None
        }
        return datetime.timedelta(**parts)

    text = formatted_string
    if text.endswith("Z"):
        text = text[:-1]
    try:
        if "T" in text or (len(text) > 10 and text[10] in "+-"):
            value = datetime.datetime.fromisoformat(text)
            return value.replace(tzinfo=None) if value.tzinfo is not None else value
        if len(text) == 10:
            return datetime.date.fromisoformat(text)
        parsed_time = datetime.time.fromisoformat(text)
        return parsed_time.replace(tzinfo=None) if parsed_time.tzinfo is not None else parsed_time
    except ValueError as exc:
        raise ValueError(f"Invalid datetime value {formatted_string}") from exc


def to_ISO8601(value):
    """Convert a date/time to an ISO-8601 string, keeping milliseconds."""
    if hasattr(value, "microsecond") and value.microsecond:
        return value.isoformat(timespec="milliseconds")
    return value.isoformat()
