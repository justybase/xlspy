"""The names of Excel's table totals functions, without their cell encoders."""

from __future__ import annotations

ILTA_NONE = 0
ILTA_SUM = 6
ILTA_CUSTOM = 9

ILTA_BY_NAME = {
    "none": ILTA_NONE,
    "average": 1,
    "count": 2,
    "countNums": 3,
    "max": 4,
    "min": 5,
    "sum": ILTA_SUM,
    "stdDev": 7,
    "var": 8,
    "custom": ILTA_CUSTOM,
}
TOTALS_FUNCTION_NAMES = frozenset(ILTA_BY_NAME) | {"label"}
