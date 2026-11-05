"""Cases shared by the Excel COM conformance tests.

A case is data only, so the same spec drives three checks:

- xlspy writes every case as XLSX and XLSB, and Excel reads it back
  (structure, values, number formats);
- Excel writes every case, and xlspy reads it back;
- an updater replaces a sheet of a template and Excel reads the result.

A cell is a plain value, a ``datetime``/``date``, or a dict:
``{"value": ...}``, ``{"date": "YYYY-MM-DDTHH:MM:SS"}`` and/or
``{"numberFormat": "..."}``. A sheet may hold ``{"generator": {"rows": n,
"columns": m}}`` instead of ``rows`` for a large sheet; ``sample_every`` limits
the per-cell checks of such a sheet (structure and the first/last row are always
checked).
"""

from __future__ import annotations

import datetime as dt

CASES: list[dict] = [
    {
        "id": "values",
        "sheets": [
            {
                "name": "Values",
                "rows": [
                    ["Text", "Number", "Float", "Bool", "Date", "DateTime", "Empty"],
                    ["plain", 42, 3.5, True, dt.date(2024, 1, 31),
                     dt.datetime(2024, 1, 31, 13, 45, 30), None],
                    ["zażółć gęślą jaźń", -7, -0.25, False, dt.date(1999, 12, 31),
                     dt.datetime(1999, 12, 31, 23, 59, 59), None],
                    ["with space ", 0, 1e-7, True, None, None, None],
                ],
            }
        ],
    },
    {
        "id": "formats",
        "sheets": [
            {
                "name": "Formats",
                "rows": [
                    ["Amount", "Percent", "Date", "Integer", "Decimal"],
                    [
                        {"value": 1234.5, "numberFormat": "#,##0.00"},
                        {"value": 0.256, "numberFormat": "0.0%"},
                        {"date": "2024-01-31T00:00:00", "numberFormat": "yyyy-mm-dd"},
                        {"value": 1234567, "numberFormat": "#,##0"},
                        {"value": 12.5, "numberFormat": "0.00"},
                    ],
                ],
            }
        ],
    },
    {
        "id": "sheets",
        "sheets": [
            {"name": "First", "rows": [["a", "b"], [1, 2]]},
            {"name": "Hidden", "hidden": True, "rows": [["x"], [7]]},
            {"name": "Last", "rows": [["z"], [9]]},
        ],
    },
    {
        "id": "autofilter",
        "sheets": [
            {
                "name": "Filtered",
                "autoFilter": True,
                "rows": [["Name", "Qty"], ["a", 1], ["b", 2], ["c", 3]],
            }
        ],
    },
    {
        "id": "ragged",
        "sheets": [
            {
                "name": "Ragged",
                "rows": [
                    ["h1", "h2", "h3"],
                    [1],
                    [],
                    [2, 3],
                    [None, None, None, 4],
                ],
            }
        ],
    },
    {
        "id": "empty",
        "sheets": [{"name": "Empty", "rows": []}],
    },
    {
        "id": "unicode",
        "sheets": [
            {
                "name": "Unicode",
                "rows": [
                    ["text"],
                    ["zażółć gęślą jaźń"],
                    ["emoji \U0001F600\U0001F680"],
                    ["  leading and trailing  "],
                    ["tab\there"],
                    ["line\nbreak"],
                    ["a&b<c>d"],
                ],
            }
        ],
    },
    {
        "id": "generated",
        # Excel authoring this case would take 10,000 COM calls per format; the
        # reader side is covered by the other cases and the xlspy-authored file.
        "excel_authored": False,
        "sheets": [
            {
                "name": "Generated",
                "rows": {"generator": {"rows": 2000, "columns": 5}},
                "sample_every": 250,
            }
        ],
    },
]

# A template and the rows a replace_sheet_data call puts in it. The header row
# stays; the body is replaced, formulas included.
UPDATE_TEMPLATE: dict = {
    "id": "update_template",
    "sheets": [
        {
            "name": "Data",
            "rows": [["Name", "Qty", "Price"], ["old", 1, 1.0], ["old", 2, 2.0]],
        }
    ],
}

UPDATE_RESULT: dict = {
    "id": "update_result",
    "sheets": [
        {
            "name": "Data",
            "rows": [
                ["Name", "Qty", "Price"],
                ["alpha", 10, 1.5],
                ["beta", 20, 2.5],
                ["gamma", 30, 3.5],
            ],
        }
    ],
}
