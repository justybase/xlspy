"""Pivot tables written with ``XlsxWriter``: the pre-rendered report, the parts
and the checks that refuse a pivot Excel could not build.

The rendered layouts were compared with Excel's own output through COM (see
``tools/validate_excel_com.ps1 -RefreshPivots``).
"""

import io
import zipfile

import openpyxl
import pytest

from xlspy import ExcelTable, PivotTable, PivotValue, XlsxWriter
from xlspy.workbook._loaded import sheet_has_pivots

ROWS = [
    ["Region", "Product", "Quarter", "Sales", "Units"],
    ["North", "Coffee", 1, 1200, 40],
    ["North", "Tea", 1, 800, 50],
    ["South", "Coffee", 1, 1500, 55],
    ["South", "Tea", 2, 700, 30],
    ["East", "Coffee", 2, 900, 33],
    ["East", "Tea", 2, 1100, 60],
    ["North", "Coffee", 2, 1300, 44],
    ["South", "Coffee", 2, 1600, 58],
]


def pivot(values="Sales", *, location="A12", **kwargs):
    """A pivot over the ``SalesTable`` the tests write, below its data."""
    kwargs.setdefault("rows", "Region")
    return PivotTable(source="SalesTable", values=values, location=location, **kwargs)


def make(pivots, *, summary=None):
    """Write ``ROWS`` as ``SalesTable`` on "Data" with ``pivots``.

    ``summary`` adds a second sheet: ``"empty"`` leaves it unwritten and
    ``"written"`` writes an empty row stream to it.
    """
    out = io.BytesIO()
    with XlsxWriter(out) as w:
        w.add_sheet("Data")
        w.write_sheet(ROWS, tables=[ExcelTable("SalesTable", "A1:E9")], pivots=pivots)
        if summary is not None:
            w.add_sheet("Summary")
            if summary == "written":
                w.write_sheet([])
    return out.getvalue()


def cells(data: bytes, sheet: str, ref: str) -> list[list]:
    workbook = openpyxl.load_workbook(io.BytesIO(data))
    return [[cell.value for cell in row] for row in workbook[sheet][ref]]


def parts(data: bytes) -> dict[str, str]:
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        return {
            name: zf.read(name).decode("utf-8")
            for name in zf.namelist()
            if name.endswith((".xml", ".rels"))
        }


ROW_ONLY = [
    ["Row Labels", "Sum of Sales"],
    ["East", 2000],
    ["North", 3300],
    ["South", 3800],
    ["Grand Total", 9100],
]


def test_row_field_with_one_value_renders_the_report():
    assert cells(make([pivot()]), "Data", "A12:B16") == ROW_ONLY


def test_several_values_add_one_column_each():
    data = make([pivot(["Sales", "Units"])])
    assert cells(data, "Data", "A12:C16") == [
        ["Row Labels", "Sum of Sales", "Sum of Units"],
        ["East", 2000, 93],
        ["North", 3300, 134],
        ["South", 3800, 143],
        ["Grand Total", 9100, 370],
    ]


@pytest.mark.parametrize(
    ("function", "caption", "expected"),
    [
        ("count", "Count of Sales", [2, 3, 3, 8]),
        ("min", "Min of Sales", [900, 800, 700, 700]),
        ("max", "Max of Sales", [1100, 1300, 1600, 1600]),
    ],
)
def test_count_min_and_max(function, caption, expected):
    data = make([pivot([PivotValue("Sales", function)])])
    rendered = cells(data, "Data", "A12:B16")
    assert rendered[0] == ["Row Labels", caption]
    assert [row[1] for row in rendered[1:]] == expected


def test_average_is_not_a_sum():
    data = make([pivot([PivotValue("Units", "average")])])
    rendered = cells(data, "Data", "A12:B16")
    assert rendered[0] == ["Row Labels", "Average of Units"]
    assert rendered[1] == ["East", pytest.approx(46.5)]
    assert rendered[2] == ["North", pytest.approx(134 / 3)]
    assert rendered[3] == ["South", pytest.approx(143 / 3)]
    assert rendered[4] == ["Grand Total", pytest.approx(46.25)]


def test_row_and_column_fields_with_one_value():
    data = make([pivot(columns="Quarter")])
    assert cells(data, "Data", "A12:D17") == [
        ["Sum of Sales", "Column Labels", None, None],
        ["Row Labels", 1, 2, "Grand Total"],
        ["East", None, 2000, 2000],
        ["North", 2000, 1300, 3300],
        ["South", 1500, 2300, 3800],
        ["Grand Total", 3500, 5600, 9100],
    ]


def test_row_and_column_fields_with_several_values():
    data = make([pivot(["Sales", "Units"], columns="Quarter")])
    assert cells(data, "Data", "A12:G18") == [
        [None, "Column Labels", None, None, None, None, None],
        [None, 1, None, 2, None, "Total: Sum of Sales", "Total: Sum of Units"],
        ["Row Labels", "Sum of Sales", "Sum of Units", "Sum of Sales", "Sum of Units", None, None],
        ["East", None, None, 2000, 93, 2000, 93],
        ["North", 2000, 90, 1300, 44, 3300, 134],
        ["South", 1500, 55, 2300, 88, 3800, 143],
        ["Grand Total", 3500, 145, 5600, 225, 9100, 370],
    ]


def test_filter_is_a_page_field_above_the_report():
    data = make([pivot(filters="Product")])
    assert cells(data, "Data", "A10:B10") == [["Product", "(All)"]]
    assert cells(data, "Data", "A12:B16") == ROW_ONLY


def test_pivot_on_another_sheet_renders_there():
    out = io.BytesIO()
    with XlsxWriter(out) as w:
        w.add_sheet("Data")
        w.write_sheet(
            ROWS,
            tables=[ExcelTable("SalesTable", "A1:E9")],
            pivots=[pivot(location="A3", sheet="Summary")],
        )
        w.add_sheet("Summary")
        w.write_sheet([])
    data = out.getvalue()
    assert cells(data, "Summary", "A3:B7") == ROW_ONLY
    assert cells(data, "Data", "A1:A1") == [["Region"]]


def test_pivot_on_an_unwritten_sheet_is_placed_on_it():
    data = make([pivot(location="A3", sheet="Summary")], summary="empty")
    assert cells(data, "Summary", "A3:B7") == ROW_ONLY


def test_column_field_cache_carries_numeric_range():
    cache = next(
        body for name, body in parts(make([pivot(columns="Quarter")])).items()
        if name.startswith("xl/pivotCache/pivotCacheDefinition")
    )
    assert 'minValue="1" maxValue="2"' in cache
    assert 'containsNumber="1"' in cache


def test_pivot_parts_follow_the_cache_without_records():
    files = parts(make([pivot(), pivot(["Units"], name="Units_pivot", location="A20")]))
    caches = [body for name, body in files.items() if "pivotCacheDefinition" in name
              and name.endswith(".xml")]
    assert len(caches) == 2
    for cache in caches:
        assert 'saveData="0"' in cache
        assert 'refreshOnLoad="1"' in cache
        assert 'worksheetSource name="SalesTable"' in cache
    assert not any("pivotCacheRecords" in name for name in files)
    assert 'Target="../pivotCache/pivotCacheDefinition1.xml"' in files[
        "xl/pivotTables/_rels/pivotTable1.xml.rels"
    ]
    assert 'name="Units_pivot"' in files["xl/pivotTables/pivotTable2.xml"]
    assert 'hideValuesRow="1"' in files["xl/pivotTables/pivotTable1.xml"]
    workbook = files["xl/workbook.xml"]
    assert workbook.count('<pivotCache cacheId=') == 2
    assert workbook.index("<calcPr") < workbook.index("<pivotCaches>")
    content_types = files["[Content_Types].xml"]
    assert "/xl/pivotTables/pivotTable1.xml" in content_types
    assert "/xl/pivotCache/pivotCacheDefinition2.xml" in content_types
    sheet_rels = files["xl/worksheets/_rels/sheet1.xml.rels"]
    assert "../pivotTables/pivotTable1.xml" in sheet_rels


def test_repeated_column_items_name_the_data_field():
    # As Excel writes them: the item that repeats the quarter holds the data field's index.
    files = parts(make([pivot(["Sales", "Units"], columns="Quarter")]))
    table = files["xl/pivotTables/pivotTable1.xml"]
    assert '<i r="1" i="1"><x v="1"/></i>' in table
    assert '<i r="1" i="1"><x v="0"/></i>' not in table


def test_workbook_without_pivots_has_no_pivot_parts():
    files = parts(make([]))
    assert not any("pivot" in name.lower() for name in files)
    assert "pivotCaches" not in files["xl/workbook.xml"]


def test_write_dataframe_passes_pivots_through():
    pd = pytest.importorskip("pandas")
    frame = pd.DataFrame(ROWS[1:], columns=ROWS[0])
    out = io.BytesIO()
    with XlsxWriter(out) as w:
        w.add_sheet("Data")
        w.write_dataframe(
            frame,
            tables=[ExcelTable("SalesTable", "A1:E9")],
            pivots=[pivot()],
        )
    assert cells(out.getvalue(), "Data", "A12:B16") == ROW_ONLY


def test_removing_a_pivot_sheet_from_a_loaded_workbook_is_refused():
    package = {
        "xl/worksheets/sheet1.xml": b"<worksheet/>",
        "xl/worksheets/_rels/sheet1.xml.rels": (
            b'<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
            b'officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/>'
            b"</Relationships>"
        ),
    }

    class Package:
        def names(self):
            return list(package)

        def get(self, name):
            return package[name]

    assert sheet_has_pivots(Package(), "xl/worksheets/sheet1.xml")
    assert not sheet_has_pivots(Package(), "xl/worksheets/sheet2.xml")


# -- refusals -----------------------------------------------------------------


def test_one_row_field_is_required():
    with pytest.raises(ValueError, match="exactly one row field"):
        PivotTable(source="SalesTable", rows=["Region", "Product"], values="Sales")
    with pytest.raises(ValueError, match="exactly one row field"):
        PivotTable(source="SalesTable", rows=(), values="Sales")


def test_one_column_field_at_most():
    with pytest.raises(ValueError, match="at most one column field"):
        PivotTable(source="SalesTable", rows="Region", columns=["Quarter", "Product"],
                   values="Sales")


def test_a_field_cannot_sit_on_two_axes():
    with pytest.raises(ValueError, match="two axes"):
        PivotTable(source="SalesTable", rows="Region", columns="Region", values="Sales")


def test_unknown_function_is_refused():
    with pytest.raises(ValueError, match="pivot value function"):
        PivotValue("Sales", "median")


def test_location_must_be_one_cell():
    with pytest.raises(ValueError, match="single cell"):
        PivotTable(source="SalesTable", rows="Region", values="Sales", location="A1:B2")


def test_invalid_pivot_name_is_refused():
    with pytest.raises(ValueError, match="invalid pivot table name"):
        PivotTable(source="SalesTable", rows="Region", values="Sales", name="bad name")


def test_unknown_source_table_is_refused():
    with pytest.raises(ValueError, match="not a table written"):
        make([PivotTable(source="Nope", rows="Region", values="Sales", location="A12")])


def test_unknown_field_is_refused():
    with pytest.raises(ValueError, match="not a column"):
        make([pivot("Revenue")])


def test_duplicate_pivot_names_are_refused():
    with pytest.raises(ValueError, match="duplicate pivot table name"):
        make([pivot(name="Same"), pivot(["Units"], name="same", location="A20")])


def test_pivot_on_the_source_sheet_must_sit_below_the_data():
    with pytest.raises(ValueError, match="below the sheet's data"):
        make([pivot(location="A5")])


def test_overlapping_pivots_are_refused():
    with pytest.raises(ValueError, match="overlap"):
        make([pivot(location="A12"), pivot(["Units"], location="A13")])


def test_target_sheet_written_before_its_source_is_refused():
    out = io.BytesIO()
    with pytest.raises(ValueError, match="written before its source"):
        with XlsxWriter(out) as w:
            w.add_sheet("Summary")
            w.write_sheet([])
            w.add_sheet("Data")
            w.write_sheet(
                ROWS,
                tables=[ExcelTable("SalesTable", "A1:E9")],
                pivots=[pivot(location="A3", sheet="Summary")],
            )


def test_missing_target_sheet_is_refused_when_saved():
    with pytest.raises(ValueError, match="never added to the workbook"):
        make([pivot(location="A3", sheet="Missing")])
