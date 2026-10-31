"""Capture how Excel stores sparklines, table slicers and pivot tables in XLSB, one option at a time.

Builds one workbook per variant through Excel COM (pywin32), saves it as XLSB,
and prints the records of interest with the bytes that differ from the base
variant marked: BrtBeginSparklineGroup (0x0411) and BrtSparkline (0x0413), or
with ``--slicers`` the slicer (0x043B) and slicer cache (0x0435, 0x043D, 0x081D)
records, or with ``--pivots`` every record of the pivot table part and the pivot
cache definition. Use it to map record fields before changing ``_xlsb_sparklines``,
``_xlsb_slicers`` or ``_xlsb_pivots``.

Usage:
    py -3.14 tools/capture_xlsb.py OUT_DIR [variant ...]
    py -3.14 tools/capture_xlsb.py OUT_DIR --slicers [variant ...]
    py -3.14 tools/capture_xlsb.py OUT_DIR --pivots [variant ...]

The XLSB files stay in OUT_DIR for inspection. Excel must be installed.
"""
import os
import sys
import zipfile

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "../src")))

from xlspy.biff_utils import iter_records  # noqa: E402

GROUP = 0x0411
SPARKLINE = 0x0413
XL_EXCEL12 = 50


def _group(sheet, location="E2", data="Sheet1!B2:D2", kind=1):
    return sheet.Range(location).SparklineGroups.Add(kind, data)


def _set(attribute):
    def apply(group):
        target = group
        *path, last = attribute.split(".")
        for name in path:
            target = getattr(target, name)
        setattr(target, last, True)
    return apply


def _custom_axis(group):
    group.Axes.Vertical.MinScaleType = 3  # xlSparkScaleCustom
    group.Axes.Vertical.CustomMinScaleValue = -5
    group.Axes.Vertical.MaxScaleType = 3
    group.Axes.Vertical.CustomMaxScaleValue = 50


def _custom_min(group):
    group.Axes.Vertical.MinScaleType = 3
    group.Axes.Vertical.CustomMinScaleValue = -5


def _shared_axis(group):
    group.Axes.Vertical.MinScaleType = 1  # xlSparkScaleGroup (2 is xlSparkScaleSingle)
    group.Axes.Vertical.MaxScaleType = 1


VARIANTS = {
    "base": {},
    "column": {"kind": 2},
    "win_loss": {"kind": 3},
    "two_rows": {"location": "E2:E3", "data": "Sheet1!B2:D3"},
    "other_sheet": {"data": "Data!B2:D2"},
    "markers": {"edit": _set("Points.Markers.Visible")},
    "high": {"edit": _set("Points.Highpoint.Visible")},
    "low": {"edit": _set("Points.Lowpoint.Visible")},
    "first": {"edit": _set("Points.Firstpoint.Visible")},
    "last": {"edit": _set("Points.Lastpoint.Visible")},
    "negative": {"edit": _set("Points.Negative.Visible")},
    "series_color": {"edit": lambda g: setattr(g.SeriesColor, "Color", 0x332211)},  # BGR
    "line_weight": {"edit": lambda g: setattr(g, "LineWeight", 2.25)},
    "empty_zero": {"edit": lambda g: setattr(g, "DisplayBlanksAs", 2)},
    "empty_connect": {"edit": lambda g: setattr(g, "DisplayBlanksAs", 3)},
    "show_hidden": {"edit": lambda g: setattr(g, "DisplayHidden", True)},
    "show_axis": {"edit": lambda g: setattr(g.Axes.Horizontal.Axis, "Visible", True)},
    "right_to_left": {"edit": lambda g: setattr(g.Axes.Horizontal, "RightToLeftPlotOrder", True)},
    "custom_axis": {"edit": _custom_axis},
    "shared_axis": {"edit": _shared_axis},
    "shared_axis_two": {"location": "E2:E3", "data": "Sheet1!B2:D3", "edit": _shared_axis},
    "custom_min": {"edit": _custom_min},
}


SLICER_VARIANTS = {
    "slicer_base": {},
    "slicer_named": {"name": "MySlicer", "caption": "My caption"},
    "slicer_columns": {"columns": 2},
    "slicer_second_column": {"column": "Month"},
    "slicer_no_caption": {"show_caption": False},
    "slicer_second_table": {"second_table": True},
    "slicer_two": {"two": True},
}
SLICER_RECORDS = {
    "xl/slicers/slicer1.bin": (0x043B,),
    "xl/slicerCaches/slicerCache1.bin": (0x0435, 0x043D, 0x081D),
}


def build_slicer(excel, out_dir, name, spec):
    import pythoncom

    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Name = "Sheet1"
        sheet.Activate()  # Slicers.Add needs the active, visible sheet
        rows = [("Region", "Month", "Sales"), ("North", "Jan", 10), ("South", "Feb", 7)]
        for row, values in enumerate(rows, start=1):
            for column, value in enumerate(values, start=1):
                sheet.Cells(row, column).Value = value
        table = sheet.ListObjects.Add(1, sheet.Range("A1:C3"), None, 1)
        table.Name = "Sales"
        if spec.get("second_table"):
            for row, values in enumerate(rows, start=6):
                for column, value in enumerate(values, start=1):
                    sheet.Cells(row, column).Value = value
            table = sheet.ListObjects.Add(1, sheet.Range("A6:C8"), None, 1)
            table.Name = "Costs"
        cache = book.SlicerCaches.Add2(table, spec.get("column", "Region"))
        slicer = cache.Slicers.Add(sheet, pythoncom.Missing, pythoncom.Missing, pythoncom.Missing,
                                   20, 300, 150, 150)
        if spec.get("two"):
            book.SlicerCaches.Add2(table, "Month").Slicers.Add(
                sheet, pythoncom.Missing, pythoncom.Missing, pythoncom.Missing, 20, 500, 150, 150)
        if "name" in spec:
            slicer.Name = spec["name"]
        if "caption" in spec:
            slicer.Caption = spec["caption"]
        if "columns" in spec:
            slicer.NumberOfColumns = spec["columns"]
        if "show_caption" in spec:
            slicer.DisplayHeader = spec["show_caption"]
        path = os.path.join(out_dir, f"{name}.xlsb")
        if os.path.exists(path):
            os.remove(path)
        book.SaveAs(path, XL_EXCEL12)
    finally:
        book.Close(False)
    return path


def slicer_records(path):
    found = []
    with zipfile.ZipFile(path) as archive:
        for part, ids in SLICER_RECORDS.items():
            data = archive.read(part)
            found += [(record.record_id, bytes(data[record.data_start:record.data_end]))
                      for record in iter_records(data) if record.record_id in ids]
    return found


PIVOT_ROWS = [
    ("Region", "Product", "Quarter", "Sales", "Units"),
    ("North", "Coffee", 1, 1200, 40), ("North", "Tea", 1, 800, 50),
    ("South", "Coffee", 1, 1500, 55), ("South", "Tea", 2, 700, 30),
    ("East", "Coffee", 2, 900, 33), ("East", "Tea", 2, 1100, 60),
    ("North", "Coffee", 2, 1300, 44), ("South", "Coffee", 2, 1600, 58),
]
XL_ROW, XL_COLUMN, XL_PAGE, XL_DATA = 1, 2, 3, 4
XL_SUM, XL_COUNT, XL_AVERAGE, XL_MAX, XL_MIN = -4157, -4112, -4106, -4136, -4139


def _pivot_variant(*, rows=("Region",), columns=(), filters=(), values=(("Sales", XL_SUM),),
                   location="A3", number_format=None):
    return {"rows": rows, "columns": columns, "filters": filters, "values": values,
            "location": location, "number_format": number_format}


PIVOT_VARIANTS = {
    "pivot_base": _pivot_variant(),
    "pivot_column": _pivot_variant(columns=("Quarter",)),
    "pivot_text_column": _pivot_variant(columns=("Product",)),
    "pivot_two_values": _pivot_variant(values=(("Sales", XL_SUM), ("Units", XL_SUM))),
    "pivot_column_two_values": _pivot_variant(
        columns=("Quarter",), values=(("Sales", XL_SUM), ("Units", XL_SUM))),
    "pivot_filter": _pivot_variant(filters=("Product",)),
    "pivot_average": _pivot_variant(values=(("Units", XL_AVERAGE),)),
    "pivot_count": _pivot_variant(values=(("Sales", XL_COUNT),)),
    "pivot_min": _pivot_variant(values=(("Sales", XL_MIN),)),
    "pivot_max": _pivot_variant(values=(("Sales", XL_MAX),)),
    "pivot_number_format": _pivot_variant(number_format="#,##0"),
    "pivot_moved": _pivot_variant(location="C6"),
}
# The pivot table part and the cache definition; the first record of each is a
# random revision id (0x0C00), which would differ in every file.
PIVOT_PARTS = ("xl/pivotTables/pivotTable1.bin", "xl/pivotCache/pivotCacheDefinition1.bin")
UID_RECORD = 0x0C00


def build_pivot(excel, out_dir, name, spec):
    """A table on "Data" and a pivot over it on "Summary", saved without cache records."""
    book = excel.Workbooks.Add()
    try:
        data = book.Worksheets(1)
        data.Name = "Data"
        for row, values in enumerate(PIVOT_ROWS, start=1):
            for column, value in enumerate(values, start=1):
                data.Cells(row, column).Value = value
        data.ListObjects.Add(1, data.Range("A1:E9"), None, 1).Name = "SalesTable"
        summary = book.Worksheets.Add(After=data)
        summary.Name = "Summary"
        cache = book.PivotCaches().Create(1, "SalesTable")
        cache.RefreshOnFileOpen = True
        pivot = cache.CreatePivotTable(summary.Range(spec["location"]), "PT")
        for orientation, fields in ((XL_ROW, spec["rows"]), (XL_COLUMN, spec["columns"]),
                                    (XL_PAGE, spec["filters"])):
            for field in fields:
                pivot.PivotFields(field).Orientation = orientation
        for field, function in spec["values"]:
            target = pivot.PivotFields(field)
            target.Orientation = XL_DATA
            target.Function = function
            if spec["number_format"]:
                target.NumberFormat = spec["number_format"]
        pivot.SaveData = False
        path = os.path.join(out_dir, f"{name}.xlsb")
        if os.path.exists(path):
            os.remove(path)
        book.SaveAs(path, XL_EXCEL12)
    finally:
        book.Close(False)
    return path


def pivot_records(path):
    found = []
    with zipfile.ZipFile(path) as archive:
        for part in PIVOT_PARTS:
            data = archive.read(part)
            found += [(record.record_id, bytes(data[record.data_start:record.data_end]))
                      for record in iter_records(data) if record.record_id != UID_RECORD]
    return found


def build(excel, out_dir, name, spec):
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Name = "Sheet1"
        data = book.Worksheets.Add(After=sheet)
        data.Name = "Data"
        for target in (sheet, data):
            for row, values in enumerate(([3, -1, 4], [1, 5, -9]), start=2):
                for column, value in enumerate(values, start=2):
                    target.Cells(row, column).Value = value
        group = _group(sheet, spec.get("location", "E2"), spec.get("data", "Sheet1!B2:D2"),
                       spec.get("kind", 1))
        if "edit" in spec:
            spec["edit"](group)
        path = os.path.join(out_dir, f"{name}.xlsb")
        if os.path.exists(path):
            os.remove(path)
        book.SaveAs(path, XL_EXCEL12)
    finally:
        book.Close(False)
    return path


def sparkline_records(path):
    found = []
    with zipfile.ZipFile(path) as archive:
        for name in sorted(archive.namelist()):
            if not (name.startswith("xl/worksheets/") and name.endswith(".bin")):
                continue
            data = archive.read(name)
            found += [(record.record_id, bytes(data[record.data_start:record.data_end]))
                      for record in iter_records(data) if record.record_id in (GROUP, SPARKLINE)]
    return found


def show(name, records, base):
    print(f"== {name}")
    for index, (record_id, payload) in enumerate(records):
        reference = base[index][1] if index < len(base) and base[index][0] == record_id else b""
        marks = "".join(
            "^^ " if index_ >= len(reference) or reference[index_] != byte else "   "
            for index_, byte in enumerate(payload))
        print(f"  0x{record_id:04X} len={len(payload)}")
        for start in range(0, len(payload), 24):
            print(f"    {start:3d}  {payload[start:start + 24].hex(' ')}")
            if reference and marks[start * 3:(start + 24) * 3].strip():
                print(f"         {marks[start * 3:(start + 24) * 3]}")


def main(argv=None):
    import win32com.client

    argv = sys.argv[1:] if argv is None else argv
    out_dir = os.path.abspath(argv[0])
    slicers = "--slicers" in argv
    pivots = "--pivots" in argv
    argv = [arg for arg in argv if arg not in ("--slicers", "--pivots")]
    if slicers:
        variants, builder, reader, base_name = (
            SLICER_VARIANTS, build_slicer, slicer_records, "slicer_base")
    elif pivots:
        variants, builder, reader, base_name = (
            PIVOT_VARIANTS, build_pivot, pivot_records, "pivot_base")
    else:
        variants, builder, reader, base_name = VARIANTS, build, sparkline_records, "base"
    names = list(dict.fromkeys([base_name, *(argv[1:] or variants)]))
    os.makedirs(out_dir, exist_ok=True)
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = slicers
    excel.DisplayAlerts = False
    try:
        paths = {name: builder(excel, out_dir, name, variants[name]) for name in names}
    finally:
        excel.Quit()
    base = reader(paths[base_name])
    for name in names:
        show(name, reader(paths[name]), base if name != base_name else [])


if __name__ == "__main__":
    main()
