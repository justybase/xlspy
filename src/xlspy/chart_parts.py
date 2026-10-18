"""Shared DrawingML chart/drawing builders for the XLSX and XLSB writers.

Both package formats store drawings and charts as the same XML parts
(``xl/drawings/drawingN.xml``, ``xl/charts/chartN.xml``); only the sheet
records that link to them differ (``<drawing>`` element vs. ``BrtDrawing``
record). Keeping the XML generation in one place keeps the writers in sync.

The XLSB path (``xlsb=True``) additionally mirrors what real Excel files
contain: cached values (``strCache``/``numCache``) inside every reference,
text categories via ``strRef``, and an explicit ``varyColors`` flag. Excel's
XLSB loader evidently depends on the caches for series structure; without
them single-series charts over vertical ranges render transposed (one series
per row, category names as legend). The XLSX path stays byte-identical to
the openpyxl oracle and is intentionally left cache-free.
"""

from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal
from xml.sax import saxutils

from .styles import parse_cell_range

_CHART_TYPES = ("bar", "line", "pie")

_CHART_AXES_XML = (
    "<catAx><axId val=\"10\"/><scaling><orientation val=\"minMax\"/></scaling>"
    "<axPos val=\"l\"/><majorTickMark val=\"none\"/><minorTickMark val=\"none\"/>"
    "<crossAx val=\"100\"/><lblOffset val=\"100\"/></catAx>"
    "<valAx><axId val=\"100\"/><scaling><orientation val=\"minMax\"/></scaling>"
    "<axPos val=\"l\"/><majorGridlines/><majorTickMark val=\"none\"/>"
    "<minorTickMark val=\"none\"/><crossAx val=\"10\"/></valAx>"
)

# The XLSB loader in Excel does not apply the same attribute defaults as the
# XLSX one: without explicit delete/axPos/tickLblPos the axis (and therefore
# its tick labels) is not drawn. Real Excel-generated XLSB charts spell these
# out, so the XLSB path mirrors them. The XLSX path stays byte-identical to
# the openpyxl oracle, which renders correctly without them.
_XLSB_CHART_AXES_XML = (
    "<catAx><axId val=\"10\"/><scaling><orientation val=\"minMax\"/></scaling>"
    "<delete val=\"0\"/><axPos val=\"b\"/>"
    "<majorTickMark val=\"none\"/><minorTickMark val=\"none\"/>"
    "<tickLblPos val=\"nextTo\"/><crossAx val=\"100\"/><lblOffset val=\"100\"/></catAx>"
    "<valAx><axId val=\"100\"/><scaling><orientation val=\"minMax\"/></scaling>"
    "<delete val=\"0\"/><axPos val=\"l\"/><majorGridlines/>"
    "<majorTickMark val=\"none\"/><minorTickMark val=\"none\"/>"
    "<tickLblPos val=\"nextTo\"/><crossAx val=\"10\"/></valAx>"
)


def column_letters() -> list[str]:
    """Column labels ``A`` … ``XFD`` (shared with the writers)."""
    letters = []
    for i in range(26):
        letters.append(chr(65 + i))
    for i in range(26):
        for j in range(26):
            letters.append(chr(65 + i) + chr(65 + j))
    for i in range(24):
        for j in range(26):
            for k in range(26):
                if i == 23 and j > 24:
                    break
                letters.append(chr(65 + i) + chr(65 + j) + chr(65 + k))
    return letters


def quote_sheet_name(sheet_name: str) -> str:
    return "'" + sheet_name.replace("'", "''") + "'"


def chart_ref(letters: list[str], sheet_name: str, ref: str, *,
              absolute: bool = True) -> str:
    first_col, first_row, last_col, last_row = parse_cell_range(ref)
    if absolute:
        start = f"${letters[first_col]}${first_row + 1}"
        end = f"${letters[last_col]}${last_row + 1}"
        range_ref = f"{start}:{end}"
    elif (first_col, first_row) == (last_col, last_row):
        range_ref = f"{letters[first_col]}{first_row + 1}"
    else:
        range_ref = (
            f"{letters[first_col]}{first_row + 1}"
            f":{letters[last_col]}{last_row + 1}"
        )
    return f"{quote_sheet_name(sheet_name)}!{range_ref}"


def _range_values(rows: list[list] | None, ref: str) -> list:
    """Cell values covered by ``ref`` in row-major order (``None`` if unknown)."""
    if rows is None:
        raise ValueError("chart caches need the sheet rows; pass rows=...")
    first_col, first_row, last_col, last_row = parse_cell_range(ref)
    values = []
    for row_idx in range(first_row, last_row + 1):
        row = rows[row_idx] if 0 <= row_idx < len(rows) else []
        for col_idx in range(first_col, last_col + 1):
            values.append(row[col_idx] if 0 <= col_idx < len(row) else None)
    return values


def _chart_number_text(value) -> str:
    """Shortest plain-text literal for a cached chart number."""
    from .excel_dates import datetime_to_serial

    if isinstance(value, bool):
        return "1" if value else "0"
    if isinstance(value, Decimal):
        value = float(value)
    if isinstance(value, (datetime, date)):
        value = datetime_to_serial(value)
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def _chart_text(value) -> str:
    """Plain-text form of a cached chart string."""
    from .formats import _unwrap_cell

    if value is None:
        return ""
    value = _unwrap_cell(value)
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    return str(value)


def _str_cache_xml(values: list) -> str:
    points = "".join(
        f"<pt idx=\"{i}\"><v>{saxutils.escape(_chart_text(v))}</v></pt>"
        for i, v in enumerate(values)
    )
    return f"<strCache><ptCount val=\"{len(values)}\"/>{points}</strCache>"


def _num_cache_xml(values: list) -> str:
    from .formats import _unwrap_cell

    points = []
    for i, raw in enumerate(values):
        value = _unwrap_cell(raw) if raw is not None else None
        if value is None:
            text = "0"
        elif isinstance(value, str):
            try:
                text = _chart_number_text(float(value))
            except ValueError:
                text = "0"
        else:
            text = _chart_number_text(value)
        points.append(f"<pt idx=\"{i}\"><v>{text}</v></pt>")
    joined = "".join(points)
    return (
        "<numCache><formatCode>General</formatCode>"
        f"<ptCount val=\"{len(values)}\"/>{joined}</numCache>"
    )


def _is_numeric_range(values: list) -> bool:
    from .formats import _unwrap_cell

    seen = False
    for raw in values:
        value = _unwrap_cell(raw) if raw is not None else None
        if value is None:
            continue
        if isinstance(value, (str, bool)):
            return False
        seen = True
    return seen


def chart_series_xml(letters: list[str], series, sheet_name: str,
                     index: int, chart_type: str, rows=None,
                     xlsb: bool = False) -> str:
    if series.title_cell is not None:
        formula = chart_ref(
            letters, sheet_name, series.title_cell, absolute=False,
        )
        if xlsb:
            title_value = _range_values(rows, series.title_cell)[0]
            title = (
                "<tx><strRef><f>"
                f"{formula}"
                "</f>"
                f"{_str_cache_xml([title_value])}</strRef></tx>"
            )
        else:
            title = (
                "<tx><strRef><f>"
                f"{formula}"
                "</f></strRef></tx>"
            )
    elif series.title is not None:
        title = (
            "<tx><v>"
            f"{saxutils.escape(series.title)}"
            "</v></tx>"
        )
    else:
        title = ""
    categories = ""
    if chart_type != "pie" and series.categories is not None:
        formula = chart_ref(letters, sheet_name, series.categories)
        if xlsb:
            cat_values = _range_values(rows, series.categories)
            if _is_numeric_range(cat_values):
                categories = (
                    "<cat><numRef><f>"
                    f"{formula}"
                    "</f>"
                    f"{_num_cache_xml(cat_values)}</numRef></cat>"
                )
            else:
                categories = (
                    "<cat><strRef><f>"
                    f"{formula}"
                    "</f>"
                    f"{_str_cache_xml(cat_values)}</strRef></cat>"
                )
        else:
            categories = (
                "<cat><numRef><f>"
                f"{formula}"
                "</f></numRef></cat>"
            )
    marker = ""
    if chart_type == "line":
        marker = (
            "<marker><symbol val=\"none\"/>"
            "<spPr><a:ln><a:prstDash val=\"solid\"/></a:ln></spPr></marker>"
        )
    values_formula = chart_ref(letters, sheet_name, series.values)
    if xlsb:
        values_xml = (
            "<val><numRef><f>"
            f"{values_formula}"
            "</f>"
            f"{_num_cache_xml(_range_values(rows, series.values))}</numRef></val>"
        )
    else:
        values_xml = (
            "<val><numRef><f>"
            f"{values_formula}"
            "</f></numRef></val>"
        )
    return (
        f"<ser><idx val=\"{index}\"/><order val=\"{index}\"/>{title}"
        "<spPr><a:ln><a:prstDash val=\"solid\"/></a:ln></spPr>"
        f"{marker}{categories}"
        f"{values_xml}</ser>"
    )


def chart_xml(letters: list[str], chart, sheet_name: str, rows=None,
              xlsb: bool = False) -> str:
    parts = []
    if chart.style is not None:
        parts.append(f"<style val=\"{chart.style}\"/>")
    parts.append("<chart>")
    if chart.title is not None:
        parts.append(
            "<title><tx><rich><a:bodyPr/><a:p><a:pPr><a:defRPr/></a:pPr>"
            f"<a:r><a:t>{saxutils.escape(chart.title)}</a:t></a:r></a:p></rich></tx></title>"
        )
    series_xml = "".join(
        chart_series_xml(
            letters, item, sheet_name, index, chart.chart_type,
            rows=rows, xlsb=xlsb,
        )
        for index, item in enumerate(chart.series)
    )
    vary = "<varyColors val=\"0\"/>" if xlsb else ""
    axes = _XLSB_CHART_AXES_XML if xlsb else _CHART_AXES_XML
    parts.append("<plotArea>")
    if chart.chart_type == "bar":
        parts.append(
            "<barChart><barDir val=\"col\"/><grouping val=\"clustered\"/>"
            f"{vary}{series_xml}"
            "<gapWidth val=\"150\"/><axId val=\"10\"/><axId val=\"100\"/></barChart>"
            f"{axes}"
        )
    elif chart.chart_type == "line":
        parts.append(
            "<lineChart><grouping val=\"standard\"/>"
            f"{vary}{series_xml}"
            "<axId val=\"10\"/><axId val=\"100\"/></lineChart>"
            f"{axes}"
        )
    else:
        parts.append(
            f"<pieChart><varyColors val=\"1\"/>{series_xml}"
            "<firstSliceAng val=\"0\"/></pieChart>"
        )
    parts.append("</plotArea>")
    parts.append(
        "<legend><legendPos val=\"r\"/></legend>"
        "<plotVisOnly val=\"1\"/><dispBlanksAs val=\"gap\"/>"
    )
    parts.append("</chart>")
    return (
        "<chartSpace "
        "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" "
        "xmlns=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">"
        f"{''.join(parts)}</chartSpace>"
    )


def drawing_xml(charts: list, chart_ids: list[int]) -> str:
    frames = []
    for position, (chart, chart_id) in enumerate(zip(charts, chart_ids), start=1):
        first_col, first_row, _, _ = parse_cell_range(chart.anchor)
        width_emu = round(chart.width * 360000)
        height_emu = round(chart.height * 360000)
        frames.append(
            "<oneCellAnchor>"
            f"<from><col>{first_col}</col><colOff>0</colOff>"
            f"<row>{first_row}</row><rowOff>0</rowOff></from>"
            f"<ext cx=\"{width_emu}\" cy=\"{height_emu}\"/>"
            "<graphicFrame><nvGraphicFramePr>"
            f"<cNvPr id=\"{position}\" name=\"Chart {position}\"/>"
            "<cNvGraphicFramePr/></nvGraphicFramePr><xfrm/>"
            "<a:graphic><a:graphicData "
            "uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">"
            f"<c:chart r:id=\"rId{position}\"/>"
            "</a:graphicData></a:graphic></graphicFrame>"
            "<clientData/></oneCellAnchor>"
        )
    return (
        "<wsDr "
        "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" "
        "xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" "
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" "
        "xmlns=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\">"
        f"{''.join(frames)}</wsDr>"
    )


def drawing_rels_xml(chart_ids: list[int]) -> str:
    parts = "".join(
        f'<Relationship Id="rId{position}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" '
        f'Target="../charts/chart{chart_id}.xml"/>'
        for position, chart_id in enumerate(chart_ids, start=1)
    )
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
{parts}
</Relationships>'''

# Standard Office theme (Excel 365) bundled with chart-enabled workbooks.
# Chart default styles and accent colors resolve through the theme; without
# it the XLSX and XLSB parsers in Excel fall back to different defaults
# (observed: per-point colors + category legend on style-less line charts
# in XLSB). Content-verbatim copy of ``xl/theme/theme1.xml`` from an Excel file.
THEME_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="5B9BD5"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="4472C4"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light" panose="020F0302020204030204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="Yu Gothic Light"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="等线 Light"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:majorFont><a:minorFont><a:latin typeface="Calibri" panose="020F0502020204030204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="Yu Gothic"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="等线"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/><a:extLst><a:ext uri="{05A4C25C-085E-4340-85A3-A5531E510DB2}"><thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="Office Theme" id="{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}" vid="{4A3C46E8-61CC-4603-A589-7422A47A8E4A}"/></a:ext></a:extLst></a:theme>
"""
