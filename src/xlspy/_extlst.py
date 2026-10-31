"""Namespaces and extension-list (``extLst``) fragments written by the XLSX writer.

The URIs and element names come from a workbook saved by Excel itself, so the
writer and Excel agree byte for byte on what an extension means.
"""

X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
X15 = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
XM = "http://schemas.microsoft.com/office/excel/2006/main"
# Slicer drawing objects: the Choice branch is keyed on sle15, the graphic on sle.
SLICER_DRAWING_CHOICE = "http://schemas.microsoft.com/office/drawing/2012/slicer"
SLICER_DRAWING = "http://schemas.microsoft.com/office/drawing/2010/slicer"

EXT_SPARKLINES = "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}"
EXT_SHEET_SLICER_LIST = "{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}"
EXT_WORKBOOK_SLICER_CACHES = "{46BE6895-7355-4a93-B00E-2C351335B9C9}"
EXT_TABLE_SLICER_CACHE = "{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}"

REL_SLICER = "http://schemas.microsoft.com/office/2007/relationships/slicer"
REL_SLICER_CACHE = "http://schemas.microsoft.com/office/2007/relationships/slicerCache"
CT_SLICER = "application/vnd.ms-excel.slicer+xml"
CT_SLICER_CACHE = "application/vnd.ms-excel.slicerCache+xml"


def ext(uri: str, body: str, *namespaces: tuple[str, str]) -> str:
    """One ``<ext>`` element; ``namespaces`` are (prefix, uri) pairs declared on it."""
    decl = "".join(f' xmlns:{prefix}="{value}"' for prefix, value in namespaces)
    return f'<ext uri="{uri}"{decl}>{body}</ext>'


def ext_list(exts) -> str:
    """The ``extLst`` element holding ``exts``, or nothing when there are none."""
    exts = [item for item in exts if item]
    if not exts:
        return ""
    return f"<extLst>{''.join(exts)}</extLst>"
