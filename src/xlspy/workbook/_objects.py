"""Objects added to the sheets of a loaded workbook.

Charts, images, tables, data validation, conditional formats and print setup
are written by ``XlsxWriter`` into a scratch workbook holding the same sheet.
Their parts and sheet elements are then moved into the loaded package: parts
get free names and relationship ids, differential styles are appended to
``styles.xml`` and sheet elements go where the schema puts them. The file's
own objects stay: the new ones are added next to them.
"""

from __future__ import annotations

import io
import posixpath
import re
import zipfile
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape, quoteattr

from .._metadata import CT, MAIN, PKG_REL, REL, append_child, element_pattern, replace_element
from .._xlsx_edit import _attributes, _bytes
from ..refs import parse_cell_range
from ._print import print_settings_for

_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
_CONTENT = {
    "drawing": "application/vnd.openxmlformats-officedocument.drawing+xml",
    "chart": "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
    "table": "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml",
}
# The children of <worksheet>, in schema order (CT_Worksheet).
_SHEET_ORDER = (
    "sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData",
    "sheetCalcPr", "sheetProtection", "protectedRanges", "scenarios", "autoFilter",
    "sortState", "dataConsolidate", "customSheetViews", "mergeCells", "phoneticPr",
    "conditionalFormatting", "dataValidations", "hyperlinks", "printOptions",
    "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks",
    "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing",
    "legacyDrawing", "legacyDrawingHF", "drawingHF", "picture", "oleObjects",
    "controls", "webPublishItems", "tableParts", "extLst",
)
_ANCHOR = re.compile(r"<(oneCellAnchor|twoCellAnchor|absoluteAnchor)\b.*?</\1>", re.S)
_OBJECT_LISTS = ("charts", "images", "tables", "data_validations", "conditional_formatting")


def pending_objects(sheet) -> dict[str, list]:
    """The objects added since the sheet was loaded or last saved."""
    saved = sheet._saved_object_counts
    return {
        name: list(getattr(sheet, name))[saved.get(name, 0):]
        for name in _OBJECT_LISTS
    }


def has_pending_objects(sheet) -> bool:
    return any(pending_objects(sheet).values()) or print_settings_for(sheet) is not None


def mark_objects_saved(sheet) -> None:
    sheet._saved_object_counts = {name: len(list(getattr(sheet, name))) for name in _OBJECT_LISTS}


def _scratch(sheet, objects, settings) -> dict[str, bytes]:
    """The parts of a workbook holding only this sheet and its new objects."""
    from ..xlsx_writer import XlsxWriter
    from ._writer import _cell_values

    stream = io.BytesIO()
    with XlsxWriter(stream) as writer:
        writer.add_sheet(sheet.title)
        writer.write_sheet(
            # Tables take their column names from the header cells.
            _cell_values(sheet) if objects["tables"] else [[None]],
            charts=objects["charts"] or None,
            images=objects["images"] or None,
            tables=objects["tables"] or None,
            data_validations=objects["data_validations"] or None,
            conditional_formats=objects["conditional_formatting"] or None,
            print_settings=settings,
            auto_filter=False,
        )
    with zipfile.ZipFile(io.BytesIO(stream.getvalue())) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def _rels_path(part: str) -> str:
    return posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")


def _resolve(part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(part), target))


def _with_namespaces(fragment: str, default: str, **prefixes: str) -> str:
    """Declare the namespaces on a fragment's first element, so it reads the
    same wherever it is inserted (the host may use other prefixes)."""
    declarations = f' xmlns="{default}"' + "".join(
        f' xmlns:{prefix}="{uri}"' for prefix, uri in prefixes.items() if f"{prefix}:" in fragment
    )
    return re.sub(r"^<([\w.-]+)", lambda m: f"<{m[1]}{declarations}", fragment, count=1)


class _Package:
    """Pending part changes over the updater's package."""

    def __init__(self, updater):
        self.updater = updater
        self.package = updater._package
        self.parts: dict[str, bytes] = {}
        self.names = set(self.package.names())
        self.types = ET.fromstring(self.package.get("[Content_Types].xml"))

    def get(self, path: str) -> bytes | None:
        if path in self.parts:
            return self.parts[path]
        return self.package.get(path) if path in self.names else None

    def set(self, path: str, data: bytes | str) -> None:
        self.parts[path] = data.encode("utf-8") if isinstance(data, str) else data
        self.names.add(path)

    def free_part(self, directory: str, stem: str, extension: str) -> str:
        index = 1
        while f"{directory}/{stem}{index}.{extension}" in self.names:
            index += 1
        path = f"{directory}/{stem}{index}.{extension}"
        self.names.add(path)
        return path

    def rels(self, part: str) -> ET.Element:
        data = self.get(_rels_path(part))
        return ET.fromstring(data) if data else ET.Element(f"{{{PKG_REL}}}Relationships")

    def add_rel(self, rels: ET.Element, part: str, kind: str, target: str) -> str:
        used = {node.get("Id") for node in rels}
        index = 1
        while f"rId{index}" in used:
            index += 1
        rid = f"rId{index}"
        ET.SubElement(rels, f"{{{PKG_REL}}}Relationship", {
            "Id": rid, "Type": f"{REL}/{kind}",
            "Target": posixpath.relpath(target, posixpath.dirname(part)),
        })
        return rid

    def override(self, path: str, content_type: str) -> None:
        ET.SubElement(self.types, f"{{{CT}}}Override", {"PartName": "/" + path, "ContentType": content_type})

    def default(self, extension: str, content_type: str) -> None:
        if not any((node.get("Extension") or "").lower() == extension.lower() for node in self.types):
            ET.SubElement(self.types, f"{{{CT}}}Default", {"Extension": extension, "ContentType": content_type})

    def commit(self) -> None:
        ET.register_namespace("", CT)
        self.set("[Content_Types].xml", _bytes(self.types))
        self.updater._apply_parts(self.parts)


class _Sheet:
    """String edits of a host worksheet in whatever prefix it uses."""

    def __init__(self, xml: str):
        self.xml = xml
        self.prefix = re.search(r"<([\w.-]+:)?worksheet\b", xml)[1] or ""

    def _limit(self) -> int:
        """Where the worksheet's own extLst (or the end tag) starts: the
        x14 elements inside it share local names with main elements."""
        p = re.escape(self.prefix)
        end = re.search(rf"</{p}worksheet\s*>\s*$", self.xml).start()
        start = self.xml.rfind(f"<{self.prefix}extLst")
        if start >= 0 and re.fullmatch(rf"(?s).*</{p}extLst\s*>\s*", self.xml[start:end]):
            return start
        return end

    def find(self, tag: str):
        p = re.escape(self.prefix)
        pattern = re.compile(rf"<{p}{tag}\b[^>]*?(?:/>|>.*?</{p}{tag}\s*>)", re.S)
        return pattern.search(self.xml, 0, self._limit())

    def find_all(self, tag: str) -> list:
        p = re.escape(self.prefix)
        pattern = re.compile(rf"<{p}{tag}\b[^>]*?(?:/>|>.*?</{p}{tag}\s*>)", re.S)
        return list(pattern.finditer(self.xml, 0, self._limit()))

    def insert(self, tag: str, fragment: str) -> None:
        """Insert before the first element that follows ``tag`` in the schema
        (so after any existing ``tag`` elements)."""
        limit = self._limit()
        position = limit
        p = re.escape(self.prefix)
        for successor in _SHEET_ORDER[_SHEET_ORDER.index(tag) + 1:]:
            match = re.compile(rf"<{p}{successor}[\s/>]").search(self.xml, 0, limit)
            if match:
                position = match.start()
                break
        self.xml = self.xml[:position] + fragment + self.xml[position:]

    def replace(self, match, fragment: str) -> None:
        self.xml = self.xml[:match.start()] + fragment + self.xml[match.end():]


def _merge_attributes(sheet: _Sheet, tag: str, donor: str) -> None:
    """Set the donor element's attributes on the host's (keeping the rest,
    such as a printer-settings r:id), or insert the donor element."""
    match = sheet.find(tag)
    if match is None:
        sheet.insert(tag, _with_namespaces(donor, MAIN, r=REL))
        return
    attributes = _attributes(match[0])
    attributes.update(_attributes(donor))
    text = "".join(f" {name}={quoteattr(value)}" for name, value in attributes.items())
    sheet.replace(match, f"<{sheet.prefix}{tag}{text}/>")


def _merge_header_footer(sheet: _Sheet, donor: str) -> None:
    match = sheet.find("headerFooter")
    if match is None:
        sheet.insert("headerFooter", _with_namespaces(donor, MAIN))
        return
    p = sheet.prefix
    host = match[0]
    if host.endswith("/>"):
        host = host[:-2] + f"></{p}headerFooter>"
    for tag in ("oddHeader", "oddFooter"):
        found = re.search(rf"<{tag}>(.*?)</{tag}>", donor, re.S)
        if not found:
            continue
        child = f"<{p}{tag}>{found[1]}</{p}{tag}>"
        existing = re.search(rf"<{re.escape(p)}{tag}\b[^>]*?(?:/>|>.*?</{re.escape(p)}{tag}>)", host, re.S)
        if existing:
            host = host[:existing.start()] + child + host[existing.end():]
            continue
        after = re.search(rf"</{re.escape(p)}oddHeader>", host) if tag == "oddFooter" else None
        position = after.end() if after else host.index(">") + 1
        host = host[:position] + child + host[position:]
    sheet.replace(match, host)


def _set_fit_to_page(sheet: _Sheet) -> None:
    p = sheet.prefix
    setup = sheet.find("sheetPr")
    if setup is None:
        sheet.xml = re.sub(
            rf"(<{re.escape(p)}worksheet\b[^>]*>)",
            lambda m: m[1] + f'<{p}sheetPr><{p}pageSetUpPr fitToPage="1"/></{p}sheetPr>',
            sheet.xml, count=1,
        )
        return
    body = setup[0]
    existing = re.search(rf"<{re.escape(p)}pageSetUpPr\b[^>]*/?>", body)
    if existing:
        tag = existing[0]
        attributes = _attributes(tag)
        attributes["fitToPage"] = "1"
        text = "".join(f" {name}={quoteattr(value)}" for name, value in attributes.items())
        rebuilt = f"<{p}pageSetUpPr{text}" + ("/>" if tag.endswith("/>") else ">")
        body = body[:existing.start()] + rebuilt + body[existing.end():]
    elif body.endswith("/>"):
        body = body[:-2] + f'><{p}pageSetUpPr fitToPage="1"/></{p}sheetPr>'
    else:
        # pageSetUpPr is the last child of sheetPr.
        position = body.rfind("</")
        body = body[:position] + f'<{p}pageSetUpPr fitToPage="1"/>' + body[position:]
    sheet.replace(setup, body)


def _donor_elements(xml: str, tag: str) -> list[str]:
    return [match[0] for match in element_pattern(tag).finditer(xml)]


def _sheet_rels(scratch: dict[str, bytes]) -> dict[str, tuple[str, str]]:
    data = scratch.get("xl/worksheets/_rels/sheet1.xml.rels")
    if not data:
        return {}
    return {
        node.get("Id"): (node.get("Type").rsplit("/", 1)[-1], _resolve("xl/worksheets/sheet1.xml", node.get("Target")))
        for node in ET.fromstring(data)
    }


def _add_conditional_formats(sheet: _Sheet, scratch: dict[str, bytes], donor: str, styles) -> None:
    formats = _donor_elements(donor, "conditionalFormatting")
    if not formats:
        return
    styles_xml = scratch["xl/styles.xml"].decode("utf-8")
    dxfs = [_with_namespaces(item, MAIN) for item in _donor_elements(styles_xml, "dxf")]
    first = styles.append_dxfs(dxfs) if dxfs else 0
    # Priorities are unique per sheet: the new rules come after the file's.
    offset = max((int(value) for value in re.findall(r'\bpriority="(\d+)"', sheet.xml)), default=0)
    fragments = []
    for item in formats:
        item = re.sub(r'\bdxfId="(\d+)"', lambda m: f'dxfId="{first + int(m[1])}"', item)
        item = re.sub(r'\bpriority="(\d+)"', lambda m: f'priority="{offset + int(m[1])}"', item)
        fragments.append(_with_namespaces(item, MAIN))
    sheet.insert("conditionalFormatting", "".join(fragments))


def _add_validations(sheet: _Sheet, donor: str) -> None:
    found = _donor_elements(donor, "dataValidations")
    if not found:
        return
    rules = _donor_elements(found[0], "dataValidation")
    host = sheet.find("dataValidations")
    if host is None:
        sheet.insert("dataValidations", _with_namespaces(found[0], MAIN))
        return
    body = host[0]
    p = sheet.prefix
    count = len(re.findall(rf"<{re.escape(p)}dataValidation\b", body)) + len(rules)
    if body.endswith("/>"):
        body = body[:-2] + f"></{p}dataValidations>"
    position = body.rfind("</")
    body = body[:position] + "".join(_with_namespaces(rule, MAIN) for rule in rules) + body[position:]
    body = re.sub(r'\bcount="\d+"', f'count="{count}"', body, count=1)
    sheet.replace(host, body)


def _add_print_setup(sheet: _Sheet, donor: str, settings) -> None:
    if settings.center_horizontally or settings.center_vertically:
        for item in _donor_elements(donor, "printOptions"):
            _merge_attributes(sheet, "printOptions", item)
    if settings.margins is not None:
        margins = _donor_elements(donor, "pageMargins")[0]
        host = sheet.find("pageMargins")
        if host is None:
            sheet.insert("pageMargins", _with_namespaces(margins, MAIN))
        else:
            sheet.replace(host, _with_namespaces(margins, MAIN))
    if any(value is not None for value in (
        settings.orientation, settings.paper_size, settings.scale,
        settings.fit_to_width, settings.fit_to_height,
    )):
        _merge_attributes(sheet, "pageSetup", _donor_elements(donor, "pageSetup")[0])
    if settings.fit_to_width is not None or settings.fit_to_height is not None:
        _set_fit_to_page(sheet)
    if settings.header or settings.footer:
        _merge_header_footer(sheet, _donor_elements(donor, "headerFooter")[0])


def _set_print_names(package: _Package, scratch: dict[str, bytes], sheet_index: int, settings) -> None:
    wanted = []
    if settings.print_area:
        wanted.append("_xlnm.Print_Area")
    if settings.repeat_rows or settings.repeat_columns:
        wanted.append("_xlnm.Print_Titles")
    if not wanted:
        return
    donor = scratch["xl/workbook.xml"].decode("utf-8")
    xml = package.get("xl/workbook.xml").decode("utf-8")

    def drop(match):
        attributes = _attributes(match[0])
        keep = not (attributes.get("name") in wanted and attributes.get("localSheetId") == str(sheet_index))
        return match[0] if keep else ""

    xml = element_pattern("definedName").sub(drop, xml)
    fragments = []
    for item in _donor_elements(donor, "definedName"):
        name = _attributes(item).get("name")
        if name in wanted:
            formula = re.search(r">(.*)</definedName>", item, re.S)[1]
            fragments.append(
                f'<definedName xmlns="{MAIN}" name="{name}" localSheetId="{sheet_index}">{formula}</definedName>'
            )
    if element_pattern("definedNames").search(xml):
        for fragment in fragments:
            xml = append_child(xml, "definedNames", fragment)
    else:
        xml = replace_element(
            xml, "definedNames", f'<definedNames xmlns="{MAIN}">{"".join(fragments)}</definedNames>',
            before=("calcPr", "oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr",
                    "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst"),
        )
    package.set("xl/workbook.xml", xml)


def _add_drawing(package: _Package, sheet: _Sheet, sheet_path: str, sheet_rels: ET.Element,
                 scratch: dict[str, bytes], donor_rels: dict) -> None:
    drawing = next((target for kind, target in donor_rels.values() if kind == "drawing"), None)
    if drawing is None:
        return
    donor_xml = scratch[drawing].decode("utf-8")
    donor_drawing_rels = ET.fromstring(scratch[_rels_path(drawing)])

    host_drawing = next(
        (_resolve(sheet_path, node.get("Target")) for node in sheet_rels
         if node.get("Type", "").endswith("/drawing")),
        None,
    )
    if host_drawing is None:
        host_drawing = package.free_part("xl/drawings", "drawing", "xml")
        package.override(host_drawing, _CONTENT["drawing"])
        rid = package.add_rel(sheet_rels, sheet_path, "drawing", host_drawing)
        sheet.insert("drawing", _with_namespaces(f'<drawing r:id="{rid}"/>', MAIN, r=REL))
        host_xml = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="{_XDR}"></xdr:wsDr>'
    else:
        host_xml = package.get(host_drawing).decode("utf-8")
    drawing_rels = package.rels(host_drawing)

    mapping = {}
    for node in donor_drawing_rels:
        kind = node.get("Type").rsplit("/", 1)[-1]
        source = _resolve(drawing, node.get("Target"))
        if kind == "chart":
            target = package.free_part("xl/charts", "chart", "xml")
            package.override(target, _CONTENT["chart"])
        else:
            extension = source.rsplit(".", 1)[-1]
            target = package.free_part("xl/media", "image", extension)
            donor_types = ET.fromstring(scratch["[Content_Types].xml"])
            content_type = next(
                item.get("ContentType") for item in donor_types
                if (item.get("Extension") or "").lower() == extension.lower()
            )
            package.default(extension, content_type)
        package.set(target, scratch[source])
        mapping[node.get("Id")] = package.add_rel(drawing_rels, host_drawing, kind, target)

    # Shape ids are unique within a drawing.
    offset = max((int(value) for value in re.findall(r'<(?:[\w.-]+:)?cNvPr\b[^>]*?\bid="(\d+)"', host_xml)), default=0)
    anchors = []
    for match in _ANCHOR.finditer(donor_xml):
        anchor = re.sub(r'\br:(id|embed)="([^"]+)"', lambda m: f'r:{m[1]}="{mapping[m[2]]}"', match[0])
        anchor = re.sub(
            r'(<cNvPr\b[^>]*?\bid=")(\d+)"', lambda m: f'{m[1]}{offset + int(m[2])}"', anchor,
        )
        anchors.append(_with_namespaces(anchor, _XDR, a=_A, c=_C, r=REL))
    end = re.search(r"</(?:[\w.-]+:)?wsDr\s*>\s*$", host_xml)
    if end is None:  # an empty <xdr:wsDr .../>
        root = re.search(r"<((?:[\w.-]+:)?wsDr)\b([^>]*)/>", host_xml)
        host_xml = host_xml[:root.start()] + f"<{root[1]}{root[2]}>" + "".join(anchors) + f"</{root[1]}>" + host_xml[root.end():]
    else:
        host_xml = host_xml[:end.start()] + "".join(anchors) + host_xml[end.start():]
    package.set(host_drawing, host_xml)
    ET.register_namespace("", PKG_REL)
    package.set(_rels_path(host_drawing), _bytes(drawing_rels))


def _existing_tables(package: _Package) -> list[tuple[str, dict]]:
    tables = []
    for name in sorted(package.names):
        if name.startswith("xl/tables/") and name.endswith(".xml"):
            data = package.get(name)
            if data:
                root = re.search(r"<(?:[\w.-]+:)?table\b[^>]*>", data.decode("utf-8"))
                if root:
                    tables.append((name, _attributes(root[0])))
    return tables


def _overlaps(first: str, second: str) -> bool:
    a = parse_cell_range(first)
    b = parse_cell_range(second)
    return not (a[2] < b[0] or b[2] < a[0] or a[3] < b[1] or b[3] < a[1])


def _add_tables(package: _Package, sheet: _Sheet, sheet_path: str, sheet_rels: ET.Element,
                scratch: dict[str, bytes], donor_rels: dict) -> None:
    sources = [target for kind, target in donor_rels.values() if kind == "table"]
    if not sources:
        return
    existing = _existing_tables(package)
    names = {attributes.get("displayName", "").casefold() for _, attributes in existing}
    on_sheet = {
        _resolve(sheet_path, node.get("Target")) for node in sheet_rels
        if node.get("Type", "").endswith("/table")
    }
    next_id = max((int(attributes.get("id", 0)) for _, attributes in existing), default=0) + 1
    fragments = []
    for source in sources:
        xml = scratch[source].decode("utf-8")
        attributes = _attributes(re.search(r"<table\b[^>]*>", xml)[0])
        if attributes["displayName"].casefold() in names:
            raise ValueError(f"a table named {attributes['displayName']!r} already exists in the workbook")
        for name, other in existing:
            if name in on_sheet and _overlaps(attributes["ref"], other.get("ref", "A1")):
                raise ValueError(
                    f"table {attributes['displayName']!r} ({attributes['ref']}) overlaps the "
                    f"sheet's table {other.get('displayName')!r} ({other.get('ref')})"
                )
        xml = re.sub(r'(<table\b[^>]*?\bid=")\d+"', lambda m: f'{m[1]}{next_id}"', xml, count=1)
        next_id += 1
        target = package.free_part("xl/tables", "table", "xml")
        package.set(target, xml)
        package.override(target, _CONTENT["table"])
        rid = package.add_rel(sheet_rels, sheet_path, "table", target)
        fragments.append(f'<tablePart r:id="{rid}"/>')
    host = sheet.find("tableParts")
    if host is None:
        body = f'<tableParts count="{len(fragments)}">{"".join(fragments)}</tableParts>'
        sheet.insert("tableParts", _with_namespaces(body, MAIN, r=REL))
        return
    body = host[0]
    p = sheet.prefix
    count = len(re.findall(rf"<{re.escape(p)}tablePart\b", body)) + len(fragments)
    if body.endswith("/>"):
        body = body[:-2] + f"></{p}tableParts>"
    position = body.rfind("</")
    added = "".join(_with_namespaces(fragment, MAIN, r=REL) for fragment in fragments)
    body = body[:position] + added + body[position:]
    body = re.sub(r'\bcount="\d+"', f'count="{count}"', body, count=1)
    sheet.replace(host, body)


def add_objects(updater, workbook, sheet, sheet_order: list[str]) -> None:
    """Move the sheet's pending objects into the loaded package."""
    objects = pending_objects(sheet)
    settings = print_settings_for(sheet)
    if not any(objects.values()) and settings is None:
        return
    scratch = _scratch(sheet, objects, settings)
    donor = scratch["xl/worksheets/sheet1.xml"].decode("utf-8")
    donor_rels = _sheet_rels(scratch)

    package = _Package(updater)
    sheet_path = updater._sheets[sheet.title]
    host = _Sheet(package.get(sheet_path).decode("utf-8"))
    sheet_rels = package.rels(sheet_path)

    # Tables first: their name and range checks are the steps that can fail,
    # and nothing outside this package (the style table) is changed yet.
    _add_tables(package, host, sheet_path, sheet_rels, scratch, donor_rels)
    _add_drawing(package, host, sheet_path, sheet_rels, scratch, donor_rels)
    _add_validations(host, donor)
    if settings is not None:
        _add_print_setup(host, donor, settings)
        _set_print_names(package, scratch, sheet_order.index(sheet.title), settings)
    _add_conditional_formats(host, scratch, donor, workbook._styles)

    package.set(sheet_path, host.xml)
    if len(sheet_rels):
        ET.register_namespace("", PKG_REL)
        package.set(_rels_path(sheet_path), _bytes(sheet_rels))
    if objects["conditional_formatting"]:
        package.set("xl/styles.xml", workbook._styles.to_bytes())
    package.commit()
    # The package now holds them: a save that fails later (another sheet's
    # table name, the output file) must not add them again when retried.
    mark_objects_saved(sheet)
