"""Objects added to the sheets of a loaded workbook.

Charts, images, tables, data validation, conditional formats and print setup
are written by ``XlsxWriter`` into a scratch workbook holding the same sheet.
Their parts and sheet elements are then moved into the loaded package: parts
get free names and relationship ids, differential styles are appended to
``styles.xml`` and sheet elements go where the schema puts them. The file's
own objects stay: the new ones are added next to them.
"""

from __future__ import annotations
from .._package_handle import PackageHandle

from ..exceptions import UnsupportedFeatureError
import io
import posixpath
import re

from .._untrusted import SafeZipFile
from dataclasses import fields
from xml.etree import ElementTree as ET
from .._safe_xml import fromstring as safe_fromstring
from xml.sax.saxutils import quoteattr

from .._metadata import CT, MAIN, PKG_REL, REL, append_child, element_pattern, replace_element
from .._xlsx_edit import _attributes, _bytes, _required_match
from ..refs import parse_cell_range
from ._print import WorksheetPageMargins, print_settings_for
from ..chart_reader import _A as _A
from ..chart_reader import _C as _C
from ..chart_reader import _XDR as _XDR

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


def _object_values(sheet, name: str) -> list:
    value = getattr(sheet, name)
    if name == "conditional_formatting":
        return list(value.all_rules())
    return list(value)


def pending_objects(sheet) -> dict[str, list]:
    """The objects added since the sheet was loaded or last saved."""
    from ._drawings import new_drawing_objects

    saved = sheet._saved_object_counts
    pending = {
        name: _object_values(sheet, name)[saved.get(name, 0):]
        for name in _OBJECT_LISTS
    }
    # Charts and images can be removed: the new ones are those not in the file.
    pending.update(new_drawing_objects(sheet))
    return pending


def has_pending_objects(sheet) -> bool:
    return (
        any(pending_objects(sheet).values())
        or print_settings_for(sheet) != getattr(sheet, "_saved_print_settings", None)
    )


def mark_objects_saved(sheet) -> None:
    from ._drawings import remember_drawing_objects

    remember_drawing_objects(sheet)
    sheet._saved_object_counts = {
        name: len(_object_values(sheet, name)) for name in _OBJECT_LISTS
    }
    sheet._saved_object_signatures = object_signatures(sheet)
    sheet._saved_print_settings = print_settings_for(sheet)


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
    with SafeZipFile(io.BytesIO(stream.getvalue())) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def _rels_path(part: str) -> str:
    return posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")


def _resolve_part_target(part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(part), target))


def _target(node: ET.Element) -> str:
    target = node.get("Target")
    if target is None:
        raise ValueError("relationship has no Target")
    return target


def _with_namespaces(fragment: str, default: str, **prefixes: str) -> str:
    """Declare the namespaces on a fragment's first element, so it reads the
    same wherever it is inserted (the host may use other prefixes)."""
    declarations = f' xmlns="{default}"' + "".join(
        f' xmlns:{prefix}="{uri}"' for prefix, uri in prefixes.items() if f"{prefix}:" in fragment
    )
    return re.sub(r"^<([\w.-]+)", lambda m: f"<{m[1]}{declarations}", fragment, count=1)


class _Package:
    """Pending part changes over the package of an updater."""

    def __init__(self, updater):
        handle = PackageHandle(updater)
        self.package = handle.package
        self.parts: dict[str, bytes] = {}
        self.names = set(self.package.names())
        self.removed: set[str] = set()
        self.types = safe_fromstring(self.package.get("[Content_Types].xml"))

    def get(self, path: str) -> bytes | None:
        if path in self.parts:
            return self.parts[path]
        if path in self.removed:
            return None
        return self.package.get(path) if path in self.names else None

    def require(self, path: str) -> bytes:
        data = self.get(path)
        if data is None:
            raise ValueError(f"package part {path} was not found")
        return data

    def set(self, path: str, data: bytes | str) -> None:
        self.parts[path] = data.encode("utf-8") if isinstance(data, str) else data
        self.names.add(path)
        self.removed.discard(path)

    def remove(self, path: str) -> None:
        self.parts.pop(path, None)
        self.names.discard(path)
        self.removed.add(path)

    def free_part(self, directory: str, stem: str, extension: str) -> str:
        index = 1
        while f"{directory}/{stem}{index}.{extension}" in self.names:
            index += 1
        path = f"{directory}/{stem}{index}.{extension}"
        self.names.add(path)
        return path

    def rels(self, part: str) -> ET.Element:
        data = self.get(_rels_path(part))
        return safe_fromstring(data) if data else ET.Element(f"{{{PKG_REL}}}Relationships")

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
        self.package.apply_parts(self.parts, self.removed - self.parts.keys())


class _Sheet:
    """String edits of a host worksheet in whatever prefix it uses."""

    def __init__(self, xml: str):
        self.xml = xml
        self.prefix = _required_match(re.search(r"<([\w.-]+:)?worksheet\b", xml), "worksheet element")[1] or ""

    def _limit(self) -> int:
        """Where the worksheet's own extLst (or the end tag) starts: the
        x14 elements inside it share local names with main elements."""
        p = re.escape(self.prefix)
        end = _required_match(re.search(rf"</{p}worksheet\s*>\s*$", self.xml), "worksheet end tag").start()
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

    def remove_all(self, tag: str) -> None:
        self.xml = element_pattern(tag).sub("", self.xml)

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
        node.get("Id"): (node.get("Type").rsplit("/", 1)[-1], _resolve_part_target("xl/worksheets/sheet1.xml", node.get("Target")))
        for node in safe_fromstring(data)
    }


def _add_conditional_formats(sheet: _Sheet, scratch: dict[str, bytes], donor: str, styles) -> None:
    formats = _donor_elements(donor, "conditionalFormatting")
    if not formats:
        return
    styles_xml = scratch["xl/styles.xml"].decode("utf-8")
    dxfs = [_with_namespaces(item, MAIN) for item in _donor_elements(styles_xml, "dxf")]
    indexes = styles.append_dxfs(dxfs) if dxfs else []
    # Priorities are unique per sheet: the new rules come after the file's.
    offset = max((int(value) for value in re.findall(r'\bpriority="(\d+)"', sheet.xml)), default=0)
    fragments = []
    for item in formats:
        item = re.sub(r'\bdxfId="(\d+)"', lambda m: f'dxfId="{indexes[int(m[1])]}"', item)
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


def _with_prefix(fragment: str, prefix: str) -> str:
    """Re-tag a donor fragment with the host worksheet's element prefix."""
    if not prefix:
        return fragment
    fragment = re.sub(r"^<([\w.-]+)", lambda m: f"<{prefix}{m[1]}", fragment, count=1)
    return re.sub(r"</([\w.-]+)\s*>\s*$", lambda m: f"</{prefix}{m[1]}>", fragment, count=1)


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
            # The element stays in the host part: keep its prefix and let the
            # worksheet root's default namespace cover it.
            sheet.replace(host, _with_prefix(margins, sheet.prefix))
    if any(value is not None for value in (
        settings.orientation, settings.paper_size, settings.scale,
        settings.fit_to_width, settings.fit_to_height,
    )):
        _merge_attributes(sheet, "pageSetup", _donor_elements(donor, "pageSetup")[0])
    if settings.fit_to_width is not None or settings.fit_to_height is not None:
        _set_fit_to_page(sheet)
    if settings.header or settings.footer:
        _merge_header_footer(sheet, _donor_elements(donor, "headerFooter")[0])


# PrintSettings fields kept as <pageSetup> attributes.
_PAGE_SETUP_ATTRIBUTES = {
    "orientation": "orientation",
    "paper_size": "paperSize",
    "scale": "scale",
    "fit_to_width": "fitToWidth",
    "fit_to_height": "fitToHeight",
}


def _print_value_set(settings, name: str) -> bool:
    if settings is None:
        return False
    value = getattr(settings, name)
    # fitToWidth="0" (automatic) is a value; False and "" are not.
    return value is not None and value is not False and value != ""


def _cleared_print_fields(saved, settings) -> set[str]:
    """The print settings the file had when loaded and the sheet no longer has."""
    if saved is None:
        return set()
    return {
        item.name for item in fields(saved)
        if _print_value_set(saved, item.name) and not _print_value_set(settings, item.name)
    }


def _without_attributes(tag: str, names) -> str:
    for name in names:
        tag = re.sub(rf"""\s{name}=(?:"[^"]*"|'[^']*')""", "", tag)
    return tag


def _drop_attributes(sheet: _Sheet, tag: str, names) -> None:
    match = sheet.find(tag)
    if match is None:
        return
    opening = _required_match(re.match(r"<[^>]*?(?=/?>)", match[0]), "opening tag")[0]
    sheet.replace(match, _without_attributes(opening, names) + match[0][len(opening):])


def _drop_fit_to_page(sheet: _Sheet) -> None:
    setup = sheet.find("sheetPr")
    if setup is None:
        return
    body = setup[0]
    existing = re.search(rf"<{re.escape(sheet.prefix)}pageSetUpPr\b[^>]*?(?=/?>)", body)
    if existing is None:
        return
    tag = _without_attributes(existing[0], ["fitToPage"])
    sheet.replace(setup, body[:existing.start()] + tag + body[existing.end():])


def _drop_print_names(package: _Package, sheet_index: int, names) -> None:
    xml = package.require("xl/workbook.xml").decode("utf-8")

    def drop(match):
        attributes = _attributes(match[0])
        keep = not (attributes.get("name") in names and attributes.get("localSheetId") == str(sheet_index))
        return match[0] if keep else ""

    edited = element_pattern("definedName").sub(drop, xml)
    if edited == xml:
        return
    # An emptied <definedNames> goes too.
    edited = re.sub(r"<([\w.-]+:)?definedNames\b[^>]*>\s*</\1?definedNames\s*>", "", edited)
    edited = re.sub(r"<([\w.-]+:)?definedNames\b[^>]*/>", "", edited)
    package.set("xl/workbook.xml", edited)


def _clear_print_setup(package: _Package, sheet: _Sheet, sheet_index: int, cleared: set[str], settings) -> None:
    """Remove the print settings the sheet had when loaded and no longer has."""
    names = []
    if "print_area" in cleared:
        names.append("_xlnm.Print_Area")
    if cleared & {"repeat_rows", "repeat_columns"} and not (
        _print_value_set(settings, "repeat_rows") or _print_value_set(settings, "repeat_columns")
    ):
        names.append("_xlnm.Print_Titles")
    if names:
        _drop_print_names(package, sheet_index, names)
    if "margins" in cleared:
        host = sheet.find("pageMargins")
        if host is not None:
            margins = WorksheetPageMargins()
            text = "".join(
                f' {item.name}="{getattr(margins, item.name)}"' for item in fields(WorksheetPageMargins)
            )
            sheet.replace(host, f"<{sheet.prefix}pageMargins{text}/>")
    attributes = [attribute for name, attribute in _PAGE_SETUP_ATTRIBUTES.items() if name in cleared]
    if attributes:
        _drop_attributes(sheet, "pageSetup", attributes)
    if cleared & {"fit_to_width", "fit_to_height"}:
        _drop_fit_to_page(sheet)
    centered = [
        attribute for name, attribute in (
            ("center_horizontally", "horizontalCentered"), ("center_vertically", "verticalCentered"),
        ) if name in cleared
    ]
    if centered:
        _drop_attributes(sheet, "printOptions", centered)
    match = sheet.find("headerFooter")
    if match is not None:
        body = match[0]
        p = re.escape(sheet.prefix)
        for name, tag in (("header", "oddHeader"), ("footer", "oddFooter")):
            if name in cleared:
                body = re.sub(rf"<{p}{tag}\b[^>]*?(?:/>|>.*?</{p}{tag}\s*>)", "", body, flags=re.S)
        if body != match[0]:
            sheet.replace(match, body)


def _set_print_names(package: _Package, scratch: dict[str, bytes], sheet_index: int, settings) -> None:
    wanted = []
    if settings.print_area:
        wanted.append("_xlnm.Print_Area")
    if settings.repeat_rows or settings.repeat_columns:
        wanted.append("_xlnm.Print_Titles")
    if not wanted:
        return
    donor = scratch["xl/workbook.xml"].decode("utf-8")
    xml = package.require("xl/workbook.xml").decode("utf-8")

    def drop(match):
        attributes = _attributes(match[0])
        keep = not (attributes.get("name") in wanted and attributes.get("localSheetId") == str(sheet_index))
        return match[0] if keep else ""

    xml = element_pattern("definedName").sub(drop, xml)
    fragments = []
    for item in _donor_elements(donor, "definedName"):
        name = _attributes(item).get("name")
        if name in wanted:
            formula = _required_match(re.search(r">(.*)</definedName>", item, re.S), "definedName text")[1]
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
    donor_drawing_rels = safe_fromstring(scratch[_rels_path(drawing)])

    host_drawing = next(
        (_resolve_part_target(sheet_path, _target(node)) for node in sheet_rels
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
        host_xml = package.require(host_drawing).decode("utf-8")
    drawing_rels = package.rels(host_drawing)

    mapping = {}
    for node in donor_drawing_rels:
        kind = node.get("Type").rsplit("/", 1)[-1]
        source = _resolve_part_target(drawing, node.get("Target"))
        if kind == "chart":
            target = package.free_part("xl/charts", "chart", "xml")
            package.override(target, _CONTENT["chart"])
        else:
            extension = source.rsplit(".", 1)[-1]
            target = package.free_part("xl/media", "image", extension)
            donor_types = safe_fromstring(scratch["[Content_Types].xml"])
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
        root = _required_match(re.search(r"<((?:[\w.-]+:)?wsDr)\b([^>]*)/>", host_xml), "wsDr element")
        host_xml = host_xml[:root.start()] + f"<{root[1]}{root[2]}>" + "".join(anchors) + f"</{root[1]}>" + host_xml[root.end():]
    else:
        host_xml = host_xml[:end.start()] + "".join(anchors) + host_xml[end.start():]
    package.set(host_drawing, host_xml)
    ET.register_namespace("", PKG_REL)
    package.set(_rels_path(host_drawing), _bytes(drawing_rels))


def _existing_tables(package: _Package) -> list[tuple[str, dict]]:
    tables = []
    for name in sorted(package.names):
        if name in package.removed:
            continue
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
        _resolve_part_target(sheet_path, _target(node)) for node in sheet_rels
        if node.get("Type", "").endswith("/table")
    }
    next_id = max((int(attributes.get("id", 0)) for _, attributes in existing), default=0) + 1
    fragments = []
    rids = []
    for source in sources:
        xml = scratch[source].decode("utf-8")
        attributes = _attributes(_required_match(re.search(r"<table\b[^>]*>", xml), "table element")[0])
        if attributes["displayName"].casefold() in names:
            raise ValueError(f"a table named {attributes['displayName']!r} already exists in the workbook")
        for name, other in existing:
            if name in on_sheet and _overlaps(attributes["ref"], other.get("ref", "A1")):
                raise ValueError(
                    f"table {attributes['displayName']!r} ({attributes['ref']}) overlaps the "
                    f"sheet's table {other.get('displayName')!r} ({other.get('ref')})"
                )
        table_id = next_id
        xml = re.sub(r'(<table\b[^>]*?\bid=")\d+"', rf'\g<1>{table_id}"', xml, count=1)
        next_id += 1
        target = package.free_part("xl/tables", "table", "xml")
        package.set(target, xml)
        package.override(target, _CONTENT["table"])
        rid = package.add_rel(sheet_rels, sheet_path, "table", target)
        rids.append(rid)
        fragments.append(f'<tablePart r:id="{rid}"/>')
    host = sheet.find("tableParts")
    if host is None:
        if sheet.prefix:
            p = sheet.prefix
            body = (
                f'<{p}tableParts count="{len(fragments)}">'
                + "".join(f'<{p}tablePart r:id="{rid}"/>' for rid in rids)
                + f"</{p}tableParts>"
            )
            sheet.insert("tableParts", body)
        else:
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


def _table_signature(table) -> tuple:
    style = getattr(table, "tableStyleInfo", None)
    columns = getattr(table, "tableColumns", None)
    if columns is None:
        columns = getattr(table, "columns", None) or []
    return (
        getattr(table, "displayName", getattr(table, "display_name", None)),
        getattr(table, "ref", None),
        tuple(getattr(column, "name", column) for column in columns),
        getattr(table, "headerRowCount", getattr(table, "header_row", True)),
        getattr(table, "totalsRowCount", getattr(table, "totals_row", False)),
        getattr(style, "name", getattr(table, "style_name", None)),
        getattr(style, "showFirstColumn", getattr(table, "show_first_column", False)),
        getattr(style, "showLastColumn", getattr(table, "show_last_column", False)),
        getattr(style, "showRowStripes", getattr(table, "show_row_stripes", True)),
        getattr(style, "showColumnStripes", getattr(table, "show_column_stripes", False)),
    )


def object_signatures(sheet) -> dict[str, tuple]:
    return {
        "tables": tuple(_table_signature(table) for table in sheet.tables),
        "data_validations": tuple(repr(rule) for rule in sheet.data_validations),
        "conditional_formatting": tuple(
            repr(rule) for rule in sheet.conditional_formatting.all_rules()
        ),
    }


def _as_excel_table(table):
    from ..styles import ExcelTable

    if isinstance(table, ExcelTable):
        return table
    style = getattr(table, "tableStyleInfo", None)
    return ExcelTable(
        display_name=table.displayName,
        ref=table.ref,
        columns=getattr(table, "columns", None) or None,
        header_row=bool(getattr(table, "headerRowCount", 1)),
        totals_row=bool(getattr(table, "totalsRowCount", 0)),
        style_name=(getattr(style, "name", None) or "TableStyleMedium9"),
        show_first_column=bool(getattr(style, "showFirstColumn", False)),
        show_last_column=bool(getattr(style, "showLastColumn", False)),
        show_row_stripes=bool(getattr(style, "showRowStripes", True)),
        show_column_stripes=bool(getattr(style, "showColumnStripes", False)),
    )


_UNREAD_LABELS = {
    "data_validations": "data validations",
    "conditional_formatting": "conditional formatting rules",
}


def note_unread(sheet, name: str, count: int = 1) -> None:
    """Count loaded ``name`` objects (data validations, conditional formats)
    that the model could not read, so a save cannot drop them unnoticed."""
    if count:
        unread = sheet.__dict__.setdefault("_unread_objects", {})
        unread[name] = unread.get(name, 0) + count


def check_unread(sheet, names) -> None:
    """Refuse to rewrite the ``names`` lists of a sheet that holds objects
    of those kinds the model could not read: the rewrite would lose them."""
    unread = getattr(sheet, "_unread_objects", {})
    for name in names:
        if unread.get(name):
            label = _UNREAD_LABELS[name]
            raise UnsupportedFeatureError(
                f"sheet {sheet.title!r} holds {unread[name]} {label} that xlspy cannot read; "
                f"saving changed {label} would drop them"
            )


def check_object_rewrites(sheet) -> None:
    """Before an XLSX save: the checks of ``sync_objects``, which rewrites
    the sheet's tables, validations and conditional formats together."""
    saved = getattr(sheet, "_saved_object_signatures", None)
    current = object_signatures(sheet)
    if current == saved or saved is None or _only_appended(saved, current):
        return
    check_unread(sheet, _UNREAD_LABELS)


def _mark_synced_objects(sheet) -> None:
    for name in ("tables", "data_validations", "conditional_formatting"):
        sheet._saved_object_counts[name] = len(_object_values(sheet, name))
    sheet._saved_object_signatures = object_signatures(sheet)


def _only_appended(saved, current) -> bool:
    if saved is None:
        return False
    for name in ("tables", "data_validations", "conditional_formatting"):
        previous, now = saved[name], current[name]
        if len(now) < len(previous) or now[:len(previous)] != previous:
            return False
    return True


def sync_objects(updater, workbook, sheet) -> None:
    """Rewrite tables/validations/CF when any loaded object was edited or removed."""
    saved = getattr(sheet, "_saved_object_signatures", None)
    current = object_signatures(sheet)
    if current == saved or saved is None or _only_appended(saved, current):
        return

    names = [
        getattr(table, "displayName", getattr(table, "display_name", None))
        for table in sheet.tables
    ]
    if len(names) != len(set(names)):
        duplicate = next(name for name in names if names.count(name) > 1)
        raise ValueError(
            f"a table named {duplicate!r} already exists in the workbook"
        )

    objects = {
        "tables": [_as_excel_table(table) for table in sheet.tables],
        "data_validations": list(sheet.data_validations),
        "conditional_formatting": list(sheet.conditional_formatting.all_rules()),
        "charts": [],
        "images": [],
    }
    scratch = _scratch(sheet, objects, None)
    package = _Package(updater)
    handle = PackageHandle(updater)
    sheet_path = handle.sheets[sheet.title]
    host = _Sheet(package.require(sheet_path).decode("utf-8"))
    sheet_rels = package.rels(sheet_path)

    for node in list(sheet_rels):
        if node.get("Type", "").endswith("/table"):
            target = _resolve_part_target(sheet_path, _target(node))
            package.remove(target)
            sheet_rels.remove(node)
    host.remove_all("tableParts")
    host.remove_all("dataValidations")
    host.remove_all("conditionalFormatting")

    donor = scratch["xl/worksheets/sheet1.xml"].decode("utf-8")
    donor_rels = _sheet_rels(scratch)
    _add_tables(package, host, sheet_path, sheet_rels, scratch, donor_rels)
    _add_validations(host, donor)
    _add_conditional_formats(host, scratch, donor, workbook._styles)

    package.set(sheet_path, host.xml)
    if len(sheet_rels):
        ET.register_namespace("", PKG_REL)
        package.set(_rels_path(sheet_path), _bytes(sheet_rels))
    if objects["conditional_formatting"]:
        package.set("xl/styles.xml", workbook._styles.to_bytes())
    package.commit()
    _mark_synced_objects(sheet)


def add_objects(updater, workbook, sheet, sheet_order: list[str]) -> None:
    """Move the sheet's pending objects into the loaded package."""
    objects = pending_objects(sheet)
    settings = print_settings_for(sheet)
    cleared = _cleared_print_fields(getattr(sheet, "_saved_print_settings", None), settings)
    if not any(objects.values()) and settings is None and not cleared:
        return

    package = _Package(updater)
    handle = PackageHandle(updater)
    sheet_path = handle.sheets[sheet.title]
    host = _Sheet(package.require(sheet_path).decode("utf-8"))
    sheet_rels = package.rels(sheet_path)
    sheet_index = sheet_order.index(sheet.title)

    if any(objects.values()) or settings is not None:
        scratch = _scratch(sheet, objects, settings)
        donor = scratch["xl/worksheets/sheet1.xml"].decode("utf-8")
        donor_rels = _sheet_rels(scratch)
        # Tables first: their name and range checks are the steps that can fail,
        # and nothing outside this package (the style table) is changed yet.
        _add_tables(package, host, sheet_path, sheet_rels, scratch, donor_rels)
        _add_drawing(package, host, sheet_path, sheet_rels, scratch, donor_rels)
        _add_validations(host, donor)
        if settings is not None:
            _add_print_setup(host, donor, settings)
            _set_print_names(package, scratch, sheet_index, settings)
        _add_conditional_formats(host, scratch, donor, workbook._styles)
    if cleared:
        _clear_print_setup(package, host, sheet_index, cleared, settings)

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
