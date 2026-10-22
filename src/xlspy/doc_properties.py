"""Document properties (File > Info in Excel) for XlsxWriter and XlsbWriter.

Both formats store them in the same XML parts: ``docProps/core.xml`` (title,
author, dates), ``docProps/app.xml`` (company, manager) and, for user-defined
properties, ``docProps/custom.xml``. The parts are written only after
``set_properties`` is called, so plain writes stay byte-for-byte reproducible.
"""

from __future__ import annotations

import datetime as _datetime
import math
import re
from xml.sax.saxutils import escape

# Characters XML 1.0 cannot carry at all.
_NOT_XML = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]")

_CORE_TYPE = "application/vnd.openxmlformats-package.core-properties+xml"
_APP_TYPE = "application/vnd.openxmlformats-officedocument.extended-properties+xml"
_CUSTOM_TYPE = "application/vnd.openxmlformats-officedocument.custom-properties+xml"
_CORE_REL = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
_APP_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
_CUSTOM_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
# The format id Office uses for every user-defined property.
_CUSTOM_FMTID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"

# set_properties keyword -> element in core.xml.
_CORE_FIELDS = {
    "title": "dc:title",
    "subject": "dc:subject",
    "author": "dc:creator",
    "keywords": "cp:keywords",
    "comments": "dc:description",
    "category": "cp:category",
    "status": "cp:contentStatus",
}
_APP_FIELDS = {"manager": "Manager", "company": "Company"}
# Excel limits a custom property name to 255 characters.
_MAX_CUSTOM_NAME = 255


def _check_text(name: str, value) -> str:
    if not isinstance(value, str):
        raise TypeError(f"{name} must be a str")
    if _NOT_XML.search(value):
        raise ValueError(f"{name} contains a control character XML cannot store")
    return value


def _utc(value: _datetime.datetime, name: str) -> str:
    if not isinstance(value, _datetime.datetime):
        raise TypeError(f"{name} must be a datetime.datetime")
    if value.tzinfo is not None:
        value = value.astimezone(_datetime.timezone.utc)
    # A naive datetime is taken as UTC, as XlsxWriter does.
    return value.strftime("%Y-%m-%dT%H:%M:%SZ")


def _custom_value(name: str, value) -> str:
    if isinstance(value, bool):
        return f"<vt:bool>{'true' if value else 'false'}</vt:bool>"
    if isinstance(value, int):
        if -2**31 <= value < 2**31:
            return f"<vt:i4>{value}</vt:i4>"
        value = float(value)
    if isinstance(value, float):
        if not math.isfinite(value):
            raise ValueError(f"custom property {name!r} must be a finite number")
        return f"<vt:r8>{value!r}</vt:r8>"
    if isinstance(value, _datetime.datetime):
        return f"<vt:filetime>{_utc(value, name)}</vt:filetime>"
    if isinstance(value, _datetime.date):
        return f"<vt:filetime>{value.isoformat()}T00:00:00Z</vt:filetime>"
    if isinstance(value, str):
        return f"<vt:lpwstr>{escape(_check_text(f'custom property {name!r}', value))}</vt:lpwstr>"
    raise TypeError(
        f"custom property {name!r} must be str, int, float, bool, date or datetime"
    )


class DocumentPropertiesMixin:
    """``set_properties`` and the docProps parts for a workbook writer."""

    _doc_props: dict | None = None

    def set_properties(
        self,
        *,
        title: str | None = None,
        subject: str | None = None,
        author: str | None = None,
        manager: str | None = None,
        company: str | None = None,
        category: str | None = None,
        keywords: str | None = None,
        comments: str | None = None,
        status: str | None = None,
        created: _datetime.datetime | None = None,
        custom: dict | None = None,
    ) -> None:
        """Set the document properties shown in Excel under File > Info.

        Text fields are optional strings. ``created`` defaults to the time of
        the call; a naive datetime is taken as UTC. ``custom`` maps property
        names to str, int, float, bool, date or datetime values and appears
        under File > Info > Properties > Advanced Properties > Custom.
        Calling it again replaces every earlier value.
        """
        self._require_writable()
        values = {
            "title": title, "subject": subject, "author": author,
            "manager": manager, "company": company, "category": category,
            "keywords": keywords, "comments": comments, "status": status,
        }
        props = {key: _check_text(key, value) for key, value in values.items() if value is not None}
        if created is None:
            created = _datetime.datetime.now(_datetime.timezone.utc)
        props["created"] = _utc(created, "created")
        custom_xml = []
        if custom is not None:
            if not isinstance(custom, dict):
                raise TypeError("custom must be a dict of property names to values")
            seen = set()
            for name, value in custom.items():
                _check_text("custom property name", name)
                if not name.strip() or len(name) > _MAX_CUSTOM_NAME:
                    raise ValueError(f"custom property names must be 1-{_MAX_CUSTOM_NAME} characters")
                if name.casefold() in seen:
                    raise ValueError(f"duplicate custom property {name!r} (names ignore case)")
                seen.add(name.casefold())
                custom_xml.append((name, _custom_value(name, value)))
        props["custom"] = custom_xml
        self._doc_props = props

    # Package parts, called by the writers while saving.

    def _doc_props_content_types(self) -> str:
        if self._doc_props is None:
            return ""
        parts = (
            f'<Override PartName="/docProps/core.xml" ContentType="{_CORE_TYPE}"/>'
            f'<Override PartName="/docProps/app.xml" ContentType="{_APP_TYPE}"/>'
        )
        if self._doc_props["custom"]:
            parts += f'<Override PartName="/docProps/custom.xml" ContentType="{_CUSTOM_TYPE}"/>'
        return parts

    def _doc_props_root_rels(self) -> str:
        """Root relationships after rId1 (the workbook)."""
        if self._doc_props is None:
            return ""
        rels = (
            f'<Relationship Id="rId2" Type="{_CORE_REL}" Target="docProps/core.xml"/>'
            f'<Relationship Id="rId3" Type="{_APP_REL}" Target="docProps/app.xml"/>'
        )
        if self._doc_props["custom"]:
            rels += f'<Relationship Id="rId4" Type="{_CUSTOM_REL}" Target="docProps/custom.xml"/>'
        return rels

    def _write_doc_props(self) -> None:
        props = self._doc_props
        if props is None:
            return
        core = "".join(
            f"<{tag}>{escape(props[key])}</{tag}>"
            for key, tag in _CORE_FIELDS.items() if key in props
        )
        if "author" in props:
            core += f"<cp:lastModifiedBy>{escape(props['author'])}</cp:lastModifiedBy>"
        stamp = props["created"]
        self._zf.writestr("docProps/core.xml", (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
            'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            f'{core}'
            f'<dcterms:created xsi:type="dcterms:W3CDTF">{stamp}</dcterms:created>'
            f'<dcterms:modified xsi:type="dcterms:W3CDTF">{stamp}</dcterms:modified>'
            '</cp:coreProperties>'
        ))
        app = "".join(
            f"<{tag}>{escape(props[key])}</{tag}>"
            for key, tag in _APP_FIELDS.items() if key in props
        )
        self._zf.writestr("docProps/app.xml", (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
            'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
            f'<Application>Microsoft Excel</Application>{app}</Properties>'
        ))
        if props["custom"]:
            items = "".join(
                f'<property fmtid="{_CUSTOM_FMTID}" pid="{pid}" name="{escape(name, {chr(34): "&quot;"})}">'
                f'{value}</property>'
                for pid, (name, value) in enumerate(props["custom"], start=2)
            )
            self._zf.writestr("docProps/custom.xml", (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" '
                'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
                f'{items}</Properties>'
            ))
