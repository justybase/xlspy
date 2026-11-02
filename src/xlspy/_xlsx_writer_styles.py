"""XlsxWriter style methods: fonts, fills, borders, styles.xml, conditional-format XML."""
from __future__ import annotations

import xml.sax.saxutils as saxutils
from .styles import (
    CF_BLANK_RULES,
    CF_TEXT_RULES,
    CF_TIME_PERIOD_FORMULAS,
    EXTRA_CF_RULES,
    AboveAverageRule,
    CellIsRule,
    DuplicateRule,
    TextRule,
    TimePeriodRule,
    TopBottomRule,
    ColorScaleRule,
    DifferentialStyle,
    FormatObject,
    FormulaRule,
    IconSetRule,
    Style,
    _alignment_key_from,
    _border_key,
    _fill_key_from,
    _font_is_empty,
    _font_key_from,
    _protection_key_from,
    parse_cell_range,
    split_style,
)
from ._xlsx_writer_xml import _alignment_xml, _format_number, _protection_xml
from ._xlsx_writer_host import _XlsxWriterHost
from .styles import _coerce_color


class _XlsxStylesMixin(_XlsxWriterHost):
    """Methods of the writer class, mixed into it; see the facade module."""

    def _intern_font(self, font_key) -> int:
        if _font_is_empty(font_key):
            return 0
        existing = self._font_registry.get(font_key)
        if existing is not None:
            return existing
        font_id = 2 + len(self._font_registry)
        self._font_registry[font_key] = font_id
        return font_id

    def _intern_fill(self, fill_key) -> int:
        if fill_key is None:
            return 0
        existing = self._fill_registry.get(fill_key)
        if existing is not None:
            return existing
        fill_id = 2 + len(self._fill_registry)
        self._fill_registry[fill_key] = fill_id
        return fill_id

    def _intern_border(self, border_key) -> int:
        if border_key is None:
            return 0
        existing = self._border_registry.get(border_key)
        if existing is not None:
            return existing
        border_id = 1 + len(self._border_registry)
        self._border_registry[border_key] = border_id
        return border_id

    def _register_dxf_style(self, dxf: DifferentialStyle) -> int:
        font_key = None if dxf.font is None else _font_key_from(dxf.font)
        if font_key is None or _font_is_empty(font_key):
            font_id = None
        else:
            if font_key in self._dxf_font_registry:
                font_id = self._dxf_font_registry[font_key]
            else:
                font_id = len(self._dxf_font_registry)
                self._dxf_font_registry[font_key] = font_id
        if dxf.fill is None:
            fill_id = None
        else:
            fill_key = _fill_key_from(dxf.fill)
            if fill_key is None:
                fill_id = None
            elif fill_key in self._dxf_fill_registry:
                fill_id = self._dxf_fill_registry[fill_key]
            else:
                fill_id = len(self._dxf_fill_registry)
                self._dxf_fill_registry[fill_key] = fill_id
        if dxf.border is None:
            border_id = None
        else:
            border_key = _border_key(dxf.border)
            if border_key is None:
                border_id = None
            elif border_key in self._dxf_border_registry:
                border_id = self._dxf_border_registry[border_key]
            else:
                border_id = len(self._dxf_border_registry)
                self._dxf_border_registry[border_key] = border_id
        align_key = (
            _alignment_key_from(dxf.alignment)
            if dxf.alignment is not None else None
        )
        if align_key is not None and not any(
            value is not None for value in align_key
        ):
            align_key = None
        prot_key = (
            _protection_key_from(dxf.protection)
            if dxf.protection is not None else None
        )
        if prot_key is not None and not any(
            value is not None for value in prot_key
        ):
            prot_key = None
        if (
            dxf.number_format is not None
            and dxf.number_format not in self._format_registry
        ):
            self._format_registry[dxf.number_format] = self._next_numfmt_id
            self._next_numfmt_id += 1
        dxf_entry = (
            font_id, fill_id, border_id, dxf.number_format,
            align_key, prot_key,
        )
        existing = self._dxf_index.get(dxf_entry)
        if existing is not None:
            return existing
        index = len(self._dxf_list)
        self._dxf_list.append(dxf_entry)
        self._dxf_index[dxf_entry] = index
        return index

    def _register_style(self, style: Style) -> int:
        if style in self._style_xf_map:
            return self._style_xf_map[style]
        font_key, fill_key, border_key, numfmt, _align, _prot = split_style(style)
        font_id = self._intern_font(font_key)
        fill_id = self._intern_fill(fill_key)
        border_id = self._intern_border(border_key)
        numfmt_id = 0
        if numfmt is not None:
            if numfmt not in self._format_registry:
                self._format_registry[numfmt] = self._next_numfmt_id
                self._next_numfmt_id += 1
            numfmt_id = self._format_registry[numfmt]
        xf_index = self._next_xf_index
        self._next_xf_index += 1
        self._style_xf_map[style] = xf_index
        self._xf_styles[xf_index] = style
        self._style_components[xf_index] = (
            font_id, fill_id, border_id, numfmt_id, _align, _prot,
        )
        return xf_index

    @staticmethod
    def _color_attrs(color_key) -> str:
        kind = color_key[0]
        if kind == "rgb":
            return f'rgb="{color_key[1]}"'
        if kind == "indexed":
            return f'indexed="{color_key[1]}"'
        if kind == "auto":
            return 'auto="1"'
        _, theme, tint = color_key
        if tint is None:
            return f'theme="{theme}"'
        return f'theme="{theme}" tint="{tint}"'

    def _color_xml(self, color_key) -> str:
        if color_key is None:
            return '<color theme="1"/>'
        return f'<color {self._color_attrs(color_key)}/>'

    def _font_xml(self, font_key, *, for_dxf: bool = False) -> str:
        name = font_key.name
        size = font_key.size
        parts = []
        if size is not None:
            parts.append(f'<sz val="{size}"/>')
        if for_dxf:
            # A dxf font applies only the properties it names; emitting the
            # Calibri/theme-1 defaults would override the cell's own font.
            if font_key.color is not None:
                parts.append(self._color_xml(font_key.color))
            if name is not None:
                label = saxutils.escape(name, {'"': '&quot;'})
                parts.append(f'<name val="{label}"/>')
            if font_key.family is not None:
                parts.append(f'<family val="{font_key.family}"/>')
            if font_key.scheme is not None:
                parts.append(f'<scheme val="{font_key.scheme}"/>')
        else:
            parts.append(self._color_xml(font_key.color))
            label = saxutils.escape(name, {'"': '&quot;'}) if name is not None else "Calibri"
            parts.append(f'<name val="{label}"/>')
            if font_key.family is not None:
                parts.append(f'<family val="{font_key.family}"/>')
            else:
                parts.append('<family val="2"/>')
            if font_key.scheme is not None:
                parts.append(f'<scheme val="{font_key.scheme}"/>')
            elif name is None:
                # Only the built-in Calibri default carries scheme="minor".
                # A scheme next to an explicit font name would make Excel
                # substitute the theme font for that name.
                parts.append('<scheme val="minor"/>')
        if font_key.charset is not None:
            parts.append(f'<charset val="{font_key.charset}"/>')
        if font_key.bold:
            parts.append('<b/>')
        if font_key.italic:
            parts.append('<i/>')
        if font_key.underline is not None:
            if font_key.underline == "single":
                parts.append('<u/>')
            else:
                parts.append(f'<u val="{font_key.underline}"/>')
        if font_key.strike:
            parts.append('<strike/>')
        if font_key.vert_align is not None:
            parts.append(f'<vertAlign val="{font_key.vert_align}"/>')
        if font_key.outline:
            parts.append('<outline/>')
        if font_key.shadow:
            parts.append('<shadow/>')
        if font_key.condense:
            parts.append('<condense/>')
        if font_key.extend:
            parts.append('<extend/>')
        return f"<font>{''.join(parts)}</font>"

    def _fill_xml(self, fill_key, *, for_dxf: bool = False) -> str:
        if fill_key[0] == "gradient":
            (
                _, fill_type, degree, left, right, top, bottom, stops,
            ) = fill_key
            attrs = [f'type="{fill_type}"']
            if degree is not None:
                attrs.append(f'degree="{_format_number(degree)}"')
            for name, value in (
                ("left", left), ("right", right), ("top", top),
                ("bottom", bottom),
            ):
                if value is not None:
                    attrs.append(f'{name}="{_format_number(value)}"')
            stop_xml = "".join(
                f'<stop position="{_format_number(position)}">'
                f'<color {self._color_attrs(color_key)}/></stop>'
                for color_key, position in stops
            )
            return (
                f'<fill><gradientFill {" ".join(attrs)}>'
                f"{stop_xml}</gradientFill></fill>"
            )
        _, pattern, fg_key, bg_key = fill_key
        pattern_type = pattern if pattern is not None else "none"
        if pattern_type == "solid":
            fg_xml = f'<fgColor {self._color_attrs(fg_key)}/>' if fg_key is not None else ""
            if bg_key is not None:
                bg_xml = f'<bgColor {self._color_attrs(bg_key)}/>'
            elif for_dxf:
                # A conditional format paints a solid fill with its background
                # color, so a fill given only start_color would not show.
                bg_xml = f'<bgColor {self._color_attrs(fg_key)}/>' if fg_key is not None else ""
            else:
                bg_xml = '<bgColor indexed="64"/>'
        else:
            fg_xml = f'<fgColor {self._color_attrs(fg_key)}/>' if fg_key is not None else ""
            bg_xml = f'<bgColor {self._color_attrs(bg_key)}/>' if bg_key is not None else ""
        return f'<fill><patternFill patternType="{pattern_type}">{fg_xml}{bg_xml}</patternFill></fill>'

    @staticmethod
    def _border_side_xml(tag, side_key, color_attrs) -> str:
        style, color_key = side_key
        if style is None and color_key is None:
            return f"<{tag}/>"
        inner = f"<color {color_attrs(color_key)}/>" if color_key is not None else ""
        if style is None:
            return f"<{tag}>{inner}</{tag}>"
        return f'<{tag} style="{style}">{inner}</{tag}>'

    def _border_xml(self, border_key) -> str:
        attrs = ""
        if border_key.diagonal_up:
            attrs += ' diagonalUp="1"'
        if border_key.diagonal_down:
            attrs += ' diagonalDown="1"'
        if border_key.outline is False:
            attrs += ' outline="0"'
        # Excel reads <start>/<end> as the left/right edges (in a
        # left-to-right sheet), so an empty <start/> would hide <left>:
        # write them only when set, like <vertical>/<horizontal>.
        parts = [
            self._border_side_xml(tag, getattr(border_key, tag), self._color_attrs)
            for tag in ("start", "end")
            if getattr(border_key, tag) != (None, None)
        ]
        parts += [
            self._border_side_xml("left", border_key.left, self._color_attrs),
            self._border_side_xml("right", border_key.right, self._color_attrs),
            self._border_side_xml("top", border_key.top, self._color_attrs),
            self._border_side_xml("bottom", border_key.bottom, self._color_attrs),
            self._border_side_xml("diagonal", border_key.diagonal, self._color_attrs),
        ]
        if border_key.vertical != (None, None):
            parts.append(self._border_side_xml("vertical", border_key.vertical, self._color_attrs))
        if border_key.horizontal != (None, None):
            parts.append(self._border_side_xml("horizontal", border_key.horizontal, self._color_attrs))
        return f"<border{attrs}>{''.join(parts)}</border>"

    def _create_styles_xml(self) -> str:
        numfmts = ""
        xf_entries = []

        numfmt_count = len(self._format_registry)
        if numfmt_count > 0:
            numfmts_parts = []
            for fmt_string, numfmt_id in self._format_registry.items():
                escaped_fmt = saxutils.escape(fmt_string, {'"': '&quot;'})
                numfmts_parts.append(f'<numFmt numFmtId="{numfmt_id}" formatCode="{escaped_fmt}"/>')
            numfmts = f'<numFmts count="{numfmt_count}">{"".join(numfmts_parts)}</numFmts>'

        xf_entries.append('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>')
        xf_entries.append('<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>')
        xf_entries.append('<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>')
        xf_entries.append('<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>')

        # Format XFs and style XFs share one index space, so they must be
        # emitted in index order: the position in cellXfs is the value a cell
        # references through its s= attribute.
        custom_xfs = {
            xf_index: f'<xf numFmtId="{self._format_registry[fmt_string]}" fontId="0" '
                      f'fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
            for fmt_string, xf_index in self._format_xf_map.items()
        }
        for xf_index, components in self._style_components.items():
            font_id, fill_id, border_id, numfmt_id, align, prot = components
            attrs = [f'numFmtId="{numfmt_id}"', f'fontId="{font_id}"',
                     f'fillId="{fill_id}"', f'borderId="{border_id}"', 'xfId="0"']
            if numfmt_id != 0:
                attrs.append('applyNumberFormat="1"')
            if font_id != 0:
                attrs.append('applyFont="1"')
            if fill_id != 0:
                attrs.append('applyFill="1"')
            if border_id != 0:
                attrs.append('applyBorder="1"')
            align_xml = _alignment_xml(align)
            if align_xml:
                attrs.append('applyAlignment="1"')
            prot_xml = _protection_xml(prot)
            if prot_xml:
                attrs.append('applyProtection="1"')
            custom_xfs[xf_index] = f'<xf {" ".join(attrs)}>{align_xml}{prot_xml}</xf>'

        for xf_index in sorted(custom_xfs):
            xf_entries.append(custom_xfs[xf_index])

        total_xf = len(xf_entries)

        fonts = [
            '<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>',
            '<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/><b/></font>',
        ]
        for key, font_id in sorted(self._font_registry.items(), key=lambda item: item[1]):
            if font_id < 2:
                continue
            fonts.append(self._font_xml(key))
        fills = [
            '<fill><patternFill patternType="none"/></fill>',
            '<fill><patternFill patternType="gray125"/></fill>',
        ]
        fills.extend(
            self._fill_xml(key)
            for key, _fill_id in sorted(self._fill_registry.items(), key=lambda item: item[1])
        )
        borders = [
            '<border><left/><right/><top/><bottom/><diagonal/></border>',
        ]
        borders.extend(
            self._border_xml(key)
            for key, _border_id in sorted(self._border_registry.items(), key=lambda item: item[1])
        )

        dxf_fonts: list = [None] * len(self._dxf_font_registry)
        for key, dxf_font_id in self._dxf_font_registry.items():
            dxf_fonts[dxf_font_id] = key
        dxf_fills: list = [None] * len(self._dxf_fill_registry)
        for key, dxf_fill_id in self._dxf_fill_registry.items():
            dxf_fills[dxf_fill_id] = key
        dxf_borders: list = [None] * len(self._dxf_border_registry)
        for key, dxf_border_id in self._dxf_border_registry.items():
            dxf_borders[dxf_border_id] = key
        dxf_parts = []
        for font_id, fill_id, border_id, numfmt, align, prot in self._dxf_list:
            inner = ""
            if font_id is not None:
                inner += self._font_xml(dxf_fonts[font_id], for_dxf=True)
            if numfmt is not None:
                escaped_fmt = saxutils.escape(numfmt, {'"': '&quot;'})
                numfmt_id = self._format_registry[numfmt]
                inner += (
                    f'<numFmt numFmtId="{numfmt_id}" '
                    f'formatCode="{escaped_fmt}"/>'
                )
            if fill_id is not None:
                inner += self._fill_xml(dxf_fills[fill_id], for_dxf=True)
            if align is not None:
                inner += _alignment_xml(align)
            if border_id is not None:
                inner += self._border_xml(dxf_borders[border_id])
            if prot is not None:
                inner += _protection_xml(prot)
            dxf_parts.append(f"<dxf>{inner}</dxf>")
        dxfs = f'<dxfs count="{len(dxf_parts)}">{"".join(dxf_parts)}</dxfs>' if dxf_parts else '<dxfs count="0"/>'

        return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
{numfmts}
<fonts count="{len(fonts)}">{"".join(fonts)}</fonts>
<fills count="{len(fills)}">{"".join(fills)}</fills>
<borders count="{len(borders)}">
{"".join(borders)}
</borders>
<cellStyleXfs count="1">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
</cellStyleXfs>
<cellXfs count="{total_xf}">
{"".join(xf_entries)}
</cellXfs>
<cellStyles count="1">
<cellStyle name="Normal" xfId="0" builtinId="0"/>
</cellStyles>
{dxfs}
<tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>'''

    @staticmethod
    def _format_object_xml(cfvo) -> str:
        attrs = f'type="{cfvo.type}"'
        if cfvo.value is not None:
            value = saxutils.escape(str(cfvo.value), {'"': "&quot;"})
            attrs += f' val="{value}"'
        return f"<cfvo {attrs}/>"

    def _extra_cf_rule_parts(self, rule) -> tuple[str, str | None]:
        """``(cfRule attributes, formula or None)`` for the rules beyond the original five."""
        first_col, first_row, _, _ = parse_cell_range(rule.sqref.split()[0])
        corner = f"{self._letters[first_col]}{first_row + 1}"
        if isinstance(rule, TopBottomRule):
            attrs = f'type="top10" rank="{rule.rank}"'
            if rule.percent:
                attrs += ' percent="1"'
            if rule.bottom:
                attrs += ' bottom="1"'
            return attrs, None
        if isinstance(rule, AboveAverageRule):
            attrs = 'type="aboveAverage"'
            if rule.below:
                attrs += ' aboveAverage="0"'
            if rule.equal_average:
                attrs += ' equalAverage="1"'
            if rule.std_dev is not None:
                attrs += f' stdDev="{rule.std_dev}"'
            return attrs, None
        if isinstance(rule, DuplicateRule):
            return f'type="{"uniqueValues" if rule.unique else "duplicateValues"}"', None
        if isinstance(rule, TextRule):
            kind, operator, template = CF_TEXT_RULES[rule.operator]
            quoted = '"' + rule.text.replace('"', '""') + '"'
            attrs = (
                f'type="{kind}" operator="{operator}" '
                f'text="{saxutils.escape(rule.text, {chr(34): "&quot;"})}"'
            )
            return attrs, template.format(t=quoted, c=corner)
        if isinstance(rule, TimePeriodRule):
            return (
                f'type="timePeriod" timePeriod="{rule.period}"',
                CF_TIME_PERIOD_FORMULAS[rule.period].format(c=corner),
            )
        kind, template = CF_BLANK_RULES[rule.kind]
        return f'type="{kind}"', template.format(c=corner)

    def _cf_rule_xml(self, rule, table_ranges: set[tuple]) -> str:
        priority = self._claim_cf_priority(rule)
        stop = ""
        if getattr(rule, "stop_if_true", None):
            stop = ' stopIfTrue="1"'
        dxf_part = ""
        dxf_index = None
        dxf = getattr(rule, "dxf", None)
        if dxf is not None:
            dxf_index = self._register_dxf_style(dxf)
            dxf_part = f' dxfId="{dxf_index}"'
        if isinstance(rule, CellIsRule):
            formulas = "".join(
                f"<formula>{saxutils.escape(text)}</formula>"
                for text in rule.formula
            )
            return (
                f'<cfRule type="cellIs" operator="{rule.operator}" '
                f'priority="{priority}"{dxf_part}{stop}>{formulas}</cfRule>'
            )
        if isinstance(rule, FormulaRule):
            formulas = "".join(
                f"<formula>{saxutils.escape(text)}</formula>"
                for text in rule.formula
            )
            return (
                f'<cfRule type="expression" priority="{priority}"'
                f'{dxf_part}{stop}>{formulas}</cfRule>'
            )
        if isinstance(rule, EXTRA_CF_RULES):
            attrs, formula = self._extra_cf_rule_parts(rule)
            formula_xml = f"<formula>{saxutils.escape(formula)}</formula>" if formula else ""
            return (
                f'<cfRule {attrs} priority="{priority}"{dxf_part}{stop}'
                + (f">{formula_xml}</cfRule>" if formula_xml else "/>")
            )
        if isinstance(rule, ColorScaleRule):
            cfvos: tuple[FormatObject | None, ...]
            if rule.midpoint is None:
                cfvos = (rule.minimum, rule.maximum)
            else:
                cfvos = (rule.minimum, rule.midpoint, rule.maximum)
            inner = "".join(
                self._format_object_xml(cfvo) if cfvo is not None
                else '<cfvo type="min"/>' if i == 0 else '<cfvo type="max"/>'
                for i, cfvo in enumerate(cfvos)
            )
            colors = "".join(
                f'<color rgb="{self._dxf_color_argb(color)}"/>'
                for color in rule.colors
            )
            return (
                f'<cfRule type="colorScale" priority="{priority}"{stop}>'
                f"<colorScale>{inner}{colors}</colorScale></cfRule>"
            )
        if isinstance(rule, IconSetRule):
            icon_attrs = [f'iconSet="{rule.icon_style}"']
            if rule.show_value is not None:
                icon_attrs.append(f'showValue="{1 if rule.show_value else 0}"')
            if rule.percent is not None:
                icon_attrs.append(f'percent="{1 if rule.percent else 0}"')
            if rule.reverse is not None:
                icon_attrs.append(f'reverse="{1 if rule.reverse else 0}"')
            cfvo_xml = "".join(
                self._format_object_xml(FormatObject(rule.cfvo_type, value))
                for value in rule.values
            )
            return (
                f'<cfRule type="iconSet" priority="{priority}"{stop}>'
                f'<iconSet {" ".join(icon_attrs)}>{cfvo_xml}</iconSet></cfRule>'
            )
        # DataBarRule
        min_xml = (
            self._format_object_xml(rule.minimum)
            if rule.minimum is not None else '<cfvo type="min"/>'
        )
        max_xml = (
            self._format_object_xml(rule.maximum)
            if rule.maximum is not None else '<cfvo type="max"/>'
        )
        show = "" if rule.show_value is None else (
            ' showValue="1"' if rule.show_value else ' showValue="0"'
        )
        min_len = "" if rule.min_length is None else f' minLength="{rule.min_length}"'
        max_len = "" if rule.max_length is None else f' maxLength="{rule.max_length}"'
        color = (
            f'<color rgb="{self._dxf_color_argb(rule.color)}"/>'
        )
        return (
            f'<cfRule type="dataBar" priority="{priority}"{stop}>'
            f"<dataBar{min_len}{max_len}{show}>{min_xml}{max_xml}{color}</dataBar>"
            "</cfRule>"
        )

    @staticmethod
    def _dxf_color_argb(color) -> str:
        resolved = _coerce_color(color, "color")
        if resolved is None or resolved.rgb is None:
            raise ValueError("CF colorScale/dataBar colors must be RGB colors")
        return resolved.rgb
