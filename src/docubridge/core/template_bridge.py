from __future__ import annotations

from copy import deepcopy
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from docx import Document
from docx.oxml.ns import qn


def _set_if_present(target: dict[str, Any], key: str, value: Any) -> None:
    if value is not None:
        target[key] = value


def _read_font_properties(font, r_fonts) -> dict[str, Any]:
    properties: dict[str, Any] = {}
    font_name = getattr(font, "name", None)
    if font_name:
        properties["font_name"] = font_name
        properties["font_ascii"] = font_name
        properties["font_hansi"] = font_name
    font_size = getattr(font, "size", None)
    if font_size is not None:
        properties["font_size"] = font_size.pt
    if getattr(font, "bold", None) is not None:
        properties["bold"] = bool(font.bold)
    if getattr(font, "italic", None) is not None:
        properties["italic"] = bool(font.italic)
    if r_fonts is None:
        return properties
    _set_if_present(properties, "font_ascii", r_fonts.get(qn("w:ascii")))
    _set_if_present(properties, "font_hansi", r_fonts.get(qn("w:hAnsi")))
    _set_if_present(properties, "font_east_asia", r_fonts.get(qn("w:eastAsia")))
    _set_if_present(properties, "font_cs", r_fonts.get(qn("w:cs")))
    return properties


def _read_paragraph_properties(paragraph_format) -> dict[str, Any]:
    properties: dict[str, Any] = {}
    first_line_indent = getattr(paragraph_format, "first_line_indent", None)
    if first_line_indent is not None:
        properties["first_line_indent_pt"] = first_line_indent.pt
    left_indent = getattr(paragraph_format, "left_indent", None)
    if left_indent is not None:
        properties["left_indent_pt"] = left_indent.pt
    right_indent = getattr(paragraph_format, "right_indent", None)
    if right_indent is not None:
        properties["right_indent_pt"] = right_indent.pt
    space_before = getattr(paragraph_format, "space_before", None)
    if space_before is not None:
        properties["space_before_pt"] = space_before.pt
    space_after = getattr(paragraph_format, "space_after", None)
    if space_after is not None:
        properties["space_after_pt"] = space_after.pt
    alignment = getattr(paragraph_format, "alignment", None)
    if alignment is not None:
        properties["alignment"] = int(alignment)
    line_spacing = getattr(paragraph_format, "line_spacing", None)
    if line_spacing is not None:
        if hasattr(line_spacing, "pt"):
            properties["line_spacing_pt"] = line_spacing.pt
        else:
            try:
                properties["line_spacing"] = float(line_spacing)
            except (TypeError, ValueError):
                pass
    return properties


def _read_style_properties(style) -> dict[str, Any]:
    properties: dict[str, Any] = {}
    properties.update(
        _read_font_properties(
            getattr(style, "font", None),
            getattr(getattr(style.element, "rPr", None), "rFonts", None),
        )
    )
    properties.update(_read_paragraph_properties(getattr(style, "paragraph_format", None)))
    return properties


def _style_numbering_reference(style) -> tuple[int, int] | None:
    current = style
    while current is not None:
        properties = getattr(current.element, "pPr", None)
        num_pr = getattr(properties, "numPr", None)
        if num_pr is not None and getattr(num_pr, "numId", None) is not None:
            ilvl = getattr(num_pr, "ilvl", None)
            return int(num_pr.numId.val), int(ilvl.val if ilvl is not None else 0)
        current = getattr(current, "base_style", None)
    return None


def _parse_numbering_definitions(document: Document) -> dict[int, dict[str, Any]]:
    try:
        numbering = document.part.numbering_part.element
    except KeyError:
        return {}

    abstract_numbering: dict[int, dict[int, dict[str, Any]]] = {}
    for abstract_node in numbering.xpath("./*[local-name()='abstractNum']"):
        abstract_id = int(abstract_node.get(qn("w:abstractNumId")))
        levels: dict[int, dict[str, Any]] = {}
        for level_node in abstract_node.xpath("./*[local-name()='lvl']"):
            level = int(level_node.get(qn("w:ilvl")))
            start_nodes = level_node.xpath("./*[local-name()='start']/@*[local-name()='val']")
            fmt_nodes = level_node.xpath("./*[local-name()='numFmt']/@*[local-name()='val']")
            text_nodes = level_node.xpath("./*[local-name()='lvlText']/@*[local-name()='val']")
            levels[level] = {
                "start": int(start_nodes[0]) if start_nodes else 1,
                "num_fmt": fmt_nodes[0] if fmt_nodes else "decimal",
                "lvl_text": text_nodes[0] if text_nodes else f"%{level + 1}.",
            }
        abstract_numbering[abstract_id] = levels

    numbering_instances: dict[int, dict[str, Any]] = {}
    for num_node in numbering.xpath("./*[local-name()='num']"):
        num_id = int(num_node.get(qn("w:numId")))
        abstract_refs = num_node.xpath("./*[local-name()='abstractNumId']/@*[local-name()='val']")
        if not abstract_refs:
            continue
        numbering_instances[num_id] = {
            "abstract_id": int(abstract_refs[0]),
            "levels": deepcopy(abstract_numbering.get(int(abstract_refs[0]), {})),
        }
    return numbering_instances


def _extract_document_defaults(document: Document) -> dict[str, Any]:
    defaults: dict[str, Any] = {}
    styles_element = document.styles.element
    run_defaults = styles_element.xpath(
        "./*[local-name()='docDefaults']/*[local-name()='rPrDefault']/*[local-name()='rPr']"
    )
    if run_defaults:
        r_pr = run_defaults[0]
        r_fonts = next((child for child in r_pr if child.tag == qn("w:rFonts")), None)
        defaults.update(_read_font_properties(type("FontProxy", (), {"name": None, "size": None, "bold": None, "italic": None})(), r_fonts))
        size_nodes = r_pr.xpath("./*[local-name()='sz']/@*[local-name()='val']")
        if size_nodes:
            defaults["font_size"] = int(size_nodes[0]) / 2.0

    paragraph_defaults = styles_element.xpath(
        "./*[local-name()='docDefaults']/*[local-name()='pPrDefault']/*[local-name()='pPr']"
    )
    if paragraph_defaults:
        p_pr = paragraph_defaults[0]
        ind_nodes = p_pr.xpath("./*[local-name()='ind']")
        if ind_nodes:
            ind = ind_nodes[0]
            first_line = ind.get(qn("w:firstLine"))
            left = ind.get(qn("w:left"))
            right = ind.get(qn("w:right"))
            if first_line is not None:
                defaults["first_line_indent_pt"] = int(first_line) / 20.0
            if left is not None:
                defaults["left_indent_pt"] = int(left) / 20.0
            if right is not None:
                defaults["right_indent_pt"] = int(right) / 20.0
        spacing_nodes = p_pr.xpath("./*[local-name()='spacing']")
        if spacing_nodes:
            spacing = spacing_nodes[0]
            before = spacing.get(qn("w:before"))
            after = spacing.get(qn("w:after"))
            if before is not None:
                defaults["space_before_pt"] = int(before) / 20.0
            if after is not None:
                defaults["space_after_pt"] = int(after) / 20.0
            line = spacing.get(qn("w:line"))
            line_rule = spacing.get(qn("w:lineRule"))
            if line is not None:
                try:
                    line_value = int(line)
                    if line_rule == "auto":
                        defaults["line_spacing"] = line_value / 240.0
                    elif line_rule in ("exact", "atLeast"):
                        defaults["line_spacing_pt"] = line_value / 20.0
                    else:
                        defaults["line_spacing"] = line_value / 240.0
                except (TypeError, ValueError):
                    pass
    return defaults


@dataclass(slots=True)
class TemplateView:
    available_styles: dict[str, dict[str, Any]] = field(default_factory=dict)
    available_numberings: dict[int, dict[str, Any]] = field(default_factory=dict)
    style_numbering_map: dict[str, tuple[int, int]] = field(default_factory=dict)
    document_defaults: dict[str, Any] = field(default_factory=dict)
    has_template: bool = False

    def get_style(self, name: str) -> dict[str, Any]:
        style = self.available_styles.get(name, {})
        return deepcopy(style)


def load_template_view(path: Path | None) -> TemplateView:
    if path is None:
        return TemplateView()
    if not path.exists():
        raise FileNotFoundError(2, "Template file not found", str(path))

    document = Document(str(path))
    available_styles: dict[str, dict[str, Any]] = {}
    style_numbering_map: dict[str, tuple[int, int]] = {}
    for style in document.styles:
        name = getattr(style, "name", "")
        if not name:
            continue
        available_styles[name] = _read_style_properties(style)
        reference = _style_numbering_reference(style)
        if reference is not None:
            style_numbering_map[name] = reference
            available_styles[name].setdefault("numbering_style", name)

    return TemplateView(
        available_styles=available_styles,
        available_numberings=_parse_numbering_definitions(document),
        style_numbering_map=style_numbering_map,
        document_defaults=_extract_document_defaults(document),
        has_template=True,
    )
