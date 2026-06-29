from __future__ import annotations

from copy import deepcopy
from dataclasses import dataclass
from typing import Any

from docubridge.core.style_schema import StyleProfile
from docubridge.core.template_bridge import TemplateView


@dataclass(slots=True)
class ResolvedStyle:
    element_name: str
    word_style_name: str
    resolved_properties: dict[str, Any]
    source_map: dict[str, str]


def _implicit_word_style_name(element_name: str) -> str:
    if element_name.startswith("heading") and element_name[7:].isdigit():
        return f"Heading {element_name[7:]}"
    return "Normal"


def _is_heading_element(element_name: str) -> bool:
    return element_name.startswith("heading") and element_name[7:].isdigit()


def resolve_element_style(
    profile: StyleProfile,
    template: TemplateView,
    element_name: str,
) -> ResolvedStyle:
    if element_name not in profile.elements:
        raise KeyError(f"Unknown style element: {element_name}")

    element = profile.elements[element_name]
    template_style = element.get("template_style")
    if isinstance(template_style, str) and template_style.strip():
        word_style_name = template_style.strip()
        if template.has_template and word_style_name not in template.available_styles:
            raise KeyError(f"Unknown template style: {word_style_name}")
    else:
        word_style_name = _implicit_word_style_name(element_name)

    resolved_properties: dict[str, Any] = deepcopy(profile.defaults)
    source_map: dict[str, str] = {key: "defaults" for key in resolved_properties}

    for key, value in template.document_defaults.items():
        resolved_properties[key] = deepcopy(value)
        source_map[key] = "template_document_default"

    template_properties = template.get_style(word_style_name)
    for key, value in template_properties.items():
        resolved_properties[key] = deepcopy(value)
        source_map[key] = "template_style"

    for key, value in element.items():
        if key == "template_style":
            continue
        resolved_properties[key] = deepcopy(value)
        source_map[key] = "yaml"

    if "numbering_style" not in resolved_properties and word_style_name in template.style_numbering_map:
        resolved_properties["numbering_style"] = word_style_name
        source_map["numbering_style"] = "template_style"

    return ResolvedStyle(
        element_name=element_name,
        word_style_name=word_style_name,
        resolved_properties=resolved_properties,
        source_map=source_map,
    )


def resolve_effective_element_style(
    profile: StyleProfile,
    template: TemplateView,
    element_name: str,
) -> ResolvedStyle:
    if element_name in profile.elements:
        return resolve_element_style(profile, template, element_name)
    if _is_heading_element(element_name) and "heading1" in profile.elements:
        fallback = resolve_element_style(profile, template, "heading1")
        return ResolvedStyle(
            element_name=element_name,
            word_style_name=_implicit_word_style_name(element_name),
            resolved_properties=deepcopy(fallback.resolved_properties),
            source_map=deepcopy(fallback.source_map),
        )
    raise KeyError(f"Unknown style element: {element_name}")
