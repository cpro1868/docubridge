from __future__ import annotations

from difflib import SequenceMatcher
from pathlib import Path
from typing import Any

from docubridge.core.style_schema import StyleProfile
from docubridge.core.template_bridge import TemplateView, load_template_view


_EXACT_ALIASES: dict[str, str] = {
    "heading 1": "heading1",
    "heading 2": "heading2",
    "heading 3": "heading3",
    "heading 4": "heading4",
    "heading 5": "heading5",
    "heading 6": "heading6",
    "标题 1": "heading1",
    "标题 2": "heading2",
    "标题 3": "heading3",
    "标题 4": "heading4",
    "标题 5": "heading5",
    "标题 6": "heading6",
    "normal": "paragraph",
    "正文": "paragraph",
    "list number": "ordered_list",
    "list bullet": "unordered_list",
    "quote": "quote",
    "block quote": "quote",
    "blockquote": "quote",
    "table grid": "table",
    "light grid": "table",
    "medium grid": "table",
    "dark grid": "table",
    "code": "code_block",
    "preformatted": "code_block",
}

_SIMILAR_CANDIDATES: dict[str, str] = {
    "title": "heading1",
    "subtitle": "heading2",
    "body text": "paragraph",
    "bullet": "unordered_list",
    "numbering": "ordered_list",
    "caption": "paragraph",
}

_ELEMENT_ORDER = [
    "heading1",
    "heading2",
    "heading3",
    "heading4",
    "heading5",
    "heading6",
    "paragraph",
    "ordered_list",
    "unordered_list",
    "quote",
    "table",
    "code_block",
]

_ALIGNMENT_LABELS: dict[int, str] = {
    0: "left",
    1: "center",
    2: "right",
    3: "justify",
}


def _normalize_name(name: str) -> str:
    return name.strip().lower()


def _is_exact_alias(name: str) -> bool:
    return _normalize_name(name) in _EXACT_ALIASES


def _map_style_name(name: str) -> str | None:
    normalized = _normalize_name(name)
    if normalized in _EXACT_ALIASES:
        return _EXACT_ALIASES[normalized]
    best_candidate: str | None = None
    best_score = 0.0
    for candidate, element in _SIMILAR_CANDIDATES.items():
        score = SequenceMatcher(None, normalized, candidate).ratio()
        if score > best_score:
            best_score = score
            best_candidate = element
    for element in _ELEMENT_ORDER:
        score = SequenceMatcher(None, normalized, element).ratio()
        if score > best_score:
            best_score = score
            best_candidate = element
    if best_score >= 0.6:
        return best_candidate
    return None


def _assign_element_mappings(style_names: list[str]) -> dict[str, str]:
    """Map Word style names to element slots, preferring exact aliases."""
    assignments: dict[str, str] = {}

    # First pass: exact aliases take precedence.
    for name in style_names:
        if _is_exact_alias(name):
            element = _map_style_name(name)
            if element is not None:
                assignments[name] = element

    # Second pass: similarity-based mapping only for unassigned elements.
    assigned_elements = set(assignments.values())
    for name in style_names:
        if name in assignments:
            continue
        element = _map_style_name(name)
        if element is None:
            continue
        if element in assigned_elements:
            continue
        assignments[name] = element
        assigned_elements.add(element)

    return assignments


def _normalize_property_value(key: str, value: Any) -> Any:
    if key == "alignment" and isinstance(value, int):
        return _ALIGNMENT_LABELS.get(value, value)
    return value


def _filter_properties(properties: dict[str, Any]) -> dict[str, Any]:
    filtered: dict[str, Any] = {}
    for key, value in properties.items():
        if value is None:
            continue
        if isinstance(value, bool) and not value:
            continue
        if isinstance(value, str) and not value.strip():
            continue
        if key == "numbering_style":
            continue
        filtered[key] = _normalize_property_value(key, value)
    return filtered


def _deduplicate_properties(
    element_props: dict[str, Any],
    defaults: dict[str, Any],
) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in element_props.items():
        default_value = defaults.get(key)
        if value == default_value:
            continue
        result[key] = value
    return result


def _ordered_elements(elements: dict[str, dict[str, Any]]) -> dict[str, dict[str, Any]]:
    ordered: dict[str, dict[str, Any]] = {}
    for name in _ELEMENT_ORDER:
        if name in elements:
            ordered[name] = elements[name]
    for name in sorted(elements):
        if name not in ordered:
            ordered[name] = elements[name]
    return ordered


def _build_profile_data(
    template: TemplateView,
    source_path: Path,
) -> tuple[dict[str, Any], list[str]]:
    defaults = _filter_properties(template.document_defaults)

    mappings = _assign_element_mappings(list(template.available_styles.keys()))

    elements: dict[str, dict[str, Any]] = {}
    unmapped: dict[str, dict[str, Any]] = {}

    for style_name, properties in template.available_styles.items():
        element_name = mappings.get(style_name)
        filtered = _filter_properties(properties)
        if element_name is None:
            unmapped[style_name] = filtered
            continue
        element_props = _deduplicate_properties(filtered, defaults)
        element_props["template_style"] = style_name
        elements[element_name] = element_props

    data: dict[str, Any] = {
        "meta": {
            "name": source_path.stem,
            "source": source_path.as_posix(),
            "version": 1,
        },
        "defaults": defaults,
        "elements": _ordered_elements(elements),
    }

    if unmapped:
        data["compat"] = {
            "extracted_styles": dict(sorted(unmapped.items())),
        }

    return data, sorted(unmapped)


def extract_style_profile(input_path: Path) -> tuple[StyleProfile, list[str]]:
    template = load_template_view(input_path)
    data, unmapped = _build_profile_data(template, input_path)
    return StyleProfile.model_validate(data), unmapped
