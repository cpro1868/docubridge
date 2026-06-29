from pytest import raises

from docubridge.core.style_resolver import resolve_effective_element_style, resolve_element_style
from docubridge.core.style_schema import StyleProfile
from docubridge.core.template_bridge import TemplateView


def test_yaml_overrides_template_values_and_tracks_source_map() -> None:
    profile = StyleProfile(
        defaults={
            "font_name": "Times New Roman",
            "font_size": 12,
            "line_spacing": 1.5,
        },
        elements={
            "heading1": {
                "template_style": "Heading 1",
                "font_size": 18,
                "bold": False,
            }
        },
    )
    template = TemplateView(
        available_styles={
            "Heading 1": {
                "font_name": "Arial",
                "font_size": 16,
                "bold": True,
            }
        }
    )

    resolved = resolve_element_style(profile, template, "heading1")

    assert resolved.word_style_name == "Heading 1"
    assert resolved.resolved_properties == {
        "font_name": "Arial",
        "font_size": 18,
        "line_spacing": 1.5,
        "bold": False,
    }
    assert resolved.source_map == {
        "font_name": "template_style",
        "font_size": "yaml",
        "line_spacing": "defaults",
        "bold": "yaml",
    }


def test_resolved_style_does_not_alias_profile_nested_defaults() -> None:
    profile = StyleProfile(
        defaults={
            "paragraph_format": {
                "spacing": {"before": 12, "after": 6},
            }
        },
        elements={
            "heading1": {
                "template_style": "Heading 1",
            }
        },
    )
    template = TemplateView(
        available_styles={
            "Heading 1": {
                "paragraph_format": {
                    "spacing": {"before": 24},
                }
            }
        }
    )

    resolved = resolve_element_style(profile, template, "heading1")
    resolved.resolved_properties["paragraph_format"]["spacing"]["before"] = 99

    assert profile.defaults["paragraph_format"]["spacing"]["before"] == 12


def test_resolve_element_style_uses_implicit_heading_style_when_template_style_is_blank() -> None:
    profile = StyleProfile(
        defaults={"font_size": 12},
        elements={
            "heading1": {
                "template_style": "   ",
                "font_size": 18,
            }
        },
    )
    template = TemplateView()

    resolved = resolve_element_style(profile, template, "heading1")

    assert resolved.word_style_name == "Heading 1"


def test_resolve_element_style_uses_implicit_heading_style_name() -> None:
    profile = StyleProfile(
        defaults={"font_size": 12},
        elements={
            "heading2": {
                "font_size": 16,
            }
        },
    )
    template = TemplateView()

    resolved = resolve_element_style(profile, template, "heading2")

    assert resolved.word_style_name == "Heading 2"


def test_resolve_element_style_keeps_paragraph_default_as_normal() -> None:
    profile = StyleProfile(
        defaults={"font_size": 12},
        elements={
            "paragraph": {
                "font_size": 12,
            }
        },
    )
    template = TemplateView()

    resolved = resolve_element_style(profile, template, "paragraph")

    assert resolved.word_style_name == "Normal"


def test_resolve_effective_element_style_supports_heading_fallback() -> None:
    profile = StyleProfile(
        defaults={"font_size": 12},
        elements={
            "heading1": {
                "font_size": 18,
            }
        },
    )
    template = TemplateView()

    resolved = resolve_effective_element_style(profile, template, "heading3")

    assert resolved.element_name == "heading3"
    assert resolved.word_style_name == "Heading 3"
    assert resolved.resolved_properties["font_size"] == 18
    assert resolved.source_map["font_size"] == "yaml"


def test_resolve_element_style_rejects_missing_element_name() -> None:
    profile = StyleProfile(defaults={"font_size": 12}, elements={})
    template = TemplateView()

    with raises(KeyError):
        resolve_element_style(profile, template, "heading1")

def test_resolve_element_style_rejects_missing_explicit_template_style() -> None:
    profile = StyleProfile(
        defaults={"font_size": 12},
        elements={
            "heading1": {
                "template_style": "Missing Style",
                "font_size": 18,
            }
        },
    )
    template = TemplateView(available_styles={}, has_template=True)

    with raises(KeyError):
        resolve_element_style(profile, template, "heading1")


def test_resolve_effective_style_merges_template_defaults_style_and_yaml() -> None:
    profile = StyleProfile.model_validate(
        {
            "defaults": {"font_size": 10},
            "elements": {
                "paragraph": {
                    "template_style": "Normal",
                    "font_east_asia": "仿宋",
                }
            },
        }
    )
    template = TemplateView(
        document_defaults={"font_ascii": "Calibri", "space_after_pt": 0},
        available_styles={"Normal": {"space_after_pt": 12, "font_east_asia": "宋体"}},
    )

    resolved = resolve_effective_element_style(profile, template, "paragraph")

    assert resolved.resolved_properties["font_ascii"] == "Calibri"
    assert resolved.resolved_properties["font_east_asia"] == "仿宋"
    assert resolved.resolved_properties["space_after_pt"] == 12
    assert resolved.source_map["font_ascii"] == "template_document_default"
    assert resolved.source_map["font_east_asia"] == "yaml"
    assert resolved.source_map["space_after_pt"] == "template_style"
