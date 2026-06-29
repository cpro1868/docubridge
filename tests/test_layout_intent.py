from docubridge.application.render_service import build_layout_intents
from docubridge.core.nodes import ListItemNode, ListNode, ParagraphNode, TextSpan
from docubridge.core.style_resolver import ResolvedStyle


def _resolved_style(name: str, **properties):
    return ResolvedStyle(
        element_name=name,
        word_style_name=properties.pop("word_style_name", "Normal"),
        resolved_properties=properties,
        source_map={key: "yaml" for key in properties},
    )


def test_build_layout_intents_maps_ordered_lists_to_numbering_intents() -> None:
    nodes = [
        ListNode(
            kind="ordered",
            items=[ListItemNode(inlines=[TextSpan(text="First")], level=0, kind="ordered")],
        )
    ]
    styles = {
        "paragraph": _resolved_style("paragraph"),
        "ordered_list": _resolved_style(
            "ordered_list",
            word_style_name="List Number",
            numbering_style="List Number",
        ),
    }

    intents = build_layout_intents(nodes, styles)

    assert len(intents) == 1
    assert intents[0].element_name == "ordered_list"
    assert intents[0].resolved_style_name == "List Number"
    assert intents[0].numbering is not None
    assert intents[0].numbering.numbering_role == "ordered_list"
    assert intents[0].numbering.level == 0
    assert intents[0].numbering.preferred_template_style == "List Number"


def test_build_layout_intents_maps_unordered_lists_to_numbering_intents_when_style_requests_numbering() -> None:
    nodes = [
        ListNode(
            kind="unordered",
            items=[ListItemNode(inlines=[TextSpan(text="Bullet")], level=0, kind="unordered")],
        )
    ]
    styles = {
        "paragraph": _resolved_style("paragraph"),
        "unordered_list": _resolved_style(
            "unordered_list",
            word_style_name="List Bullet",
            numbering_style="List Bullet",
        ),
    }

    intents = build_layout_intents(nodes, styles)

    assert len(intents) == 1
    assert intents[0].element_name == "unordered_list"
    assert intents[0].resolved_style_name == "List Bullet"
    assert intents[0].numbering is not None
    assert intents[0].numbering.numbering_role == "unordered_list"
    assert intents[0].numbering.level == 0
    assert intents[0].numbering.preferred_template_style == "List Bullet"
    assert intents[0].prefix_text == "- "


def test_build_layout_intents_preserves_inline_run_flags() -> None:
    nodes = [
        ParagraphNode(
            inlines=[
                TextSpan(text="Plain "),
                TextSpan(text="bold", bold=True),
                TextSpan(text="OpenAI", href="https://openai.com"),
            ]
        )
    ]
    styles = {
        "paragraph": _resolved_style("paragraph", word_style_name="Normal"),
    }

    intents = build_layout_intents(nodes, styles)

    assert len(intents) == 1
    assert intents[0].runs[0].text == "Plain "
    assert intents[0].runs[1].bold is True
    assert intents[0].runs[2].href == "https://openai.com"


def test_build_layout_intents_marks_ordered_list_restart_and_start_value() -> None:
    nodes = [
        ListNode(
            kind="ordered",
            start=3,
            items=[
                ListItemNode(inlines=[TextSpan(text="Third")], level=0, kind="ordered"),
                ListItemNode(inlines=[TextSpan(text="Fourth")], level=0, kind="ordered"),
            ],
        )
    ]
    styles = {
        "paragraph": _resolved_style("paragraph"),
        "ordered_list": _resolved_style(
            "ordered_list",
            word_style_name="List Number",
            numbering_style="List Number",
        ),
    }

    intents = build_layout_intents(nodes, styles)

    assert len(intents) == 2
    assert intents[0].numbering is not None
    assert intents[0].numbering.continue_sequence is False
    assert intents[0].numbering.start_at == 3
    assert intents[1].numbering is not None
    assert intents[1].numbering.continue_sequence is True
    assert intents[1].numbering.start_at is None


def test_build_layout_intents_uses_per_level_sequence_start_for_nested_ordered_lists() -> None:
    nodes = [
        ListNode(
            kind="ordered",
            start=1,
            items=[
                ListItemNode(inlines=[TextSpan(text="Parent One")], level=0, kind="ordered", sequence_start=1),
                ListItemNode(inlines=[TextSpan(text="Child A")], level=1, kind="ordered", sequence_start=3),
                ListItemNode(inlines=[TextSpan(text="Child B")], level=1, kind="ordered"),
                ListItemNode(inlines=[TextSpan(text="Parent Two")], level=0, kind="ordered"),
                ListItemNode(inlines=[TextSpan(text="Child C")], level=1, kind="ordered", sequence_start=7),
            ],
        )
    ]
    styles = {
        "paragraph": _resolved_style("paragraph"),
        "ordered_list": _resolved_style(
            "ordered_list",
            word_style_name="List Number",
            numbering_style="List Number",
        ),
    }

    intents = build_layout_intents(nodes, styles)

    assert len(intents) == 5
    assert intents[0].numbering is not None
    assert intents[0].numbering.level == 0
    assert intents[0].numbering.continue_sequence is False
    assert intents[0].numbering.start_at == 1
    assert intents[1].numbering is not None
    assert intents[1].numbering.level == 1
    assert intents[1].numbering.continue_sequence is False
    assert intents[1].numbering.start_at == 3
    assert intents[2].numbering is not None
    assert intents[2].numbering.level == 1
    assert intents[2].numbering.continue_sequence is True
    assert intents[2].numbering.start_at is None
    assert intents[3].numbering is not None
    assert intents[3].numbering.level == 0
    assert intents[3].numbering.continue_sequence is True
    assert intents[3].numbering.start_at is None
    assert intents[4].numbering is not None
    assert intents[4].numbering.level == 1
    assert intents[4].numbering.continue_sequence is False
    assert intents[4].numbering.start_at == 7
