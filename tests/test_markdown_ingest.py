from docubridge.core.markdown_ingest import parse_markdown_file

from pathlib import Path
import base64


FIXTURE_PATH = Path(__file__).resolve().parent / "fixtures" / "sample.md"
_PNG_1X1 = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aRX0AAAAASUVORK5CYII="
)


def test_parse_markdown_file_returns_heading_and_paragraph_nodes() -> None:
    nodes = parse_markdown_file(FIXTURE_PATH)

    assert nodes[0].type == "heading"
    assert nodes[0].level == 1
    assert nodes[0].inlines[0].text == "Sample Title"
    assert nodes[1].type == "paragraph"
    assert "".join(span.text for span in nodes[1].inlines) == "Plain paragraph with emphasis, link, and code."
    assert nodes[2].type == "paragraph"
    assert "".join(span.text for span in nodes[2].inlines) == "Line one\nLine two with diagram alt."


def test_parse_markdown_file_preserves_inline_spans_inside_headings(tmp_path: Path) -> None:
    heading_path = tmp_path / "heading-inline.md"
    heading_path.write_text(
        "# Title **bold** *italic* ~~gone~~ `code` [OpenAI](https://openai.com)\n",
        encoding="utf-8",
    )

    nodes = parse_markdown_file(heading_path)

    assert len(nodes) == 1
    heading = nodes[0]
    assert heading.type == "heading"
    assert heading.level == 1
    assert [(span.text, span.bold, span.italic, span.strike, span.code, span.href) for span in heading.inlines] == [
        ("Title ", False, False, False, False, None),
        ("bold", True, False, False, False, None),
        (" ", False, False, False, False, None),
        ("italic", False, True, False, False, None),
        (" ", False, False, False, False, None),
        ("gone", False, False, True, False, None),
        (" ", False, False, False, False, None),
        ("code", False, False, False, True, None),
        (" ", False, False, False, False, None),
        ("OpenAI", False, False, False, False, "https://openai.com"),
    ]


def test_task_list_items_capture_checked_state() -> None:
    nodes = parse_markdown_file(FIXTURE_PATH)
    task_list = next(node for node in nodes if node.type == "list")

    assert task_list.items[0].task is True
    assert task_list.items[0].checked is True
    assert task_list.items[1].task is True
    assert task_list.items[1].checked is False


def test_parse_markdown_file_preserves_nested_list_levels(tmp_path: Path) -> None:
    nested_list_path = tmp_path / "nested-list.md"
    nested_list_path.write_text(
        "\n".join(
            [
                "- Parent",
                "  - Child",
                "1. Ordered parent",
                "   1. Ordered child",
            ]
        )
        + "\n",
        encoding="utf-8",
    )

    nodes = parse_markdown_file(nested_list_path)

    assert [node.type for node in nodes] == ["list", "list"]
    assert nodes[0].kind == "unordered"
    assert [(item.inlines[0].text, item.level, item.kind) for item in nodes[0].items] == [
        ("Parent", 0, "unordered"),
        ("Child", 1, "unordered"),
    ]
    assert nodes[1].kind == "ordered"
    assert [(item.inlines[0].text, item.level, item.kind) for item in nodes[1].items] == [
        ("Ordered parent", 0, "ordered"),
        ("Ordered child", 1, "ordered"),
    ]


def test_parse_markdown_file_preserves_nested_ordered_list_start_values(tmp_path: Path) -> None:
    nested_start_path = tmp_path / "nested-ordered-start.md"
    nested_start_path.write_text(
        "\n".join(
            [
                "1. Parent One",
                "   1. Child A",
                "   2. Child B",
                "2. Parent Two",
                "   1. Child C",
                "   2. Child D",
            ]
        )
        + "\n",
        encoding="utf-8",
    )

    nodes = parse_markdown_file(nested_start_path)

    assert len(nodes) == 1
    list_node = nodes[0]
    assert list_node.type == "list"
    assert list_node.kind == "ordered"
    assert [(item.inlines[0].text, item.level, item.sequence_start) for item in list_node.items] == [
        ("Parent One", 0, 1),
        ("Child A", 1, 1),
        ("Child B", 1, None),
        ("Parent Two", 0, None),
        ("Child C", 1, 1),
        ("Child D", 1, None),
    ]


def test_parse_markdown_file_preserves_inline_spans_inside_list_items(tmp_path: Path) -> None:
    list_path = tmp_path / "list-inline.md"
    list_path.write_text(
        "- Plain **bold** *italic* `code` [OpenAI](https://openai.com)\n",
        encoding="utf-8",
    )

    nodes = parse_markdown_file(list_path)

    assert len(nodes) == 1
    list_node = nodes[0]
    assert list_node.type == "list"
    assert [(span.text, span.bold, span.italic, span.code, span.href) for span in list_node.items[0].inlines] == [
        ("Plain ", False, False, False, None),
        ("bold", True, False, False, None),
        (" ", False, False, False, None),
        ("italic", False, True, False, None),
        (" ", False, False, False, None),
        ("code", False, False, True, None),
        (" ", False, False, False, None),
        ("OpenAI", False, False, False, "https://openai.com"),
    ]


def test_parse_markdown_file_preserves_task_list_inline_spans(tmp_path: Path) -> None:
    list_path = tmp_path / "task-list-inline.md"
    list_path.write_text(
        "- [x] **done** ~~old~~ `code` [ref](https://example.com)\n",
        encoding="utf-8",
    )

    nodes = parse_markdown_file(list_path)

    assert len(nodes) == 1
    list_node = nodes[0]
    assert list_node.items[0].task is True
    assert list_node.items[0].checked is True
    assert [(span.text, span.bold, span.strike, span.code, span.href) for span in list_node.items[0].inlines] == [
        ("done", True, False, False, None),
        (" ", False, False, False, None),
        ("old", False, True, False, None),
        (" ", False, False, False, None),
        ("code", False, False, True, None),
        (" ", False, False, False, None),
        ("ref", False, False, False, "https://example.com"),
    ]


def test_parse_markdown_file_extracts_gfm_tables(tmp_path: Path) -> None:
    table_path = tmp_path / "table.md"
    table_path.write_text(
        "\n".join(
            [
                "| Name | Value |",
                "| --- | --- |",
                "| Alpha | 1 |",
                "| Beta | 2 |",
            ]
        )
        + "\n",
        encoding="utf-8",
    )

    nodes = parse_markdown_file(table_path)

    assert len(nodes) == 1
    table = nodes[0]
    assert table.type == "table"
    assert ["".join(span.text for span in cell) for cell in table.headers] == ["Name", "Value"]
    assert [
        ["".join(span.text for span in cell) for cell in row]
        for row in table.rows
    ] == [["Alpha", "1"], ["Beta", "2"]]


def test_parse_markdown_file_preserves_inline_spans_inside_table_cells(tmp_path: Path) -> None:
    table_path = tmp_path / "table-inline.md"
    table_path.write_text(
        "\n".join(
            [
                "| Name | Value |",
                "| --- | --- |",
                "| **Alpha** | `42` [ref](https://example.com) |",
            ]
        )
        + "\n",
        encoding="utf-8",
    )

    nodes = parse_markdown_file(table_path)

    assert len(nodes) == 1
    table = nodes[0]
    assert [(span.text, span.bold, span.italic, span.code, span.href) for span in table.rows[0][0]] == [
        ("Alpha", True, False, False, None),
    ]
    assert [(span.text, span.bold, span.italic, span.code, span.href) for span in table.rows[0][1]] == [
        ("42", False, False, True, None),
        (" ", False, False, False, None),
        ("ref", False, False, False, "https://example.com"),
    ]


def test_parse_markdown_file_extracts_fenced_code_blocks(tmp_path: Path) -> None:
    code_path = tmp_path / "code.md"
    code_path.write_text(
        "\n".join(
            [
                "```python",
                "print('hello')",
                "print('world')",
                "```",
            ]
        )
        + "\n",
        encoding="utf-8",
    )

    nodes = parse_markdown_file(code_path)

    assert len(nodes) == 1
    code_block = nodes[0]
    assert code_block.type == "code_block"
    assert code_block.language == "python"
    assert code_block.content == "print('hello')\nprint('world')\n"


def test_parse_markdown_file_extracts_standalone_images(tmp_path: Path) -> None:
    image_path = tmp_path / "diagram.png"
    image_path.write_bytes(_PNG_1X1)
    markdown_path = tmp_path / "image.md"
    markdown_path.write_text(f"![diagram]({image_path.name})\n", encoding="utf-8")

    nodes = parse_markdown_file(markdown_path)

    assert len(nodes) == 1
    image = nodes[0]
    assert image.type == "image"
    assert image.raw_path == image_path.name
    assert image.alt_text == "diagram"


def test_parse_markdown_file_extracts_blockquotes(tmp_path: Path) -> None:
    quote_path = tmp_path / "quote.md"
    quote_path.write_text("> Quoted line\n", encoding="utf-8")

    nodes = parse_markdown_file(quote_path)

    assert len(nodes) == 1
    quote = nodes[0]
    assert quote.type == "quote"
    assert quote.inlines[0].text == "Quoted line"


def test_parse_markdown_file_preserves_inline_spans_inside_blockquotes(tmp_path: Path) -> None:
    quote_path = tmp_path / "quote-inline.md"
    quote_path.write_text(
        "> Quote **bold** *italic* `code` [OpenAI](https://openai.com)\n",
        encoding="utf-8",
    )

    nodes = parse_markdown_file(quote_path)

    assert len(nodes) == 1
    quote = nodes[0]
    assert quote.type == "quote"
    assert [(span.text, span.bold, span.italic, span.code, span.href) for span in quote.inlines] == [
        ("Quote ", False, False, False, None),
        ("bold", True, False, False, None),
        (" ", False, False, False, None),
        ("italic", False, True, False, None),
        (" ", False, False, False, None),
        ("code", False, False, True, None),
        (" ", False, False, False, None),
        ("OpenAI", False, False, False, "https://openai.com"),
    ]


def test_parse_markdown_file_extracts_horizontal_rules(tmp_path: Path) -> None:
    rule_path = tmp_path / "rule.md"
    rule_path.write_text("---\n", encoding="utf-8")

    nodes = parse_markdown_file(rule_path)

    assert len(nodes) == 1
    rule = nodes[0]
    assert rule.type == "horizontal_rule"


def test_parse_markdown_file_preserves_inline_bold_italic_and_code_spans(tmp_path: Path) -> None:
    inline_path = tmp_path / "inline.md"
    inline_path.write_text("Plain **bold** *italic* `code`\n", encoding="utf-8")

    nodes = parse_markdown_file(inline_path)

    assert len(nodes) == 1
    paragraph = nodes[0]
    assert paragraph.type == "paragraph"
    assert [(span.text, span.bold, span.italic, span.code) for span in paragraph.inlines] == [
        ("Plain ", False, False, False),
        ("bold", True, False, False),
        (" ", False, False, False),
        ("italic", False, True, False),
        (" ", False, False, False),
        ("code", False, False, True),
    ]


def test_parse_markdown_file_preserves_inline_links(tmp_path: Path) -> None:
    link_path = tmp_path / "link.md"
    link_path.write_text("Visit [OpenAI](https://openai.com)\n", encoding="utf-8")

    nodes = parse_markdown_file(link_path)

    assert len(nodes) == 1
    paragraph = nodes[0]
    assert paragraph.type == "paragraph"
    assert [(span.text, span.href) for span in paragraph.inlines] == [
        ("Visit ", None),
        ("OpenAI", "https://openai.com"),
    ]


def test_parse_markdown_file_preserves_inline_strikethrough(tmp_path: Path) -> None:
    strike_path = tmp_path / "strike.md"
    strike_path.write_text("Keep ~~remove~~ text\n", encoding="utf-8")

    nodes = parse_markdown_file(strike_path)

    assert len(nodes) == 1
    paragraph = nodes[0]
    assert paragraph.type == "paragraph"
    assert [(span.text, span.strike) for span in paragraph.inlines] == [
        ("Keep ", False),
        ("remove", True),
        (" text", False),
    ]
