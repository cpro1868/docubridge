from pathlib import Path

from docubridge.adapters.markdown_adapter import load_markdown_tokens
from docubridge.core.nodes import CodeBlockNode, HeadingNode, HorizontalRuleNode, ImageBlockNode, ListItemNode, ListNode, ParagraphNode, QuoteNode, TableNode, TextSpan


def _extract_visible_text(inline_token) -> str:
    parts: list[str] = []
    for child in getattr(inline_token, "children", []) or []:
        if child.type == "text":
            parts.append(child.content)
            continue
        if child.type == "code_inline":
            parts.append(child.content)
            continue
        if child.type in {"softbreak", "hardbreak"}:
            parts.append("\n")
            continue
        if child.type == "image":
            if getattr(child, "children", None):
                parts.append(_extract_visible_text(child))
            elif child.content:
                parts.append(child.content)
            continue
        if getattr(child, "children", None):
            parts.append(_extract_visible_text(child))
    return "".join(parts)


def _make_inline_spans_from_text(text: str) -> list[TextSpan]:
    return [TextSpan(text=text)] if text else []


def _make_inline_spans(inline_token) -> list[TextSpan]:
    spans: list[TextSpan] = []
    bold = False
    italic = False
    strike = False
    href: str | None = None
    for child in getattr(inline_token, "children", []) or []:
        if child.type == "strong_open":
            bold = True
            continue
        if child.type == "strong_close":
            bold = False
            continue
        if child.type == "em_open":
            italic = True
            continue
        if child.type == "em_close":
            italic = False
            continue
        if child.type == "s_open":
            strike = True
            continue
        if child.type == "s_close":
            strike = False
            continue
        if child.type == "link_open":
            href = child.attrGet("href")
            continue
        if child.type == "link_close":
            href = None
            continue
        if child.type == "text":
            spans.append(TextSpan(text=child.content, bold=bold, italic=italic, strike=strike, href=href))
            continue
        if child.type == "softbreak" or child.type == "hardbreak":
            spans.append(TextSpan(text="\n", bold=bold, italic=italic, strike=strike, href=href))
            continue
        if child.type == "code_inline":
            spans.append(TextSpan(text=child.content, code=True, strike=strike, href=href))
            continue
        if child.type == "image":
            alt_text = _extract_visible_text(child) or child.content
            if alt_text:
                spans.append(TextSpan(text=alt_text, bold=bold, italic=italic, strike=strike, href=href))
            continue
        if getattr(child, "children", None):
            spans.extend(_make_inline_spans(child))
    return [span for span in spans if span.text]


def _extract_standalone_image(inline_token) -> ImageBlockNode | None:
    children = list(getattr(inline_token, "children", []) or [])
    if len(children) != 1:
        return None
    child = children[0]
    if child.type != "image":
        return None
    raw_path = child.attrGet("src") or ""
    title = child.attrGet("title")
    alt_text = _extract_visible_text(child) or child.content or ""
    return ImageBlockNode(raw_path=raw_path, alt_text=alt_text, title=title)


def _coerce_table_cell(inline_token) -> list[TextSpan]:
    return _make_inline_spans(inline_token)


def _parse_list_item_text(content: str) -> tuple[bool, bool | None, str]:
    lowered = content.lower()
    if lowered.startswith("[x] "):
        return True, True, content[4:]
    if lowered.startswith("[ ] "):
        return True, False, content[4:]
    return False, None, content


def _normalize_list_item_spans(inline_token) -> tuple[bool, bool | None, list[TextSpan]]:
    spans = _make_inline_spans(inline_token)
    visible_text = "".join(span.text for span in spans)
    task, checked, text = _parse_list_item_text(visible_text)
    if not task:
        return False, None, spans

    remaining = 4
    normalized_spans: list[TextSpan] = []
    for span in spans:
        if remaining <= 0:
            normalized_spans.append(span)
            continue
        if len(span.text) <= remaining:
            remaining -= len(span.text)
            continue
        normalized_spans.append(
            TextSpan(
                text=span.text[remaining:],
                bold=span.bold,
                italic=span.italic,
                strike=span.strike,
                code=span.code,
                href=span.href,
                type=span.type,
            )
        )
        remaining = 0

    if not normalized_spans and text:
        normalized_spans = _make_inline_spans_from_text(text)
    return task, checked, normalized_spans


def _parse_list(tokens: list[object], index: int, level: int = 0) -> tuple[ListNode, int]:
    open_token = tokens[index]
    kind = "ordered" if open_token.type == "ordered_list_open" else "unordered"
    start = 1
    if kind == "ordered":
        raw_start = open_token.attrGet("start")
        if raw_start is not None:
            try:
                start = max(int(raw_start), 1)
            except (TypeError, ValueError):
                start = 1
    close_type = open_token.type.replace("_open", "_close")
    items: list[ListItemNode] = []
    first_item_in_list = True
    index += 1

    while index < len(tokens) and tokens[index].type != close_type:
        token = tokens[index]
        if token.type != "list_item_open":
            index += 1
            continue

        item_inlines: list[TextSpan] | None = None
        task = False
        checked: bool | None = None
        nested_items: list[ListItemNode] = []
        index += 1

        while index < len(tokens) and tokens[index].type != "list_item_close":
            current = tokens[index]
            if current.type == "inline" and item_inlines is None:
                task, checked, item_inlines = _normalize_list_item_spans(current)
                index += 1
                continue
            if current.type in {"bullet_list_open", "ordered_list_open"}:
                nested_list, index = _parse_list(tokens, index, level + 1)
                nested_items.extend(nested_list.items)
                continue
            index += 1

        items.append(
            ListItemNode(
                inlines=item_inlines or [],
                task=task,
                checked=checked,
                level=level,
                kind=kind,
                sequence_start=start if kind == "ordered" and first_item_in_list else None,
            )
        )
        first_item_in_list = False
        items.extend(nested_items)
        index += 1

    return ListNode(kind=kind, start=start, items=items), index + 1


def parse_markdown_file(path: Path):
    tokens = load_markdown_tokens(path)
    nodes = []
    index = 0
    while index < len(tokens):
        token = tokens[index]
        if token.type == "heading_open":
            inline = tokens[index + 1]
            level = int(token.tag[1:])
            nodes.append(HeadingNode(level=level, inlines=_make_inline_spans(inline)))
            index += 3
            continue
        if token.type == "hr":
            nodes.append(HorizontalRuleNode())
            index += 1
            continue
        if token.type == "fence":
            info = (token.info or "").strip()
            language = info.split()[0] if info else None
            nodes.append(CodeBlockNode(content=token.content, language=language))
            index += 1
            continue
        if token.type == "blockquote_open":
            index += 1
            while index < len(tokens) and tokens[index].type != "blockquote_close":
                if tokens[index].type == "paragraph_open":
                    inline = tokens[index + 1]
                    nodes.append(QuoteNode(inlines=_make_inline_spans(inline)))
                    index += 3
                    continue
                index += 1
            index += 1
            continue
        if token.type == "paragraph_open":
            inline = tokens[index + 1]
            standalone_image = _extract_standalone_image(inline)
            if standalone_image is not None:
                nodes.append(standalone_image)
                index += 3
                continue
            nodes.append(ParagraphNode(inlines=_make_inline_spans(inline)))
            index += 3
            continue
        if token.type in {"bullet_list_open", "ordered_list_open"}:
            list_node, index = _parse_list(tokens, index)
            nodes.append(list_node)
            continue
        if token.type == "table_open":
            headers: list[list[TextSpan]] = []
            rows: list[list[list[TextSpan]]] = []
            index += 1
            while index < len(tokens) and tokens[index].type != "table_close":
                token = tokens[index]
                if token.type == "thead_open":
                    index += 1
                    while index < len(tokens) and tokens[index].type != "thead_close":
                        if tokens[index].type == "tr_open":
                            row: list[list[TextSpan]] = []
                            index += 1
                            while index < len(tokens) and tokens[index].type != "tr_close":
                                if tokens[index].type == "inline":
                                    row.append(_coerce_table_cell(tokens[index]))
                                index += 1
                            headers = row
                        index += 1
                    index += 1
                    continue
                if token.type == "tbody_open":
                    index += 1
                    while index < len(tokens) and tokens[index].type != "tbody_close":
                        if tokens[index].type == "tr_open":
                            row: list[list[TextSpan]] = []
                            index += 1
                            while index < len(tokens) and tokens[index].type != "tr_close":
                                if tokens[index].type == "inline":
                                    row.append(_coerce_table_cell(tokens[index]))
                                index += 1
                            if row:
                                rows.append(row)
                        index += 1
                    index += 1
                    continue
                index += 1
            nodes.append(TableNode(headers=headers, rows=rows))
            index += 1
            continue
        index += 1
    return nodes
