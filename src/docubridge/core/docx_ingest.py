from __future__ import annotations

from pathlib import Path
import re

from docx import Document
from docx.document import Document as DocumentObject
from docx.oxml.ns import qn
from docx.oxml.text.run import CT_R
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.text.run import Run


def _iter_block_items(document: DocumentObject):
    body = document.element.body
    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, document)
        elif child.tag == qn("w:tbl"):
            yield Table(child, document)


def _export_image(
    document: DocumentObject,
    rel_id: str,
    media_dir: Path,
    exported: dict[str, tuple[str, str]],
) -> tuple[str, str]:
    if rel_id in exported:
        return exported[rel_id]

    image_part = document.part.related_parts[rel_id]
    partname = Path(str(image_part.partname))
    target_name = f"{len(exported) + 1:03d}-{partname.name}"
    target_path = media_dir / target_name
    target_path.write_bytes(image_part.blob)
    alt_text = Path(partname.name).stem
    exported[rel_id] = (target_name, alt_text)
    return exported[rel_id]


def _extract_paragraph_images(
    paragraph: Paragraph,
    document: DocumentObject,
    media_dir: Path | None,
    exported: dict[str, tuple[str, str]],
) -> list[str]:
    image_refs: list[str] = []
    if media_dir is None:
        return image_refs

    for blip in paragraph._p.xpath(".//*[local-name()='blip']"):
        rel_id = blip.get(qn("r:embed"))
        if not rel_id:
            continue
        alt_text = "image"
        picture_names = blip.xpath(
            "./ancestor::*[local-name()='pic']/*[local-name()='nvPicPr']/*[local-name()='cNvPr']/@name"
        )
        if picture_names:
            alt_text = Path(picture_names[0]).stem or "image"
        target_name, _ = _export_image(document, rel_id, media_dir, exported)
        image_refs.append(f"![{alt_text}](assets/{target_name})")
    return image_refs


def _append_markdown_table(lines: list[str], table) -> None:
    rows = [[_table_cell_to_markdown(cell) for cell in row.cells] for row in table.rows]
    rows = [row for row in rows if any(cell for cell in row)]
    if not rows:
        return

    header = rows[0]
    lines.append("| " + " | ".join(header) + " |")
    lines.append("| " + " | ".join("---" for _ in header) + " |")
    for row in rows[1:]:
        padded = row + [""] * (len(header) - len(row))
        lines.append("| " + " | ".join(padded[: len(header)]) + " |")
    lines.append("")


def _wrap_markdown_text(
    text: str,
    *,
    bold: bool = False,
    italic: bool = False,
    strike: bool = False,
) -> str:
    if not text:
        return ""
    if bold and italic:
        text = f"***{text}***"
    elif bold:
        text = f"**{text}**"
    elif italic:
        text = f"*{text}*"
    if strike:
        text = f"~~{text}~~"
    return text


def _is_code_run(run: Run) -> bool:
    font_name = (run.font.name or "").strip().lower()
    return font_name in {"consolas", "courier new", "courier"}


def _run_to_markdown(run: Run) -> str:
    text = _wrap_markdown_text(
        run.text,
        bold=bool(run.bold),
        italic=bool(run.italic),
        strike=bool(run.font.strike),
    )
    if text and _is_code_run(run):
        return f"`{text}`"
    return text


def _inline_fragment_from_run(run: Run) -> tuple[str, bool, bool, bool, bool]:
    return (
        run.text,
        bool(run.bold),
        bool(run.italic),
        bool(run.font.strike),
        _is_code_run(run),
    )


def _append_fragment(
    fragments: list[tuple[str, bool, bool, bool, bool]],
    fragment: tuple[str, bool, bool, bool, bool],
) -> None:
    text, bold, italic, strike, code = fragment
    if not text:
        return
    if fragments and fragments[-1][1:] == (bold, italic, strike, code):
        previous_text, _, _, _, _ = fragments[-1]
        fragments[-1] = (previous_text + text, bold, italic, strike, code)
        return
    fragments.append(fragment)


def _normalise_fragments(
    fragments: list[tuple[str, bool, bool, bool, bool]],
) -> list[tuple[str, bool, bool, bool, bool]]:
    if not fragments:
        return fragments

    normalised: list[tuple[str, bool, bool, bool, bool]] = [fragments[0]]
    for text, bold, italic, strike, code in fragments[1:]:
        previous_text, previous_bold, previous_italic, previous_strike, previous_code = normalised[-1]
        if not (previous_bold or previous_italic or previous_strike or previous_code) or (
            bold or italic or strike or code
        ):
            normalised.append((text, bold, italic, strike, code))
            continue

        punctuation, remainder = _split_leading_punctuation(text)
        if punctuation:
            normalised[-1] = (
                previous_text + punctuation,
                previous_bold,
                previous_italic,
                previous_strike,
                previous_code,
            )
        if remainder:
            normalised.append((remainder, bold, italic, strike, code))
    return normalised


def _split_leading_punctuation(text: str) -> tuple[str, str]:
    index = 0
    while index < len(text) and text[index] in "：:：，,。、；;）)]】》/·-":
        index += 1
    return text[:index], text[index:]


def _fragments_to_markdown(fragments: list[tuple[str, bool, bool, bool, bool]]) -> str:
    parts: list[str] = []
    for text, bold, italic, strike, code in _normalise_fragments(fragments):
        rendered = _wrap_markdown_text(text, bold=bold, italic=italic, strike=strike)
        if rendered and code:
            rendered = f"`{rendered}`"
        if parts and _needs_inline_separator(parts[-1], text):
            parts.append(" ")
        parts.append(rendered)
    return "".join(parts)


def _needs_inline_separator(previous_rendered: str, current_text: str) -> bool:
    if not previous_rendered or not current_text:
        return False
    if current_text[0].isspace():
        return False
    if not _starts_with_word_like(current_text[0]):
        return False
    return previous_rendered.endswith(("**", "*", "~~", "`"))


def _starts_with_word_like(char: str) -> bool:
    if char.isalnum():
        return True
    return "\u4e00" <= char <= "\u9fff"


def _hyperlink_to_markdown(paragraph: Paragraph, hyperlink) -> str:
    rel_id = hyperlink.get(qn("r:id"))
    href = ""
    if rel_id:
        relationship = paragraph.part.rels.get(rel_id)
        if relationship is not None:
            href = getattr(relationship, "target_ref", "")
    visible_text = "".join(
        run.text for run in (Run(child, paragraph) for child in hyperlink.iterchildren() if isinstance(child, CT_R))
    )
    if href and visible_text:
        return f"[{visible_text}]({href})"
    return visible_text


def _paragraph_to_markdown_text(paragraph: Paragraph) -> str:
    fragments: list[tuple[str, bool, bool, bool, bool]] = []
    for child in paragraph._p.iterchildren():
        if isinstance(child, CT_R):
            _append_fragment(fragments, _inline_fragment_from_run(Run(child, paragraph)))
            continue
        if child.tag == qn("w:hyperlink"):
            hyperlink_markdown = _hyperlink_to_markdown(paragraph, child)
            if hyperlink_markdown:
                _append_fragment(fragments, (hyperlink_markdown, False, False, False, False))
    return _fragments_to_markdown(fragments).strip()


def _table_cell_to_markdown(cell) -> str:
    parts = [_paragraph_to_markdown_text(paragraph) for paragraph in cell.paragraphs]
    parts = [part for part in parts if part]
    return _normalise_table_cell_text("<br>".join(parts).strip())


def _normalise_table_cell_text(text: str) -> str:
    if not text:
        return ""
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = text.replace("\n", "<br>")
    text = text.replace("|", r"\|")
    return text


def _list_indent_prefix_for_level(level: int) -> str:
    return "  " * max(level, 0)


def _list_indent_prefix(paragraph: Paragraph) -> str:
    return _list_indent_prefix_for_level(_list_level(paragraph))


def _list_level(paragraph: Paragraph) -> int:
    left_indent = paragraph.paragraph_format.left_indent
    if left_indent is None:
        return 0
    points = left_indent.pt
    if points <= 0:
        return 0
    return max(int(round(points / 36.0)), 0)


def _style_numbering_reference(style) -> tuple[int, int] | None:
    current = style
    while current is not None:
        properties = getattr(current.element, "pPr", None)
        num_pr = getattr(properties, "numPr", None)
        if num_pr is not None and getattr(num_pr, "numId", None) is not None:
            ilvl = getattr(num_pr, "ilvl", None)
            return int(num_pr.numId.val), int(ilvl.val if ilvl is not None else 0)
        current = current.base_style
    return None


def _paragraph_numbering_reference(paragraph: Paragraph) -> tuple[int, int] | None:
    properties = paragraph._p.pPr
    num_pr = getattr(properties, "numPr", None) if properties is not None else None
    if num_pr is not None and getattr(num_pr, "numId", None) is not None:
        ilvl = getattr(num_pr, "ilvl", None)
        return int(num_pr.numId.val), int(ilvl.val if ilvl is not None else 0)
    return _style_numbering_reference(paragraph.style)


def _parse_numbering_definitions(
    document: DocumentObject,
) -> tuple[dict[int, dict[int, dict[str, object]]], dict[int, dict[str, object]]]:
    try:
        numbering = document.part.numbering_part.element
    except KeyError:
        return {}, {}

    abstract_numbering: dict[int, dict[int, dict[str, object]]] = {}
    for abstract_node in numbering.xpath("./*[local-name()='abstractNum']"):
        abstract_id = int(abstract_node.get(qn("w:abstractNumId")))
        levels: dict[int, dict[str, object]] = {}
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

    numbering_instances: dict[int, dict[str, object]] = {}
    for num_node in numbering.xpath("./*[local-name()='num']"):
        num_id = int(num_node.get(qn("w:numId")))
        abstract_refs = num_node.xpath("./*[local-name()='abstractNumId']/@*[local-name()='val']")
        if not abstract_refs:
            continue

        start_overrides: dict[int, int] = {}
        for override_node in num_node.xpath("./*[local-name()='lvlOverride']"):
            level = int(override_node.get(qn("w:ilvl")))
            start_nodes = override_node.xpath("./*[local-name()='startOverride']/@*[local-name()='val']")
            if start_nodes:
                start_overrides[level] = int(start_nodes[0])
        numbering_instances[num_id] = {
            "abstract_id": int(abstract_refs[0]),
            "start_overrides": start_overrides,
        }
    return abstract_numbering, numbering_instances


def _render_numbering_marker(lvl_text: str, counters: dict[int, int], level: int) -> str:
    rendered = lvl_text
    for index in range(level + 1):
        rendered = rendered.replace(f"%{index + 1}", str(counters.get(index, 1)))
    rendered = re.sub(r"%\\d+", "", rendered)
    return rendered.strip() or f"{counters.get(level, 1)}."


def _resolve_numbering_marker(
    paragraph: Paragraph,
    numbering_state: dict[int, dict[int, int]],
    abstract_numbering: dict[int, dict[int, dict[str, object]]],
    numbering_instances: dict[int, dict[str, object]],
) -> tuple[str | None, int | None, int | None]:
    reference = _paragraph_numbering_reference(paragraph)
    if reference is None:
        return None, None, None

    num_id, level = reference
    instance = numbering_instances.get(num_id)
    if instance is None:
        return None, level, num_id

    abstract_levels = abstract_numbering.get(int(instance["abstract_id"]), {})
    level_definition = abstract_levels.get(level)
    if level_definition is None:
        return None, level, num_id

    if str(level_definition.get("num_fmt", "decimal")) == "bullet":
        return None, level, num_id

    counters = numbering_state.setdefault(num_id, {})
    for existing_level in list(counters):
        if existing_level > level:
            counters.pop(existing_level, None)

    if level not in counters:
        overrides = instance.get("start_overrides", {})
        start_value = int(overrides.get(level, level_definition.get("start", 1)))
        counters[level] = start_value
    else:
        counters[level] += 1

    marker = _render_numbering_marker(str(level_definition.get("lvl_text", f"%{level + 1}.")), counters, level)
    return marker, level, num_id


def _is_heading_style(style_name: str) -> bool:
    return style_name.startswith("Heading ")


def _heading_level(style_name: str) -> int:
    level_text = style_name.removeprefix("Heading ").strip()
    try:
        return max(int(level_text), 1)
    except ValueError:
        return 1


def _is_bullet_style(style_name: str) -> bool:
    return style_name.startswith("List Bullet")


def _is_numbered_style(style_name: str) -> bool:
    return style_name.startswith("List Number")


def _list_level_from_paragraph_or_numbering(paragraph: Paragraph, numbering_level: int | None) -> int:
    level = _list_level(paragraph)
    if level > 0:
        return level
    return max(numbering_level or 0, 0)


def _normalise_list_marker(marker: str) -> str:
    return marker if marker.endswith((".", ")")) else f"{marker}."


def _format_numbered_item(text: str, marker: str, level: int) -> str:
    return f"{_list_indent_prefix_for_level(level)}{_normalise_list_marker(marker)} {text}"


def _format_heading(text: str, marker: str | None, level: int) -> str:
    prefix = f"{marker} " if marker else ""
    return f"{'#' * max(level, 1)} {prefix}{text}"


def _marker_number(marker: str) -> int | None:
    match = re.match(r"^\s*(\d+)", marker)
    if match is None:
        return None
    return int(match.group(1))


def parse_docx_file(path: Path, media_dir: Path | None = None) -> str:
    document = Document(path)
    lines: list[str] = []
    exported: dict[str, tuple[str, str]] = {}
    abstract_numbering, numbering_instances = _parse_numbering_definitions(document)
    numbering_state: dict[int, dict[int, int]] = {}
    numbered_counters: dict[int, int] = {}

    if media_dir is not None:
        media_dir.mkdir(parents=True, exist_ok=True)

    for block in _iter_block_items(document):
        if isinstance(block, Table):
            numbered_counters.clear()
            _append_markdown_table(lines, block)
            continue

        paragraph = block
        text = _paragraph_to_markdown_text(paragraph)
        image_refs = _extract_paragraph_images(paragraph, document, media_dir, exported)
        if not text and not image_refs:
            continue

        if not text:
            numbered_counters.clear()
            lines.extend(image_refs)
            lines.append("")
            continue

        style_name = paragraph.style.name if paragraph.style is not None else ""
        numbering_marker, numbering_level, numbering_num_id = _resolve_numbering_marker(
            paragraph,
            numbering_state,
            abstract_numbering,
            numbering_instances,
        )

        if _is_heading_style(style_name):
            numbered_counters.clear()
            lines.append(_format_heading(text, numbering_marker, _heading_level(style_name)))
            lines.extend(image_refs)
            lines.append("")
            continue

        if _is_bullet_style(style_name) or (numbering_level is not None and numbering_marker is None):
            numbered_counters.clear()
            level = _list_level_from_paragraph_or_numbering(paragraph, numbering_level)
            lines.append(f"{_list_indent_prefix_for_level(level)}- {text}")
            lines.extend(image_refs)
            continue

        if numbering_marker is not None:
            level = _list_level_from_paragraph_or_numbering(paragraph, numbering_level)
            for existing_level in list(numbered_counters):
                if existing_level > level:
                    numbered_counters.pop(existing_level, None)
            if numbering_level is not None and level > numbering_level:
                if numbering_num_id is not None:
                    counters = numbering_state.get(numbering_num_id)
                    if counters is not None and numbering_level in counters:
                        counters[numbering_level] = max(counters[numbering_level] - 1, 0)
                index = numbered_counters.get(level, 0) + 1
                numbered_counters[level] = index
                lines.append(_format_numbered_item(text, f"{index}.", level))
            else:
                marker_value = _marker_number(numbering_marker)
                if marker_value is not None:
                    numbered_counters[level] = marker_value
                lines.append(_format_numbered_item(text, numbering_marker, level))
            lines.extend(image_refs)
            continue

        if _is_numbered_style(style_name):
            level = _list_level(paragraph)
            for existing_level in list(numbered_counters):
                if existing_level > level:
                    numbered_counters.pop(existing_level, None)
            index = numbered_counters.get(level, 0) + 1
            numbered_counters[level] = index
            lines.append(f"{_list_indent_prefix(paragraph)}{index}. {text}")
            lines.extend(image_refs)
            continue

        numbered_counters.clear()
        lines.append(text)
        lines.extend(image_refs)
        lines.append("")

    while lines and lines[-1] == "":
        lines.pop()
    return "\n".join(lines) + ("\n" if lines else "")
