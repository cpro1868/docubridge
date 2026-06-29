from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from pydantic import ValidationError

from docubridge.application.models import RenderRequest
from docubridge.core.diagnostics import Diagnostic, Severity
from docubridge.core.layout_intent import NumberingIntent, ParagraphLayoutIntent, RunStyleIntent
from docubridge.core.markdown_ingest import parse_markdown_file
from docubridge.core.nodes import HeadingNode, ListNode, ParagraphNode, QuoteNode
from docubridge.core.style_resolver import resolve_effective_element_style
from docubridge.core.style_schema import load_style_profile
from docubridge.core.template_bridge import load_template_view
from docubridge.core.word_renderer import render_nodes_to_document


@dataclass(slots=True)
class RenderResult:
    success: bool
    output_path: Path
    diagnostics: list[dict] = field(default_factory=list)


def _build_run_intents(inlines: list[object]) -> list[RunStyleIntent]:
    return [
        RunStyleIntent(
            text=getattr(inline, "text", ""),
            bold=bool(getattr(inline, "bold", False)),
            italic=bool(getattr(inline, "italic", False)),
            strike=bool(getattr(inline, "strike", False)),
            code=bool(getattr(inline, "code", False)),
            href=getattr(inline, "href", None),
        )
        for inline in inlines
    ]


def _list_prefix_text(item: object, kind: str) -> str:
    level = max(getattr(item, "level", 0), 0)
    indent = "  " * level
    if getattr(item, "task", False):
        checked = getattr(item, "checked", None)
        marker = "\u2611" if checked is True else "\u2610"
        return f"{indent}{marker} "
    if kind == "unordered":
        return f"{indent}- "
    return ""


def build_layout_intents(nodes: list[object], styles: dict[str, object]) -> list[ParagraphLayoutIntent]:
    intents: list[ParagraphLayoutIntent] = []
    for node in nodes:
        if isinstance(node, HeadingNode):
            style = styles[f"heading{node.level}"]
            intents.append(
                ParagraphLayoutIntent(
                    element_name=f"heading{node.level}",
                    runs=_build_run_intents(list(node.inlines)),
                    resolved_style_name=style.word_style_name,
                    resolved_properties=dict(getattr(style, "resolved_properties", {})),
                )
            )
            continue
        if isinstance(node, ParagraphNode):
            style = styles["paragraph"]
            intents.append(
                ParagraphLayoutIntent(
                    element_name="paragraph",
                    runs=_build_run_intents(list(node.inlines)),
                    resolved_style_name=style.word_style_name,
                    resolved_properties=dict(getattr(style, "resolved_properties", {})),
                )
            )
            continue
        if isinstance(node, QuoteNode):
            style = styles.get("quote", styles["paragraph"])
            intents.append(
                ParagraphLayoutIntent(
                    element_name="quote",
                    runs=_build_run_intents(list(node.inlines)),
                    resolved_style_name=style.word_style_name,
                    resolved_properties=dict(getattr(style, "resolved_properties", {})),
                    prefix_text="> ",
                )
            )
            continue
        if isinstance(node, ListNode):
            ordered_counters: dict[int, int] = {}
            raw_start = getattr(node, "start", 1)
            try:
                ordered_list_start = max(int(raw_start), 1)
            except (TypeError, ValueError):
                ordered_list_start = 1
            first_ordered_item = True
            for item in node.items:
                item_kind = getattr(item, "kind", None) or node.kind
                element_name = "ordered_list" if item_kind == "ordered" else "unordered_list"
                style = styles.get(element_name, styles["paragraph"])
                numbering = None
                prefix_text = ""
                if item_kind == "ordered":
                    item_level = max(getattr(item, "level", 0), 0)
                    raw_sequence_start = getattr(item, "sequence_start", None)
                    try:
                        sequence_start = max(int(raw_sequence_start), 1) if raw_sequence_start is not None else None
                    except (TypeError, ValueError):
                        sequence_start = None
                    if sequence_start is None and first_ordered_item:
                        sequence_start = ordered_list_start
                    ordered_counters = {
                        level: counter for level, counter in ordered_counters.items() if level <= item_level
                    }
                    if sequence_start is not None:
                        ordered_counters[item_level] = sequence_start - 1
                    elif item_level == 0 and 0 not in ordered_counters:
                        ordered_counters[0] = ordered_list_start - 1
                    display_index = ordered_counters.get(item_level, 0) + 1
                    ordered_counters[item_level] = display_index
                    numbering = NumberingIntent(
                        numbering_role="ordered_list",
                        level=item_level,
                        continue_sequence=sequence_start is None,
                        start_at=sequence_start,
                        preferred_template_style=getattr(style, "resolved_properties", {}).get("numbering_style"),
                    )
                    if first_ordered_item:
                        first_ordered_item = False
                    prefix_text = f"{'  ' * item_level}{display_index}. "
                else:
                    item_level = max(getattr(item, "level", 0), 0)
                    numbering_style = getattr(style, "resolved_properties", {}).get("numbering_style")
                    if numbering_style and not getattr(item, "task", False):
                        numbering = NumberingIntent(
                            numbering_role="unordered_list",
                            level=item_level,
                            preferred_template_style=numbering_style,
                        )
                    prefix_text = _list_prefix_text(item, item_kind)
                runs = _build_run_intents(list(item.inlines))
                intents.append(
                    ParagraphLayoutIntent(
                        element_name=element_name,
                        runs=runs,
                        resolved_style_name=style.word_style_name,
                        resolved_properties=dict(getattr(style, "resolved_properties", {})),
                        prefix_text=prefix_text,
                        numbering=numbering,
                    )
                )
    return intents


def _resolve_styles(nodes: list[object], request: RenderRequest) -> dict[str, object]:
    profile = load_style_profile(request.style_path, request.overrides)
    template = load_template_view(request.template_path)
    paragraph_style = resolve_effective_element_style(profile, template, "paragraph")

    styles: dict[str, object] = {"paragraph": paragraph_style}
    for element_name in ("quote", "ordered_list", "unordered_list", "table", "code_block", "image"):
        if element_name in profile.elements:
            styles[element_name] = resolve_effective_element_style(profile, template, element_name)
    heading_levels = {
        node.level for node in nodes if isinstance(node, HeadingNode)
    }
    if heading_levels:
        for level in sorted(heading_levels):
            element_name = f"heading{level}"
            resolved = resolve_effective_element_style(profile, template, element_name)
            if not template.has_template and "font_color" not in resolved.resolved_properties:
                resolved.resolved_properties["font_color"] = "000000"
                resolved.source_map["font_color"] = "default_no_template"
            styles[element_name] = resolved

    required_numbering_elements: set[str] = set()
    for node in nodes:
        if not isinstance(node, ListNode):
            continue
        for item in node.items:
            item_kind = getattr(item, "kind", None) or node.kind
            if item_kind == "ordered":
                required_numbering_elements.add("ordered_list")
            elif item_kind == "unordered" and not getattr(item, "task", False):
                required_numbering_elements.add("unordered_list")
    for element_name in sorted(required_numbering_elements):
        if element_name not in styles:
            continue
        numbering_style = getattr(styles[element_name], "resolved_properties", {}).get("numbering_style")
        if numbering_style and numbering_style not in template.style_numbering_map:
            raise KeyError(f"Missing template numbering resource: {numbering_style}")
    return styles


def _failure_result(
    request: RenderRequest,
    code: str,
    message: str,
) -> RenderResult:
    diagnostic = Diagnostic(
        severity=Severity.ERROR,
        code=code,
        message=message,
    )
    return RenderResult(
        success=False,
        output_path=request.output_path,
        diagnostics=[diagnostic.to_dict()],
    )


def _failure_result_from_key_error(
    request: RenderRequest,
    exc: KeyError,
) -> RenderResult:
    message = exc.args[0] if exc.args else str(exc)
    code = (
        "TEMPLATE_VALIDATION_ERROR"
        if "template style" in message.lower()
        else "STYLE_VALIDATION_ERROR"
    )
    return _failure_result(request, code, message)


def _failure_result_from_os_error(
    request: RenderRequest,
    exc: OSError,
) -> RenderResult:
    if isinstance(exc, FileNotFoundError):
        code = "INPUT_FILE_NOT_FOUND"
    else:
        code = "INPUT_IO_ERROR"
    message = exc.strerror or str(exc)
    if exc.filename:
        message = f"{message}: {exc.filename}"
    return _failure_result(request, code, message)


def _failure_result_from_template_error(
    request: RenderRequest,
    exc: OSError,
) -> RenderResult:
    if isinstance(exc, FileNotFoundError):
        message = f"Template file not found: {request.template_path}"
    else:
        message = exc.strerror or str(exc)
        if exc.filename:
            message = f"{message}: {exc.filename}"
    return _failure_result(request, "TEMPLATE_VALIDATION_ERROR", message)


def _failure_result_from_unicode_error(
    request: RenderRequest,
    exc: UnicodeDecodeError,
    code: str,
    source_label: str,
    source_path: Path | None,
) -> RenderResult:
    message = f"Unable to decode {source_label} as UTF-8: {exc.reason}"
    if source_path is not None:
        message = f"{message}: {source_path}"
    return _failure_result(request, code, message)


def validate_render_styles(
    nodes: list[object],
    request: RenderRequest,
) -> dict[str, object]:
    return _resolve_styles(nodes, request)


def run_render(request: RenderRequest) -> RenderResult:
    try:
        nodes = parse_markdown_file(request.input_path)
    except UnicodeDecodeError as exc:
        return _failure_result_from_unicode_error(
            request,
            exc,
            "INPUT_ENCODING_ERROR",
            "markdown input",
            request.input_path,
        )
    except OSError as exc:
        return _failure_result_from_os_error(request, exc)

    try:
        styles = validate_render_styles(nodes, request)
        layout_intents = build_layout_intents(nodes, styles)
    except ValidationError as exc:
        return _failure_result(request, "STYLE_VALIDATION_ERROR", str(exc))
    except TypeError as exc:
        return _failure_result(request, "STYLE_VALIDATION_ERROR", str(exc))
    except UnicodeDecodeError as exc:
        return _failure_result_from_unicode_error(
            request,
            exc,
            "STYLE_VALIDATION_ERROR",
            "style file",
            request.style_path,
        )
    except KeyError as exc:
        return _failure_result_from_key_error(request, exc)
    except OSError as exc:
        if request.template_path is not None and exc.filename == str(request.template_path):
            return _failure_result_from_template_error(request, exc)
        return _failure_result_from_os_error(request, exc)

    try:
        document = render_nodes_to_document(
            nodes,
            styles,
            base_dir=request.input_path.parent,
            template_path=request.template_path,
            layout_intents=layout_intents,
        )
        request.output_path.parent.mkdir(parents=True, exist_ok=True)
        document.save(request.output_path)
    except OSError as exc:
        return _failure_result_from_os_error(request, exc)
    return RenderResult(success=True, output_path=request.output_path)
