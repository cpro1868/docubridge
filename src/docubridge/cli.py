from __future__ import annotations

from importlib import resources
import json
from pathlib import Path

import typer

from docubridge.adapters.yaml_adapter import write_yaml
from docubridge.application.models import RenderRequest
from docubridge.application.parse_service import run_parse
from docubridge.application.render_service import build_layout_intents, run_render, validate_render_styles
from docubridge.core.doctor_checks import scan_markdown_warnings
from docubridge.core.markdown_ingest import parse_markdown_file
from docubridge.core.style_extractor import extract_style_profile
from docubridge.core.style_resolver import resolve_effective_element_style
from docubridge.core.style_schema import load_style_profile
from docubridge.core.template_bridge import TemplateView, load_template_view


app = typer.Typer(no_args_is_help=True)
style_app = typer.Typer(
    no_args_is_help=True,
    help="Inspect and validate style profiles.",
)

app.add_typer(style_app, name="style")


def _emit_doctor_json(
    *,
    success: bool,
    summary: str,
    checks: list[dict[str, str]],
    warnings: list[str],
    error: dict[str, str] | None = None,
) -> None:
    payload: dict[str, object] = {
        "success": success,
        "summary": summary,
        "checks": checks,
        "warnings": warnings,
    }
    if error is not None:
        payload["error"] = error
    typer.echo(json.dumps(payload, ensure_ascii=False))


def _display_path(path: Path) -> str:
    return path.as_posix()


def _emit_render_json(*, success: bool, output_path: Path, diagnostics: list[dict]) -> None:
    typer.echo(
        json.dumps(
            {
                "success": success,
                "output_path": _display_path(output_path),
                "diagnostics": diagnostics,
            },
            ensure_ascii=False,
        )
    )


def _emit_parse_json(*, success: bool, output_path: Path, diagnostics: list[dict]) -> None:
    _emit_json(
        {
            "success": success,
            "output_path": _display_path(output_path),
            "diagnostics": diagnostics,
        }
    )


def _emit_json(payload: object, *, pretty: bool = False) -> None:
    typer.echo(
        json.dumps(
            payload,
            ensure_ascii=False,
            indent=2 if pretty else None,
        )
    )


def _numbering_diagnostics(
    *,
    resolved_properties: dict[str, object],
    source_map: dict[str, str],
    template: TemplateView,
) -> dict[str, object] | None:
    requested_style = resolved_properties.get("numbering_style")
    if not isinstance(requested_style, str) or not requested_style:
        return None
    available_in_template = requested_style in template.style_numbering_map
    return {
        "requested_style": requested_style,
        "source": source_map.get("numbering_style", "unknown"),
        "available_in_template": available_in_template,
        "fallback_mode": "native" if available_in_template else "text-prefix",
    }


def _style_resolution_detail(nodes: list[object], styles: dict[str, object], template: TemplateView) -> str | None:
    layout_intents = build_layout_intents(nodes, styles)
    ordered_intents = [intent for intent in layout_intents if intent.numbering is not None]
    if not ordered_intents:
        return None
    numbering_style = ordered_intents[0].numbering.preferred_template_style
    if not numbering_style:
        return "ordered lists: no explicit numbering style"
    if numbering_style in template.style_numbering_map:
        return f"native numbering: {numbering_style}"
    return f"text-prefix fallback: {numbering_style}"


def _builtin_style_names() -> list[str]:
    package = resources.files("docubridge").joinpath("builtin_styles")
    return sorted(
        entry.stem
        for entry in package.iterdir()
        if entry.is_file() and entry.suffix == ".yaml"
    )


def _builtin_style_path(name: str):
    package = resources.files("docubridge").joinpath("builtin_styles")
    path = package.joinpath(f"{name}.yaml")
    if not path.is_file():
        raise typer.BadParameter(f"Unknown builtin style: {name}")
    return path


def _parse_override_args(values: list[str] | None) -> dict[str, str]:
    overrides: dict[str, str] = {}
    for value in values or []:
        key, sep, raw = value.partition("=")
        if not sep or not key:
            raise typer.BadParameter(f"Invalid --set value: {value}")
        overrides[key] = raw
    return overrides


@app.command()
def parse(
    input_path: Path,
    output_path: Path = typer.Option(..., "-o", "--output", help="Write Markdown output to this path."),
    json_output: bool = typer.Option(False, "--json", help="Emit a structured JSON result instead of text."),
) -> None:
    """Parse supported documents to Markdown. Currently supports .docx, .xlsx, and .pptx input."""
    result = run_parse(input_path, output_path)
    if json_output:
        _emit_parse_json(
            success=result.success,
            output_path=result.output_path,
            diagnostics=result.diagnostics,
        )
    if not result.success:
        diagnostic = result.diagnostics[0] if result.diagnostics else None
        if diagnostic is not None and not json_output:
            typer.echo(diagnostic["message"], err=True)
        raise typer.Exit(code=5)


@app.command()
def render(
    input_path: Path,
    output_path: Path = typer.Option(..., "-o", "--output"),
    style_path: Path = typer.Option(
        ...,
        "--style",
        help="Use a YAML style profile for explicit formatting rules.",
    ),
    template_path: Path | None = typer.Option(
        None,
        "--template",
        help="Use a .docx template as the host document. YAML still defines explicit style intent.",
    ),
    json_output: bool = typer.Option(False, "--json"),
) -> None:
    """Render Markdown to .docx."""
    request = RenderRequest(
        input_path=input_path,
        output_path=output_path,
        style_path=style_path,
        template_path=template_path,
    )
    result = run_render(request)
    if json_output:
        _emit_render_json(
            success=result.success,
            output_path=result.output_path,
            diagnostics=result.diagnostics,
        )
    if not result.success:
        diagnostic = result.diagnostics[0] if result.diagnostics else None
        if diagnostic is not None:
            if not json_output:
                typer.echo(diagnostic["message"], err=True)
            code = diagnostic.get("code", "")
            raise typer.Exit(code=4 if code.startswith(("STYLE_", "TEMPLATE_")) else 5)
        raise typer.Exit(code=5)


@app.command()
def extract_styles(
    input_path: Path,
    output_path: Path = typer.Option(..., "-o", "--output"),
    pretty: bool = typer.Option(False, "--pretty", help="Write indented, human-readable YAML."),
    strict: bool = typer.Option(False, "--strict", help="Fail if any Word style cannot be mapped to a known element."),
    json_output: bool = typer.Option(False, "--json", help="Emit a structured JSON result instead of text."),
) -> None:
    """Extract Word styles from a .docx into a docubridge style profile YAML."""
    try:
        profile, unmapped = extract_style_profile(input_path)
    except FileNotFoundError as exc:
        if json_output:
            _emit_json({"success": False, "output_path": _display_path(output_path), "diagnostics": [{"code": "INPUT_FILE_NOT_FOUND", "message": str(exc)}]})
        else:
            typer.echo(str(exc), err=True)
        raise typer.Exit(code=5)
    except OSError as exc:
        if json_output:
            _emit_json({"success": False, "output_path": _display_path(output_path), "diagnostics": [{"code": "INPUT_IO_ERROR", "message": str(exc)}]})
        else:
            typer.echo(str(exc), err=True)
        raise typer.Exit(code=5)

    if strict and unmapped:
        message = f"Unmapped styles: {', '.join(unmapped)}"
        if json_output:
            _emit_json({"success": False, "output_path": _display_path(output_path), "diagnostics": [{"code": "STYLE_EXTRACTION_ERROR", "message": message}]})
        else:
            typer.echo(message, err=True)
        raise typer.Exit(code=4)

    try:
        write_yaml(output_path, profile.model_dump(by_alias=True, exclude_defaults=False), pretty=pretty)
    except OSError as exc:
        if json_output:
            _emit_json({"success": False, "output_path": _display_path(output_path), "diagnostics": [{"code": "OUTPUT_IO_ERROR", "message": str(exc)}]})
        else:
            typer.echo(str(exc), err=True)
        raise typer.Exit(code=5)

    diagnostics: list[dict] = []
    if unmapped:
        template = load_template_view(input_path)
        diagnostics.append({"code": "UNMAPPED_STYLES", "message": f"extracted {len(template.available_styles)} styles, {len(unmapped)} unmapped styles saved to compat.extracted_styles"})

    if json_output:
        _emit_json({"success": True, "output_path": _display_path(output_path), "diagnostics": diagnostics})
        return

    if unmapped:
        template = load_template_view(input_path)
        typer.echo(f"extracted {len(template.available_styles)} styles, {len(unmapped)} unmapped styles saved to compat.extracted_styles")
    typer.echo(f"Extracted styles written to {_display_path(output_path)}")


@style_app.command("list")
def style_list() -> None:
    for name in _builtin_style_names():
        typer.echo(name)


@style_app.command("show")
def style_show(
    name: str,
    json_output: bool = typer.Option(False, "--json"),
) -> None:
    path = _builtin_style_path(name)
    content = path.read_text(encoding="utf-8")
    if json_output:
        _emit_json(
            {
                "name": name,
                "path": path.name,
                "content": content,
            }
        )
        return
    typer.echo(content)


@style_app.command("validate")
def style_validate(
    style_path: Path,
    json_output: bool = typer.Option(False, "--json"),
) -> None:
    try:
        load_style_profile(style_path)
    except (FileNotFoundError, OSError, TypeError, ValueError, KeyError) as exc:
        if json_output:
            _emit_json(
                {
                    "success": False,
                    "style_path": _display_path(style_path),
                    "message": str(exc),
                }
            )
        else:
            typer.echo(str(exc), err=True)
        raise typer.Exit(code=4)
    if json_output:
        _emit_json(
            {
                "success": True,
                "style_path": _display_path(style_path),
                "message": "Style OK",
            }
        )
        return
    typer.echo("Style OK")


@style_app.command("explain")
def style_explain(
    style_path: Path,
    element_name: str,
    template_path: Path | None = typer.Option(
        None,
        "--template",
        help="Resolve styles against a .docx template when inspecting template-backed profiles.",
    ),
    pretty: bool = typer.Option(False, "--pretty"),
) -> None:
    try:
        profile = load_style_profile(style_path)
        template = load_template_view(template_path)
        resolved = resolve_effective_element_style(profile, template, element_name)
    except (FileNotFoundError, OSError, TypeError, ValueError, KeyError) as exc:
        typer.echo(str(exc), err=True)
        raise typer.Exit(code=4)

    payload = {
        "element_name": resolved.element_name,
        "word_style_name": resolved.word_style_name,
        "resolved_properties": resolved.resolved_properties,
        "source_map": resolved.source_map,
    }
    numbering = _numbering_diagnostics(
        resolved_properties=resolved.resolved_properties,
        source_map=resolved.source_map,
        template=template,
    )
    if numbering is not None:
        payload["numbering"] = numbering

    _emit_json(payload, pretty=pretty)


@style_app.command("merge")
def style_merge(
    style_path: Path,
    overrides: list[str] = typer.Option(None, "--set"),
    pretty: bool = typer.Option(False, "--pretty"),
) -> None:
    try:
        profile = load_style_profile(style_path, _parse_override_args(overrides))
    except (FileNotFoundError, OSError, TypeError, ValueError, KeyError) as exc:
        typer.echo(str(exc), err=True)
        raise typer.Exit(code=4)

    _emit_json(profile.model_dump(mode="python"), pretty=pretty)


@app.command()
def doctor(
    input_path: Path | None = typer.Argument(None),
    style: Path | None = typer.Option(None, "--style"),
    template: Path | None = typer.Option(
        None,
        "--template",
        help="Resolve styles against a .docx template when validating render readiness.",
    ),
    json_output: bool = typer.Option(False, "--json"),
) -> None:
    """Run environment and task-level preflight checks."""
    checks: list[dict[str, str]] = [{"name": "environment", "status": "ok"}]
    if not json_output:
        typer.echo("environment: ok")
    if input_path is None and style is None:
        if json_output:
            _emit_doctor_json(
                success=True,
                summary="Environment OK",
                checks=checks,
                warnings=[],
            )
        else:
            typer.echo("Environment OK")
        return

    nodes = None
    markdown_warnings: list[str] = []
    if input_path is not None:
        try:
            nodes = parse_markdown_file(input_path)
            markdown_warnings = scan_markdown_warnings(input_path)
        except (FileNotFoundError, OSError, UnicodeDecodeError) as exc:
            checks.append({"name": "markdown", "status": "error"})
            if json_output:
                _emit_doctor_json(
                    success=False,
                    summary="Doctor failed",
                    checks=checks,
                    warnings=[],
                    error={"check": "markdown", "message": str(exc)},
                )
            else:
                typer.echo(f"markdown: {exc}", err=True)
            raise typer.Exit(code=5)
        checks.append({"name": "markdown", "status": "ok", "detail": _display_path(input_path)})
        if not json_output:
            typer.echo(f"markdown: ok ({input_path})")
    if style is not None:
        try:
            load_style_profile(style)
        except (FileNotFoundError, OSError, UnicodeDecodeError) as exc:
            checks.append({"name": "style", "status": "error"})
            if json_output:
                _emit_doctor_json(
                    success=False,
                    summary="Doctor failed",
                    checks=checks,
                    warnings=[],
                    error={"check": "style", "message": str(exc)},
                )
            else:
                typer.echo(f"style: {exc}", err=True)
            raise typer.Exit(code=5)
        except (TypeError, ValueError, KeyError) as exc:
            checks.append({"name": "style", "status": "error"})
            if json_output:
                _emit_doctor_json(
                    success=False,
                    summary="Doctor failed",
                    checks=checks,
                    warnings=[],
                    error={"check": "style", "message": str(exc)},
                )
            else:
                typer.echo(f"style: {exc}", err=True)
            raise typer.Exit(code=4)
        checks.append({"name": "style", "status": "ok", "detail": _display_path(style)})
        if not json_output:
            typer.echo(f"style: ok ({style})")
        if nodes is not None:
            try:
                request = RenderRequest(
                    input_path=input_path,
                    output_path=Path("doctor-preview.docx"),
                    style_path=style,
                    template_path=template,
                )
                styles = validate_render_styles(nodes, request)
            except (OSError, TypeError, ValueError, KeyError) as exc:
                checks.append({"name": "style-resolution", "status": "error"})
                if json_output:
                    _emit_doctor_json(
                        success=False,
                        summary="Doctor failed",
                        checks=checks,
                        warnings=[],
                        error={"check": "style-resolution", "message": str(exc)},
                    )
                else:
                    typer.echo(f"style-resolution: {exc}", err=True)
                raise typer.Exit(code=4)
            detail = _style_resolution_detail(nodes, styles, load_template_view(template))
            style_resolution_check: dict[str, str] = {"name": "style-resolution", "status": "ok"}
            if detail is not None:
                style_resolution_check["detail"] = detail
            checks.append(style_resolution_check)
            if not json_output:
                if detail is not None:
                    typer.echo(f"style-resolution: ok ({detail})")
                else:
                    typer.echo("style-resolution: ok")

    if json_output:
        _emit_doctor_json(
            success=True,
            summary="Task OK",
            checks=checks,
            warnings=markdown_warnings,
        )
        return

    if markdown_warnings:
        typer.echo(f"warnings: {len(markdown_warnings)}")
        for warning in markdown_warnings:
            typer.echo(f"warning: {warning}")
    typer.echo("Task OK")
