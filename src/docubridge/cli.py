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


from docubridge import __version__


def _version_callback(value: bool) -> None:
    if value:
        typer.echo(f"docubridge {__version__}")
        raise typer.Exit()


app = typer.Typer(
    no_args_is_help=True,
    rich_markup_mode="rich",
    help="""docubridge: bidirectional Markdown <-> Office document toolkit.

Examples:
  docubridge render draft.md -o out.docx --style default
  docubridge render draft.md -o out.docx --style style.yaml --template corp.docx
  docubridge parse report.docx -o report.md
  docubridge extract-styles corp.docx -o corp-style.yaml
  docubridge doctor draft.md --style style.yaml --template corp.docx
""",
    add_completion=False,
)
style_app = typer.Typer(
    no_args_is_help=True,
    rich_markup_mode="rich",
    help="Inspect, validate and preview style profiles.",
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


@app.callback()
def _main(
    version: bool = typer.Option(
        False,
        "--version",
        "-v",
        help="Show the version and exit.",
        callback=_version_callback,
        is_eager=True,
    ),
) -> None:
    """Common entry-point options."""
    pass


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


@app.command(help="""Parse .docx / .xlsx / .pptx into Markdown.

Examples:
  docubridge parse report.docx -o report.md
  docubridge parse slides.pptx -o slides.md
  docubridge parse data.xlsx -o sheet.md --json
""")
def parse(
    input_path: Path = typer.Argument(..., help="Source document path (.docx, .xlsx or .pptx)."),
    output_path: Path = typer.Option(..., "-o", "--output", help="Write Markdown output to this path."),
    json_output: bool = typer.Option(False, "--json", help="Emit a structured JSON result instead of text."),
) -> None:
    """Convert .docx / .xlsx / .pptx files to Markdown. Images and other assets are exported to an 'assets/' folder next to the output file."""
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


@app.command(help="""Render Markdown into a styled .docx.

Examples:
  docubridge render input.md -o out.docx --style style.yaml
  docubridge render input.md -o out.docx --style default
  docubridge render input.md -o out.docx --style style.yaml --template corp.docx
  docubridge render input.md -o out.docx --style $(docubridge style show default --path)
""")
def render(
    input_path: Path = typer.Argument(..., help="Markdown source file."),
    output_path: Path = typer.Option(..., "-o", "--output", help="Write the resulting .docx to this path."),
    style_path: Path = typer.Option(
        ...,
        "--style",
        help="YAML style profile (builtin names like 'default' are accepted and resolved automatically).",
    ),
    template_path: Path | None = typer.Option(
        None,
        "--template",
        help="Optional .docx template used as the host document. YAML values override template values.",
    ),
    json_output: bool = typer.Option(False, "--json", help="Emit structured JSON result instead of text."),
) -> None:
    """Render Markdown to a styled .docx."""
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


@app.command(help="""Extract Word styles from a .docx into a YAML profile.

Examples:
  docubridge extract-styles contract.docx -o contract-style.yaml
  docubridge extract-styles contract.docx -o contract-style.yaml --pretty
  docubridge extract-styles contract.docx -o contract-style.yaml --strict

The generated YAML can be passed to 'render --style'.
Unmapped styles are kept under 'compat.extracted_styles' unless --strict is used.
""")
def extract_styles(
    input_path: Path = typer.Argument(..., help="Source .docx to extract styles from."),
    output_path: Path = typer.Option(..., "-o", "--output", help="Write the extracted style YAML to this path."),
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


@style_app.command("list", help="""List built-in style profile names.

Examples:
  docubridge style list
""")
def style_list() -> None:
    """List built-in style profile names."""
    for name in _builtin_style_names():
        typer.echo(name)


@style_app.command("show", help="""Print a built-in style profile.

Examples:
  docubridge style show default
  docubridge style show academic --json
""")
def style_show(
    name: str = typer.Argument(..., help="Built-in profile name (e.g. default, academic, business)."),
    json_output: bool = typer.Option(False, "--json", help="Output structured JSON instead of raw YAML."),
) -> None:
    """Print a built-in style profile to stdout."""
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


@style_app.command("validate", help="""Validate a style profile YAML file.

Examples:
  docubridge style validate style.yaml
  docubridge style validate style.yaml --json
""")
def style_validate(
    style_path: Path = typer.Argument(..., help="Path to the YAML style profile to validate."),
    json_output: bool = typer.Option(False, "--json", help="Output structured JSON instead of text."),
) -> None:
    """Validate a style profile YAML file."""
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


@style_app.command("explain", help="""Show the resolved style for a single element.

Examples:
  docubridge style explain style.yaml heading1
  docubridge style explain style.yaml paragraph --pretty
  docubridge style explain style.yaml ordered_list --template corp.docx --pretty
""")
def style_explain(
    style_path: Path = typer.Argument(..., help="Path to the YAML style profile."),
    element_name: str = typer.Argument(..., help="Element to inspect (e.g. heading1, paragraph, ordered_list)."),
    template_path: Path | None = typer.Option(
        None,
        "--template",
        help="Resolve styles against a .docx template when inspecting template-backed profiles.",
    ),
    pretty: bool = typer.Option(False, "--pretty", help="Pretty-print the JSON output."),
) -> None:
    """Show the resolved style for a single element. Use --template to see how the profile interacts with a host Word file."""
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


@style_app.command("merge", help="""Merge a style profile with optional overrides and print it.

Examples:
  docubridge style merge style.yaml --pretty
  docubridge style merge style.yaml --set heading1.font_size=20 --pretty
  docubridge style merge style.yaml --set document.toc.depth=4 --pretty
""")
def style_merge(
    style_path: Path = typer.Argument(..., help="Path to the YAML style profile."),
    overrides: list[str] = typer.Option(None, "--set", help="Override values using dot-path notation, e.g. --set heading1.font_size=20"),
    pretty: bool = typer.Option(False, "--pretty", help="Pretty-print the JSON output."),
) -> None:
    """Merge a style profile with optional overrides and print the result."""
    try:
        profile = load_style_profile(style_path, _parse_override_args(overrides))
    except (FileNotFoundError, OSError, TypeError, ValueError, KeyError) as exc:
        typer.echo(str(exc), err=True)
        raise typer.Exit(code=4)

    _emit_json(profile.model_dump(mode="python"), pretty=pretty)


@app.command(help="""Run environment and task-level preflight checks.

Examples:
  docubridge doctor
  docubridge doctor input.md --style style.yaml
  docubridge doctor input.md --style style.yaml --template corp.docx
""")
def doctor(
    input_path: Path | None = typer.Argument(None, help="Optional Markdown source file to validate."),
    style: Path | None = typer.Option(None, "--style", help="Optional YAML style profile to validate against."),
    template: Path | None = typer.Option(
        None,
        "--template",
        help="Resolve styles against a .docx template when validating render readiness.",
    ),
    json_output: bool = typer.Option(False, "--json", help="Emit structured JSON result instead of text."),
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
