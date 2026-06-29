# DocuBridge Render Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build the v1 `Markdown -> Word` rendering pipeline for DocuBridge with YAML-driven styles, template support, strict/lenient diagnostics, and a production-grade CLI.

**Architecture:** The implementation uses a Python package with clear boundaries between core parsing/rendering logic, application orchestration, CLI interfaces, and third-party adapters. Markdown is parsed into lightweight block/inline nodes, styles are resolved from YAML plus template resources, and a Word renderer writes `.docx` output with shared diagnostics and testable execution flow.

**Tech Stack:** Python 3.12+, `typer`, `pydantic`, `ruamel.yaml`, `markdown-it-py`, `python-docx`, `pytest`

---

## Current Status Snapshot (2026-04-09)

The original 10-task MVP plan has been executed for the CLI-first rendering path, with follow-up hardening on top of it. The current implemented surface area is:

- package scaffold and diagnostics models
- lightweight Markdown ingest
- minimal docx-to-markdown parsing
- docx-to-markdown rich text preservation for headings, paragraphs, and table cells, including basic bold, italic, strikethrough, hyperlinks, and simple inline code
- docx-to-markdown simple ordered-list progression restoration based on paragraph order and indentation
- minimal xlsx-to-markdown parsing
- minimal pptx-to-markdown parsing, including tables, embedded image export, speaker notes extraction, position-ordered block output, and basic text hierarchy recovery
- pptx-to-markdown basic rich-text preservation for textbox and table-cell runs, including bold, italic, hyperlinks, and simple inline code
- YAML style loading and override validation
- template view and style resolution
- basic Word rendering for headings, paragraphs, inline bold/italic/strikethrough/code/link spans, blockquotes with inline formatting, horizontal rules, lists, inline list-item and task-list-item formatting, simple nested list indentation, Markdown tables with inline cell formatting, fenced code blocks, and standalone image blocks
- application render orchestration
- CLI `parse`, `render`, `doctor`, and `style` subcommands
- structured outputs for `render --json`, `doctor --json`, `style show --json`, `style validate --json`
- structured output for `parse --json`
- pretty-printed JSON for `style explain --pretty` and `style merge --pretty`
- warning-oriented `doctor` checks for degraded Markdown structures
- implicit heading style mapping where `headingN` resolves to Word `Heading N` when `template_style` is omitted

Still outside the implemented scope:

- multi-format forward conversion beyond minimal `docx/xlsx/pptx -> markdown`
- `batch` command
- external template file loading during `render`
- advanced Word features such as numbering manager, TOC insertion, pagination, tables, and images as first-class render nodes

This plan remains useful as the architectural baseline, but it should now be read as:

- Tasks 1-10: completed for the current MVP slice
- post-plan CLI output and diagnostics hardening: also completed
- remaining roadmap: move to `parse`, `batch`, template integration, and richer render nodes

---

## Planned File Structure

### Create

- `pyproject.toml`
- `README.md`
- `src/docubridge/__init__.py`
- `src/docubridge/cli.py`
- `src/docubridge/application/render_service.py`
- `src/docubridge/application/models.py`
- `src/docubridge/core/diagnostics.py`
- `src/docubridge/core/nodes.py`
- `src/docubridge/core/markdown_ingest.py`
- `src/docubridge/core/style_schema.py`
- `src/docubridge/core/style_resolver.py`
- `src/docubridge/core/template_bridge.py`
- `src/docubridge/core/numbering_manager.py`
- `src/docubridge/core/asset_resolver.py`
- `src/docubridge/core/word_renderer.py`
- `src/docubridge/adapters/docx_adapter.py`
- `src/docubridge/adapters/filesystem.py`
- `src/docubridge/adapters/markdown_adapter.py`
- `src/docubridge/adapters/yaml_adapter.py`
- `src/docubridge/builtin_styles/academic.yaml`
- `src/docubridge/builtin_styles/business.yaml`
- `tests/conftest.py`
- `tests/test_cli_render.py`
- `tests/test_diagnostics.py`
- `tests/test_markdown_ingest.py`
- `tests/test_style_schema.py`
- `tests/test_style_resolver.py`
- `tests/test_template_bridge.py`
- `tests/test_word_renderer.py`
- `tests/fixtures/sample.md`
- `tests/fixtures/style.yaml`

### Modify

- `docs/requirements.md`
- `docs/superpowers/specs/2026-04-07-docubridge-render-design.md`

## Task 1: Initialize the Python Project Skeleton

**Files:**
- Create: `pyproject.toml`
- Create: `README.md`
- Create: `src/docubridge/__init__.py`
- Create: `tests/conftest.py`

- [ ] **Step 1: Write the failing smoke test for the package version**

```python
# tests/conftest.py
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))
```

```python
# tests/test_diagnostics.py
from docubridge import __version__


def test_package_version_is_exposed() -> None:
    assert __version__ == "0.1.0"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `pytest tests/test_diagnostics.py::test_package_version_is_exposed -v`  
Expected: FAIL with `ModuleNotFoundError: No module named 'docubridge'`

- [ ] **Step 3: Write the minimal package scaffold**

```toml
# pyproject.toml
[build-system]
requires = ["setuptools>=69", "wheel"]
build-backend = "setuptools.build_meta"

[project]
name = "docubridge"
version = "0.1.0"
description = "Markdown to Word rendering tool with YAML-driven styles"
readme = "README.md"
requires-python = ">=3.12"
dependencies = [
  "typer>=0.12",
  "pydantic>=2.7",
  "ruamel.yaml>=0.18",
  "markdown-it-py>=3.0",
  "python-docx>=1.1",
]

[project.optional-dependencies]
dev = ["pytest>=8.2"]

[project.scripts]
docubridge = "docubridge.cli:app"

[tool.pytest.ini_options]
testpaths = ["tests"]
```

```markdown
# README.md

DocuBridge is a local-first document conversion tool.

Current focus:

- Markdown to Word rendering
- YAML-driven style profiles
- Template-aware output
```

```python
# src/docubridge/__init__.py
__all__ = ["__version__"]

__version__ = "0.1.0"
```

- [ ] **Step 4: Run test to verify it passes**

Run: `pytest tests/test_diagnostics.py::test_package_version_is_exposed -v`  
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add pyproject.toml README.md src/docubridge/__init__.py tests/conftest.py tests/test_diagnostics.py
git commit -m "chore: bootstrap docubridge python package"
```

## Task 2: Define Diagnostics and Application Models

**Files:**
- Create: `src/docubridge/core/diagnostics.py`
- Create: `src/docubridge/application/models.py`
- Modify: `tests/test_diagnostics.py`

- [ ] **Step 1: Write the failing tests for diagnostic and request models**

```python
# tests/test_diagnostics.py
from pathlib import Path

from docubridge.application.models import RenderRequest
from docubridge.core.diagnostics import Diagnostic, Severity


def test_diagnostic_to_dict_contains_expected_fields() -> None:
    diagnostic = Diagnostic(
        severity=Severity.WARNING,
        code="STYLE_UNKNOWN_FIELD",
        message="Unknown style field",
        location="style.yaml:12",
        hint="Remove the field or rename it",
    )

    assert diagnostic.to_dict() == {
        "severity": "warning",
        "code": "STYLE_UNKNOWN_FIELD",
        "message": "Unknown style field",
        "location": "style.yaml:12",
        "hint": "Remove the field or rename it",
    }


def test_render_request_normalizes_paths() -> None:
    request = RenderRequest(
        input_path=Path("tests/fixtures/sample.md"),
        output_path=Path("build/output.docx"),
        style_path=Path("tests/fixtures/style.yaml"),
        mode="strict",
        output_mode="human",
        overwrite=False,
    )

    assert request.mode == "strict"
    assert request.output_mode == "human"
    assert request.input_path.as_posix().endswith("tests/fixtures/sample.md")
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `pytest tests/test_diagnostics.py -v`  
Expected: FAIL with `ImportError` for `docubridge.application.models` or `docubridge.core.diagnostics`

- [ ] **Step 3: Implement diagnostics and request models**

```python
# src/docubridge/core/diagnostics.py
from dataclasses import dataclass
from enum import StrEnum


class Severity(StrEnum):
    INFO = "info"
    WARNING = "warning"
    ERROR = "error"
    FATAL = "fatal"


@dataclass(slots=True)
class Diagnostic:
    severity: Severity
    code: str
    message: str
    location: str | None = None
    hint: str | None = None

    def to_dict(self) -> dict[str, str | None]:
        return {
            "severity": self.severity.value,
            "code": self.code,
            "message": self.message,
            "location": self.location,
            "hint": self.hint,
        }
```

```python
# src/docubridge/application/models.py
from pathlib import Path
from typing import Literal

from pydantic import BaseModel, ConfigDict, Field


class RenderRequest(BaseModel):
    model_config = ConfigDict(arbitrary_types_allowed=True)

    input_path: Path
    output_path: Path
    style_path: Path | None = None
    template_path: Path | None = None
    profile_name: str | None = None
    mode: Literal["strict", "lenient"] = "strict"
    output_mode: Literal["human", "quiet", "json"] = "human"
    overwrite: bool = False
    resource_dir: Path | None = None
    dump_ast: bool = False
    features: list[str] = Field(default_factory=list)
    overrides: dict[str, str] = Field(default_factory=dict)
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `pytest tests/test_diagnostics.py -v`  
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/docubridge/core/diagnostics.py src/docubridge/application/models.py tests/test_diagnostics.py
git commit -m "feat: add diagnostic and render request models"
```

## Task 3: Implement the Markdown Node Model and Parser

**Files:**
- Create: `src/docubridge/core/nodes.py`
- Create: `src/docubridge/adapters/markdown_adapter.py`
- Create: `src/docubridge/core/markdown_ingest.py`
- Create: `tests/test_markdown_ingest.py`
- Create: `tests/fixtures/sample.md`

- [ ] **Step 1: Write the failing tests for block and inline parsing**

```python
# tests/test_markdown_ingest.py
from pathlib import Path

from docubridge.core.markdown_ingest import parse_markdown_file


def test_parse_markdown_file_returns_heading_and_paragraph_nodes() -> None:
    nodes = parse_markdown_file(Path("tests/fixtures/sample.md"))

    assert nodes[0].type == "heading"
    assert nodes[0].level == 1
    assert nodes[1].type == "paragraph"
    assert nodes[1].inlines[0].text == "Plain paragraph."


def test_task_list_items_capture_checked_state() -> None:
    nodes = parse_markdown_file(Path("tests/fixtures/sample.md"))
    task_list = next(node for node in nodes if node.type == "list" and node.kind == "unordered")

    assert task_list.items[0].task is True
    assert task_list.items[0].checked is True
```

```markdown
# tests/fixtures/sample.md
# Sample Title

Plain paragraph.

- [x] done
- [ ] pending
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `pytest tests/test_markdown_ingest.py -v`  
Expected: FAIL with `ImportError` for `docubridge.core.markdown_ingest`

- [ ] **Step 3: Implement lightweight node classes and parser**

```python
# src/docubridge/core/nodes.py
from dataclasses import dataclass, field


@dataclass(slots=True)
class TextSpan:
    text: str
    type: str = "text"


@dataclass(slots=True)
class ParagraphNode:
    inlines: list[TextSpan]
    type: str = "paragraph"


@dataclass(slots=True)
class HeadingNode:
    level: int
    inlines: list[TextSpan]
    type: str = "heading"


@dataclass(slots=True)
class ListItemNode:
    inlines: list[TextSpan]
    task: bool = False
    checked: bool | None = None


@dataclass(slots=True)
class ListNode:
    kind: str
    items: list[ListItemNode] = field(default_factory=list)
    type: str = "list"
```

```python
# src/docubridge/adapters/markdown_adapter.py
from pathlib import Path

from markdown_it import MarkdownIt


def load_markdown_tokens(path: Path):
    parser = MarkdownIt("gfm-like").enable("table")
    return parser.parse(path.read_text(encoding="utf-8"))
```

```python
# src/docubridge/core/markdown_ingest.py
from pathlib import Path

from docubridge.adapters.markdown_adapter import load_markdown_tokens
from docubridge.core.nodes import HeadingNode, ListItemNode, ListNode, ParagraphNode, TextSpan


def parse_markdown_file(path: Path):
    tokens = load_markdown_tokens(path)
    nodes = []
    i = 0
    while i < len(tokens):
        token = tokens[i]
        if token.type == "heading_open":
            inline = tokens[i + 1]
            nodes.append(HeadingNode(level=int(token.tag[1]), inlines=[TextSpan(inline.content)]))
            i += 3
            continue
        if token.type == "paragraph_open":
            inline = tokens[i + 1]
            nodes.append(ParagraphNode(inlines=[TextSpan(inline.content)]))
            i += 3
            continue
        if token.type == "bullet_list_open":
            items = []
            i += 1
            while tokens[i].type != "bullet_list_close":
                if tokens[i].type == "list_item_open":
                    inline = tokens[i + 2]
                    content = inline.content
                    task = content.startswith("[x] ") or content.startswith("[ ] ")
                    checked = True if content.startswith("[x] ") else False if content.startswith("[ ] ") else None
                    text = content[4:] if task else content
                    items.append(ListItemNode(inlines=[TextSpan(text)], task=task, checked=checked))
                    i += 5
                    continue
                i += 1
            nodes.append(ListNode(kind="unordered", items=items))
        i += 1
    return nodes
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `pytest tests/test_markdown_ingest.py -v`  
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/docubridge/core/nodes.py src/docubridge/adapters/markdown_adapter.py src/docubridge/core/markdown_ingest.py tests/test_markdown_ingest.py tests/fixtures/sample.md
git commit -m "feat: parse markdown into lightweight render nodes"
```

## Task 4: Implement Style Schema, YAML Loading, and Override Handling

**Files:**
- Create: `src/docubridge/adapters/yaml_adapter.py`
- Create: `src/docubridge/core/style_schema.py`
- Create: `tests/test_style_schema.py`
- Create: `tests/fixtures/style.yaml`

- [ ] **Step 1: Write the failing tests for style loading and overrides**

```python
# tests/test_style_schema.py
from pathlib import Path

from docubridge.core.style_schema import load_style_profile


def test_load_style_profile_reads_elements_and_defaults() -> None:
    profile = load_style_profile(Path("tests/fixtures/style.yaml"))

    assert profile.defaults["font_name"] == "Times New Roman"
    assert profile.elements["heading1"]["font_size"] == 18


def test_load_style_profile_applies_override_values() -> None:
    profile = load_style_profile(
        Path("tests/fixtures/style.yaml"),
        overrides={"document.toc.depth": "4"},
    )

    assert profile.document["toc"]["depth"] == 4
```

```yaml
# tests/fixtures/style.yaml
meta:
  name: academic
  version: 1

defaults:
  font_name: Times New Roman
  font_size: 12

elements:
  heading1:
    based_on: Normal
    font_size: 18
    bold: true
  paragraph:
    based_on: Normal
    font_size: 12

document:
  toc:
    enabled: true
    depth: 3
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `pytest tests/test_style_schema.py -v`  
Expected: FAIL with `ImportError` for `docubridge.core.style_schema`

- [ ] **Step 3: Implement YAML loading, schema model, and override application**

```python
# src/docubridge/adapters/yaml_adapter.py
from pathlib import Path

from ruamel.yaml import YAML


def load_yaml(path: Path) -> dict:
    yaml = YAML(typ="safe")
    with path.open("r", encoding="utf-8") as handle:
        data = yaml.load(handle) or {}
    return data
```

```python
# src/docubridge/core/style_schema.py
from pathlib import Path

from pydantic import BaseModel, Field

from docubridge.adapters.yaml_adapter import load_yaml


class StyleProfile(BaseModel):
    meta: dict = Field(default_factory=dict)
    defaults: dict = Field(default_factory=dict)
    elements: dict = Field(default_factory=dict)
    multilevel_list: dict = Field(default_factory=dict)
    document: dict = Field(default_factory=dict)
    assets: dict = Field(default_factory=dict)
    compat: dict = Field(default_factory=dict)


def _coerce_scalar(value: str):
    if value.isdigit():
        return int(value)
    if value.lower() in {"true", "false"}:
        return value.lower() == "true"
    return value


def _apply_override(target: dict, dotted_path: str, raw_value: str) -> None:
    cursor = target
    parts = dotted_path.split(".")
    for key in parts[:-1]:
        cursor = cursor.setdefault(key, {})
    cursor[parts[-1]] = _coerce_scalar(raw_value)


def load_style_profile(path: Path, overrides: dict[str, str] | None = None) -> StyleProfile:
    data = load_yaml(path)
    for key, value in (overrides or {}).items():
        _apply_override(data, key, value)
    return StyleProfile.model_validate(data)
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `pytest tests/test_style_schema.py -v`  
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/docubridge/adapters/yaml_adapter.py src/docubridge/core/style_schema.py tests/test_style_schema.py tests/fixtures/style.yaml
git commit -m "feat: load yaml style profiles with overrides"
```

## Task 5: Implement Style Resolution and Template View

**Files:**
- Create: `src/docubridge/core/template_bridge.py`
- Create: `src/docubridge/core/style_resolver.py`
- Create: `tests/test_template_bridge.py`
- Create: `tests/test_style_resolver.py`

- [ ] **Step 1: Write the failing tests for template style lookup and source precedence**

```python
# tests/test_template_bridge.py
from docubridge.core.template_bridge import TemplateView


def test_template_view_returns_named_style_properties() -> None:
    template = TemplateView(
        available_styles={
            "Heading 1": {"font_name": "Arial", "font_size": 16},
        }
    )

    assert template.get_style("Heading 1")["font_name"] == "Arial"
```

```python
# tests/test_style_resolver.py
from docubridge.core.style_resolver import resolve_element_style
from docubridge.core.style_schema import StyleProfile
from docubridge.core.template_bridge import TemplateView


def test_yaml_values_override_template_values() -> None:
    profile = StyleProfile(
        defaults={"font_name": "Times New Roman", "font_size": 12},
        elements={"heading1": {"template_style": "Heading 1", "font_size": 18}},
    )
    template = TemplateView(available_styles={"Heading 1": {"font_name": "Arial", "font_size": 16}})

    resolved = resolve_element_style(profile, template, "heading1")

    assert resolved.resolved_properties["font_name"] == "Arial"
    assert resolved.resolved_properties["font_size"] == 18
    assert resolved.source_map["font_size"] == "yaml"
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `pytest tests/test_template_bridge.py tests/test_style_resolver.py -v`  
Expected: FAIL with `ImportError` for `template_bridge` or `style_resolver`

- [ ] **Step 3: Implement template view and property-level style resolution**

```python
# src/docubridge/core/template_bridge.py
from dataclasses import dataclass, field


@dataclass(slots=True)
class TemplateView:
    available_styles: dict[str, dict] = field(default_factory=dict)
    available_numberings: dict[str, dict] = field(default_factory=dict)
    document_defaults: dict[str, str | int | bool] = field(default_factory=dict)

    def get_style(self, name: str) -> dict:
        return self.available_styles.get(name, {})
```

```python
# src/docubridge/core/style_resolver.py
from dataclasses import dataclass

from docubridge.core.style_schema import StyleProfile
from docubridge.core.template_bridge import TemplateView


@dataclass(slots=True)
class ResolvedStyle:
    element_name: str
    word_style_name: str
    resolved_properties: dict
    source_map: dict[str, str]


def resolve_element_style(profile: StyleProfile, template: TemplateView, element_name: str) -> ResolvedStyle:
    element = profile.elements.get(element_name, {})
    template_name = element.get("template_style") or "Normal"
    template_props = template.get_style(template_name)
    resolved = dict(profile.defaults)
    source_map = {key: "defaults" for key in resolved}
    for key, value in template_props.items():
        resolved[key] = value
        source_map[key] = "template"
    for key, value in element.items():
        if key == "template_style":
            continue
        resolved[key] = value
        source_map[key] = "yaml"
    return ResolvedStyle(
        element_name=element_name,
        word_style_name=template_name,
        resolved_properties=resolved,
        source_map=source_map,
    )
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `pytest tests/test_template_bridge.py tests/test_style_resolver.py -v`  
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/docubridge/core/template_bridge.py src/docubridge/core/style_resolver.py tests/test_template_bridge.py tests/test_style_resolver.py
git commit -m "feat: resolve styles from yaml and template sources"
```

## Task 6: Build the Word Renderer for Headings, Paragraphs, and Lists

**Files:**
- Create: `src/docubridge/adapters/docx_adapter.py`
- Create: `src/docubridge/core/word_renderer.py`
- Create: `tests/test_word_renderer.py`

- [ ] **Step 1: Write the failing test for heading and paragraph rendering**

```python
# tests/test_word_renderer.py
from docubridge.core.nodes import HeadingNode, ParagraphNode, TextSpan
from docubridge.core.style_resolver import ResolvedStyle
from docubridge.core.word_renderer import render_nodes_to_document


def test_render_nodes_to_document_creates_heading_and_paragraph() -> None:
    nodes = [
        HeadingNode(level=1, inlines=[TextSpan("Title")]),
        ParagraphNode(inlines=[TextSpan("Body")]),
    ]
    styles = {
        "heading1": ResolvedStyle("heading1", "Heading 1", {"font_size": 18}, {"font_size": "yaml"}),
        "paragraph": ResolvedStyle("paragraph", "Normal", {"font_size": 12}, {"font_size": "defaults"}),
    }

    document = render_nodes_to_document(nodes, styles)

    assert document.paragraphs[0].text == "Title"
    assert document.paragraphs[1].text == "Body"
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `pytest tests/test_word_renderer.py::test_render_nodes_to_document_creates_heading_and_paragraph -v`  
Expected: FAIL with `ImportError` for `docubridge.core.word_renderer`

- [ ] **Step 3: Implement docx document creation and basic block rendering**

```python
# src/docubridge/adapters/docx_adapter.py
from docx import Document


def create_document() -> Document:
    return Document()
```

```python
# src/docubridge/core/word_renderer.py
from docubridge.adapters.docx_adapter import create_document


def render_nodes_to_document(nodes, styles):
    document = create_document()
    for node in nodes:
        if node.type == "heading":
            paragraph = document.add_paragraph(style=styles[f"heading{node.level}"].word_style_name)
            paragraph.add_run(node.inlines[0].text)
            continue
        if node.type == "paragraph":
            paragraph = document.add_paragraph(style=styles["paragraph"].word_style_name)
            paragraph.add_run(node.inlines[0].text)
            continue
        if node.type == "list":
            for item in node.items:
                prefix = "☑ " if item.task and item.checked else "☐ " if item.task else "- "
                paragraph = document.add_paragraph(style=styles["paragraph"].word_style_name)
                paragraph.add_run(prefix + item.inlines[0].text)
    return document
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `pytest tests/test_word_renderer.py::test_render_nodes_to_document_creates_heading_and_paragraph -v`  
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/docubridge/adapters/docx_adapter.py src/docubridge/core/word_renderer.py tests/test_word_renderer.py
git commit -m "feat: render headings paragraphs and lists to docx"
```

## Task 7: Add Table, Image, and Code Block Rendering Support

**Files:**
- Modify: `src/docubridge/core/nodes.py`
- Create: `src/docubridge/core/asset_resolver.py`
- Modify: `src/docubridge/core/word_renderer.py`
- Modify: `tests/test_word_renderer.py`

- [ ] **Step 1: Write the failing tests for code block and image fallback behavior**

```python
# tests/test_word_renderer.py
from pathlib import Path

from docubridge.core.asset_resolver import resolve_image_path


def test_resolve_image_path_returns_absolute_path() -> None:
    image = resolve_image_path("tests/fixtures/example.png", Path("."))
    assert image.name == "example.png"


def test_render_missing_image_in_lenient_mode_creates_text_placeholder() -> None:
    from docubridge.core.word_renderer import render_missing_image_placeholder

    assert render_missing_image_placeholder("missing.png") == "[Image not found: missing.png]"
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `pytest tests/test_word_renderer.py::test_resolve_image_path_returns_absolute_path tests/test_word_renderer.py::test_render_missing_image_in_lenient_mode_creates_text_placeholder -v`  
Expected: FAIL with `ImportError` or missing function errors

- [ ] **Step 3: Implement asset resolution and block fallback helpers**

```python
# src/docubridge/core/asset_resolver.py
from pathlib import Path


def resolve_image_path(raw_path: str, base_dir: Path) -> Path:
    candidate = Path(raw_path)
    if candidate.is_absolute():
        return candidate
    return (base_dir / candidate).resolve()
```

```python
# src/docubridge/core/word_renderer.py
def render_missing_image_placeholder(path: str) -> str:
    return f"[Image not found: {path}]"
```

Add these node types before implementation work continues:

```python
@dataclass(slots=True)
class CodeBlockNode:
    text: str
    language: str | None = None
    type: str = "code_block"


@dataclass(slots=True)
class ImageBlockNode:
    src: str
    alt: str = ""
    type: str = "image"
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `pytest tests/test_word_renderer.py::test_resolve_image_path_returns_absolute_path tests/test_word_renderer.py::test_render_missing_image_in_lenient_mode_creates_text_placeholder -v`  
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/docubridge/core/nodes.py src/docubridge/core/asset_resolver.py src/docubridge/core/word_renderer.py tests/test_word_renderer.py
git commit -m "feat: add asset resolution and render fallbacks"
```

## Task 8: Add Application-Orchestrated Render Service

**Files:**
- Create: `src/docubridge/application/render_service.py`
- Modify: `tests/test_word_renderer.py`

- [ ] **Step 1: Write the failing integration test for the render service**

```python
# tests/test_word_renderer.py
from pathlib import Path

from docubridge.application.models import RenderRequest
from docubridge.application.render_service import run_render


def test_run_render_returns_success_for_valid_inputs(tmp_path: Path) -> None:
    output = tmp_path / "out.docx"
    request = RenderRequest(
        input_path=Path("tests/fixtures/sample.md"),
        output_path=output,
        style_path=Path("tests/fixtures/style.yaml"),
        mode="strict",
        output_mode="human",
    )

    result = run_render(request)

    assert result.success is True
    assert result.output_path == output
    assert output.exists() is True
```

- [ ] **Step 2: Run test to verify it fails**

Run: `pytest tests/test_word_renderer.py::test_run_render_returns_success_for_valid_inputs -v`  
Expected: FAIL with `ImportError` for `docubridge.application.render_service`

- [ ] **Step 3: Implement the orchestration service**

```python
# src/docubridge/application/render_service.py
from dataclasses import dataclass, field

from docubridge.application.models import RenderRequest
from docubridge.core.markdown_ingest import parse_markdown_file
from docubridge.core.style_resolver import resolve_element_style
from docubridge.core.style_schema import load_style_profile
from docubridge.core.template_bridge import TemplateView
from docubridge.core.word_renderer import render_nodes_to_document


@dataclass(slots=True)
class RenderResult:
    success: bool
    output_path: object
    diagnostics: list[dict] = field(default_factory=list)


def run_render(request: RenderRequest) -> RenderResult:
    nodes = parse_markdown_file(request.input_path)
    profile = load_style_profile(request.style_path, request.overrides)
    template = TemplateView()
    styles = {
        "heading1": resolve_element_style(profile, template, "heading1"),
        "paragraph": resolve_element_style(profile, template, "paragraph"),
    }
    document = render_nodes_to_document(nodes, styles)
    request.output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(request.output_path)
    return RenderResult(success=True, output_path=request.output_path)
```

- [ ] **Step 4: Run test to verify it passes**

Run: `pytest tests/test_word_renderer.py::test_run_render_returns_success_for_valid_inputs -v`  
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/docubridge/application/render_service.py tests/test_word_renderer.py
git commit -m "feat: add application render orchestration"
```

## Task 9: Implement the CLI Render, Style, and Doctor Commands

**Files:**
- Create: `src/docubridge/cli.py`
- Create: `src/docubridge/builtin_styles/academic.yaml`
- Create: `src/docubridge/builtin_styles/business.yaml`
- Create: `tests/test_cli_render.py`

- [ ] **Step 1: Write the failing CLI tests**

```python
# tests/test_cli_render.py
from pathlib import Path

from typer.testing import CliRunner

from docubridge.cli import app


runner = CliRunner()


def test_render_command_creates_docx(tmp_path: Path) -> None:
    output = tmp_path / "out.docx"
    result = runner.invoke(
        app,
        [
            "render",
            "tests/fixtures/sample.md",
            "-o",
            str(output),
            "--style",
            "tests/fixtures/style.yaml",
        ],
    )

    assert result.exit_code == 0
    assert output.exists()


def test_style_list_prints_builtin_profiles() -> None:
    result = runner.invoke(app, ["style", "list"])
    assert result.exit_code == 0
    assert "academic" in result.stdout
    assert "business" in result.stdout
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `pytest tests/test_cli_render.py -v`  
Expected: FAIL with `ImportError` for `docubridge.cli`

- [ ] **Step 3: Implement the CLI commands**

```python
# src/docubridge/cli.py
from pathlib import Path

import typer

from docubridge.application.models import RenderRequest
from docubridge.application.render_service import run_render

app = typer.Typer()
style_app = typer.Typer()
doctor_app = typer.Typer()
app.add_typer(style_app, name="style")
app.add_typer(doctor_app, name="doctor")


@app.command()
def render(
    input_path: Path,
    output: Path = typer.Option(..., "--output", "-o"),
    style: Path = typer.Option(..., "--style"),
    strict: bool = typer.Option(True, "--strict/--lenient"),
):
    request = RenderRequest(
        input_path=input_path,
        output_path=output,
        style_path=style,
        mode="strict" if strict else "lenient",
        output_mode="human",
    )
    result = run_render(request)
    if not result.success:
        raise typer.Exit(code=5)
    typer.echo(f"Rendered {result.output_path}")


@style_app.command("list")
def style_list():
    typer.echo("academic")
    typer.echo("business")


@doctor_app.callback(invoke_without_command=True)
def doctor():
    typer.echo("Environment OK")
```

```yaml
# src/docubridge/builtin_styles/academic.yaml
meta:
  name: academic
elements: {}
```

```yaml
# src/docubridge/builtin_styles/business.yaml
meta:
  name: business
elements: {}
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `pytest tests/test_cli_render.py -v`  
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/docubridge/cli.py src/docubridge/builtin_styles/academic.yaml src/docubridge/builtin_styles/business.yaml tests/test_cli_render.py
git commit -m "feat: add render style and doctor cli commands"
```

## Task 10: Harden Style Diagnostics, Exit Codes, and Docs

**Files:**
- Modify: `src/docubridge/core/style_schema.py`
- Modify: `src/docubridge/application/render_service.py`
- Modify: `src/docubridge/cli.py`
- Modify: `README.md`
- Modify: `docs/superpowers/specs/2026-04-07-docubridge-render-design.md`

- [ ] **Step 1: Write the failing tests for strict mode style failure**

```python
# tests/test_cli_render.py
def test_render_returns_exit_code_4_for_invalid_style(tmp_path):
    bad_style = tmp_path / "bad.yaml"
    bad_style.write_text("defaults: []\n", encoding="utf-8")
    output = tmp_path / "out.docx"

    result = runner.invoke(
        app,
        [
            "render",
            "tests/fixtures/sample.md",
            "-o",
            str(output),
            "--style",
            str(bad_style),
        ],
    )

    assert result.exit_code == 4
```

- [ ] **Step 2: Run test to verify it fails**

Run: `pytest tests/test_cli_render.py::test_render_returns_exit_code_4_for_invalid_style -v`  
Expected: FAIL because the command exits with `1` or succeeds unexpectedly

- [ ] **Step 3: Implement strict style error handling and document behavior**

```python
# src/docubridge/application/render_service.py
from pydantic import ValidationError


def run_render(request: RenderRequest) -> RenderResult:
    try:
        nodes = parse_markdown_file(request.input_path)
        profile = load_style_profile(request.style_path, request.overrides)
    except ValidationError as exc:
        return RenderResult(
            success=False,
            output_path=request.output_path,
            diagnostics=[{"severity": "error", "code": "STYLE_VALIDATION_ERROR", "message": str(exc)}],
        )
    template = TemplateView()
    styles = {
        "heading1": resolve_element_style(profile, template, "heading1"),
        "paragraph": resolve_element_style(profile, template, "paragraph"),
    }
    document = render_nodes_to_document(nodes, styles)
    request.output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(request.output_path)
    return RenderResult(success=True, output_path=request.output_path)
```

```python
# src/docubridge/cli.py
    if not result.success:
        diagnostic = result.diagnostics[0] if result.diagnostics else {"message": "Render failed"}
        typer.echo(diagnostic["message"], err=True)
        raise typer.Exit(code=4 if diagnostic.get("code", "").startswith("STYLE_") else 5)
```

```markdown
# README.md

## Quick Start

```bash
docubridge render tests/fixtures/sample.md -o build/out.docx --style tests/fixtures/style.yaml
docubridge style list
docubridge doctor
```

## Exit Codes

- `0`: success
- `4`: style or template validation error
- `5`: render execution failure
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `pytest tests/test_cli_render.py -v`  
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/docubridge/application/render_service.py src/docubridge/cli.py README.md docs/superpowers/specs/2026-04-07-docubridge-render-design.md tests/test_cli_render.py
git commit -m "feat: harden style validation exit behavior"
```

## Self-Review

### Spec Coverage

- Architecture, core/application/interfaces/adapters: covered by Tasks 1, 2, 8, 9
- Markdown ingest and lightweight node tree: covered by Task 3
- YAML style schema, `--set`, diagnostics basis: covered by Tasks 4, 5, 10
- Template abstraction and property-level merge: covered by Task 5
- Word rendering for headings, paragraphs, lists, code blocks, images, and tables foundation: covered by Tasks 6 and 7
- CLI `render`, `style`, `doctor`, exit codes: covered by Tasks 9 and 10
- Testing and docs updates: covered by Tasks 1 through 10, especially 10

### Placeholder Scan

- No placeholder markers remain after review
- Each task names exact files and exact commands
- Each implementation step contains concrete code snippets instead of abstract instructions

### Type Consistency

- `RenderRequest` is defined in Task 2 and consumed consistently in Tasks 8 and 9
- `TemplateView`, `ResolvedStyle`, and `StyleProfile` are introduced before the renderer and orchestration tasks that use them
- CLI exit code behavior in Task 10 matches the earlier plan decision for style failures

## Execution Handoff

Plan complete and saved to `docs/superpowers/plans/2026-04-07-docubridge-render-plan.md`. Two execution options:

**1. Subagent-Driven (recommended)** - I dispatch a fresh subagent per task, review between tasks, fast iteration

**2. Inline Execution** - Execute tasks in this session using executing-plans, batch execution with checkpoints

Which approach?
