# DocuBridge v1 Release Focus Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Ship the first user-facing DocuBridge version by focusing release readiness on `Word (.docx) -> Markdown`, `Markdown -> Word (.docx)`, and template/YAML customization quality.

**Architecture:** The release plan does not expand product scope. It narrows the execution path to the two core document flows and the style/template system that makes `Markdown -> Word` usable in real user scenarios. All remaining work should either strengthen these three areas or be explicitly deferred.

**Tech Stack:** Python 3.12+, `typer`, `python-docx`, `pytest`, Markdown parsing core, YAML style system

---

## Release Scope

`v1.0` release-blocking scope:

- `.docx -> .md`
- `.md -> .docx`
- YAML style configuration
- `--template` template integration
- CLI usability for the above paths
- docs and examples needed to operate the above paths

Explicitly non-blocking for `v1.0`:

- `.xlsx -> .md`
- `.pptx -> .md`
- `.pdf -> .md`
- `batch`
- GUI
- advanced layout fidelity beyond current Word/document structure goals

## Release Exit Criteria

The release is ready only when all items below are true:

- `.docx -> markdown` core scenarios are covered by automated tests and a real-sample checklist
- `markdown -> .docx` core scenarios are covered by automated tests and a real-sample checklist
- template loading and template/YAML precedence are implemented, tested, and documented
- style diagnostics are understandable enough for a user to fix invalid configuration without reading source code
- at least 2-3 reusable style examples are available and documented
- CLI help and README reflect actual supported scope

## Planned File Structure

### Create

- `docs/superpowers/plans/2026-04-10-docubridge-v1-release-focus.md`

### Modify

- `docs/requirements.md`
- `docs/superpowers/specs/2026-04-07-docubridge-render-design.md`
- `README.md`
- `src/docubridge/application/render_service.py`
- `src/docubridge/core/template_bridge.py`
- `src/docubridge/core/style_resolver.py`
- `src/docubridge/core/word_renderer.py`
- `src/docubridge/core/docx_ingest.py`
- `src/docubridge/cli.py`
- `tests/test_cli_render.py`
- `tests/test_docx_ingest.py`
- `tests/test_style_resolver.py`
- `tests/test_template_bridge.py`
- `tests/test_word_renderer.py`

## Task 1: Lock Release Scope in Product Docs

**Files:**
- Modify: `docs/requirements.md`
- Modify: `docs/superpowers/specs/2026-04-07-docubridge-render-design.md`
- Modify: `README.md`

- [ ] **Step 1: Verify the docs currently overstate `v1.0` scope**

Run: `rg -n "pdf|pptx|xlsx|batch|GUI|发布|release|v1.0" docs/requirements.md docs/superpowers/specs/2026-04-07-docubridge-render-design.md README.md`
Expected: matches show `xlsx/pptx/pdf` and other roadmap features still mixed into `v1.0` language.

- [ ] **Step 2: Rewrite the release scope language**

Add the following product statements:

```markdown
- `v1.0` release-blocking scope is limited to `.docx -> .md`, `.md -> .docx`, and template/YAML customization.
- `.xlsx/.pptx/.pdf` remain roadmap or non-blocking capabilities for the first release.
- Template customization is a core deliverable, not an optional extra.
```

- [ ] **Step 3: Verify the docs now use a consistent release gate**

Run: `rg -n "release-blocking|发布门槛|非阻塞|template|docx -> markdown|markdown -> docx" docs/requirements.md docs/superpowers/specs/2026-04-07-docubridge-render-design.md README.md`
Expected: every document reflects the narrowed `v1.0` scope.

## Task 2: Audit `.docx -> markdown` Against Real Release Cases

**Files:**
- Modify: `tests/test_docx_ingest.py`
- Modify: `src/docubridge/core/docx_ingest.py`

- [ ] **Step 1: Add one failing release-case test at a time**

Focus only on release-blocking structures:

```python
def test_parse_docx_release_case_preserves_heading_paragraph_list_table_and_image():
    ...
```

```python
def test_parse_docx_release_case_preserves_basic_inline_styles():
    ...
```

- [ ] **Step 2: Run each new test to confirm a real failure before implementation**

Run: `pytest tests/test_docx_ingest.py -q`
Expected: the newly added release-case test fails for a concrete missing behavior, not for fixture or syntax errors.

- [ ] **Step 3: Implement the minimal parser change needed**

Keep the parser focused on:

```text
heading order
paragraph order
list nesting
table extraction
image references
bold / italic / strikethrough / link / inline code
```

- [ ] **Step 4: Re-run the parser test file**

Run: `pytest tests/test_docx_ingest.py -q`
Expected: all `.docx -> markdown` parser tests pass.

## Task 3: Close the Template Integration Gap in `render`

**Files:**
- Modify: `src/docubridge/application/render_service.py`
- Modify: `src/docubridge/core/template_bridge.py`
- Modify: `src/docubridge/core/style_resolver.py`
- Modify: `tests/test_template_bridge.py`
- Modify: `tests/test_style_resolver.py`
- Modify: `tests/test_cli_render.py`

- [ ] **Step 1: Write a failing template-entry test at the service or CLI layer**

```python
def test_render_command_uses_explicit_template_file(tmp_path: Path) -> None:
    ...
```

```python
def test_render_service_reports_template_file_not_found() -> None:
    ...
```

- [ ] **Step 2: Run only the new template tests**

Run: `pytest tests/test_template_bridge.py tests/test_style_resolver.py tests/test_cli_render.py -q`
Expected: failure proves `--template` is not fully wired into the render path or diagnostics path.

- [ ] **Step 3: Implement the smallest end-to-end template flow**

Required behaviors:

```text
CLI passes template_path into RenderRequest
application layer loads template view from file
resolver merges template properties with YAML according to precedence
template failures produce deterministic diagnostics and exit code 4
```

- [ ] **Step 4: Re-run the template-focused tests**

Run: `pytest tests/test_template_bridge.py tests/test_style_resolver.py tests/test_cli_render.py -q`
Expected: all template flow tests pass.

## Task 4: Harden `Markdown -> Word` Rendering for Release Structures

**Files:**
- Modify: `src/docubridge/core/word_renderer.py`
- Modify: `tests/test_word_renderer.py`

- [ ] **Step 1: Add one failing release-case render test per missing structure**

Cover only release-blocking output:

```python
def test_render_release_case_applies_heading_paragraph_and_list_styles() -> None:
    ...
```

```python
def test_render_release_case_applies_table_code_block_and_image_output() -> None:
    ...
```

- [ ] **Step 2: Run the renderer tests to confirm the exact gap**

Run: `pytest tests/test_word_renderer.py -q`
Expected: failures identify concrete rendering gaps in release-critical structures.

- [ ] **Step 3: Implement the minimal renderer changes**

Keep scope limited to:

```text
style application
list continuity
table rendering
code block rendering
image placement
inline formatting retention
```

- [ ] **Step 4: Re-run the renderer tests**

Run: `pytest tests/test_word_renderer.py -q`
Expected: renderer tests are fully green.

## Task 5: Make Template Customization Usable, Not Just Present

**Files:**
- Modify: `README.md`
- Modify: `src/docubridge/cli.py`
- Modify: `tests/test_cli_render.py`

- [ ] **Step 1: Add failing CLI/help assertions for template and style guidance**

```python
def test_render_help_describes_style_and_template_roles() -> None:
    ...
```

- [ ] **Step 2: Run the CLI help test**

Run: `pytest tests/test_cli_render.py::test_render_help_describes_style_and_template_roles -q`
Expected: FAIL because current help text does not clearly explain the release-critical customization path.

- [ ] **Step 3: Update help and README examples**

Required guidance:

```text
`--style` defines explicit formatting intent
`--template` provides host Word styles/resources
YAML overrides template on conflicting explicit properties
```

- [ ] **Step 4: Re-run the CLI help test**

Run: `pytest tests/test_cli_render.py::test_render_help_describes_style_and_template_roles -q`
Expected: PASS

## Task 6: Run Full Release Verification

**Files:**
- Modify: none

- [ ] **Step 1: Run focused release-path tests**

Run: `pytest tests/test_docx_ingest.py tests/test_word_renderer.py tests/test_template_bridge.py tests/test_style_resolver.py tests/test_cli_render.py -q`
Expected: all release-path tests pass.

- [ ] **Step 2: Run the full suite**

Run: `pytest -q`
Expected: full suite passes.

- [ ] **Step 3: Record release evidence in docs or release notes**

Capture:

```text
test count
date
release scope
known non-blocking limitations
```

## Self-Review

Spec coverage:

- product scope narrowing is handled in Task 1
- `.docx -> markdown` release readiness is handled in Task 2
- template integration and precedence are handled in Task 3
- `markdown -> .docx` release readiness is handled in Task 4
- usability/documentation for template customization is handled in Task 5
- verification evidence is handled in Task 6

Placeholder scan:

- no `TODO` or `TBD` markers remain
- every task has exact files and commands

Type consistency:

- paths and component names match the current codebase naming (`render_service`, `template_bridge`, `style_resolver`, `word_renderer`, `docx_ingest`)

## Execution Handoff

Plan complete and saved to `docs/superpowers/plans/2026-04-10-docubridge-v1-release-focus.md`. Two execution options:

**1. Subagent-Driven (recommended)** - I dispatch a fresh subagent per task, review between tasks, fast iteration

**2. Inline Execution** - Execute tasks in this session using executing-plans, batch execution with checkpoints

Which approach?
