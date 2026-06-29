# DocuBridge Template Engine Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Upgrade `Markdown -> Word (.docx)` from a lightweight style mapper into a mixed-mode template layout engine that applies real Word template styles, paragraph properties, font slots, and numbering definitions.

**Architecture:** Introduce a new rendering pipeline in front of the current Word writer: template extraction builds a richer `TemplateView`, Markdown nodes are translated into layout intent objects, and a merge layer resolves template defaults plus YAML overrides into a final layout plan. The Word writer then applies paragraph styles, run fonts, paragraph properties, and native numbering bindings instead of simulating lists with text prefixes.

**Tech Stack:** Python 3.12+, `python-docx`, OOXML access via `docx.oxml`, `typer`, `pydantic`, `pytest`

---

## Planned File Structure

### Create

- `docs/superpowers/plans/2026-04-14-docubridge-template-engine-implementation-plan.md`
- `src/docubridge/core/layout_intent.py`
- `tests/test_layout_intent.py`

### Modify

- `src/docubridge/core/template_bridge.py`
- `src/docubridge/core/style_schema.py`
- `src/docubridge/core/style_resolver.py`
- `src/docubridge/core/word_renderer.py`
- `src/docubridge/application/render_service.py`
- `src/docubridge/cli.py`
- `tests/test_template_bridge.py`
- `tests/test_style_resolver.py`
- `tests/test_word_renderer.py`
- `tests/test_cli_render.py`
- `README.md`
- `README_CN.md`

## Task 1: Freeze the Baseline With Failing Template-Engine Tests

**Files:**
- Modify: `tests/test_template_bridge.py`
- Modify: `tests/test_style_resolver.py`
- Modify: `tests/test_word_renderer.py`
- Modify: `tests/test_cli_render.py`

- [ ] **Step 1: Add a failing template extraction test for paragraph properties and font slots**

```python
def test_load_template_view_extracts_font_slots_and_paragraph_properties(tmp_path: Path) -> None:
    template_path = tmp_path / "template.docx"
    document = Document()
    normal = document.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    normal.paragraph_format.first_line_indent = Pt(21)
    normal.paragraph_format.space_after = Pt(6)
    document.save(template_path)

    template = load_template_view(template_path)

    normal_style = template.get_style("Normal")
    assert normal_style["font_ascii"] == "Times New Roman"
    assert normal_style["font_east_asia"] == "宋体"
    assert normal_style["first_line_indent_pt"] == 21
    assert normal_style["space_after_pt"] == 6
```

- [ ] **Step 2: Add a failing numbering extraction test**

```python
def test_load_template_view_extracts_numbering_definitions(tmp_path: Path) -> None:
    template_path = tmp_path / "numbering-template.docx"
    document = Document()
    first = document.add_paragraph("One", style="List Number")
    second = document.add_paragraph("Sub", style="List Number 2")
    document.save(template_path)

    template = load_template_view(template_path)

    assert template.available_numberings
    assert "List Number" in template.style_numbering_map
```

- [ ] **Step 3: Add a failing resolver test for template defaults + YAML override precedence**

```python
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
```

- [ ] **Step 4: Add a failing renderer test for native numbering and paragraph properties**

```python
def test_render_nodes_to_document_binds_native_numbering_and_paragraph_properties(tmp_path: Path) -> None:
    template_path = tmp_path / "template.docx"
    _write_numbered_template(template_path)
    nodes = [
        ListNode(kind="ordered", items=[ListItemNode(inlines=[TextSpan(text="A")], level=0)]),
    ]
    styles = {
        "paragraph": resolved_style("Normal"),
        "ordered_list": resolved_style(
            "List Number",
            first_line_indent_pt=21,
            left_indent_pt=28,
            numbering_style="List Number",
        ),
    }

    document = render_nodes_to_document(nodes, styles, template_path=template_path)
    paragraph = document.paragraphs[0]

    assert paragraph.style.name == "List Number"
    assert _paragraph_numbering_reference(paragraph) is not None
    assert paragraph.text == "A"
```

- [ ] **Step 5: Add a failing CLI explain test for new template-engine diagnostics**

```python
def test_style_explain_outputs_font_slots_paragraph_properties_and_numbering(tmp_path: Path) -> None:
    ...
    assert payload["resolved_properties"]["font_east_asia"] == "宋体"
    assert payload["resolved_properties"]["first_line_indent_pt"] == 21
    assert payload["resolved_properties"]["numbering_style"] == "Heading 1"
```

- [ ] **Step 6: Run only the new baseline tests**

Run: `rtk pytest tests/test_template_bridge.py tests/test_style_resolver.py tests/test_word_renderer.py tests/test_cli_render.py -q`
Expected: FAIL with missing properties, missing numbering metadata, and simulated list rendering behavior.

## Task 2: Expand `TemplateView` Into a Real Template Metadata Model

**Files:**
- Modify: `src/docubridge/core/template_bridge.py`
- Modify: `tests/test_template_bridge.py`

- [ ] **Step 1: Introduce richer template data structures**

Add focused dataclasses in `src/docubridge/core/template_bridge.py`:

```python
@dataclass(slots=True)
class TemplateStyle:
    name: str
    style_type: str
    based_on: str | None = None
    properties: dict[str, Any] = field(default_factory=dict)


@dataclass(slots=True)
class NumberingLevel:
    ilvl: int
    num_fmt: str
    lvl_text: str
    start: int = 1


@dataclass(slots=True)
class NumberingDefinition:
    num_id: int
    abstract_id: int
    levels: dict[int, NumberingLevel] = field(default_factory=dict)
```

- [ ] **Step 2: Extend `TemplateView` to expose styles, document defaults, and style-to-numbering mapping**

Implement the view shape:

```python
@dataclass(slots=True)
class TemplateView:
    available_styles: dict[str, TemplateStyle] = field(default_factory=dict)
    available_numberings: dict[int, NumberingDefinition] = field(default_factory=dict)
    style_numbering_map: dict[str, tuple[int, int]] = field(default_factory=dict)
    document_defaults: dict[str, Any] = field(default_factory=dict)

    def get_style(self, name: str) -> dict[str, Any]:
        style = self.available_styles.get(name)
        return deepcopy(style.properties if style is not None else {})
```

- [ ] **Step 3: Read document default run/paragraph properties from the template**

Implement helper functions:

```python
def _read_document_defaults(document: Document) -> dict[str, Any]:
    ...


def _read_style_properties(style) -> dict[str, Any]:
    ...
```

Required extracted fields for the first phase:

```text
font_ascii
font_hansi
font_east_asia
font_cs
font_size
bold
italic
first_line_indent_pt
left_indent_pt
right_indent_pt
space_before_pt
space_after_pt
alignment
```

- [ ] **Step 4: Read numbering definitions and style-linked numbering references**

Implement helpers in `template_bridge.py`:

```python
def _parse_numbering_part(document: Document) -> tuple[dict[int, NumberingDefinition], dict[str, tuple[int, int]]]:
    ...
```

Behavior:

```text
read abstract numbering levels
read numbering instances
capture style-linked numId/ilvl pairs
populate TemplateView.available_numberings
populate TemplateView.style_numbering_map
```

- [ ] **Step 5: Re-run template bridge tests**

Run: `rtk pytest tests/test_template_bridge.py -q`
Expected: PASS with extracted font slots, paragraph properties, and numbering metadata.

## Task 3: Introduce Layout Intent Objects Between Markdown and Word

**Files:**
- Create: `src/docubridge/core/layout_intent.py`
- Modify: `src/docubridge/application/render_service.py`
- Create: `tests/test_layout_intent.py`

- [ ] **Step 1: Create layout intent dataclasses**

Add `src/docubridge/core/layout_intent.py`:

```python
@dataclass(slots=True)
class RunStyleIntent:
    text: str
    bold: bool = False
    italic: bool = False
    strike: bool = False
    code: bool = False
    href: str | None = None


@dataclass(slots=True)
class NumberingIntent:
    numbering_role: str
    level: int = 0
    continue_sequence: bool = True
    preferred_template_style: str | None = None


@dataclass(slots=True)
class ParagraphLayoutIntent:
    element_name: str
    runs: list[RunStyleIntent]
    resolved_style_name: str
    resolved_properties: dict[str, Any]
    numbering: NumberingIntent | None = None
```

- [ ] **Step 2: Add a translator from existing Markdown nodes into layout intents**

Implement in `render_service.py`:

```python
def build_layout_intents(nodes: list[object], styles: dict[str, ResolvedStyle]) -> list[ParagraphLayoutIntent]:
    ...
```

Rules:

```text
paragraph -> one paragraph intent
heading -> one paragraph intent with heading style
ordered/unordered list items -> one paragraph intent per item with numbering metadata
quote -> paragraph intent
table/code/image remain handled by renderer-specific structures in the same list or companion wrappers
```

- [ ] **Step 3: Write tests for the translator**

Add to `tests/test_layout_intent.py`:

```python
def test_build_layout_intents_maps_ordered_lists_to_numbering_intents() -> None:
    ...
    assert intents[0].numbering.numbering_role == "ordered_list"
    assert intents[0].numbering.level == 0
```

```python
def test_build_layout_intents_preserves_inline_run_flags() -> None:
    ...
    assert intents[0].runs[0].bold is True
    assert intents[0].runs[1].href == "https://example.com"
```

- [ ] **Step 4: Re-run layout intent tests**

Run: `rtk pytest tests/test_layout_intent.py -q`
Expected: PASS and `run_render` can build layout intents before document writing.

## Task 4: Replace Style Resolution With Template Defaults + Style + YAML Merge

**Files:**
- Modify: `src/docubridge/core/style_schema.py`
- Modify: `src/docubridge/core/style_resolver.py`
- Modify: `tests/test_style_resolver.py`

- [ ] **Step 1: Extend the style schema to accept first-phase template-engine fields**

Update the README-facing schema expectations by allowing these properties in `defaults` and `elements`:

```python
SUPPORTED_LAYOUT_FIELDS = {
    "font_ascii",
    "font_hansi",
    "font_east_asia",
    "font_cs",
    "font_size",
    "bold",
    "italic",
    "first_line_indent_pt",
    "left_indent_pt",
    "right_indent_pt",
    "space_before_pt",
    "space_after_pt",
    "alignment",
    "numbering_style",
}
```

- [ ] **Step 2: Replace the current shallow merge with explicit precedence**

Implement in `style_resolver.py`:

```python
def _merge_properties(
    *,
    defaults: dict[str, Any],
    template_document_defaults: dict[str, Any],
    template_style_properties: dict[str, Any],
    yaml_properties: dict[str, Any],
) -> tuple[dict[str, Any], dict[str, str]]:
    ...
```

Expected source labels:

```text
defaults
template_document_default
template_style
yaml
yaml_override
```

- [ ] **Step 3: Attach numbering hints from the template style map**

Behavior:

```text
if YAML explicitly sets numbering_style, keep it
else if resolved Word style name has a template numbering reference, store that as numbering_style
else leave numbering_style unset
```

- [ ] **Step 4: Re-run resolver tests**

Run: `rtk pytest tests/test_style_resolver.py -q`
Expected: PASS with correct precedence and source maps for font slots, paragraph properties, and numbering.

## Task 5: Teach the Word Writer To Apply Native Paragraph Properties and Numbering

**Files:**
- Modify: `src/docubridge/core/word_renderer.py`
- Modify: `tests/test_word_renderer.py`

- [ ] **Step 1: Stop writing ordered-list text prefixes for numbered paragraphs**

Replace the current block:

```python
paragraph.add_run(
    f"{_list_item_indent(item)}{_list_item_prefix(item, item_kind, display_index)} "
)
```

with logic that:

```python
paragraph = document.add_paragraph(style=style_name)
_apply_paragraph_properties(paragraph, resolved_properties)
_bind_numbering(paragraph, template_view, numbering_intent)
_add_inline_runs(paragraph, item.inlines)
```

- [ ] **Step 2: Implement paragraph property application helpers**

Add helpers in `word_renderer.py`:

```python
def _apply_paragraph_properties(paragraph, properties: Mapping[str, object]) -> None:
    ...


def _apply_run_fonts(run, properties: Mapping[str, object]) -> None:
    ...
```

Required behavior:

```text
first_line_indent_pt -> paragraph.paragraph_format.first_line_indent
left/right indent -> paragraph.paragraph_format.left_indent/right_indent
space_before_pt/space_after_pt -> paragraph spacing
alignment -> paragraph.alignment
font_east_asia/font_ascii/font_hansi/font_cs -> direct OOXML rFonts assignment
```

- [ ] **Step 3: Implement numbering binding helpers using template numbering metadata**

Add helper shape:

```python
def _bind_numbering(paragraph, template: TemplateView, numbering: NumberingIntent | None) -> None:
    ...
```

Behavior:

```text
lookup numbering_style or resolved style name in template.style_numbering_map
write w:numPr with numId and ilvl onto the paragraph
do not inject textual "1." prefixes for ordered lists
```

- [ ] **Step 4: Keep unordered list fallback explicit**

If the template lacks a bullet numbering binding:

```text
apply the requested paragraph style
emit a deterministic diagnostic-capable fallback path
keep "-" text prefix only for explicit fallback mode
```

- [ ] **Step 5: Re-run renderer tests**

Run: `rtk pytest tests/test_word_renderer.py -q`
Expected: PASS with native numbering references present and paragraph/run properties applied.

## Task 6: Wire the New Template Engine Through `run_render` and Diagnostics

**Files:**
- Modify: `src/docubridge/application/render_service.py`
- Modify: `src/docubridge/cli.py`
- Modify: `tests/test_cli_render.py`

- [ ] **Step 1: Make `run_render` build template view once and pass it through**

Refactor service flow:

```python
template = load_template_view(request.template_path)
styles = validate_render_styles(nodes, request, template=template)
layout_intents = build_layout_intents(nodes, styles)
document = render_nodes_to_document(
    nodes,
    styles,
    layout_intents=layout_intents,
    template_view=template,
    base_dir=request.input_path.parent,
    template_path=request.template_path,
)
```

- [ ] **Step 2: Expand `style explain` output**

Required JSON additions:

```json
{
  "resolved_properties": {
    "font_east_asia": "宋体",
    "font_ascii": "Times New Roman",
    "first_line_indent_pt": 21,
    "space_after_pt": 6,
    "numbering_style": "Heading 1"
  },
  "source_map": {
    "font_east_asia": "template_style",
    "numbering_style": "template_style"
  }
}
```

- [ ] **Step 3: Add deterministic diagnostics for missing numbering resources**

Expected behavior:

```text
doctor --template <file> reports TEMPLATE_VALIDATION_ERROR when a required numbering style is requested but not found
style explain shows numbering_style missing and source map details
render in strict mode fails for missing required native numbering resources
render in lenient mode may fall back with explicit diagnostic text
```

- [ ] **Step 4: Re-run CLI and service tests**

Run: `rtk pytest tests/test_cli_render.py -q`
Expected: PASS with richer explain output and numbering-related failures surfaced deterministically.

## Task 7: Document the First-Phase Template Engine Contract

**Files:**
- Modify: `README.md`
- Modify: `README_CN.md`

- [ ] **Step 1: Update template documentation to reflect the new first-phase contract**

Add the following points in both READMEs:

```markdown
- Template defaults, paragraph properties, and native numbering are now part of the render contract.
- Chinese and English fonts can be configured separately.
- `--template` provides style and numbering resources; `--style` maps Markdown elements and applies explicit overrides.
- Missing native numbering resources fail fast in strict mode.
```

- [ ] **Step 2: Add a template YAML example that shows split fonts and paragraph properties**

Document example:

```yaml
defaults:
  font_ascii: "Times New Roman"
  font_east_asia: "宋体"
  font_size: 12

elements:
  paragraph:
    template_style: Normal
    first_line_indent_pt: 21
    space_after_pt: 6
  heading1:
    template_style: Heading 1
  ordered_list:
    template_style: List Number
```

- [ ] **Step 3: Verify README references and examples**

Run: `rtk rg -n "font_east_asia|native numbering|strict mode|template defaults" README.md README_CN.md`
Expected: both docs mention the first-phase template engine rules.

## Task 8: Full Verification and Packaging

**Files:**
- Modify: none

- [ ] **Step 1: Run focused template-engine tests**

Run: `rtk pytest tests/test_template_bridge.py tests/test_layout_intent.py tests/test_style_resolver.py tests/test_word_renderer.py tests/test_cli_render.py -q`
Expected: all focused template-engine tests pass.

- [ ] **Step 2: Run the full suite**

Run: `rtk pytest -q`
Expected: full suite passes.

- [ ] **Step 3: Rebuild the wheel**

Run: `rtk python -m pip wheel . -w dist --no-deps`
Expected: `dist/docubridge-0.1.0-py3-none-any.whl` is rebuilt successfully.

- [ ] **Step 4: Record final evidence**

Capture in the implementation notes or release docs:

```text
test count
date
wheel filename
supported first-phase template capabilities
known non-goals still deferred
```

## Self-Review

Spec coverage:

- template style defaults and YAML merge are covered by Tasks 1, 2, and 4
- Chinese/English font slot support is covered by Tasks 1, 2, 4, and 5
- paragraph property support is covered by Tasks 1, 2, 4, and 5
- native numbering and template numbering reuse are covered by Tasks 1, 2, 5, and 6
- diagnosable `doctor` / `style explain` output is covered by Task 6
- updated user-facing contract is covered by Task 7

Placeholder scan:

- no `TODO`, `TBD`, or “implement later” placeholders remain
- each task names exact files, concrete commands, and expected behavior

Type consistency:

- `TemplateView`, `ResolvedStyle`, `ParagraphLayoutIntent`, and `NumberingIntent` are named consistently across tasks
- the plan reuses current codebase names such as `render_service`, `template_bridge`, `style_resolver`, and `word_renderer`

## Execution Handoff

Plan complete and saved to `docs/superpowers/plans/2026-04-14-docubridge-template-engine-implementation-plan.md`. Two execution options:

**1. Subagent-Driven (recommended)** - I dispatch a fresh subagent per task, review between tasks, fast iteration

**2. Inline Execution** - Execute tasks in this session using executing-plans, batch execution with checkpoints

Which approach?
