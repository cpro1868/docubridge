# Your AI-Era Document Bridge (docubridge)

**The document format translator that speaks both Markdown and Word — and actually gets along with AI.**

> *Because you shouldn't have to manually reformat documents when AI can do it for you.*

---

## What is this thing?

Ever tried to feed a Word document to an AI? Good luck. AIs *love* Markdown — it's clean, readable, and doesn't come with invisible binary baggage. But the real world runs on `.docx` files, and your boss wants everything in Word.

**Docubridge is the bridge between these two worlds.**

It does two things, and does them well:

| What it does | Why you need it |
|--------------|-----------------|
| **Word/Excel/PPT → Markdown** | Convert office docs to AI-friendly Markdown for analysis, RAG pipelines, or feeding to LLMs |
| **Markdown → Word** | Turn AI-generated Markdown into properly formatted Word documents with real styles, fonts, and numbering |

---

## Who is this for?

### 📝 The AI-Powered Writer

You use AI to draft reports, proposals, or documentation. The AI spits out Markdown. But your company expects Word. Now you can render that Markdown into a properly formatted `.docx` with your organization's house style — without touching the mouse.

### 🔬 The Researcher

Got a stack of Word documents to feed into an AI research assistant? Parse them to Markdown first. Keep headings, lists, tables, and images. Lose the binary headaches.

### 🏢 The Office Worker

Need to convert a colleague's `.docx` into something AI-readable? Or turn an AI-generated draft into a document that looks like *you* actually formatted it? Docubridge has your back.

### 👨‍💻 The Developer

Integrate document conversion into your scripts, CI/CD pipelines, or RAG preprocessing. It's CLI-first, scriptable, and doesn't phone home.

---

## Installation

**Requirements: Python 3.12 or higher**

---

### Option A: Install from downloaded package (Recommended for production use) ⭐

Download the wheel file from the release page, then install it:

```bash
pip install ./docubridge-0.2.0-py3-none-any.whl
```

Or if you prefer a specific version:

```powershell
# Windows
pip install .\docubridge-0.2.0-py3-none-any.whl

# macOS / Linux
pip install ./docubridge-0.2.0-py3-none-any.whl
```

After installation, verify:

```bash
docubridge --help
```

---

### Option B: Install from source (For development and testing) 🔧

If you want to modify the code or run tests, install from source:

#### Windows

```powershell
# 1. Create a virtual environment
python -m venv .venv

# 2. Activate it
.\.venv\Scripts\Activate.ps1

# 3. Install dependencies
python -m pip install --upgrade pip
pip install -e .
```

#### macOS / Linux

```bash
# 1. Create a virtual environment
python3 -m venv .venv

# 2. Activate it
source .venv/bin/activate

# 3. Install dependencies
python -m pip install --upgrade pip
pip install -e .
```

#### Install dev dependencies (for running tests)

```bash
pip install -e .[dev]
```

---

### First-run Verification

```bash
docubridge doctor
```

If you see "Environment OK", you're good to go.

---

## Quick Start

### The "I got a Word doc from my colleague" workflow

```bash
# Parse the Word doc to Markdown (AI food!)
docubridge parse colleague-report.docx -o colleague-report.md

# Now you can feed it to any AI tool
```

### The "AI wrote my draft, now I need a real Word doc" workflow

```bash
# Option 1: Use a built-in style
docubridge render ai-draft.md -o final-report.docx --style academic

# Option 2: Extract styles from your company's template
docubridge extract-styles company-template.docx -o company-style.yaml
docubridge render ai-draft.md -o final-report.docx --style company-style.yaml --template company-template.docx
```

### The "Let me see what this thing can do" workflow

```bash
# List built-in styles
docubridge style list

# Show me the academic style
docubridge style show academic

# Render a test document
docubridge render tests/fixtures/sample.md -o test-output.docx --style tests/fixtures/style.yaml
```

---

## Core Features

### 🎯 `parse` — Convert Office Docs to Markdown

Turn Word, Excel, and PowerPoint files into clean, AI-readable Markdown.

```bash
docubridge parse document.docx -o output.md
docubridge parse spreadsheet.xlsx -o output.md
docubridge parse slides.pptx -o output.md
```

**What it extracts:**

- 📑 **Headings** — `Heading 1` becomes `#`, `Heading 2` becomes `##`, etc.
- 📝 **Paragraphs** — Including basic formatting (bold, italic, links, code)
- 📋 **Lists** — Ordered and unordered, including simple nesting
- 📊 **Tables** — Converted to Markdown pipe tables
- 🖼️ **Images** — Extracted to an `assets/` folder, referenced in Markdown
- 📑 **Excel Sheets** — Each sheet becomes a section with a table

**Output is JSON-parseable if you need it:**

```bash
docubridge parse document.docx -o output.md --json
```

---

### ✨ `render` — Convert Markdown to Word (Full Guide)

Transform Markdown into a properly formatted Word document with real styles. This is where the magic happens.

#### Basic Usage

```bash
# Use a built-in style
docubridge render draft.md -o final.docx --style academic

# Use a custom style file
docubridge render draft.md -o final.docx --style my-style.yaml

# Use both style AND template for best results
docubridge render draft.md -o final.docx --style my-style.yaml --template company-template.docx
```

---

#### Understanding `--style` and `--template`

This is the most important concept to understand. These two flags work together but have different jobs:

| Flag | What it does | Think of it as |
|------|--------------|----------------|
| `--style` | YAML configuration file | **The rulebook** — tells Docubridge how to map Markdown to Word styles |
| `--template` | Word document (.docx) | **The style library** — provides actual Word styles to reuse |

**The simple explanation:**

```
Markdown = content structure (what to write)
YAML     = mapping rules (which style for which element)
Template = Word styles (what those styles actually look like)
```

**Example: How they work together**

When your Markdown has:
```markdown
# Chapter 1 Introduction
```

The YAML says:
```yaml
heading1:
  template_style: Heading 1
```

The Template provides the actual `Heading 1` style with your company's fonts, colors, and spacing.

Result: A properly styled heading in your company's brand.

---

#### The Three Ways to Use `render`

**Method 1: Built-in Styles (Simplest)**

```bash
docubridge render draft.md -o output.docx --style academic
```

Available built-in styles: `academic`, `business`, `default`

Great for: Quick results, testing, simple documents.

---

**Method 2: Custom YAML Only (No Template)**

```bash
docubridge render draft.md -o output.docx --style my-style.yaml
```

You write a YAML file defining your styles. Docubridge creates Word styles dynamically.

Great for: Portable configurations, sharing styles across machines.

**Minimal YAML example:**

```yaml
defaults:
  font_name: Times New Roman
  font_size: 12

elements:
  heading1:
    template_style: Heading 1
    font_size: 18
    bold: true
  paragraph:
    template_style: Normal
    font_size: 12
```

---

**Method 3: YAML + Template (Recommended for production) ⭐**

```bash
docubridge render draft.md -o output.docx --style my-style.yaml --template company-template.docx
```

Combines your YAML mapping rules with your company's existing Word template.

Great for: Company documents, academic papers, anything that needs consistent branding.

**Why this is the best approach:**

1. **Template** provides reusable Word styles (fonts, colors, numbering, spacing)
2. **YAML** specifies which Markdown elements use which styles
3. **Together**: Professional results that match your organization's standards

---

#### Priority Rule

When both YAML and template define the same property:

```
YAML explicit value  >  Template value  >  Default value
```

In plain English: If you explicitly write something in YAML, it wins. Otherwise, Docubridge uses the template. If neither has it, default kicks in.

**Example:**
```yaml
# YAML says font_size = 14
heading1:
  font_size: 14
```

Even if the template's `Heading 1` has font_size = 16, the output will use 14 because YAML takes priority.

---

#### 🎨 `extract-styles` — Your Secret Weapon ⭐

This is the new hotness. **Extract styles from any Word document and reuse them.**

Got a Word template with your company's perfect formatting? Use it as a style source:

```bash
# Extract styles from your company's template
docubridge extract-styles company-template.docx -o company-style.yaml

# Now use it to render any Markdown with company branding
docubridge render ai-draft.md -o company-doc.docx --style company-style.yaml --template company-template.docx
```

**What it does:**

- Scans Word styles (Heading 1, Normal, List Number, etc.)
- Maps them to Markdown elements automatically
- Handles Chinese Office styles (`标题 1` → `heading1`, `正文` → `paragraph`)
- Saves unmapped styles to `compat.extracted_styles` for manual review
- Generates a YAML file you can tweak and reuse

**Flags:**

| Flag | What it does |
|------|--------------|
| `--pretty` | Output human-readable YAML with indentation |
| `--strict` | Fail if any Word style can't be mapped (for perfectionists) |
| `--json` | Get structured JSON output instead of text |

---

#### Recommended Workflow for Template Rendering

For best results, follow this sequence:

**Step 1: Extract styles from your template (first time only)**

```bash
docubridge extract-styles company-template.docx -o company-style.yaml
```

**Step 2: Run doctor to check everything**

```bash
docubridge doctor draft.md --style company-style.yaml --template company-template.docx
```

**Step 3: Inspect specific elements if needed**

```bash
docubridge style explain company-style.yaml heading1 --template company-template.docx --pretty
docubridge style explain company-style.yaml paragraph --template company-template.docx --pretty
```

**Step 4: Render the final document**

```bash
docubridge render draft.md -o final.docx --style company-style.yaml --template company-template.docx
```

---

#### Troubleshooting Template Issues

**"My template has styles but they didn't apply"**

Run `style explain` to debug:
```bash
docubridge style explain my-style.yaml heading1 --template template.docx --pretty
```

Look for:
- `word_style_name`: Is it the style you expected?
- `source_map`: Where did the value come from? (`yaml` wins over `template`)

**"The output doesn't match my template exactly"**

Remember Docubridge's priorities:
1. YAML explicit values
2. Template values
3. Defaults

Check if YAML is overriding template values.

**"Which elements should I check first?"**

Start with these in order:
1. `heading1` — Most visible element
2. `paragraph` — Body text
3. `ordered_list` or `unordered_list` — Lists
4. `table` — Tables

---

### 📋 `style` — Style Management

Built-in styles included:

| Style | Best For |
|-------|----------|
| `academic` | Research papers, theses |
| `business` | Reports, memos, proposals |
| `default` | General purpose (Chinese-friendly defaults) |

**Commands:**

```bash
# List all built-in styles
docubridge style list

# Show a style's YAML
docubridge style show academic

# Validate your custom style file
docubridge style validate my-style.yaml

# Explain how a specific element will render
docubridge style explain my-style.yaml heading1 --template template.docx

# Merge overrides into a style (great for scripting)
docubridge style merge my-style.yaml --set document.toc.depth=4
```

---

### 🔍 `doctor` — Pre-flight Check

Before rendering an important document, run the diagnostic:

```bash
docubridge doctor draft.md --style academic --template template.docx
```

This checks:

- ✅ Environment is ready
- ✅ Markdown is readable
- ✅ Style file is valid
- ✅ Styles resolve correctly
- ✅ Template numbering resources are available

---

## Real-World Scenarios

### Scenario 1: AI-Generated Research Summary

```
You: "ChatGPT, write me a 10-page research summary on quantum computing."

ChatGPT: *outputs beautiful Markdown*

You: *copy-paste to file: quantum-summary.md*

You: docubridge render quantum-summary.md -o client-ready.docx --style academic

Result: A properly formatted Word document your client can actually use. ✓
```

### Scenario 2: Analyzing Competitor's RFP

```
You: "Extract all text from this RFP doc."

Docubridge: *parses the 50-page Word doc to clean Markdown*

You: *feed to AI for analysis*

Result: AI-friendly content in seconds. No copy-paste hell. ✓
```

### Scenario 3: Company Template Magic

```
IT: "Here's our new branded Word template with all the correct styles."

You: docubridge extract-styles corporate-template.docx -o corp-style.yaml

You: docubridge render markdown-draft.md -o final.docx --style corp-style.yaml --template corporate-template.docx

Result: AI-generated content in company-approved format. IT is impressed. ✓
```

---

## What's Coming Next

- 🔜 **PDF parsing** — Text extraction, OCR support for scanned documents, and multi-column layout recovery
- 🔜 **Batch processing** — Convert folders of files at once
- 🔜 **GUI** — Desktop application for non-technical users
- 🔜 **Real-time collaboration** — Cloud sync and multi-user editing

The core `.docx ↔ Markdown` bridge is solid and production-ready.

---

## Technical Details

**Built with:**
- Python 3.12+
- `python-docx` for Word manipulation
- `openpyxl` for Excel parsing
- `python-pptx` for PowerPoint parsing
- `markdown-it-py` for Markdown parsing
- `ruamel.yaml` for YAML handling
- `typer` for the CLI

**Exit codes:**

| Code | Meaning |
|------|---------|
| 0 | Success |
| 4 | Style or template validation error |
| 5 | Execution failure (file not found, etc.) |

---

## Getting Help

Chinese documentation: `README_CN.md`

Additional guides:
- Template Usage Guide: `docs/superpowers/specs/2026-04-14-docubridge-template-guide-cn.md`
- Style Configuration Reference: `docs/superpowers/specs/2026-04-14-docubridge-template-reference-cn.md`
- 5-Minute Quickstart (Chinese): `docs/superpowers/specs/2026-04-12-docubridge-5min-quickstart-cn.md`

---

**Made for people who work with AI and still have to deliver Word documents.**

*Because life's too short to manually format footnotes.*
