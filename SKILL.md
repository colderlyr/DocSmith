---
name: DocSmith
description: >
  Create professional .docx (Word) documents with native LaTeX equation rendering (OMML) and
  Chinese/Western font control. Use this skill whenever the user asks to create a Word document,
  .docx file, academic paper, technical report, or any document containing math formulas.
  Also use when the user mentions 制作Word文档, 生成docx, 写报告, 排版, 数学公式,
  LaTeX转Word, or wants to produce a formatted document from text content.
  This skill provides a structured three-phase workflow: requirements collection,
  format confirmation, and code-driven document generation.
---

# DocSmith

Craft professional .docx documents from Markdown content with native LaTeX equation
rendering (OMML) and precise Chinese/Western font control.

**Design principle**: Template-first. If the user provides a template `.doc`/`.docx`,
extract formatting from it. If not, default to single-column — the most universal layout.

**v3 highlights**: ordered lists, image placeholders, abstract/acknowledgment sections,
TOC field, page-aware equation tab stops, preset JSON config files, `requirements.txt`.

## Three-Phase Workflow

```
Phase 1: Requirements →  Template? → Extract format specs
                          No template? → Default single-column
Phase 2: Format Confirmation → Output config table, wait approval
Phase 3: Code Execution      → Run Python to generate .docx
```

**Never skip phases.** Always complete Phase 1 and get Phase 2 approval before
writing any code.

---

## Phase 1: Requirements Collection

### 1.0 Template Detection (First Priority)

**Before asking any formatting questions, check if the user has a template:**

- "Do you have a template .doc/.docx file you want to follow?"
- If YES: analyze the template to extract margins, column layout, font sizes,
  heading styles, equation format. Map these directly to the config JSON.
  Skip the manual formatting questions below — the template is the authority.
- If NO: default to **single-column A4**, then ask the following questions
  to fill in remaining details.

### 1.1 Document Topic & Reference

- What is the document's core topic?
- Is there any reference material (outline, example document)?
- What output format is needed? (Single-column is the universal default.)

### 1.2 Heading Hierarchy

Ask the user to choose one:
- **Two-level** (H1, H2)
- **Three-level** (H1, H2, H3) — recommended for academic papers
- **Four-level** (H1, H2, H3, H4)

### 1.3 Heading Format (per level)

For each heading level, collect:
- **Numbering style**: e.g., `1.`, `1.1`, `I.` (Roman), `A.` (letters), `一、`, `(1)`, or none

  Supported numbering families and their automatic expansion:

  | H1 Style | H2 Auto | H3 Auto | H4 Auto | Use Case |
  |----------|---------|---------|---------|----------|
  | `1.`     | `1.1`   | `1.1.1` | —       | Standard academic |
  | `I.`     | `A.`    | `1)`    | `a)`    | IEEE papers |
  | `i.`     | `a.`    | —       | —       | Law / humanities |
  | `(1)`    | `(a)`   | —       | —       | Legal documents |
  | `1)`     | `a)`    | `i)`    | —       | Technical reports |
  | `一、`   | `（一）`| —       | —       | Chinese documents |

- **Font**: e.g., 黑体 (SimHei), 宋体 (SimSun), Arial, Times New Roman
- **Font size**: e.g., 16pt, 14pt, 三号 (16pt), 四号 (14pt)
- **Line spacing**: e.g., 1.5x, 2x, fixed 22pt

### 1.4 Body Text Format

- **Font**: Chinese body font + Western body font
- **Font size**: e.g., 12pt, 小四 (12pt)
- **Line spacing**: e.g., 1.5x, fixed 20pt
- **First-line indent**: e.g., 2 characters, 0.5 inch, none
- **Alignment**: e.g., justified (两端对齐), left

### 1.5 Page Layout

- **Page size**: A4 (default) or US Letter
- **Margins**: top, bottom, left, right (default 2.54cm / 1 inch)
- **Columns**: **Always 1 (single column) by default.** Only use 2-column if:
  - The user explicitly requests it, OR
  - A provided template specifies two-column layout
- **Column gap**: only relevant for 2-column (typically 0.25 inch / 0.63 cm)

### 1.6 Math Equation Specification (Critical)

#### LaTeX Sources (three input formats supported)

1. **`$$...$$` blocks** — standard display equations. Best for simple, well-formed LaTeX.
2. **Code blocks (` ``` `)** — equations in code fences. Best for papers with
   many equations, Unicode math symbols, or equations scraped from PDF.
   Each line is processed independently; equation numbers like `    (1)` are
   auto-extracted.
3. **Inline `$...$`** — standard inline math.

#### Equation rendering pipeline

When equations come from code blocks (non-LaTeX sources), they go through
an automatic cleanup pipeline:

```
Unicode math (α, ρ, ∇, ≤, ×...)
    ↓ latex_utils.unicode_to_latex()
Clean LaTeX with _{...} and ^{...} wrapping
    ↓ Named subscript pre-processing (_max → _{\max})
No double subscripts (V_{f,\max} not V_{f}_{max})
    ↓ latex2word
Native OMML (Word equation — double-click to edit)
```

#### Equation format

- **Display equations**: centered, with/without right-side numbering
- **Inline equations**: standard inline rendering
- **Text constraints**: lines like `(governing equations (1)-(3))` that have
  equation numbers but no math operators are auto-detected and rendered as text
  (not OMML), preserving the equation number alignment.

#### Known equation pitfalls

| Problem | Cause | Fix |
|---------|-------|-----|
| `[Eq: ...]` red placeholder | latex2word failed to parse | Check LaTeX for invalid patterns like double subscripts `V_{f}_{max}` |
| Double subscripts | Named words like `_max` after another `_` | Pre-processed: `_max`→`_{\max}`, `_min`→`_{\min}` |
| `\text{}` not working | Not actually broken — latex2word supports `\text{}`, `\textrm{}`, `\mathrm{}` | Use `\text{subject to:}` for text within equations |
| Unicode not rendering | Greek letters not converted to LaTeX | `unicode_to_latex()` converts α→\alpha, ρ→\rho, etc. |
| Equation number wrong | Multi-equation code block parsed as single line | Each line is now processed independently |

### 1.7 References (参考文献)

- Ask: "Do you need a references section at the end of the document?"
- If yes: generate 3-5 references in **GB/T 7714** format based on the document topic
- Format references in the content markdown as:
  ```
  ## 参考文献

  [1] Author. Title[J]. Journal, Year, Volume(Issue): Pages.
  [2] Author. Title[C]//Conference, Year.
  [3] Author. Title[M]. Publisher, Year.
  ```
- "参考文献" heading will be rendered as **Heading 2** style (matching `##` markdown level)
- Each reference entry uses hanging indent (2-char indent matching body font size)

---

## Phase 2: Format Confirmation

After collecting all requirements, output this **exact table** and ask the user to confirm:

```
| Configuration Item | Specification |
|-------------------|---------------|
| Document Topic    | [fill in]     |
| Reference Material| [fill in, or "None"] |
| Page Layout       | [A4/Letter, margins, 1-col/2-col] |
| Heading Structure | [e.g., Three-level] |
| H1 Format         | Numbering: [ ], Font: [ ], Size: [ ], Line spacing: [ ] |
| H2 Format         | Numbering: [ ], Font: [ ], Size: [ ], Line spacing: [ ] |
| H3 Format         | Numbering: [ ], Font: [ ], Size: [ ], Line spacing: [ ] |
| H4 Format (if any)| [fill in, or "N/A"] |
| Body Text         | Font: [ ], Size: [ ], Line spacing: [ ], Indent/Align: [ ] |
| Math Equations    | [None / Inline+Display, Display centered, with/without numbering] |
| References         | [Yes: GB/T 7714, N items / No] |
```

Ask: "Please confirm the above format requirements are correct. If so, provide the
document content or outline, and I will generate the formatted document."

**Do not proceed to Phase 3 until the user confirms and provides content.**

---

## Phase 3: Code Execution & Document Generation

When the user confirms the table and provides content, generate the .docx file by
running the bundled Python script.

### Usage

```bash
cd docsmith/
python3 -m scripts --output <output-file.docx> --config <config.json> --content <content.md>
```

The script accepts:
- `--output`: path for the generated .docx file
- `--config`: JSON file with all format settings from Phase 2
- `--content`: Markdown file with the document content

### Creating the config and content files

Before running the script, write two temporary files:

**1. Config JSON** (`/tmp/docx_config.json`):
```json
{
  "page": {
    "size": "A4",
    "margin_top": 2.54,
    "margin_bottom": 2.54,
    "margin_left": 2.54,
    "margin_right": 2.54,
    "columns": 1,
    "column_gap": 0.63
  },
  "headings": {
    "h1": { "numbering": "1.", "font": "SimHei", "font_west": "Arial", "size": 16, "bold": true, "line_spacing": 1.5 },
    "h2": { "numbering": "1.1", "font": "SimHei", "font_west": "Arial", "size": 14, "bold": true, "line_spacing": 1.5 },
    "h3": { "numbering": "1.1.1", "font": "SimHei", "font_west": "Arial", "size": 13, "bold": true, "line_spacing": 1.5 }
  },
  "body": {
    "font": "SimSun",
    "font_west": "Times New Roman",
    "size": 12,
    "line_spacing": 1.5,
    "first_line_indent": "2chars",
    "alignment": "justify"
  },
  "equations": {
    "display": "center",
    "numbering": false
  }
}
```

**2. Content Markdown** (`/tmp/docx_content.md`):

Write the user's content as Markdown. Three equation formats are supported:

```markdown
# Standard LaTeX display equations
$$\sum_{i=1}^{n} x_i$$

# Code block equations (for papers with many equations / Unicode math)
```
ρ(u · ∇)u = -∇p + μ∇²u - α(γ)u          (1)
∇ · u = 0                                  (2)
```

# Inline equations
Einstein's famous equation $E = mc^2$ changed physics forever.
```

### Technical Implementation

The script uses three critical techniques:

#### Technique 1: Precise Font Control (python-docx + OXML)

For Chinese-Western font mixing, the script injects `w:eastAsia` font identifiers
into the underlying XML. This ensures Chinese text renders in the specified Chinese font
while Western text renders in the specified Western font — they won't interfere.

```python
from docx.oxml.ns import qn

def set_run_font(run, font_cn, font_west, size_pt):
    """Set both Chinese and Western fonts on a run."""
    run.font.size = Pt(size_pt)
    run.font.name = font_west
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), font_cn)
    rFonts.set(qn('w:ascii'), font_west)
    rFonts.set(qn('w:hAnsi'), font_west)
```

#### Technique 2: LaTeX → OMML Native Equation Rendering

Equations are converted using `latex2word` (a pure-Python library), which produces
Word-native OMML XML:

- **Fully editable** in Word (not images)
- **Render natively** with Word's equation engine
- **Survive format conversion** and copy-paste

When equations come from non-LaTeX sources (code blocks with Unicode math),
the `latex_utils` module pre-processes them:

1. Unicode Greek/math symbols → LaTeX commands (α→\alpha, ≤→\leq, etc.)
2. Named subscripts use proper operators (_max→_{\max}, _min→_{\min})
3. Consecutive subscripts are joined (V_{f}_{max}→V_{f,\max})
4. Text operators are recognized (min:, subject to: → \text{...})

```python
from latex2word import LatexToWordElement

def insert_equation(paragraph, latex_str, display=True):
    element = LatexToWordElement(latex_str)
    element.add_latex_to_paragraph(paragraph, display=display)
```

#### Technique 3: Multi-Column Layout (OXML)

For IEEE/journal two-column layouts, columns are injected via OXML `w:cols` on the
section properties. Title/author/abstract paragraphs are forced to span both columns
using `w:cnt` elements.

```python
# Two-column layout
cols = OxmlElement('w:cols')
cols.set(qn('w:num'), '2')
cols.set(qn('w:space'), '360')  # 0.25" gap in twips
sectPr.append(cols)

# Single-column span for title area
cnt = OxmlElement('w:cnt')
para._p.get_or_add_pPr().append(cnt)
```

**Never write equations as plain text or images.** Always use native OMML conversion.

### After Generation

1. Run the script
2. Verify the output file was created
3. Tell the user the file path
4. Remind them: "Please open the document in Word to verify formatting and equations.
   You can double-click any equation to edit it natively."

---

## Dependencies

Install before first use:

```bash
pip3 install -r requirements.txt
```

Or manually: `pip3 install python-docx latex2word lxml`

---

## Preset Styles

When the user has NO template and is unsure about formatting, offer these.
**Presets A and B use single-column (universal default).**

### Preset A: Academic Paper (学术论文) — single column, recommended default

| Item | Setting |
|------|---------|
| Page | A4, margins 1 inch, single column |
| H1 | 黑体, 16pt, bold, centered, "1." |
| H2 | 黑体, 14pt, bold, left, "1.1" |
| H3 | 黑体, 13pt, bold, left, "1.1.1" |
| Body | 宋体 + Times New Roman, 12pt, 1.5x line spacing, first-line indent 2 chars, justified |
| Equations | Display centered, numbered on right |

### Preset B: Standard Report (标准报告)

| Item | Setting |
|------|---------|
| Page | A4, margins 1 inch, single column |
| H1 | 黑体, 15pt, bold, left, "一、" |
| H2 | 黑体, 13pt, bold, left, "（一）" |
| H3 | 楷体, 12pt, bold, left, "1." |
| Body | 宋体 + Arial, 11pt, 1.5x line spacing, first-line indent 2 chars, justified |
| Equations | Display centered, no numbering |

### Preset C: IEEE Conference Paper (IEEE 会议论文) — two-column, use ONLY when template or user requires it

| Item | Setting |
|------|---------|
| Page | US Letter, margins top=0.75"/bottom=1"/sides=0.625", **two columns** (3.5" each, 0.25" gap) |
| Title | Times New Roman, 24pt, bold, centered (single-column span) |
| Authors | Times New Roman, 11pt, centered (single-column span) |
| Affiliations | Times New Roman, 10pt, italic, centered (single-column span) |
| Abstract | 9pt bold heading + 10pt body (single-column span) |
| H1 | Times New Roman, 10pt, bold, centered, "I." (Roman) |
| H2 | Times New Roman, 10pt, bold italic, left, "A." (letter) |
| H3 | Times New Roman, 10pt, bold, left, "1)" |
| Body | Times New Roman, 10pt, justified, first-line indent 0.14", 1.0x spacing |
| Equations | Display centered, numbered right-aligned in parentheses |
| References | Times New Roman, 8pt, hanging indent |

```json
{
  "page": {
    "size": "letter",
    "margin_top": 1.91,
    "margin_bottom": 2.54,
    "margin_left": 1.59,
    "margin_right": 1.59,
    "columns": 2,
    "column_gap": 0.63
  },
  "headings": {
    "h1": { "numbering": "I.", "font": "Times New Roman", "font_west": "Times New Roman", "size": 10, "bold": true, "line_spacing": 1.0 },
    "h2": { "numbering": "A.", "font": "Times New Roman", "font_west": "Times New Roman", "size": 10, "bold": true, "line_spacing": 1.0 },
    "h3": { "numbering": "1)", "font": "Times New Roman", "font_west": "Times New Roman", "size": 10, "bold": true, "line_spacing": 1.0 }
  },
  "body": {
    "font": "Times New Roman",
    "font_west": "Times New Roman",
    "size": 10,
    "line_spacing": 1.0,
    "first_line_indent": "0.14in",
    "alignment": "justify"
  },
  "equations": {
    "display": "center",
    "numbering": true,
    "numbering_format": "({n})",
    "numbering_font": "Times New Roman",
    "numbering_size": 10
  }
}
```

---

## Architecture

```
DocSmith/
├── SKILL.md                      # This file — skill definition & user guide
├── README.md                     # Project overview & installation
├── requirements.txt              # Python dependencies
├── presets/
│   ├── academic_paper.json       # 学术论文 preset
│   ├── standard_report.json      # 标准报告 preset
│   └── ieee_conference.json      # IEEE 会议论文 preset
├── scripts/
│   ├── generate_docx.py          # CLI entry point + document builder
│   ├── parser.py                 # Markdown → structured blocks + DocContext
│   ├── elements.py               # Page setup, headings, tables, TOC, sections
│   ├── omml.py                   # LaTeX → OMML equation rendering (page-aware)
│   ├── latex_utils.py            # Unicode math → LaTeX preprocessing
│   └── fonts.py                  # Font size names, run/paragraph formatting
```

### Module responsibilities

| Module | What it does |
|--------|-------------|
| `parser.py` | Parses markdown into blocks: headings, $$math$$, \`\`\`code blocks, tables, ordered/unordered lists, images, references. Detects abstract/acknowledgment sections. |
| `latex_utils.py` | Preprocesses non-LaTeX equations: Unicode→LaTeX, subscript normalization, equation detection |
| `omml.py` | Converts LaTeX strings to Word-native OMML. Page-aware tab stops for equation numbering. |
| `elements.py` | Renders page setup (multi-column), headings (auto-numbering with fallback), tables, TOC, abstract/acknowledgment sections, references |
| `fonts.py` | Chinese font size names (三号→16pt), run/paragraph formatting with east-asia font support. CSS-like length parsing. |
| `generate_docx.py` | Orchestrates parsing → rendering, CLI interface. Equation label pre-scan, multi-column span management. |

### State Flag Management (parser pattern)

When parsing sections with multiple special modes (abstract, acknowledgments,
references), each transition MUST explicitly reset conflicting flags:

```python
# BUG: in_acknowledgment stays True → references leak into acknowledgments
if heading.startswith('reference'):
    in_references = True

# FIX: explicitly reset predecessor flag
if heading.startswith('reference'):
    in_acknowledgment = False  # ← critical!
    in_references = True
```

The `elif` chain order matters: `in_acknowledgment` is checked before `in_references`,
so if both flags are True, all content goes to the wrong collector.
