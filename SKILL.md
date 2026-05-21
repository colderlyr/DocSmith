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

## Three-Phase Workflow

```
Phase 1: Requirements Collection  →  Ask questions, gather specs
Phase 2: Format Confirmation      →  Output config table, wait for approval
Phase 3: Code Execution           →  Run Python to generate the .docx file
```

**Never skip phases.** Always complete Phase 1 and get Phase 2 approval before
writing any code.

---

## Phase 1: Requirements Collection

Ask the user these five categories of questions. If the user doesn't know specific
formatting parameters, proactively offer 1–2 preset styles (e.g., "Academic Paper"
or "Standard Report").

### 1.1 Document Topic & Reference

- What is the document's core topic?
- Is there any reference material (template, outline, example document)?

### 1.2 Heading Hierarchy

Ask the user to choose one:
- **Two-level** (H1, H2)
- **Three-level** (H1, H2, H3) — recommended for academic papers
- **Four-level** (H1, H2, H3, H4)

### 1.3 Heading Format (per level)

For each heading level, collect:
- **Numbering style**: e.g., `1.`, `1.1`, `一、`, `(1)`, or none
- **Font**: e.g., 黑体 (SimHei), 宋体 (SimSun), Arial, Times New Roman
- **Font size**: e.g., 16pt, 14pt, 三号 (16pt), 四号 (14pt)
- **Line spacing**: e.g., 1.5x, 2x, fixed 22pt

### 1.4 Body Text Format

- **Font**: Chinese body font + Western body font
- **Font size**: e.g., 12pt, 小四 (12pt)
- **Line spacing**: e.g., 1.5x, fixed 20pt
- **First-line indent**: e.g., 2 characters, 0.5 inch, none
- **Alignment**: e.g., justified (两端对齐), left

### 1.6 References (参考文献)

- Ask: "Do you need a references section at the end of the document?"
- If yes: generate 3-5 references in **GB/T 7714** format based on the document topic
- Format references in the content markdown as:
  ```
  ## 参考文献

  [1] Author. Title[J]. Journal, Year, Volume(Issue): Pages.
  [2] Author. Title[C]//Conference, Year.
  [3] Author. Title[M]. Publisher, Year.
  ```
- "参考文献" heading will be rendered as **Heading 1** style (黑体, 三号, black, no indent)
- Each reference entry uses hanging indent (2-char indent matching body font size)

### 1.5 Math Equation Specification (Critical)

Ask whether the document contains equations. If yes, confirm:
- **Display equations**: centered, with/without right-side numbering
- **Inline equations**: standard inline rendering

Instruct the user: "When providing content, wrap all formulas in standard LaTeX:
`$...$` for inline and `$$...$$` for display equations."

---

## Phase 2: Format Confirmation

After collecting all requirements, output this **exact table** and ask the user to confirm:

```
| Configuration Item | Specification |
|-------------------|---------------|
| Document Topic    | [fill in]     |
| Reference Material| [fill in, or "None"] |
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
cd DocSmith/
python3 -m scripts.generate_docx \
  --output <output-file.docx> \
  --config <config.json> \
  --content <content.md>
```

The script accepts:
- `--output`: path for the generated .docx file
- `--config`: JSON file with all format settings from Phase 2
- `--content`: Markdown file with the document content (LaTeX formulas in `$...$` or `$$...$$`)

### Creating the config and content files

Before running the script, write two temporary files:

**1. Config JSON** (`/tmp/docx_config.json`):
```json
{
  "page": {
    "size": "A4",
    "margin_top": 1440,
    "margin_bottom": 1440,
    "margin_left": 1440,
    "margin_right": 1440
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

Write the user's content as Markdown with LaTeX formulas:
- `# Heading 1`, `## Heading 2`, `### Heading 3` for headings
- `$E = mc^2$` for inline equations
- `$$\sum_{i=1}^{n} x_i$$` for display equations
- Standard Markdown for paragraphs, bold, italic, lists, tables

### Technical Implementation

The script uses two critical techniques:

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
    # Critical: inject east-asia font for Chinese rendering
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
Word-native OMML XML. This means equations:

- Are **fully editable** in Word (not images)
- Render **natively** with Word's equation engine
- Survive **format conversion** and copy-paste

```python
from latex2word import LatexToWordElement

def insert_equation(paragraph, latex_str, display=True):
    """Insert a LaTeX equation as native OMML into a paragraph."""
    element = LatexToWordElement(latex_str)
    element.add_latex_to_paragraph(paragraph, display=display)
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
pip3 install python-docx latex2word
```

---

## Preset Styles

When the user is unsure about formatting, offer these presets:

### Preset A: Academic Paper (学术论文)

| Item | Setting |
|------|---------|
| Page | A4, margins 1 inch |
| H1 | 黑体, 16pt, bold, centered, "1." |
| H2 | 黑体, 14pt, bold, left, "1.1" |
| H3 | 黑体, 13pt, bold, left, "1.1.1" |
| Body | 宋体 + Times New Roman, 12pt, 1.5x line spacing, first-line indent 2 chars, justified |
| Equations | Display centered, numbered on right |

### Preset B: Standard Report (标准报告)

| Item | Setting |
|------|---------|
| Page | A4, margins 1 inch |
| H1 | 黑体, 15pt, bold, left, "一、" |
| H2 | 黑体, 13pt, bold, left, "（一）" |
| H3 | 楷体, 12pt, bold, left, "1." |
| Body | 宋体 + Arial, 11pt, 1.5x line spacing, first-line indent 2 chars, justified |
| Equations | Display centered, no numbering |
