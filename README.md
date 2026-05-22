# DocSmith

> Markdown + LaTeX → 排版精美的 Word 文档

A Claude Code skill for crafting professional `.docx` documents with **native LaTeX equation rendering** (OMML), **precise Chinese/Western font control**, **multi-column layouts**, **cross-references**, and **table/equation auto-numbering**.

## What's New (v3)

- **Ordered lists** — `1.` / `1)` / `a.` / `(1)` style lists with per-level auto-numbering
- **Image placeholders** — `![alt](path)` rendered as gray placeholder text
- **Abstract & Acknowledgments** — Auto-detected `## 摘要` / `## 致谢` sections
- **TOC (Table of Contents)** — Native Word `TOC` field code (`"toc": true` in config)
- **Page-aware equation tab stops** — Correctly aligned numbers for any page size, margins, or column count
- **Same-line `$$...$$` display math** — Fixed critical parser bug that swallowed content after inline display equations
- **Preset config files** — `presets/academic_paper.json`, `standard_report.json`, `ieee_conference.json`
- **`requirements.txt`** — Standard dependency declaration

## Why This Skill?

Existing docx tools for AI assistants have three critical gaps:

| Gap | Existing Solutions | This Skill |
|-----|-------------------|------------|
| **Equations** | Rendered as images (uneditable) or plain text | Native OMML — double-click to edit in Word |
| **Chinese fonts** | Garbled or wrong font for mixed Chinese/English | OXML `w:eastAsia` injection per run |
| **UX workflow** | One-shot generation, no format confirmation | Three-phase: requirements → confirm → generate |
| **Unicode math** | Broken when pasting equations from PDF/papers | Auto-converted: α→\alpha, ∇→\nabla, etc. |
| **Column layouts** | No support for journal two-column format | OXML w:cols + w:cnt single-column spans |

## Features

- **Three-phase workflow**: Requirements collection → format confirmation table → code generation
- **LaTeX to native OMML**: Converts `$...$`, `$$...$$`, AND `\`\`\`` code blocks to editable Word equations via `latex2word`
- **Unicode math preprocessing**: Auto-converts Greek letters, math symbols, sub/superscripts to proper LaTeX
- **Chinese font segregation**: `w:eastAsia` font control so Chinese/English text each use correct fonts
- **Format configuration table**: Structured checkpoint before generation prevents rework
- **Three presets**: "Academic Paper" (学术论文), "Standard Report" (标准报告), "IEEE Conference Paper" (IEEE 会议论文)
- **Multi-column layouts**: IEEE two-column, journal formats via OXML
- **Extended numbering**: Roman (I., II.), letters (A., B.), and mixed families
- **Ordered & unordered lists**: Auto-numbered `1.` / `a.` / `(1)` style lists and bullet lists
- **Image placeholders**: `![alt](path)` markdown → gray placeholder with path
- **Abstract & Acknowledgment**: Auto-detected sections with single-column span in multi-col layouts

## Installation

```bash
# Install Python dependencies
pip3 install python-docx latex2word lxml
```

## Quick Start

```
User: 帮我写一篇关于Transformer注意力机制的技术报告，包含数学公式

Claude (skill triggers):
  [Phase 1] Asks about heading structure, fonts, sizes, equation preferences...
  [Phase 2] Outputs format confirmation table
  [Phase 3] Generates .docx with native equations
```

For IEEE papers with complex equations:

```
User: 把这个论文markdown按IEEE模板排版成docx

Claude (skill triggers):
  [Phase 1] Reads template .doc, extracts IEEE specs
  [Phase 2] Confirms two-column format, Roman numerals, equation numbering
  [Phase 3] Generates .docx — Unicode math→LaTeX→OMML, images→placeholders
```

## How It Works

### Equation Pipeline

```
Source equations ──┬── $$...$$ LaTeX ──────────────┐
                   │                                ▼
                   └── ``` Unicode math ──► latex_utils.preprocess()
                         • α→\alpha, ∇→\nabla      │
                         • _max→_{\max}             ▼
                         • V_f_max→V_{f,\max}   latex2word
                                                    │
                                                    ▼
                                              Native OMML
                                      (double-click to edit in Word)
```

### Font Pipeline

```python
# Each text run gets both fonts injected:
rFonts.set(qn('w:eastAsia'), 'SimSun')     # Chinese font
rFonts.set(qn('w:ascii'), 'Times New Roman')  # Western font
rFonts.set(qn('w:hAnsi'), 'Times New Roman')  # Western font
```

This ensures: 中文用宋体 and English uses Times New Roman — in the same paragraph, without font conflict.

### Column Pipeline

```python
# Two-column layout via OXML
cols = OxmlElement('w:cols')
cols.set(qn('w:num'), '2')
sectPr.append(cols)

# Title spans both columns
cnt = OxmlElement('w:cnt')
para._p.get_or_add_pPr().append(cnt)
```

## Architecture

```
DocSmith/
├── SKILL.md                   # Skill definition & user guide
├── README.md                  # This file
├── requirements.txt           # Python dependencies
├── presets/
│   ├── academic_paper.json    # 学术论文
│   ├── standard_report.json   # 标准报告
│   └── ieee_conference.json   # IEEE 会议论文
├── scripts/
│   ├── generate_docx.py       # CLI + document builder
│   ├── parser.py              # Markdown → structured blocks
│   ├── elements.py            # Page setup, headings, tables, TOC, sections
│   ├── omml.py                # LaTeX → OMML (latex2word)
│   ├── latex_utils.py         # Unicode math → LaTeX preprocessing
│   └── fonts.py               # Font utilities
```

## Comparison with Other Skills

| Feature | anthropic docx skill | Brise322 latex-convert | **This Skill** |
|---------|---------------------|----------------------|----------------|
| Create new docx | Yes (JavaScript) | No (edit only) | Yes (Python) |
| LaTeX → OMML | No | Yes (Pandoc) | Yes (Pure Python) |
| Unicode math→LaTeX | No | No | **Yes** |
| Chinese fonts | No | No | Yes |
| Multi-column layout | No | No | **Yes** |
| Format confirmation | No | No | Yes |
| IEEE preset | No | No | **Yes** |
| Preset styles | No | No | Yes |
| Editable equations | N/A | Yes | Yes |

## License

MIT

## Acknowledgments

- [latex2word](https://github.com/Gu-f/LatexToWord) — Pure Python LaTeX → OMML conversion
- [python-docx](https://github.com/python-openxml/python-docx) — Word document generation
- IEEE paper generation experience — Source of Unicode math preprocessing, multi-column layout, and state flag management patterns
