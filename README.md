# DocSmith

> Markdown + LaTeX → 排版精美的 Word 文档

A Claude Code skill for crafting professional `.docx` documents with **native LaTeX equation rendering** (OMML), **precise Chinese/Western font control**, **cross-references**, and **table/equation auto-numbering**.

## Why This Skill?

Existing docx tools for AI assistants have three critical gaps:

| Gap | Existing Solutions | This Skill |
|-----|-------------------|------------|
| **Equations** | Rendered as images (uneditable) or plain text | Native OMML — double-click to edit in Word |
| **Chinese fonts** | Garbled or wrong font for mixed Chinese/English | OXML `w:eastAsia` injection per run |
| **UX workflow** | One-shot generation, no format confirmation | Three-phase: requirements → confirm → generate |

## Features

- **Three-phase workflow**: Requirements collection → format confirmation table → code generation
- **LaTeX to native OMML**: Converts `$...$` and `$$...$$` to editable Word equations via `latex2word`
- **Chinese font segregation**: `w:eastAsia` font control so Chinese/English text each use correct fonts
- **Format configuration table**: Structured checkpoint before generation prevents rework
- **Two academic presets**: "Academic Paper" (学术论文) and "Standard Report" (标准报告)
- **Customizable**: Every aspect — heading numbering, fonts, sizes, spacing, page size, equation alignment — is configurable via JSON

## Installation

```bash
# Clone the skill
git clone https://github.com/YOUR_USERNAME/DocSmith.git ~/.claude/skills/DocSmith

# Install Python dependencies
pip3 install python-docx latex2word lxml
```

Then in Claude Code, the skill auto-triggers when you ask to create a Word document with equations or Chinese formatting.

## Quick Start

```
User: 帮我写一篇关于Transformer注意力机制的技术报告，包含数学公式

Claude (skill triggers):
  [Phase 1] Asks about heading structure, fonts, sizes, equation preferences...
  [Phase 2] Outputs format confirmation table
  [Phase 3] Generates .docx with native equations
```

## How It Works

### Equation Pipeline

```
LaTeX string ($E=mc^2$)
    ↓ latex2mathml
MathML
    ↓ mathml2omml (SAX handler)
OMML XML (m:oMath / m:oMathPara)
    ↓ lxml
Native Word equation element → appended to paragraph._element
```

### Font Pipeline

```python
# Each text run gets both fonts injected:
rFonts.set(qn('w:eastAsia'), 'SimSun')     # Chinese font
rFonts.set(qn('w:ascii'), 'Times New Roman')  # Western font
rFonts.set(qn('w:hAnsi'), 'Times New Roman')  # Western font
```

This ensures: 中文用宋体 and English uses Times New Roman — in the same paragraph, without font conflict.

## Comparison with Other Skills

| Feature | anthropic docx skill | Brise322 latex-convert | **This Skill** |
|---------|---------------------|----------------------|----------------|
| Create new docx | ✅ (JavaScript) | ❌ (edit only) | ✅ (Python) |
| LaTeX → OMML | ❌ | ✅ (Pandoc) | ✅ (Pure Python) |
| Chinese fonts | ❌ | ❌ | ✅ |
| Format confirmation | ❌ | ❌ | ✅ |
| Preset styles | ❌ | ❌ | ✅ |
| Editable equations | N/A | ✅ | ✅ |

## License

MIT

## Author

[Your Name]

## Acknowledgments

- [latex2word](https://github.com/Gu-f/LatexToWord) — Pure Python LaTeX → OMML conversion
- [latex2mathml](https://github.com/roniemartinez/latex2mathml) — LaTeX → MathML conversion
- [python-docx](https://github.com/python-openxml/python-docx) — Word document generation
