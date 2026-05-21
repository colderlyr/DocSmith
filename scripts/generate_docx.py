#!/usr/bin/env python3
"""
Generate a .docx file from Markdown content with:
- Native OMML equation rendering (via latex2word)
- Chinese/Western font segregation (via OXML eastAsia injection)
- Configurable heading, body, and page formatting

Usage:
    python3 generate_docx.py --output out.docx --config config.json --content content.md
"""

import json
import re
import sys
import argparse
from pathlib import Path

from docx import Document
from docx.shared import Pt, Cm, Inches, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml, OxmlElement
from lxml import etree

try:
    from latex2word import LatexToWordElement
    HAS_LATEX2WORD = True
except ImportError:
    HAS_LATEX2WORD = False
    print("WARNING: latex2word not installed. Equations will be rendered as plain text.")
    print("Install with: pip3 install latex2word")


# ---------------------------------------------------------------------------
# Font & OXML utilities
# ---------------------------------------------------------------------------

def set_run_font_cn(run, font_cn, font_west, size_pt, bold=False, italic=False):
    """Set both Chinese and Western fonts on a run via OXML injection."""
    run.font.size = Pt(size_pt)
    run.font.name = font_west
    run.bold = bold
    run.italic = italic
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), font_cn)
    rFonts.set(qn('w:ascii'), font_west)
    rFonts.set(qn('w:hAnsi'), font_west)
    rFonts.set(qn('w:cs'), font_west)


def set_paragraph_spacing(paragraph, line_spacing, first_line_indent=None, alignment=None):
    """Configure paragraph spacing, indent, and alignment."""
    pf = paragraph.paragraph_format
    pf.line_spacing = line_spacing

    if first_line_indent == "2chars":
        # Approximate 2 Chinese characters at current font size
        pf.first_line_indent = Cm(0.74)
    elif first_line_indent:
        pf.first_line_indent = first_line_indent

    align_map = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    }
    if alignment and alignment in align_map:
        paragraph.alignment = align_map[alignment]


def apply_body_style(paragraph, body_cfg):
    """Apply body text font and spacing to all runs in a paragraph."""
    set_paragraph_spacing(
        paragraph,
        body_cfg.get("line_spacing", 1.5),
        body_cfg.get("first_line_indent"),
        body_cfg.get("alignment", "justify"),
    )
    for run in paragraph.runs:
        set_run_font_cn(
            run,
            body_cfg["font"],
            body_cfg.get("font_west", "Times New Roman"),
            body_cfg.get("size", 12),
        )


# ---------------------------------------------------------------------------
# Equation handling
# ---------------------------------------------------------------------------

MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"


def insert_omml_equation(paragraph, latex_str, display=True):
    """Convert a LaTeX string to native OMML and append to paragraph.

    For display equations, wraps in m:oMathPara (centered, numbered if configured).
    For inline equations, appends m:oMath directly.
    """
    if not HAS_LATEX2WORD:
        run = paragraph.add_run(f"[Equation: {latex_str}]")
        run.font.color.rgb = RGBColor(0xCC, 0x00, 0x00)
        return

    try:
        eq = LatexToWordElement(latex_str)
        omml_element = eq.element()  # This is an lxml m:oMath element

        if display:
            # Wrap m:oMath in m:oMathPara for display equations
            omath_para = etree.SubElement(
                etree.Element(f"{{{MATH_NS}}}oMathPara"),
                f"{{{MATH_NS}}}oMath",
            )
            # Copy children from the original oMath to the new one
            for child in omml_element:
                omath_para.append(child)
            paragraph._element.append(omath_para.getparent())
        else:
            # Inline equation — append oMath directly
            paragraph._element.append(omml_element)
    except Exception as e:
        run = paragraph.add_run(f"[Equation error: {latex_str}]")
        run.font.color.rgb = RGBColor(0xCC, 0x00, 0x00)
        print(f"  WARNING: Failed to convert LaTeX: {latex_str[:60]}... — {e}")


def process_inline_latex(paragraph, text, body_cfg):
    """Split text on $...$ inline LaTeX, adding runs and OMML equations."""
    parts = re.split(r'(?<!\$)\$(?!\$)(.+?)(?<!\$)\$(?!\$)', text)
    for i, part in enumerate(parts):
        if i % 2 == 0:
            # Plain text
            if part:
                run = paragraph.add_run(part)
                set_run_font_cn(
                    run,
                    body_cfg["font"],
                    body_cfg.get("font_west", "Times New Roman"),
                    body_cfg.get("size", 12),
                )
        else:
            # Inline LaTeX equation
            insert_omml_equation(paragraph, part.strip(), display=False)


# ---------------------------------------------------------------------------
# Page setup
# ---------------------------------------------------------------------------

def setup_page(doc, page_cfg):
    """Set page size and margins."""
    for section in doc.sections:
        size_map = {
            "A4": (Cm(21.0), Cm(29.7)),
            "letter": (Cm(21.59), Cm(27.94)),
        }
        if page_cfg.get("size") in size_map:
            w, h = size_map[page_cfg["size"]]
            section.page_width = w
            section.page_height = h

        section.top_margin = Cm(page_cfg.get("margin_top", 2.54))
        section.bottom_margin = Cm(page_cfg.get("margin_bottom", 2.54))
        section.left_margin = Cm(page_cfg.get("margin_left", 2.54))
        section.right_margin = Cm(page_cfg.get("margin_right", 2.54))


# ---------------------------------------------------------------------------
# Heading configuration
# ---------------------------------------------------------------------------

def get_heading_number(text, level, headings_cfg, counters):
    """Generate heading number prefix based on numbering style."""
    key = f"h{level}"
    cfg = headings_cfg.get(key, {})
    style = cfg.get("numbering", "")

    if not style:
        return ""

    # Increment counter for this level
    counters[level - 1] += 1
    # Reset lower-level counters
    for i in range(level, len(counters)):
        counters[i] = 0

    if style == "1.":
        if level == 1:
            return f"{counters[0]}. "
        elif level == 2:
            return f"{counters[0]}.{counters[1]} "
        elif level == 3:
            return f"{counters[0]}.{counters[1]}.{counters[2]} "
    elif style == "一、":
        cn = ["", "一", "二", "三", "四", "五", "六", "七", "八", "九", "十",
              "十一", "十二", "十三", "十四", "十五"]
        return f"{cn[counters[0]] if counters[0] < len(cn) else str(counters[0])}、"
    elif style == "（一）":
        cn = ["", "一", "二", "三", "四", "五", "六", "七", "八", "九", "十"]
        return f"（{cn[counters[1]] if counters[1] < len(cn) else str(counters[1])}）"

    return ""


def add_heading(doc, level, text, headings_cfg, counters):
    """Add a styled heading with numbering prefix."""
    key = f"h{level}"
    cfg = headings_cfg.get(key)
    if not cfg:
        return

    prefix = get_heading_number(text, level, headings_cfg, counters)
    full_text = prefix + text

    paragraph = doc.add_paragraph()
    run = paragraph.add_run(full_text)
    set_run_font_cn(
        run,
        cfg.get("font", "SimHei"),
        cfg.get("font_west", "Arial"),
        cfg.get("size", 16),
        bold=cfg.get("bold", True),
    )

    # Heading spacing
    pf = paragraph.paragraph_format
    pf.line_spacing = cfg.get("line_spacing", 1.5)
    pf.space_before = Pt(12)
    pf.space_after = Pt(6)

    return paragraph


# ---------------------------------------------------------------------------
# List handling
# ---------------------------------------------------------------------------

def add_list_item(doc, text, body_cfg, ordered=False, level=0):
    """Add a list item paragraph."""
    paragraph = doc.add_paragraph()
    prefix = f"{'  ' * level}{'• ' if not ordered else ''}"
    run = paragraph.add_run(prefix + text)
    set_run_font_cn(
        run,
        body_cfg["font"],
        body_cfg.get("font_west", "Times New Roman"),
        body_cfg.get("size", 12),
    )
    pf = paragraph.paragraph_format
    pf.line_spacing = body_cfg.get("line_spacing", 1.5)
    pf.left_indent = Cm(1.0 + level * 0.5)
    return paragraph


# ---------------------------------------------------------------------------
# Markdown parser
# ---------------------------------------------------------------------------

def parse_markdown_content(markdown_text):
    """Parse markdown text into a list of (type, level, text, extra) tuples.

    Types: 'heading', 'paragraph', 'display_math', 'list_item', 'blank'
    """
    lines = markdown_text.split('\n')
    blocks = []
    i = 0

    while i < len(lines):
        line = lines[i]

        # Display math $$...$$
        if line.strip().startswith('$$'):
            math_lines = []
            i += 1
            while i < len(lines):
                if lines[i].strip().endswith('$$') or lines[i].strip() == '$$':
                    math_lines.append(lines[i].rstrip('$$').strip())
                    break
                math_lines.append(lines[i])
                i += 1
            blocks.append(('display_math', 0, '\n'.join(math_lines), None))
            i += 1
            continue

        # Heading
        heading_match = re.match(r'^(#{1,4})\s+(.+)', line)
        if heading_match:
            level = len(heading_match.group(1))
            text = heading_match.group(2).strip()
            # Remove Markdown bold/italic markers from heading text
            text = re.sub(r'\*\*(.+?)\*\*', r'\1', text)
            text = re.sub(r'\*(.+?)\*', r'\1', text)
            blocks.append(('heading', level, text, None))
            i += 1
            continue

        # Unordered list
        list_match = re.match(r'^(\s*)[-*]\s+(.+)', line)
        if list_match:
            indent = len(list_match.group(1))
            level = indent // 2
            text = list_match.group(2).strip()
            blocks.append(('list_item', level, text, None))
            i += 1
            continue

        # Ordered list
        ordered_match = re.match(r'^(\s*)\d+[.)]\s+(.+)', line)
        if ordered_match:
            indent = len(ordered_match.group(1))
            level = indent // 2
            text = ordered_match.group(2).strip()
            blocks.append(('list_item', level, text, None))
            i += 1
            continue

        # Blank line
        if not line.strip():
            blocks.append(('blank', 0, '', None))
            i += 1
            continue

        # Regular paragraph
        text = line.strip()
        blocks.append(('paragraph', 0, text, None))
        i += 1

    return blocks


def process_inline_formatting(paragraph, text, body_cfg):
    """Process bold (**text**) and italic (*text*) within a paragraph, handling inline LaTeX."""
    # Split on bold and italic markers
    # First, protect inline LaTeX $...$
    latex_spans = []
    def save_latex(m):
        latex_spans.append(m.group(0))
        return f'\x00LATEX{len(latex_spans) - 1}\x00'
    text = re.sub(r'\$(.+?)\$', save_latex, text)

    # Now process bold and italic
    segments = re.split(r'(\*\*\*(.+?)\*\*\*|\*\*(.+?)\*\*|\*(.+?)\*)', text)

    for seg in segments:
        if not seg or seg.startswith('***') or seg.startswith('**') or seg.startswith('*'):
            continue

        # Check for bold+italic (***text***)
        bi_match = re.match(r'\*\*\*(.+?)\*\*\*', seg)
        # Check for bold (**text**)
        b_match = re.match(r'\*\*(.+?)\*\*', seg)
        # Check for italic (*text*)
        i_match = re.match(r'\*(.+?)\*', seg)

        content = seg
        bold = False
        italic = False

        if bi_match:
            content = bi_match.group(1)
            bold = True
            italic = True
        elif b_match:
            content = b_match.group(1)
            bold = True
        elif i_match:
            content = i_match.group(1)
            italic = True

        # Restore and process inline LaTeX
        if '\x00LATEX' in content:
            parts = re.split(r'\x00LATEX(\d+)\x00', content)
            for j, part in enumerate(parts):
                if j % 2 == 0:
                    if part:
                        run = paragraph.add_run(part)
                        set_run_font_cn(run, body_cfg["font"],
                                       body_cfg.get("font_west", "Times New Roman"),
                                       body_cfg.get("size", 12),
                                       bold=bold, italic=italic)
                else:
                    idx = int(part)
                    if idx < len(latex_spans):
                        latex_str = latex_spans[idx].strip('$')
                        insert_omml_equation(paragraph, latex_str, display=False)
        else:
            run = paragraph.add_run(content)
            set_run_font_cn(run, body_cfg["font"],
                           body_cfg.get("font_west", "Times New Roman"),
                           body_cfg.get("size", 12),
                           bold=bold, italic=italic)


# ---------------------------------------------------------------------------
# Main document builder
# ---------------------------------------------------------------------------

def build_document(config, markdown_text):
    """Build a python-docx Document from config and markdown."""
    doc = Document()

    # Page setup
    page_cfg = config.get("page", {})
    setup_page(doc, page_cfg)

    # Extract config sections
    headings_cfg = config.get("headings", {})
    body_cfg = config.get("body", {})
    eq_cfg = config.get("equations", {})

    # Parse markdown
    blocks = parse_markdown_content(markdown_text)

    # Heading counters
    counters = [0, 0, 0, 0]

    # Build paragraphs
    prev_was_blank = True  # Track paragraph spacing

    for block_type, level, text, extra in blocks:
        if block_type == 'blank':
            prev_was_blank = True
            continue

        if block_type == 'heading':
            if not prev_was_blank:
                doc.add_paragraph()  # spacing before heading
            add_heading(doc, level, text, headings_cfg, counters)
            prev_was_blank = False

        elif block_type == 'display_math':
            paragraph = doc.add_paragraph()
            if eq_cfg.get("display") == "center":
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            insert_omml_equation(paragraph, text.strip(), display=True)
            prev_was_blank = False

        elif block_type == 'list_item':
            add_list_item(doc, text, body_cfg, ordered=False, level=level)
            prev_was_blank = False

        elif block_type == 'paragraph':
            paragraph = doc.add_paragraph()
            set_paragraph_spacing(
                paragraph,
                body_cfg.get("line_spacing", 1.5),
                body_cfg.get("first_line_indent"),
                body_cfg.get("alignment", "justify"),
            )
            process_inline_formatting(paragraph, text, body_cfg)
            prev_was_blank = False

    return doc


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Generate .docx from Markdown + LaTeX")
    parser.add_argument("--output", required=True, help="Output .docx file path")
    parser.add_argument("--config", required=True, help="JSON config file with format settings")
    parser.add_argument("--content", required=True, help="Markdown content file")
    args = parser.parse_args()

    # Load config
    with open(args.config, 'r', encoding='utf-8') as f:
        config = json.load(f)

    # Load content
    with open(args.content, 'r', encoding='utf-8') as f:
        markdown_text = f.read()

    # Build document
    doc = build_document(config, markdown_text)

    # Save
    doc.save(args.output)
    print(f"Document saved to: {args.output}")
    print(f"Please open in Word to verify formatting and equations.")


if __name__ == "__main__":
    main()
