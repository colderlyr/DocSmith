#!/usr/bin/env python3
"""
DocSmith — Generate .docx from Markdown with native OMML equations, Chinese fonts,
table support, equation numbering, and cross-references.

Usage:
    python3 generate_docx.py --output out.docx --config config.json --content content.md
"""

import json
import re
import argparse
from docx import Document
from docx.shared import Pt, Cm, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from lxml import etree

try:
    from latex2word import LatexToWordElement
    HAS_LATEX2WORD = True
except ImportError:
    HAS_LATEX2WORD = False

MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"

# ---------------------------------------------------------------------------
# Chinese font size → pt
# ---------------------------------------------------------------------------
CN_FONT_SIZES = {
    "初号": 42, "小初": 36, "一号": 26, "小一": 24,
    "二号": 22, "小二": 18, "三号": 16, "小三": 15,
    "四号": 14, "小四": 12, "五号": 10.5, "小五": 9,
}

def to_pt(size):
    """Resolve font size: '小四'→12, '12'→12, 12→12."""
    if isinstance(size, (int, float)):
        return float(size)
    if size in CN_FONT_SIZES:
        return CN_FONT_SIZES[size]
    try:
        return float(size)
    except ValueError:
        return 12.0

# ---------------------------------------------------------------------------
# Font & paragraph utilities
# ---------------------------------------------------------------------------

def set_run_font(run, font_cn, font_west, size_pt, bold=False, italic=False, color=None):
    run.font.size = Pt(size_pt)
    run.font.name = font_west
    run.bold = bold
    run.italic = italic
    if color:
        run.font.color.rgb = RGBColor(*color)
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), font_cn)
    rFonts.set(qn('w:ascii'), font_west)
    rFonts.set(qn('w:hAnsi'), font_west)
    rFonts.set(qn('w:cs'), font_west)

def set_para_fmt(para, line_spacing=1.5, first_indent=None, alignment=None):
    pf = para.paragraph_format
    pf.line_spacing = float(line_spacing)
    if first_indent == "2chars":
        pf.first_line_indent = Cm(0.74)
    elif first_indent:
        pf.first_line_indent = first_indent
    amap = {"left": 0, "center": 1, "right": 2, "justify": 3}
    if alignment in amap:
        para.alignment = amap[alignment]

def add_run(para, text, font_cn, font_west, size, bold=False, italic=False):
    run = para.add_run(text)
    set_run_font(run, font_cn, font_west, to_pt(size), bold, italic)
    return run

# ---------------------------------------------------------------------------
# OMML Equations
# ---------------------------------------------------------------------------

def build_ommath_para(omml_element):
    """Wrap an m:oMath element in a properly structured m:oMathPara (display eq)."""
    omp = etree.Element(f"{{{MATH_NS}}}oMathPara")
    omp_pr = etree.SubElement(omp, f"{{{MATH_NS}}}oMathParaPr")
    jc = etree.SubElement(omp_pr, f"{{{MATH_NS}}}jc")
    jc.set(f"{{{MATH_NS}}}val", "center")
    om = etree.SubElement(omp, f"{{{MATH_NS}}}oMath")
    for child in omml_element:
        om.append(child)
    return omp


def append_omml(para, latex_str, display=True):
    """Convert LaTeX to OMML and append to paragraph.
    Returns the OMML element (oMathPara for display, oMath for inline)."""
    if not HAS_LATEX2WORD:
        run = para.add_run(f"[Equation: {latex_str}]")
        run.font.color.rgb = RGBColor(0xCC, 0x00, 0x00)
        return None
    try:
        eq = LatexToWordElement(latex_str)
        omml = eq.element()
        if display:
            omp = build_ommath_para(omml)
            para._element.append(omp)
            return omp
        else:
            para._element.append(omml)
            return omml
    except Exception as e:
        run = para.add_run(f"[Eq: {latex_str[:40]}]")
        run.font.color.rgb = RGBColor(0xCC, 0x00, 0x00)
        return None


def add_display_eq(para, latex_str, eq_cfg, eq_num, ctx):
    """Add a display equation. Numbered: tab-stop layout (eq centered, number right).
    Unnumbered: paragraph centered, OMML only."""
    if eq_cfg.get("numbering"):
        # Tab-stop layout: [TAB→center] [equation] [TAB→right] [(1)]
        para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf = para.paragraph_format
        pf.tab_stops.add_tab_stop(Cm(7.96), alignment=WD_ALIGN_PARAGRAPH.CENTER)
        pf.tab_stops.add_tab_stop(Cm(15.92), alignment=WD_ALIGN_PARAGRAPH.RIGHT)
        # TAB run (positions cursor at center tab)
        para.add_run("\t")
        # OMML equation
        append_omml(para, latex_str.strip(), display=True)
        # TAB + number run (positions at right tab)
        fmt = eq_cfg.get("numbering_format", "({n})")
        num_text = fmt.replace("{n}", str(eq_num))
        run = para.add_run(f"\t{num_text}")
        nfont = eq_cfg.get("numbering_font", "Times New Roman")
        nsize = to_pt(eq_cfg.get("numbering_size", 12))
        set_run_font(run, nfont, nfont, nsize)
        bm_id = ctx.next_bookmark(f"eq{eq_num}")
        add_bookmark_to_run(run, f"eq{eq_num}", bm_id)
    else:
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        append_omml(para, latex_str.strip(), display=True)

def add_bookmark_to_run(run, name, bmid):
    """Append bookmark start/end around the run's element."""
    parent = run._r.getparent()
    idx = list(parent).index(run._r)
    bs = OxmlElement('w:bookmarkStart')
    bs.set(qn('w:id'), str(bmid))
    bs.set(qn('w:name'), name)
    be = OxmlElement('w:bookmarkEnd')
    be.set(qn('w:id'), str(bmid))
    parent.insert(idx, bs)
    parent.insert(idx + 2, be)

def add_bookmark_to_para(para, name, bmid):
    bs = OxmlElement('w:bookmarkStart')
    bs.set(qn('w:id'), str(bmid))
    bs.set(qn('w:name'), name)
    be = OxmlElement('w:bookmarkEnd')
    be.set(qn('w:id'), str(bmid))
    para._p.append(bs)
    para._p.append(be)

def add_internal_hyperlink(para, anchor, display_text, font_cn, font_west, size):
    """Add a clickable internal cross-reference like '表 1'."""
    hl = OxmlElement('w:hyperlink')
    hl.set(qn('w:anchor'), anchor)
    r = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    rs = OxmlElement('w:rStyle')
    rs.set(qn('w:val'), 'Hyperlink')
    rPr.append(rs)
    # Font size for hyperlink
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), str(int(to_pt(size) * 2)))
    rPr.append(sz)
    r.append(rPr)
    t = OxmlElement('w:t')
    t.text = display_text
    t.set(qn('xml:space'), 'preserve')
    r.append(t)
    hl.append(r)
    para._p.append(hl)

# ---------------------------------------------------------------------------
# Page setup
# ---------------------------------------------------------------------------

def setup_page(doc, cfg):
    for s in doc.sections:
        smap = {"A4": (Cm(21.0), Cm(29.7)), "letter": (Cm(21.59), Cm(27.94))}
        if cfg.get("size") in smap:
            s.page_width, s.page_height = smap[cfg["size"]]
        s.top_margin = Cm(cfg.get("margin_top", 2.54))
        s.bottom_margin = Cm(cfg.get("margin_bottom", 2.54))
        s.left_margin = Cm(cfg.get("margin_left", 2.54))
        s.right_margin = Cm(cfg.get("margin_right", 2.54))

# ---------------------------------------------------------------------------
# Heading
# ---------------------------------------------------------------------------

def heading_prefix(level, cfg, ctx):
    key = f"h{level}"
    style = cfg.get(key, {}).get("numbering", "")
    if not style:
        return ""
    ctx.heading_counters[level - 1] += 1
    for i in range(level, len(ctx.heading_counters)):
        ctx.heading_counters[i] = 0
    c = ctx.heading_counters
    if style == "1.":
        return ".".join(str(c[i]) for i in range(level) if c[i] > 0) + (". " if level == 1 else " ")
    if style == "一、":
        cn = ["", "一", "二", "三", "四", "五", "六", "七", "八", "九", "十"]
        return f"{cn[c[0]] if c[0] < len(cn) else str(c[0])}、"
    if style == "（一）":
        cn = ["", "一", "二", "三", "四", "五"]
        return f"（{cn[c[1]] if c[1] < len(cn) else str(c[1])}）"
    return ""

def add_heading(doc, level, text, headings_cfg, ctx):
    key = f"h{level}"
    hc = headings_cfg.get(key)
    if not hc:
        return
    prefix = heading_prefix(level, headings_cfg, ctx)
    para = doc.add_paragraph()
    run = para.add_run(prefix + text)
    set_run_font(run, hc.get("font", "SimHei"), hc.get("font_west", "Arial"),
                 to_pt(hc.get("size", 16)), bold=hc.get("bold", True))
    pf = para.paragraph_format
    pf.line_spacing = hc.get("line_spacing", 1.5)
    pf.space_before = Pt(12)
    pf.space_after = Pt(6)
    return para

# ---------------------------------------------------------------------------
# Tables
# ---------------------------------------------------------------------------

def add_table(doc, headers, rows, table_cfg, caption_text, ctx, label=None):
    """Create a formatted Word table with optional caption."""
    ctx.table_counter += 1
    tnum = ctx.table_counter
    anchor = f"tab{tnum}"

    # Register label if provided
    if label:
        ctx.register_label(label, tnum, anchor)

    # Caption (above table)
    if caption_text and table_cfg.get("caption_position", "above") == "above":
        add_table_caption(doc, caption_text, table_cfg, tnum, ctx, anchor)

    # Build table
    ncols = len(headers)
    table = doc.add_table(rows=1 + len(rows), cols=ncols)
    table.style = 'Table Grid'

    # Header row
    hfont = table_cfg.get("header_font", "SimHei")
    hfont_w = table_cfg.get("header_font_west", "Arial")
    hsize = to_pt(table_cfg.get("header_size", "小四"))
    hbold = table_cfg.get("header_bold", True)
    for j, h in enumerate(headers):
        cell = table.rows[0].cells[j]
        cell.text = ""
        p = cell.paragraphs[0]
        run = p.add_run(str(h))
        set_run_font(run, hfont, hfont_w, hsize, bold=hbold)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        shading = OxmlElement('w:shd')
        shading.set(qn('w:fill'), 'D9E2F3')
        shading.set(qn('w:val'), 'clear')
        cell._tc.get_or_add_tcPr().append(shading)

    # Data rows
    bfont = table_cfg.get("body_font", "SimSun")
    bfont_w = table_cfg.get("body_font_west", "Times New Roman")
    bsize = to_pt(table_cfg.get("body_size", "小五"))

    for i, row_data in enumerate(rows):
        for j, val in enumerate(row_data):
            if j >= ncols:
                break
            cell = table.rows[i + 1].cells[j]
            cell.text = ""
            p = cell.paragraphs[0]
            # Check for inline LaTeX in cell
            if '$' in str(val):
                _add_simple_text(p, str(val), bfont, bfont_w, bsize)
            else:
                run = p.add_run(str(val))
                set_run_font(run, bfont, bfont_w, bsize)

    # Caption (below table)
    if caption_text and table_cfg.get("caption_position", "above") != "above":
        add_table_caption(doc, caption_text, table_cfg, tnum, ctx, anchor)

    return table

def add_table_caption(doc, caption_text, table_cfg, tnum, ctx, anchor=None):
    """Add '表 1 标题文本' caption paragraph."""
    if anchor is None:
        anchor = f"tab{tnum}"
    num_fmt = table_cfg.get("numbering", "表{n} ")
    prefix = num_fmt.replace("{n}", str(tnum))
    para = doc.add_paragraph()
    run = para.add_run(prefix + caption_text)
    cfont = table_cfg.get("caption_font", "SimHei")
    cfont_w = table_cfg.get("caption_font_west", "Arial")
    csize = to_pt(table_cfg.get("caption_size", "小四"))
    cbold = table_cfg.get("caption_bold", True)
    set_run_font(run, cfont, cfont_w, csize, bold=cbold)
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    # Bookmark on caption (use anchor name for label-based lookup)
    add_bookmark_to_para(para, anchor, ctx.next_bookmark(anchor))
    pf = para.paragraph_format
    pf.space_before = Pt(6)
    pf.space_after = Pt(6)

# ---------------------------------------------------------------------------
# Cross-reference processing
# ---------------------------------------------------------------------------

def resolve_cross_ref(ref_str, body_cfg, ctx):
    """Parse \\ref{type:id} and return (display_text, anchor) or (None, None)."""
    m = re.match(r'\\ref\{(tab|eq):(\S+?)\}', ref_str)
    if not m:
        return None, None
    ref_type, ref_id = m.group(1), m.group(2)
    xref = body_cfg.get("cross_ref", {})
    if ref_type == "tab":
        if ref_id in ctx.label_map:
            number, anchor = ctx.label_map[ref_id]
        else:
            number, anchor = ref_id, f"tab{ref_id}"
        display = f"{xref.get('table_prefix', '表')} {number}"
        return display, anchor
    elif ref_type == "eq":
        anchor = f"eq{ref_id}"
        fmt = body_cfg.get("equations", {}).get("numbering_format", "({n})")
        display = f"{xref.get('equation_prefix', '公式')} {fmt.replace('{n}', ref_id)}"
        return display, anchor
    return None, None


def _add_simple_text(para, text, font_cn, font_west, size, bold=False, italic=False):
    """Add text with optional inline LaTeX ($...$) — no cross-ref or bold/italic parsing."""
    parts = re.split(r'(?<!\$)\$(?!\$)(.+?)(?<!\$)\$(?!\$)', text)
    for i, part in enumerate(parts):
        if i % 2 == 0:
            if part:
                run = para.add_run(part)
                set_run_font(run, font_cn, font_west, to_pt(size), bold, italic)
        else:
            if part.strip():
                append_omml(para, part.strip(), display=False)
    """Parse \\ref{type:id} and return (display_text, anchor) or (None, None)."""
    m = re.match(r'\\ref\{(tab|eq):(\S+?)\}', ref_str)
    if not m:
        return None, None
    ref_type, ref_id = m.group(1), m.group(2)
    xref = body_cfg.get("cross_ref", {})
    if ref_type == "tab":
        if ref_id in ctx.label_map:
            number, anchor = ctx.label_map[ref_id]
        else:
            number, anchor = ref_id, f"tab{ref_id}"
        display = f"{xref.get('table_prefix', '表')} {number}"
        return display, anchor
    elif ref_type == "eq":
        anchor = f"eq{ref_id}"
        fmt = body_cfg.get("equations", {}).get("numbering_format", "({n})")
        display = f"{xref.get('equation_prefix', '公式')} {fmt.replace('{n}', ref_id)}"
        return display, anchor
    return None, None


def process_inline_formatting(para, text, body_cfg, ctx):
    """Process cross-refs, inline LaTeX, and bold/italic in a single ordered pass."""
    bfont = body_cfg.get("font", "SimSun")
    bfont_w = body_cfg.get("font_west", "Times New Roman")
    bsize = body_cfg.get("size", 12)

    # Split text into segments: \ref{...}, $latex$, and plain text
    # Use a combined pattern to split while keeping delimiters
    segments = re.split(r'(\\ref\{(?:tab|eq):\S+?\}|\$.+?\$)', text)

    for seg in segments:
        if not seg:
            continue

        # Cross-reference
        if seg.startswith('\\ref{'):
            disp, anchor = resolve_cross_ref(seg, body_cfg, ctx)
            if disp:
                add_internal_hyperlink(para, anchor, disp, bfont, bfont_w, bsize)

        # Inline LaTeX
        elif seg.startswith('$') and seg.endswith('$'):
            latex_str = seg.strip('$')
            if latex_str:
                append_omml(para, latex_str, display=False)

        # Plain text (may contain bold/italic markers)
        else:
            # Process bold/italic within this text segment
            _add_formatted_text(para, seg, bfont, bfont_w, bsize)


def _add_formatted_text(para, text, font_cn, font_west, size):
    """Add text runs with bold/italic support."""
    parts = re.split(r'(\*\*\*(.+?)\*\*\*|\*\*(.+?)\*\*|\*(.+?)\*)', text)
    for part in parts:
        if not part:
            continue
        content = part
        bold = False
        italic = False
        bi = re.match(r'\*\*\*(.+?)\*\*\*', part)
        bm = re.match(r'\*\*(.+?)\*\*', part)
        im = re.match(r'\*(.+?)\*', part)
        if bi:
            content, bold, italic = bi.group(1), True, True
        elif bm:
            content, bold = bm.group(1), True
        elif im:
            content, italic = im.group(1), True

        run = para.add_run(content)
        set_run_font(run, font_cn, font_west, to_pt(size), bold, italic)

# ---------------------------------------------------------------------------
# Context tracker
# ---------------------------------------------------------------------------

class DocContext:
    def __init__(self):
        self.heading_counters = [0, 0, 0, 0]
        self.table_counter = 0
        self.equation_counter = 0
        self.figure_counter = 0
        self._bookmark_id = 0
        self.label_map = {}  # label → (number, bookmark_anchor)

    def next_bookmark(self, name):
        self._bookmark_id += 1
        return self._bookmark_id

    def register_label(self, label, number, anchor):
        """Map a user label (e.g. 'exp_params') to its number and bookmark anchor."""
        self.label_map[label] = (number, anchor)

# ---------------------------------------------------------------------------
# Markdown Parser
# ---------------------------------------------------------------------------

def parse_markdown(text):
    """Parse markdown into blocks. Returns list of (type, level, text, meta)."""
    lines = text.split('\n')
    blocks = []
    i = 0

    while i < len(lines):
        line = lines[i]

        # Display math $$...$$
        if line.strip().startswith('$$'):
            parts = []
            i += 1
            while i < len(lines):
                if lines[i].strip() == '$$' or lines[i].strip().endswith('$$'):
                    parts.append(lines[i].strip().rstrip('$$').strip())
                    break
                parts.append(lines[i])
                i += 1
            blocks.append(('display_math', 0, '\n'.join(parts), None))
            i += 1
            continue

        # Table: collect consecutive |...| lines
        if re.match(r'^\|.+\|$', line.strip()):
            table_lines = []
            while i < len(lines) and re.match(r'^\|.+\|$', lines[i].strip()):
                table_lines.append(lines[i].strip())
                i += 1
            # Check if next line is a caption [表:label] text (skip blank lines)
            caption = ""
            label = None
            while i < len(lines) and not lines[i].strip():
                i += 1
            if i < len(lines):
                cap_match = re.match(r'^\[表:(\w+)\]\s*(.+)', lines[i].strip())
                if cap_match:
                    label = cap_match.group(1)
                    caption = cap_match.group(2).strip()
                    i += 1
            blocks.append(('table', 0, '', {
                'lines': table_lines, 'caption': caption, 'label': label
            }))
            continue

        # Table-only caption line (without preceding table lines)
        cap_match = re.match(r'^\[表:(\w+)\]\s*(.+)', line.strip())
        if cap_match:
            blocks.append(('table_caption', 0, '', {
                'label': cap_match.group(1), 'caption': cap_match.group(2).strip()
            }))
            i += 1
            continue

        # Heading
        hm = re.match(r'^(#{1,4})\s+(.+)', line)
        if hm:
            lvl = len(hm.group(1))
            txt = hm.group(2).strip()
            txt = re.sub(r'\*\*(.+?)\*\*', r'\1', txt)
            txt = re.sub(r'\*(.+?)\*', r'\1', txt)
            blocks.append(('heading', lvl, txt, None))
            i += 1
            continue

        # List
        lm = re.match(r'^(\s*)[-*]\s+(.+)', line)
        if lm:
            indent = len(lm.group(1))
            blocks.append(('list_item', indent // 2, lm.group(2).strip(), None))
            i += 1
            continue

        # Blank
        if not line.strip():
            blocks.append(('blank', 0, '', None))
            i += 1
            continue

        # Paragraph
        blocks.append(('paragraph', 0, line.strip(), None))
        i += 1

    return blocks

# ---------------------------------------------------------------------------
# Document builder
# ---------------------------------------------------------------------------

def build_document(config, md_text):
    doc = Document()
    setup_page(doc, config.get("page", {}))
    headings_cfg = config.get("headings", {})
    body_cfg = config.get("body", {})
    eq_cfg = config.get("equations", {})
    table_cfg = config.get("table", {})
    ctx = DocContext()
    blocks = parse_markdown(md_text)

    # ---- Pre-scan: register all table labels with forward-looking numbers ----
    tnum = 0
    for btype, level, text, meta in blocks:
        if btype == 'table' and meta.get('label'):
            tnum += 1
            ctx.register_label(meta['label'], tnum, f"tab{tnum}")
    # -------------------------------------------------------------------------

    prev_blank = True

    for btype, level, text, meta in blocks:
        if btype == 'blank':
            prev_blank = True
            continue

        if btype == 'heading':
            if not prev_blank:
                doc.add_paragraph()
            add_heading(doc, level, text, headings_cfg, ctx)
            prev_blank = False

        elif btype == 'display_math':
            ctx.equation_counter += 1
            para = doc.add_paragraph()
            add_display_eq(para, text.strip(), eq_cfg, ctx.equation_counter, ctx)
            prev_blank = False

        elif btype == 'table':
            lines = meta['lines']
            if len(lines) < 2:
                continue
            # Parse header and rows, skip separator line
            def parse_row(r):
                return [c.strip() for c in r.strip('|').split('|')]
            headers = parse_row(lines[0])
            if re.match(r'^[\|\s\-:]+$', lines[1]):
                rows = [parse_row(r) for r in lines[2:]]
            else:
                rows = [parse_row(r) for r in lines[1:]]
            add_table(doc, headers, rows, table_cfg, meta.get('caption', ''),
                       ctx, label=meta.get('label'))
            prev_blank = False

        elif btype == 'list_item':
            para = doc.add_paragraph()
            prefix = f"{'  ' * level}• "
            _add_simple_text(para, prefix + text,
                                body_cfg.get("font", "SimSun"),
                                body_cfg.get("font_west", "Times New Roman"),
                                body_cfg.get("size", 12), ctx=ctx)
            pf = para.paragraph_format
            pf.line_spacing = body_cfg.get("line_spacing", 1.5)
            pf.left_indent = Cm(1.0 + level * 0.5)
            prev_blank = False

        elif btype == 'paragraph':
            para = doc.add_paragraph()
            set_para_fmt(para, body_cfg.get("line_spacing", 1.5),
                        body_cfg.get("first_line_indent"),
                        body_cfg.get("alignment", "justify"))
            process_inline_formatting(para, text, body_cfg, ctx)
            prev_blank = False

    return doc

# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main():
    p = argparse.ArgumentParser(description="DocSmith — MD+LaTeX → professional .docx")
    p.add_argument("--output", required=True)
    p.add_argument("--config", required=True)
    p.add_argument("--content", required=True)
    args = p.parse_args()
    with open(args.config, 'r', encoding='utf-8') as f:
        config = json.load(f)
    with open(args.content, 'r', encoding='utf-8') as f:
        md = f.read()
    doc = build_document(config, md)
    doc.save(args.output)
    print(f"OK — {args.output}")

if __name__ == "__main__":
    main()
