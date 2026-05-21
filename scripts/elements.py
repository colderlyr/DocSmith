"""Document elements: headings, tables, inline formatting, page setup."""

import re
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from .fonts import to_pt, set_run_font, set_para_fmt, add_run
from .omml import append_omml, add_internal_hyperlink, add_bookmark_to_para
from .parser import resolve_cross_ref


# ---------------------------------------------------------------------------
# Page setup
# ---------------------------------------------------------------------------

def setup_page(doc, cfg):
    """Configure page size and margins for all sections."""
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
    """Generate auto-numbering prefix like '1.', '1.1', '一、' etc."""
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
    """Add a heading paragraph bound to Word's Heading N style."""
    key = f"h{level}"
    hc = headings_cfg.get(key)
    if not hc:
        return
    prefix = heading_prefix(level, headings_cfg, ctx)
    style_name = f"Heading {level}"
    para = doc.add_paragraph(style=style_name)
    run = para.add_run(prefix + text)
    set_run_font(run, hc.get("font", "SimHei"), hc.get("font_west", "Arial"),
                 to_pt(hc.get("size", 16)), bold=hc.get("bold", True))
    pf = para.paragraph_format
    pf.line_spacing = hc.get("line_spacing", 1.5)
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    return para


# ---------------------------------------------------------------------------
# Tables
# ---------------------------------------------------------------------------

def add_table(doc, headers, rows, table_cfg, caption_text, ctx, label=None):
    """Create a formatted Word table with optional caption (no background shading)."""
    ctx.table_counter += 1
    tnum = ctx.table_counter
    anchor = f"tab{tnum}"

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
    """Add '表 1 标题文本' caption paragraph bound to Word's Caption style."""
    if anchor is None:
        anchor = f"tab{tnum}"
    num_fmt = table_cfg.get("numbering", "表 {n} ")
    prefix = num_fmt.replace("{n}", str(tnum))
    para = doc.add_paragraph(style='Caption')
    run = para.add_run(prefix + caption_text)
    cfont = table_cfg.get("caption_font", "SimHei")
    cfont_w = table_cfg.get("caption_font_west", "Arial")
    csize = to_pt(table_cfg.get("caption_size", "小四"))
    cbold = table_cfg.get("caption_bold", True)
    set_run_font(run, cfont, cfont_w, csize, bold=cbold)
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_bookmark_to_para(para, anchor, ctx.next_bookmark(anchor))
    pf = para.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)


# ---------------------------------------------------------------------------
# Inline formatting (cross-refs, LaTeX, bold/italic)
# ---------------------------------------------------------------------------

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


def process_inline_formatting(para, text, body_cfg, ctx):
    """Process cross-refs, inline LaTeX, and bold/italic in a single ordered pass."""
    bfont = body_cfg.get("font", "SimSun")
    bfont_w = body_cfg.get("font_west", "Times New Roman")
    bsize = body_cfg.get("size", 12)

    # Split text into segments: \ref{...}, $latex$, and plain text
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
            _add_formatted_text(para, seg, bfont, bfont_w, bsize)


def _add_formatted_text(para, text, font_cn, font_west, size):
    """Add text runs with **bold** and *italic* support."""
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
