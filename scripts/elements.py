"""Document elements: page setup, headings, tables, inline formatting.

Supports:
  - Single and multi-column page layouts (via OXML w:cols)
  - Single-column paragraph spans (w:cnt) for title/abstract areas
  - Column-width-aware tab stops for equation numbering
  - Auto-numbered headings with configurable numbering styles:
    "1.", "1.1", "一、", "（一）", "I.", "II.", "A.", "B."
  - Tables with captions and cross-reference bookmarks
  - Reference sections with hanging indents
  - Abstract and acknowledgment sections with special formatting
"""

import re
from docx.shared import Pt, Cm, Inches, RGBColor
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
    """Configure page size, margins, and column layout for all sections.

    cfg keys:
      size        : "A4" or "letter" (default A4)
      margin_*    : top, bottom, left, right in cm (default 2.54)
      columns     : None (single column, default) or int (2 for two-column)
      column_gap  : gap between columns in cm (default 0.63 for 0.25 inches)

    Two-column layout uses OXML w:cols injection since python-docx
    does not natively support multi-column sections.
    """
    for s in doc.sections:
        smap = {"A4": (Cm(21.0), Cm(29.7)), "letter": (Cm(21.59), Cm(27.94))}
        if cfg.get("size") in smap:
            s.page_width, s.page_height = smap[cfg["size"]]
        s.top_margin = Cm(cfg.get("margin_top", 2.54))
        s.bottom_margin = Cm(cfg.get("margin_bottom", 2.54))
        s.left_margin = Cm(cfg.get("margin_left", 2.54))
        s.right_margin = Cm(cfg.get("margin_right", 2.54))

        # Multi-column layout
        num_cols = cfg.get("columns", None)
        if num_cols and num_cols > 1:
            sectPr = s._sectPr
            cols = OxmlElement('w:cols')
            cols.set(qn('w:num'), str(num_cols))
            cols.set(qn('w:equalWidth'), '1')
            # Column gap: default 0.63cm (~0.25 inch)
            gap_cm = cfg.get("column_gap", 0.63)
            gap_twips = str(int(gap_cm * 567))  # 1 cm ≈ 567 twips
            cols.set(qn('w:space'), gap_twips)
            for e in sectPr.findall(qn('w:cols')):
                sectPr.remove(e)
            sectPr.append(cols)


def get_column_width_cm(cfg):
    """Compute usable column width in cm from page config.
    Used to calculate correct equation tab stops."""
    page_size = cfg.get("size", "A4")
    ml = cfg.get("margin_left", 2.54)
    mr = cfg.get("margin_right", 2.54)
    num_cols = cfg.get("columns", 1) or 1
    col_gap = cfg.get("column_gap", 0.63)

    page_widths = {"A4": 21.0, "letter": 21.59}
    full_width = page_widths.get(page_size, 21.0)
    usable = full_width - ml - mr

    if num_cols > 1:
        return (usable - col_gap * (num_cols - 1)) / num_cols
    return usable


def set_para_single_column(para):
    """Force paragraph to span all columns (for title, authors, abstract).

    Uses the w:cnt (continuous) element to break out of multi-column flow.
    Paragraphs after this one will resume normal column flow.
    """
    pPr = para._p.get_or_add_pPr()
    # Remove any existing cnt element to avoid duplicates
    for old in pPr.findall(qn('w:cnt')):
        pPr.remove(old)
    cnt = OxmlElement('w:cnt')
    pPr.append(cnt)


# ---------------------------------------------------------------------------
# Heading numbering families
# ---------------------------------------------------------------------------

def _roman(n):
    """Convert integer to uppercase Roman numeral."""
    vals = [(1000, 'M'), (900, 'CM'), (500, 'D'), (400, 'CD'),
            (100, 'C'), (90, 'XC'), (50, 'L'), (40, 'XL'),
            (10, 'X'), (9, 'IX'), (5, 'V'), (4, 'IV'), (1, 'I')]
    result = ''
    for v, r in vals:
        while n >= v:
            result += r
            n -= v
    return result


def _letter(n):
    """Convert integer to uppercase letter: 1→A, 2→B, ..., 26→Z, 27→AA."""
    result = ''
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result = chr(ord('A') + rem) + result
    return result


def _cn_num(n):
    """Convert integer 1-10 to Chinese numeral; >10 returns str(n)."""
    cn = ["", "一", "二", "三", "四", "五", "六", "七", "八", "九", "十"]
    return cn[n] if n < len(cn) else str(n)


def _dot_fallback(counters, level):
    """Generate '1.', '1.1', '1.1.1' numbering from counters (generic fallback)."""
    parts = [str(counters[i]) for i in range(level) if counters[i] > 0]
    return ".".join(parts) + " "


def heading_prefix(level, cfg, ctx):
    """Generate auto-numbering prefix based on H1 numbering family.

    Supported families and their expansion:
      "1."      → H1: "1.", H2: "1.1", H3: "1.1.1", H4: "1.1.1.1"
      "一、"     → H1: "一、", H2: "（一）", H3+: dot fallback
      "I."      → H1: "I.", H2: "A.", H3: "1)", H4: "a)"
      "i."      → H1: "i.", H2: "a.", H3+: dot fallback
      "(1)"     → H1: "(1)", H2: "(a)", H3+: dot fallback
      "1)"      → H1: "1)", H2: "a)", H3: "i)", H4: dot fallback
      None/""   → no numbering

    Setting H2 numbering to something non-standard overrides the family expansion.
    """
    key = f"h{level}"
    hc = cfg.get(key, {})
    if not hc.get("numbering"):
        return ""

    # Increment counter for this level, reset deeper levels
    ctx.heading_counters[level - 1] += 1
    for i in range(level, len(ctx.heading_counters)):
        ctx.heading_counters[i] = 0
    c = ctx.heading_counters

    family = cfg.get("h1", {}).get("numbering", "")
    h2_numbering = cfg.get("h2", {}).get("numbering", "")

    # Allow h2 to override the family-derived expansion
    if level >= 2 and h2_numbering and h2_numbering != _derive_h2_family(family):
        result = _format_by_style(h2_numbering, c, level)
        if result:
            return result

    if family == "1.":
        return _dot_fallback(c, level)

    elif family == "一、":
        if level == 1:
            return f"{_cn_num(c[0])}、"
        elif level == 2:
            return f"（{_cn_num(c[1])}）"
        else:
            return _dot_fallback(c, level)

    elif family == "（一）":
        if level == 1:
            return f"（{_cn_num(c[0])}）"
        elif level == 2:
            return f"{c[1]}. "
        else:
            return _dot_fallback(c, level)

    elif family == "I.":
        if level == 1:
            return f"{_roman(c[0])}. "
        elif level == 2:
            return f"{_letter(c[1])}. "
        elif level == 3:
            return f"{c[2]}) "
        elif level == 4:
            return f"{_letter(c[3]).lower()}) "
        else:
            return _dot_fallback(c, level)

    elif family == "i.":
        if level == 1:
            return f"{_roman(c[0]).lower()}. "
        elif level == 2:
            return f"{_letter(c[1]).lower()}. "
        else:
            return _dot_fallback(c, level)

    elif family == "1)":
        if level == 1:
            return f"{c[0]}) "
        elif level == 2:
            return f"{_letter(c[1]).lower()}) "
        elif level == 3:
            return f"{_roman(c[2]).lower()}) "
        else:
            return _dot_fallback(c, level)

    elif family in ("(1)", "(a)"):
        result = _format_by_style(family, c, level)
        if result:
            return result
        return _dot_fallback(c, level)

    # Unknown family → dot fallback
    return _dot_fallback(c, level)


def _derive_h2_family(h1_family):
    """Derive the default H2 numbering from the H1 family."""
    mapping = {"1.": "1.1", "一、": "（一）", "I.": "A.", "i.": "a.",
               "(1)": "(a)", "1)": "a)", "（一）": "1."}
    return mapping.get(h1_family, "")


def _format_by_style(style, counters, level):
    """Format a counter for a specific numbering style.
    Handles: (1), (a), 1), a), i), 1., 1.1, A.
    Returns empty string for unrecognized styles."""
    c = counters
    idx = level - 1

    if style == "(1)":
        return f"({c[idx]}) "
    elif style == "(a)":
        return f"({_letter(c[idx]).lower()}) "
    elif style == "1)":
        return f"{c[idx]}) "
    elif style == "a)":
        return f"{_letter(c[idx]).lower()}) "
    elif style == "i)":
        return f"{_roman(c[idx]).lower()}) "
    elif style == "A.":
        return f"{_letter(c[idx])}. "
    elif style == "I.":
        return f"{_roman(c[idx])}. "
    elif re.match(r'^[\d.]+$', style):
        # Dot-separated numeric pattern: "1.", "1.1", etc.
        return _dot_fallback(counters, level)
    return ""


# ---------------------------------------------------------------------------
# Heading
# ---------------------------------------------------------------------------

def add_heading(doc, level, text, headings_cfg, ctx):
    """Add a heading paragraph bound to Word's Heading N style.
    Forces black font color (overrides Word's default blue) and no indent."""
    key = f"h{level}"
    hc = headings_cfg.get(key)
    if not hc:
        return
    prefix = heading_prefix(level, headings_cfg, ctx)
    style_name = f"Heading {level}"
    para = doc.add_paragraph(style=style_name)
    run = para.add_run(prefix + text)
    set_run_font(run, hc.get("font", "SimHei"), hc.get("font_west", "Arial"),
                 to_pt(hc.get("size", 16)), bold=hc.get("bold", True),
                 color=(0, 0, 0))
    pf = para.paragraph_format
    pf.line_spacing = hc.get("line_spacing", 1.5)
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.first_line_indent = Pt(0)
    pf.left_indent = Pt(0)
    return para


# ---------------------------------------------------------------------------
# Tables
# ---------------------------------------------------------------------------

def add_table(doc, headers, rows, table_cfg, caption_text, ctx, label=None):
    """Create a formatted Word table with optional caption."""
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
    """Add '表 1 标题文本' caption paragraph."""
    if anchor is None:
        anchor = f"tab{tnum}"
    num_fmt = table_cfg.get("numbering", "表 {n} ")
    prefix = num_fmt.replace("{n}", str(tnum))
    para = doc.add_paragraph(style='Caption')
    run = para.add_run(prefix + caption_text)
    cfont = table_cfg.get("caption_font", "黑体")
    cfont_w = table_cfg.get("caption_font_west", "Arial")
    csize = to_pt(table_cfg.get("caption_size", "小五"))
    cbold = table_cfg.get("caption_bold", True)
    set_run_font(run, cfont, cfont_w, csize, bold=cbold, color=(0, 0, 0))
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_bookmark_to_para(para, anchor, ctx.next_bookmark(anchor))
    pf = para.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)


# ---------------------------------------------------------------------------
# Inline formatting (cross-refs, LaTeX, bold/italic)
# ---------------------------------------------------------------------------

def _add_simple_text(para, text, font_cn, font_west, size, bold=False, italic=False):
    """Add text with optional inline LaTeX ($...$). size may be a float (pt) or
    a string like '小五' — to_pt handles both idempotently."""
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


# ---------------------------------------------------------------------------
# Abstract / Acknowledgment sections
# ---------------------------------------------------------------------------

def add_section_body_para(doc, text, body_cfg, section_type, ctx):
    """Add a body paragraph within an abstract or acknowledgment section.
    Uses body text formatting. The section heading already provides the label,
    so no prefix is prepended here."""
    para = doc.add_paragraph(style='Normal')
    bfont = body_cfg.get("font", "SimSun")
    bfont_w = body_cfg.get("font_west", "Times New Roman")
    bsize = body_cfg.get("size", 12)
    body_font_size_pt = to_pt(bsize)

    set_para_fmt(para, body_cfg.get("line_spacing", 1.5),
                 body_cfg.get("first_line_indent"),
                 body_cfg.get("alignment", "justify"),
                 font_size_pt=body_font_size_pt)

    process_inline_formatting(para, text, body_cfg, ctx)
    return para


# ---------------------------------------------------------------------------
# References (GB/T 7714)
# ---------------------------------------------------------------------------

def add_reference_section(doc, ref_items, body_cfg):
    """Add a GB/T 7714 formatted reference list at end of document.
    Title '参考文献' uses Heading 1 style (black, no indent).
    Each reference is a Normal paragraph with hanging indent."""
    headings_cfg = {
        "h1": {"font": "SimHei", "font_west": "Arial", "size": "三号", "bold": True}
    }
    from .parser import DocContext
    ctx = DocContext()
    # "参考文献" as Heading 2 (matching ## markdown level)
    level = 2
    style_name = f"Heading {level}"
    para = doc.add_paragraph(style=style_name)
    hc = headings_cfg["h1"]
    run = para.add_run("参考文献")
    set_run_font(run, hc.get("font", "SimHei"), hc.get("font_west", "Arial"),
                 to_pt(hc.get("size", 16)), bold=hc.get("bold", True),
                 color=(0, 0, 0))
    pf = para.paragraph_format
    pf.line_spacing = 1.5
    pf.space_before = Pt(12)
    pf.space_after = Pt(6)
    pf.first_line_indent = Pt(0)
    pf.left_indent = Pt(0)

    # Each reference as Normal paragraph with hanging indent
    for item in ref_items:
        ref_para = doc.add_paragraph(style='Normal')
        bfont = body_cfg.get("font", "SimSun")
        bfont_w = body_cfg.get("font_west", "Times New Roman")
        bsize_val = body_cfg.get("size", "小五")
        bsize = to_pt(bsize_val)
        run = ref_para.add_run(item)
        set_run_font(run, bfont, bfont_w, bsize, color=(0, 0, 0))
        rpf = ref_para.paragraph_format
        rpf.line_spacing = 1.5
        rpf.space_before = Pt(0)
        rpf.space_after = Pt(0)
        rpf.first_line_indent = Pt(0)
        # Hanging indent: left indent = 2 chars, first line = -2 chars
        hang = Pt(bsize * 2)
        rpf.left_indent = hang
        rpf.first_line_indent = -hang


# ---------------------------------------------------------------------------
# TOC (Table of Contents)
# ---------------------------------------------------------------------------

def add_toc(doc):
    """Insert a Word TOC field at the current cursor position.
    The TOC will populate when the user opens the document in Word
    and confirms 'Update Table of Contents'."""
    para = doc.add_paragraph(style='Normal')
    run = para.add_run()
    fldChar_begin = OxmlElement('w:fldChar')
    fldChar_begin.set(qn('w:fldCharType'), 'begin')
    run._r.append(fldChar_begin)

    run2 = para.add_run()
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = ' TOC \\o "1-3" \\h \\z '
    run2._r.append(instrText)

    run3 = para.add_run()
    fldChar_separate = OxmlElement('w:fldChar')
    fldChar_separate.set(qn('w:fldCharType'), 'separate')
    run3._r.append(fldChar_separate)

    run4 = para.add_run('[ Table of Contents — right-click and select "Update Field" in Word ]')
    set_run_font(run4, "SimSun", "Times New Roman", to_pt(10), italic=True)

    run5 = para.add_run()
    fldChar_end = OxmlElement('w:fldChar')
    fldChar_end.set(qn('w:fldCharType'), 'end')
    run5._r.append(fldChar_end)

    pf = para.paragraph_format
    pf.space_before = Pt(12)
    pf.space_after = Pt(12)
    return para
