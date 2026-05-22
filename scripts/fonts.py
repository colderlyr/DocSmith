"""Font utilities: Chinese size names, run formatting, paragraph formatting."""

import re
from docx.shared import Pt, Cm, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

CN_FONT_SIZES = {
    "初号": 42, "小初": 36, "一号": 26, "小一": 24,
    "二号": 22, "小二": 18, "三号": 16, "小三": 15,
    "四号": 14, "小四": 12, "五号": 10.5, "小五": 9,
}


def to_pt(size):
    """Resolve font size: '小四'→12, '12'→12, 12→12, '12pt'→12."""
    if isinstance(size, (int, float)):
        return float(size)
    if size in CN_FONT_SIZES:
        return CN_FONT_SIZES[size]
    try:
        return float(size)
    except ValueError:
        # Try parsing as a CSS-like length string (e.g. "12pt")
        parsed = _parse_length(size)
        if isinstance(parsed, Pt):
            return parsed.pt  # unpack Pt to raw float
        return 12.0


def set_run_font(run, font_cn, font_west, size_pt, bold=False, italic=False, color=None):
    """Set font properties with east-asia/western font separation.
    If color is None, the default (inherited from style) is kept.
    Pass color=(0,0,0) to force black."""
    run.font.size = Pt(size_pt)
    run.font.name = font_west
    run.bold = bold
    run.italic = italic
    if color is not None:
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


def set_para_fmt(para, line_spacing=1.5, first_indent=None, alignment=None,
                 space_before=0, space_after=0, font_size_pt=None):
    """Set paragraph formatting. Forces space_before/after to 0 by default.
    When first_indent='2chars', computes Pt(font_size_pt * 2) for accurate
    Chinese 2-character indent based on actual font size (e.g. 12pt → 24pt).
    Also supports string values like '0.14in', '0.5cm', '12pt'."""
    pf = para.paragraph_format
    pf.line_spacing = float(line_spacing)
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    if first_indent == "2chars":
        if font_size_pt is not None:
            pf.first_line_indent = Pt(font_size_pt * 2)
        else:
            pf.first_line_indent = Cm(0.74)  # fallback
    elif isinstance(first_indent, str):
        pf.first_line_indent = _parse_length(first_indent)
    elif first_indent is not None:
        pf.first_line_indent = first_indent
    amap = {"left": 0, "center": 1, "right": 2, "justify": 3}
    if alignment in amap:
        para.alignment = amap[alignment]


def _parse_length(s):
    """Parse a CSS-like length string into a docx.shared.Length value.
    Supports: '12pt', '0.14in', '3.5mm', '0.5cm', '24px'.
    Defaults to Pt(12) if unparseable."""
    s = s.strip().lower()
    m = re.match(r'^([\d.]+)\s*(pt|in|cm|mm|px)$', s)
    if not m:
        return Pt(12)
    value = float(m.group(1))
    unit = m.group(2)
    if unit == 'pt':
        return Pt(value)
    elif unit == 'in':
        from docx.shared import Inches
        return Inches(value)
    elif unit == 'cm':
        return Cm(value)
    elif unit == 'mm':
        return Cm(value / 10)
    elif unit == 'px':
        return Pt(value * 0.75)
    return Pt(value)


def add_run(para, text, font_cn, font_west, size, bold=False, italic=False):
    """Add a run with the given font settings. Convenience wrapper."""
    run = para.add_run(text)
    set_run_font(run, font_cn, font_west, to_pt(size), bold, italic)
    return run
