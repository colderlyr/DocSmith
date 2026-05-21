"""Font utilities: Chinese size names, run formatting, paragraph formatting."""

from docx.shared import Pt, Cm, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

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


def set_run_font(run, font_cn, font_west, size_pt, bold=False, italic=False, color=None):
    """Set font properties with east-asia/western font separation."""
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


def set_para_fmt(para, line_spacing=1.5, first_indent=None, alignment=None,
                 space_before=0, space_after=0):
    """Set paragraph formatting. Forces space_before/after to 0 by default."""
    pf = para.paragraph_format
    pf.line_spacing = float(line_spacing)
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    if first_indent == "2chars":
        pf.first_line_indent = Cm(0.74)
    elif first_indent:
        pf.first_line_indent = first_indent
    amap = {"left": 0, "center": 1, "right": 2, "justify": 3}
    if alignment in amap:
        para.alignment = amap[alignment]


def add_run(para, text, font_cn, font_west, size, bold=False, italic=False):
    """Add a run with the given font settings."""
    run = para.add_run(text)
    set_run_font(run, font_cn, font_west, to_pt(size), bold, italic)
    return run
