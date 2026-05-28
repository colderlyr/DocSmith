"""Block renderer: dispatches parsed blocks to docx element functions.

Uses a dispatch dict for extensibility — new block types can be registered
without modifying the render loop.
"""

import re
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

from .fonts import to_pt, set_run_font, set_para_fmt
from .equations import add_display_eq
from .elements import (
    add_heading, add_table, process_inline_formatting,
    add_reference_section, set_para_single_column,
    add_section_body_para
)
from .numbering import format_list_prefix


class BlockRenderer:
    """Renders parsed blocks into a python-docx Document.

    Uses a dispatch dict (_renderers) mapping block type strings to handler
    methods. New block types can be added by registering additional entries.
    """

    def __init__(self, doc, config, ctx):
        self.doc = doc
        self.config = config
        self.ctx = ctx
        self.body_cfg = config.get("body", {})
        self.eq_cfg = config.get("equations", {})
        self.table_cfg = config.get("table", {})
        self.headings_cfg = config.get("headings", {})
        self.page_cfg = config.get("page", {})

        self.body_font_size_pt = to_pt(self.body_cfg.get("size", 12))
        self.column_width_cm = _get_column_width_cm(self.page_cfg)
        self.is_multi_col = self.page_cfg.get("columns", 1) > 1

        self.prev_blank = True
        self.seen_first_heading = False
        self.current_section = None

        self._renderers = {
            'blank':             self._render_blank,
            'heading':           self._render_heading,
            'display_math':      self._render_display_math,
            'code_eq':           self._render_code_eq,
            'table':             self._render_table,
            'table_caption':     self._render_table_caption,
            'image':             self._render_image,
            'list_item':         self._render_list_item,
            'ordered_list_item': self._render_ordered_list_item,
            'paragraph':         self._render_paragraph,
            'references':        self._render_references,
        }

    def render(self, block):
        """Dispatch a single block to its renderer."""
        btype = block[0]
        renderer = self._renderers.get(btype)
        if renderer:
            renderer(block)

    # ---- Individual renderers ----

    def _render_blank(self, block):
        self.prev_blank = True

    def _render_heading(self, block):
        _, level, text, meta = block
        if not self.prev_blank:
            self.doc.add_paragraph(style='Normal')
        para = add_heading(self.doc, level, text, self.headings_cfg, self.ctx)

        if self.is_multi_col and not self.seen_first_heading:
            set_para_single_column(para)
            self.seen_first_heading = True

        section_type = meta.get('section') if meta else None
        if self.is_multi_col and section_type in ('abstract', 'acknowledgment'):
            set_para_single_column(para)
            self.current_section = section_type
        elif section_type is None:
            self.current_section = None

        self.prev_blank = False

    def _render_display_math(self, block):
        _, _, text, meta = block
        self.ctx.equation_counter += 1
        para = self.doc.add_paragraph(style='Normal')
        eq_label = meta.get('label') if meta else None
        if eq_label:
            self.ctx.register_label(eq_label, self.ctx.equation_counter,
                                    f"eq{self.ctx.equation_counter}")
        label_match = re.search(r'\\label\{eq:(\S+?)\}', text)
        if label_match:
            self.ctx.register_label(label_match.group(1), self.ctx.equation_counter,
                                    f"eq{self.ctx.equation_counter}")
        add_display_eq(para, text.strip(), self.eq_cfg, self.ctx.equation_counter,
                       self.ctx, column_width_cm=self.column_width_cm)
        self.prev_blank = False

    def _render_code_eq(self, block):
        _, _, _, meta = block
        items = meta.get('items', [])
        for etype, content, eq_num in items:
            if etype == 'eq':
                self.ctx.equation_counter += 1
                para = self.doc.add_paragraph(style='Normal')
                label_match = re.search(r'\\label\{eq:(\S+?)\}', content)
                if label_match:
                    self.ctx.register_label(label_match.group(1),
                                            self.ctx.equation_counter,
                                            f"eq{self.ctx.equation_counter}")
                display_num = eq_num if eq_num and eq_num.isdigit() else self.ctx.equation_counter
                add_display_eq(para, content, self.eq_cfg,
                               int(display_num) if isinstance(display_num, str) and display_num.isdigit() else display_num,
                               self.ctx, column_width_cm=self.column_width_cm)
            elif etype == 'text_eq':
                para = self.doc.add_paragraph(style='Normal')
                _add_text_equation(para, content, eq_num, self.eq_cfg, self.body_cfg,
                                   column_width_cm=self.column_width_cm)
        self.prev_blank = False

    def _render_table(self, block):
        _, _, _, meta = block
        lines = meta['lines']
        if len(lines) < 2:
            return

        def parse_row(r):
            return [c.strip() for c in r.strip('|').split('|')]
        headers = parse_row(lines[0])
        if re.match(r'^[\|\s\-:]+$', lines[1]):
            rows = [parse_row(r) for r in lines[2:]]
        else:
            rows = [parse_row(r) for r in lines[1:]]
        add_table(self.doc, headers, rows, self.table_cfg, meta.get('caption', ''),
                  self.ctx, label=meta.get('label'))
        self.prev_blank = False

    def _render_table_caption(self, block):
        self.prev_blank = False

    def _render_image(self, block):
        _, _, text, meta = block
        para = self.doc.add_paragraph(style='Normal')
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf = para.paragraph_format
        pf.space_before = Pt(6)
        pf.space_after = Pt(6)
        bfont = self.body_cfg.get("font", "SimSun")
        bfont_w = self.body_cfg.get("font_west", "Times New Roman")
        bsize = self.body_cfg.get("size", 12)
        path = meta.get('path', '') if meta else ''
        alt = text if text else 'Image'
        placeholder = f"[Figure: {alt}]"
        if path:
            placeholder += f" ({path})"
        run = para.add_run(placeholder)
        set_run_font(run, bfont, bfont_w, to_pt(bsize), italic=True, color=(0x99, 0x99, 0x99))
        self.prev_blank = False

    def _render_list_item(self, block):
        _, level, text, meta = block
        para = self.doc.add_paragraph(style='Normal')
        set_para_fmt(para, self.body_cfg.get("line_spacing", 1.5),
                     first_indent=None, alignment=self.body_cfg.get("alignment", "justify"))
        bfont = self.body_cfg.get("font", "SimSun")
        bfont_w = self.body_cfg.get("font_west", "Times New Roman")
        bsize = self.body_cfg.get("size", 12)
        prefix = f"{'  ' * level}• "
        run = para.add_run(prefix)
        set_run_font(run, bfont, bfont_w, to_pt(bsize))
        pf = para.paragraph_format
        pf.left_indent = Cm(1.0 + level * 0.5)
        process_inline_formatting(para, text, self.body_cfg, self.ctx)
        if self.is_multi_col and self.current_section in ('abstract', 'acknowledgment'):
            set_para_single_column(para)
        self.prev_blank = False

    def _render_ordered_list_item(self, block):
        _, level, text, meta = block
        style = meta.get('style', '1.') if meta else '1.'
        para = self.doc.add_paragraph(style='Normal')
        set_para_fmt(para, self.body_cfg.get("line_spacing", 1.5),
                     first_indent=None, alignment=self.body_cfg.get("alignment", "justify"))
        bfont = self.body_cfg.get("font", "SimSun")
        bfont_w = self.body_cfg.get("font_west", "Times New Roman")
        bsize = self.body_cfg.get("size", 12)

        self.ctx.list_counters[level] += 1
        for i in range(level + 1, len(self.ctx.list_counters)):
            self.ctx.list_counters[i] = 0
        num = self.ctx.list_counters[level]
        prefix = format_list_prefix(style, num)
        indent = f"{'  ' * level}"
        run = para.add_run(f"{indent}{prefix}")
        set_run_font(run, bfont, bfont_w, to_pt(bsize))
        pf = para.paragraph_format
        pf.left_indent = Cm(1.0 + level * 0.5)
        process_inline_formatting(para, text, self.body_cfg, self.ctx)
        if self.is_multi_col and self.current_section in ('abstract', 'acknowledgment'):
            set_para_single_column(para)
        self.prev_blank = False

    def _render_paragraph(self, block):
        _, _, text, meta = block
        section_type = meta.get('section') if meta else None
        if section_type in ('abstract', 'acknowledgment'):
            add_section_body_para(self.doc, text, self.body_cfg, section_type, self.ctx)
            if self.is_multi_col:
                set_para_single_column(self.doc.paragraphs[-1])
        else:
            para = self.doc.add_paragraph(style='Normal')
            set_para_fmt(para, self.body_cfg.get("line_spacing", 1.5),
                         self.body_cfg.get("first_line_indent"),
                         self.body_cfg.get("alignment", "justify"),
                         font_size_pt=self.body_font_size_pt)
            process_inline_formatting(para, text, self.body_cfg, self.ctx)
        self.prev_blank = False

    def _render_references(self, block):
        _, _, _, meta = block
        ref_heading_level = self.config.get("references", {}).get("heading_level", 2)
        add_reference_section(self.doc, meta['items'], self.body_cfg,
                              headings_cfg=self.headings_cfg,
                              ref_heading_level=ref_heading_level)
        self.prev_blank = False


# ---- Helpers ----

def _get_column_width_cm(cfg):
    """Compute usable column width in cm from page config."""
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


def _add_text_equation(para, text, eq_num, eq_cfg, body_cfg, column_width_cm=None):
    """Render a text constraint as a centered equation-like line with optional number."""
    if column_width_cm is None:
        column_width_cm = 15.92
    bfont = body_cfg.get("font", "SimSun")
    bfont_w = body_cfg.get("font_west", "Times New Roman")
    bsize = body_cfg.get("size", 12)
    nfont = eq_cfg.get("numbering_font", "Times New Roman")
    nsize = to_pt(eq_cfg.get("numbering_size", 12))
    if eq_num:
        para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf = para.paragraph_format
        center_stop = Cm(column_width_cm / 2.0)
        right_stop = Cm(column_width_cm)
        pf.tab_stops.add_tab_stop(center_stop, alignment=WD_ALIGN_PARAGRAPH.CENTER)
        pf.tab_stops.add_tab_stop(right_stop, alignment=WD_ALIGN_PARAGRAPH.RIGHT)
        para.add_run("\t")
        run = para.add_run(text)
        set_run_font(run, bfont, bfont_w, to_pt(bsize), italic=True)
        run2 = para.add_run(f"\t({eq_num})")
        set_run_font(run2, nfont, nfont, nsize)
    else:
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = para.add_run(text)
        set_run_font(run, bfont, bfont_w, to_pt(bsize), italic=True)
