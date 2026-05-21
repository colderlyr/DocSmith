#!/usr/bin/env python3
"""
DocSmith — Generate .docx from Markdown with native OMML equations, Chinese fonts,
table support, equation numbering, and cross-references.

Usage:
    python3 -m scripts.generate_docx --output out.docx --config config.json --content content.md
    (run from DocSmith/ directory)
"""

import json
import re
import argparse
from docx import Document
from docx.shared import Pt, Cm

from .fonts import to_pt, set_run_font, set_para_fmt, add_run
from .omml import add_display_eq
from .parser import parse_markdown, DocContext
from .elements import (setup_page, add_heading, add_table, add_table_caption,
                       process_inline_formatting, _add_simple_text)


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
                doc.add_paragraph(style='Normal')
            add_heading(doc, level, text, headings_cfg, ctx)
            prev_blank = False

        elif btype == 'display_math':
            ctx.equation_counter += 1
            para = doc.add_paragraph(style='Normal')
            add_display_eq(para, text.strip(), eq_cfg, ctx.equation_counter, ctx)
            prev_blank = False

        elif btype == 'table':
            lines = meta['lines']
            if len(lines) < 2:
                continue
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
            para = doc.add_paragraph(style='Normal')
            set_para_fmt(para, body_cfg.get("line_spacing", 1.5),
                        first_indent=None, alignment=body_cfg.get("alignment", "justify"))
            # Add bullet prefix
            bfont = body_cfg.get("font", "SimSun")
            bfont_w = body_cfg.get("font_west", "Times New Roman")
            bsize = body_cfg.get("size", 12)
            prefix = f"{'  ' * level}• "
            run = para.add_run(prefix)
            set_run_font(run, bfont, bfont_w, to_pt(bsize))
            pf = para.paragraph_format
            pf.left_indent = Cm(1.0 + level * 0.5)
            # Process list item text with full inline formatting
            process_inline_formatting(para, text, body_cfg, ctx)
            prev_blank = False

        elif btype == 'paragraph':
            para = doc.add_paragraph(style='Normal')
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
