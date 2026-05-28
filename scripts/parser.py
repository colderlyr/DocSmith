"""Markdown parser and document context tracker.

Parses markdown into structured blocks. Supports:
  - Headings (H1-H4) with auto-numbering
  - Display math: $$...$$ (LaTeX) and ``` code blocks (equations)
  - Tables with optional [表:label] captions
  - Lists: unordered (-/*) and ordered (1./1)/a./a)/(1)/(a))
  - Images: ![alt](path) → placeholder blocks
  - Abstract / Acknowledgment sections
  - Paragraphs, references
  - Inline LaTeX: $...$ within paragraphs

State Flag Management (critical pattern — DO NOT break):
  When transitioning between special sections (abstract, acknowledgments,
  references), conflicting flags MUST be explicitly reset. Example:
    entering references section → set in_acknowledgment = False
  The elif chain order matters: in_acknowledgment is checked BEFORE
  in_references, so if both flags are True, references leak into acknowledgments.
  Always reset predecessor flags when switching modes.
"""

import re

from .context import DocContext
from .equations import parse_equation_block


def parse_markdown(text):
    """Parse markdown into blocks. Returns list of (type, level, text, meta).

    Block types:
      - 'heading'             : (heading, level, text, {'section': s})
                                section is 'abstract', 'acknowledgment', or None
      - 'display_math'        : (display_math, 0, latex_str, None)    — $$...$$
      - 'code_eq'             : (code_eq, 0, '', {'items': [...]})   — ``` equations
                                items: [('eq', latex, num) or ('text_eq', text, num)]
      - 'table'               : (table, 0, '', {'lines':..., 'caption':..., 'label':...})
      - 'list_item'           : (list_item, level, text, {'section': s})
      - 'ordered_list_item'   : (ordered_list_item, level, text, {'style': str, 'section': s})
      - 'image'               : (image, 0, alt_text, {'path': str})
      - 'paragraph'           : (paragraph, 0, text, {'section': s})
      - 'blank'               : (blank, 0, '', None)
      - 'references'          : (references, 0, '', {'items': [...]})
    """
    lines = text.split('\n')
    blocks = []
    i = 0
    current_section = None  # 'abstract', 'acknowledgment', or None

    while i < len(lines):
        line = lines[i]

        # Display math $$...$$
        if line.strip().startswith('$$'):
            stripped = line.strip()
            # Same-line open and close: $$ formula $$
            if stripped != '$$':
                m_close = re.search(r'(?<!\$)\$\$\s*$', stripped[2:])
                if m_close:
                    content = stripped[2:2 + m_close.start()].strip()
                    blocks.append(('display_math', 0, content, None))
                    i += 1
                    continue
            # Multi-line display math: $$ ... $$
            parts = []
            i += 1
            while i < len(lines):
                stripped = lines[i].strip()
                # Standalone $$ line closes the block
                if stripped == '$$':
                    break
                # Trailing $$ not preceded by another $ also closes
                m = re.search(r'(?<!\$)\$\$\s*$', stripped)
                if m:
                    parts.append(stripped[:m.start()].strip())
                    break
                parts.append(lines[i])
                i += 1
            blocks.append(('display_math', 0, '\n'.join(parts), None))
            i += 1
            continue

        # Code block: ``` at start of line
        if line.strip().startswith('```'):
            fence_match = re.match(r'^```\s*(\S+)?', line.strip())
            lang = fence_match.group(1) if fence_match else None

            if lang:
                # Named language block — collect verbatim, skip
                i += 1
                while i < len(lines) and not lines[i].strip().startswith('```'):
                    i += 1
                i += 1  # skip closing ```
                continue

            # Bare ``` fence — treat as equation block
            eq_lines = []
            i += 1
            while i < len(lines):
                if lines[i].strip() == '```':
                    break
                eq_lines.append(lines[i])
                i += 1
            eq_items = parse_equation_block(eq_lines)
            blocks.append(('code_eq', 0, '', {'items': eq_items}))
            i += 1  # skip closing ```
            continue

        # Image: ![alt text](path)
        img_match = re.match(r'^!\[([^\]]*)\]\(([^)]+)\)', line.strip())
        if img_match:
            blocks.append(('image', 0, img_match.group(1).strip(), {
                'path': img_match.group(2).strip()
            }))
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

            # Detect special sections by heading text
            txt_lower = txt.lower()
            if any(kw in txt_lower for kw in ['abstract', '摘要']):
                current_section = 'abstract'
            elif any(kw in txt_lower for kw in [
                'acknowledgment', 'acknowledgement', 'acknowledgments',
                'acknowledgements', '致谢', '鸣谢'
            ]):
                current_section = 'acknowledgment'
            elif '参考文献' in txt:
                current_section = None  # handled below
            else:
                # Non-special, non-reference heading resets section context
                current_section = None

            # Detect "参考文献" heading followed by numbered references
            if lvl == 2 and '参考文献' in txt:
                i += 1
                ref_items = []
                while i < len(lines):
                    ref_line = lines[i].strip()
                    ref_match = re.match(r'^\[(\d+)\]\s+(.+)', ref_line)
                    if ref_match:
                        ref_items.append(ref_match.group(2).strip())
                        i += 1
                    elif not ref_line:
                        i += 1
                    else:
                        break
                if ref_items:
                    blocks.append(('references', 0, '', {'items': ref_items}))
                else:
                    blocks.append(('heading', lvl, txt, {'section': None}))
                continue

            blocks.append(('heading', lvl, txt, {
                'section': current_section
            }))
            i += 1
            continue

        # Ordered list: 1. / 1) / (1) / a. / a) / (a)
        ol_match = re.match(
            r'^(\s*)(\d+[.)]\s+|[a-zA-Z][.)]\s+|\(\d+\)\s+|\([a-zA-Z]\)\s+)(.+)',
            line
        )
        if ol_match:
            indent = len(ol_match.group(1))
            style = ol_match.group(2).strip()
            content = ol_match.group(3).strip()
            unit = 2 if indent < 4 else 4
            level = indent // unit
            blocks.append(('ordered_list_item', level, content, {
                'style': style, 'section': current_section
            }))
            i += 1
            continue

        # Unordered list: - text, * text
        lm = re.match(r'^(\s*)[-*]\s+(.+)', line)
        if lm:
            indent = len(lm.group(1))
            unit = 2 if indent < 4 else 4
            level = indent // unit
            blocks.append(('list_item', level, lm.group(2).strip(), {
                'section': current_section
            }))
            i += 1
            continue

        # Blank
        if not line.strip():
            blocks.append(('blank', 0, '', None))
            i += 1
            continue

        # Paragraph
        blocks.append(('paragraph', 0, line.strip(), {
            'section': current_section
        }))
        i += 1

    return blocks


def resolve_cross_ref(ref_str, body_cfg, ctx):
    """Parse \\ref{type:id} and return (display_text, anchor) or (None, None)."""
    m = re.match(r'\\ref\{(tab|eq):(\S+?)\}', ref_str)
    if not m:
        return None, None
    ref_type, ref_id = m.group(1), m.group(2)
    xref = body_cfg.get("cross_ref", {})
    fmt = body_cfg.get("equations", {}).get("numbering_format", "({n})")

    if ref_type == "tab":
        if ref_id in ctx.label_map:
            number, anchor = ctx.label_map[ref_id]
        else:
            number, anchor = ref_id, f"tab{ref_id}"
        display = f"{xref.get('table_prefix', '表')} {number}"
        return display, anchor

    elif ref_type == "eq":
        if ref_id in ctx.label_map:
            number, anchor = ctx.label_map[ref_id]
        else:
            number, anchor = ref_id, f"eq{ref_id}"
        prefix = xref.get('equation_prefix', '公式')
        display = f"{prefix} {fmt.replace('{n}', str(number))}"
        return display, anchor

    return None, None
