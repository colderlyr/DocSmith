"""Markdown parser and document context tracker."""

import re


class DocContext:
    """Tracks heading counters, table/equation numbers, bookmarks, and labels."""

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
            # Detect "参考文献" heading (level 2) followed by numbered references
            if lvl == 2 and '参考文献' in txt:
                i += 1
                # Collect numbered reference items [1] ... [2] ...
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
                    # Only add references block — add_reference_section creates its own Heading 1
                    blocks.append(('references', 0, '', {'items': ref_items}))
                else:
                    # No reference items found — treat as normal heading
                    blocks.append(('heading', lvl, txt, None))
                continue
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
