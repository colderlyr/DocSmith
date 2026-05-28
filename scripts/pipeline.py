"""Two-pass document build pipeline.

Pass 1: Walk all blocks, register labels and numbers into DocContext.
Pass 2: Walk all blocks, render each into the python-docx Document.

This eliminates the fragile pre-scan (Bug #1) — there is exactly ONE counter
progression, executed identically in both passes.
"""

import re
from docx import Document

from .context import DocContext
from .parser import parse_markdown
from .renderer import BlockRenderer
from .elements import setup_page, add_toc


class TwoPassBuilder:
    """Orchestrates label registration (Pass 1) then rendering (Pass 2)."""

    def __init__(self, config, md_text):
        self.config = config
        self.blocks = parse_markdown(md_text)
        self.ctx = DocContext()

    def build(self, output_path):
        doc = Document()
        setup_page(doc, self.config.get("page", {}))

        if self.config.get("toc", False):
            add_toc(doc)

        self._pass1_register()
        self._pass2_render(doc)
        doc.save(output_path)
        return output_path

    # ---- Pass 1: Label Registration ----

    def _pass1_register(self):
        """Walk all blocks. Register equation/table labels and numbers.
        This is the SINGLE source of truth for counter progression."""
        for btype, level, text, meta in self.blocks:
            if btype == 'table' and meta and meta.get('label'):
                self.ctx.table_counter += 1
                self.ctx.register_label(meta['label'], self.ctx.table_counter,
                                        f"tab{self.ctx.table_counter}")

            elif btype == 'display_math':
                self.ctx.equation_counter += 1
                if meta and meta.get('label'):
                    self.ctx.register_label(meta['label'], self.ctx.equation_counter,
                                            f"eq{self.ctx.equation_counter}")
                self._register_latex_labels(text)

            elif btype == 'code_eq':
                items = meta.get('items', []) if meta else []
                for etype, content, eq_num in items:
                    if etype == 'eq':
                        self.ctx.equation_counter += 1
                        self._register_latex_labels(content)

            elif btype == 'figure_caption' and meta and meta.get('label'):
                self.ctx.figure_counter += 1
                self.ctx.register_label(meta['label'], self.ctx.figure_counter,
                                        f"fig{self.ctx.figure_counter}")

        self.ctx.reset_counters()

    def _register_latex_labels(self, latex_str):
        """Extract \\label{eq:...} from a LaTeX string and register."""
        for m in re.finditer(r'\\label\{eq:(\S+?)\}', latex_str):
            self.ctx.register_label(m.group(1), self.ctx.equation_counter,
                                    f"eq{self.ctx.equation_counter}")

    # ---- Pass 2: Render ----

    def _pass2_render(self, doc):
        """Walk all blocks and render each into the Document."""
        renderer = BlockRenderer(doc, self.config, self.ctx)
        for block in self.blocks:
            renderer.render(block)
