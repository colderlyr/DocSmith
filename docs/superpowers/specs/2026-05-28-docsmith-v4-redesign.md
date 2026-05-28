# DocSmith v4 Redesign Spec

## Motivation

Fix 5 known bugs and restructure the codebase for maintainability and extensibility.
The core architectural change is replacing the fragile pre-scan + render loop with a
clean two-pass pipeline, while splitting the god module `generate_docx.py` into
single-responsibility modules.

## Current vs Target Architecture

```
BEFORE (6 files)                          AFTER (9 files)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
scripts/                                  scripts/
├── generate_docx.py  (360L god)          ├── __main__.py        CLI only
├── parser.py         (285L)              ├── pipeline.py        [NEW] Two-pass orchestrator
├── elements.py       (550L)              ├── context.py         [NEW] DocContext extracted
├── omml.py           (141L)              ├── parser.py          Markdown → blocks
├── latex_utils.py    (239L)              ├── renderer.py        [NEW] blocks → docx
└── fonts.py          (105L)              ├── elements.py        docx primitives
                                          ├── equations.py       [MERGE] omml + latex_utils
                                          ├── numbering.py       [NEW] heading/list numbers
                                          └── fonts.py           unchanged
```

### Responsibility Boundaries

| Module | LOC (est.) | Responsibility | Depends On |
|--------|-----------|----------------|------------|
| `__main__.py` | ~30 | argparse, read files, call pipeline | pipeline |
| `pipeline.py` | ~80 | TwoPassBuilder: Pass1 labels → Pass2 render → save | parser, context, renderer |
| `context.py` | ~40 | DocContext: counters, label_map, bookmarks (pure data) | nothing |
| `parser.py` | ~250 | Markdown text → list of blocks | equations (for code_eq parsing) |
| `renderer.py` | ~150 | Block dispatch + render loop | elements, equations, numbering, context, fonts |
| `elements.py` | ~400 | Page setup, headings, tables, TOC, references, sections | fonts, numbering |
| `equations.py` | ~320 | Unicode→LaTeX→OMML pipeline | fonts |
| `numbering.py` | ~120 | heading_prefix(), list formatting, roman/letter/cn | nothing |
| `fonts.py` | ~90 | set_run_font, set_para_fmt, to_pt | nothing |

---

## Bug Fixes

### Bug 1: Equation Label Pre-scan (HIGH)

**Root cause**: `generate_docx.py:64-89` has a pre-scan loop that increments
counters and registers labels, then the render loop (lines 102-278) increments
counters AGAIN. The pre-scan handles `display_math` labels with `pass` (no-op).
`code_eq` items are iterated twice (once for `meta['label']`, once for `\label{}`
inside the LaTeX), causing double-counting.

**Fix**: Eliminate the pre-scan entirely. Replace with two-pass pipeline:

```
Pass 1 (register_labels):
  for block in blocks:
      if block is equation → ctx.equation_counter += 1, register \label
      if block is table    → ctx.table_counter += 1,    register [表:label]
      if block is figure   → ctx.figure_counter += 1,   register [图:label]
  ctx.reset_counters()

Pass 2 (render_blocks):
  for block in blocks:
      dispatch to render function (counters increment again, matching Pass 1)
```

No code duplication. One source of truth for counter progression.

### Bug 2: Aggressive `}_{` Replacement (HIGH)

**Root cause**: `latex_utils.py:159` — `re.sub(r'\}_\{', r',', s)` replaces ALL
`}_{` with `,`, which destroys legitimate nested subscripts like `a_{b_{c}}`.

**Fix**: Replace with a targeted pattern that only merges same-level consecutive
subscripts:

```python
# Before: s = re.sub(r'\}_\{', r',', s)  # destroys a_{b_{c}}

# After: only merge when the pattern is word_{sub1}_{sub2} (no nested braces)
s = re.sub(
    r'([a-zA-Z\\]+)_\{([^{}]+)\}_\{([^{}]+)\}',
    r'\1_{\2,\3}',
    s
)
```

This matches `V_{f}_{max}` → `V_{f,\max}` but leaves `a_{b_{c}}` alone.

### Bug 3: SKILL.md Path Case (MEDIUM)

**Fix**: `SKILL.md:197` — `cd DocSmith/` → `cd docsmith/`.

### Bug 4: Hardcoded Reference Heading Level (MEDIUM)

**Root cause**: `elements.py:478` hardcodes `level = 2` for the "参考文献" heading.

**Fix**: Read from config:

```python
ref_heading_level = config.get("references", {}).get("heading_level", 2)
```

Default to 2 for backward compatibility.

### Bug 5: List Indentation Assumes 2-space (MEDIUM)

**Root cause**: `parser.py:227` uses `indent // 2` which counts 2 spaces = 1 level.
Markdown allows 4-space indentation.

**Fix**: Detect the indentation unit from the first indented list item:

```python
# Detect tab_size from first indented item (2 or 4)
if self._list_tab_size is None and indent > 0:
    self._list_tab_size = indent  # first indented level sets the unit
level = indent // self._list_tab_size if self._list_tab_size else 0
```

---

## Extension Points

Three registries/hooks designed for future features without modifying core loops.

### Extension Point 1: Block Type Registry

`renderer.py` uses a dispatch dict instead of if/elif chain:

```python
BLOCK_RENDERERS: dict[str, Callable] = {
    'heading':          _render_heading,
    'display_math':     _render_display_math,
    'code_eq':          _render_code_eq,
    'table':            _render_table,
    'image':            _render_image,
    'list_item':        _render_list_item,
    'ordered_list_item': _render_ordered_list_item,
    'paragraph':        _render_paragraph,
    'references':       _render_references,
    'blank':            _render_blank,
}
```

To add a new block type (e.g., `'figure'`, `'citation'`, `'footnote'`):
1. Add parsing logic in `parser.py`
2. Register the renderer in `BLOCK_RENDERERS`
3. Add Pass 1 label registration in `pipeline.py`

No changes to the render loop itself.

### Extension Point 2: Inline Processor Pipeline

`renderer.py` processes inline content through an ordered list of processors:

```python
INLINE_PROCESSORS = [
    cross_ref_processor,     # \ref{tab:xxx}, \ref{eq:xxx}, \ref{fig:xxx}
    citation_processor,      # [@smith2023] → superscript reference number
    inline_math_processor,   # $...$ → OMML
    formatting_processor,    # **bold**, *italic*
]
```

Adding citation support:
1. Implement `citation_processor(text, ctx, body_cfg)` — matches `[@...]` patterns
2. Insert into `INLINE_PROCESSORS` list (before or after cross_ref, depending on desired parse order)

### Extension Point 3: Label Registration Hook

`pipeline.py` Pass 1 uses a registration hook per block type:

```python
LABEL_EXTRACTORS = {
    'display_math': _extract_eq_label,
    'code_eq':      _extract_code_eq_labels,
    'table':        _extract_table_label,
    'image':        _extract_figure_label,       # ready for future use
}
```

Adding figure numbering:
1. Parser already produces `('image', ...)` blocks
2. Implement `_extract_figure_label(block, ctx)` — parse `[图:label]` caption
3. Add to `LABEL_EXTRACTORS`
4. Renderer produces "图 1" caption instead of gray placeholder

### Future Feature: Inline Citations

Design sketch (not implementing now, just the interface):

```
Content:  "As shown by Smith[@smith2023], the method..."
                          ↓ parser detects [@...]
                          ↓ Pass 1: register citation key → number
                          ↓ Pass 2: inline processor replaces with [1] superscript
                          ↓ References section auto-generated from registered keys
Output:   "As shown by Smith[1], the method..."
```

The `DocContext` gets a `citation_map: dict[str, int]` for key→number resolution.
The parser adds a new inline pattern for `[@key]`. No new block type needed —
citations are inline elements, same as `\ref{...}`.

### Future Feature: Figure Support

Currently images render as gray `[Figure: alt]` placeholders. Full support needs:

1. **Caption parsing** — `[图:label] caption text` (mirrors existing `[表:label]`)
2. **Numbering** — `图 1`, `图 2` via `ctx.figure_counter`
3. **Cross-reference** — `\ref{fig:label}` resolves to `图 1`
4. **Image embedding** — `python-docx` can embed PNG/JPG natively

Parser already produces `('image', 0, alt_text, {'path': str})` blocks.
The extension just needs a caption line detector (like table captions) and
a PIL/pillow-based image embedder in `elements.py`.

---

## TwoPassBuilder Interface

```python
class TwoPassBuilder:
    """Orchestrates label registration (Pass 1) and rendering (Pass 2)."""

    def __init__(self, config: dict, md_text: str):
        self.config = config
        self.blocks = parse_markdown(md_text)
        self.ctx = DocContext()

    def build(self, output_path: str) -> str:
        doc = Document()
        setup_page(doc, self.config.get("page", {}))
        if self.config.get("toc"):
            add_toc(doc)
        self._pass1_register()
        self._pass2_render(doc)
        doc.save(output_path)
        return output_path

    def _pass1_register(self) -> None:
        """Walk blocks. Register labels and numbers. Do NOT touch docx."""
        for btype, level, text, meta in self.blocks:
            extractor = LABEL_EXTRACTORS.get(btype)
            if extractor:
                extractor(self.ctx, btype, text, meta)

        # Reset for Pass 2 (render loop will re-increment)
        self.ctx.reset_counters()

    def _pass2_render(self, doc: Document) -> None:
        """Walk blocks. Render each to the Document."""
        renderer = BlockRenderer(doc, self.config, self.ctx)
        for block in self.blocks:
            renderer.render(block)
```

---

## Migration Plan

1. Create new modules (`context.py`, `pipeline.py`, `renderer.py`, `numbering.py`, `equations.py`)
2. Move code from old modules into new ones WITHOUT changing logic
3. Fix the 5 bugs in the new modules
4. Wire up `__main__.py` to call `TwoPassBuilder`
5. Delete old `generate_docx.py`, `omml.py`, `latex_utils.py`
6. Update `SKILL.md` Phase 3 instructions and path case
7. Create symlink `~/.claude/skills/docsmith → ~/myFiles/docsmith`
8. Run existing presets through the pipeline to verify no regressions

### Risk Mitigation

- **Regression testing**: Before deleting old files, run all 3 presets (academic_paper, standard_report, ieee_conference) through both old and new pipeline, diff the docx XML
- **Backward compatibility**: Config JSON format unchanged. Content markdown format unchanged. CLI interface unchanged (`--output`, `--config`, `--content`).
- **Rollback**: Old files kept until new pipeline passes regression tests

---

## Not in Scope

- PIL/pillow image embedding (keep placeholder behavior)
- Inline citation `[@key]` parsing and resolution
- Figure caption `[图:label]` support
- Header/footer support
- Page break directive
- First-page-different layout option
