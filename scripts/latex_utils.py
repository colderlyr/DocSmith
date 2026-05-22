"""
LaTeX equation utilities for DocSmith.

Handles:
  - Unicode math symbol → LaTeX command conversion
  - Subscript/superscript Unicode characters
  - ASCII _ and ^ pattern normalization for equations
  - Named subscript pre-processing (max, min, total, ref, etc.)
  - Consecutive subscript joining (prevents _{X}_{Y} double-subscript errors)
  - Equation vs text constraint detection for mixed blocks

Key lessons from IEEE paper generation:
  1. latex2word supports \\text{}, \\textrm{}, \\mathrm{} for text in equations
  2. Unicode Greek (α, ρ, μ...) must be converted to LaTeX BEFORE latex2word
  3. ASCII _X patterns need {}-wrapping: _max → _{\\max}, _f → _{f}
  4. Consecutive subscripts V_f_max → V_{f}_{max} (INVALID) → V_{f,\\max} (valid)
  5. Named subscripts should use proper LaTeX operators: _{\\max}, _{\\min}
  6. Lines without math operators (=, ≤, +, etc.) are text constraints, not equations
"""

import re

# ---------------------------------------------------------------------------
# Unicode → LaTeX symbol mappings
# ---------------------------------------------------------------------------

# Greek and math symbols (sorted longest-first for safe replacement)
UNICODE_MATH_SYMBOLS = [
    # Operators / relations
    ('∥', r'\parallel '), ('…', r'\dots '), ('∑', r'\sum '), ('∏', r'\prod '),
    ('√', r'\sqrt '), ('⟨', r'\langle '), ('⟩', r'\rangle '),
    ('∇', r'\nabla '), ('∂', r'\partial '), ('∫', r'\int '),
    # Greek lowercase
    ('α', r'\alpha '), ('β', r'\beta '), ('γ', r'\gamma '),
    ('δ', r'\delta '), ('ε', r'\epsilon '), ('ζ', r'\zeta '),
    ('η', r'\eta '), ('θ', r'\theta '), ('ι', r'\iota '),
    ('κ', r'\kappa '), ('λ', r'\lambda '), ('μ', r'\mu '),
    ('ν', r'\nu '), ('ξ', r'\xi '), ('π', r'\pi '),
    ('ρ', r'\rho '), ('σ', r'\sigma '), ('τ', r'\tau '),
    ('υ', r'\upsilon '), ('φ', r'\phi '), ('χ', r'\chi '),
    ('ψ', r'\psi '), ('ω', r'\omega '),
    # Greek uppercase
    ('Γ', r'\Gamma '), ('Δ', r'\Delta '), ('Θ', r'\Theta '),
    ('Λ', r'\Lambda '), ('Ξ', r'\Xi '), ('Π', r'\Pi '),
    ('Σ', r'\Sigma '), ('Φ', r'\Phi '), ('Ψ', r'\Psi '),
    ('Ω', r'\Omega '),
    # Operators
    ('·', r'\cdot '), ('×', r'\times '),
    ('≤', r'\leq '), ('≥', r'\geq '), ('∈', r'\in '),
    ('∞', r'\infty '), ('→', r'\rightarrow '), ('⇒', r'\Rightarrow '),
    ('≠', r'\neq '), ('≈', r'\approx '), ('°', r'^{\circ}'),
    # Arrows
    ('←', r'\leftarrow '), ('↑', r'\uparrow '), ('↓', r'\downarrow '),
    ('↔', r'\leftrightarrow '), ('⇐', r'\Leftarrow '), ('⇑', r'\Uparrow '),
]

# Unicode subscript characters
UNICODE_SUBSCRIPTS = {
    '₀': '_{0}', '₁': '_{1}', '₂': '_{2}', '₃': '_{3}', '₄': '_{4}',
    '₅': '_{5}', '₆': '_{6}', '₇': '_{7}', '₈': '_{8}', '₉': '_{9}',
    'ₐ': '_{a}', 'ₑ': '_{e}', 'ₕ': '_{h}', 'ᵢ': '_{i}',
    'ⱼ': '_{j}', 'ₖ': '_{k}', 'ₗ': '_{l}', 'ₘ': '_{m}',
    'ₙ': '_{n}', 'ₒ': '_{o}', 'ₚ': '_{p}', 'ᵣ': '_{r}',
    'ₛ': '_{s}', 'ₜ': '_{t}', 'ᵤ': '_{u}', 'ᵥ': '_{v}',
    'ₓ': '_{x}', '₊': '_{+}', '₋': '_{-}',
}

# Unicode superscript characters
UNICODE_SUPERSCRIPTS = {
    '⁰': '^{0}', '¹': '^{1}', '²': '^{2}', '³': '^{3}', '⁴': '^{4}',
    '⁵': '^{5}', '⁶': '^{6}', '⁷': '^{7}', '⁸': '^{8}', '⁹': '^{9}',
    '⁺': '^{+}', '⁻': '^{-}', '⁼': '^{=}',
    'ⁿ': '^{n}', 'ⁱ': '^{i}',
}

# Named subscripts that should use LaTeX operators or text
# Applied BEFORE general _ processing to prevent double-subscript bugs
# Pattern: '_word' → proper LaTeX
NAMED_SUBSCRIPTS = [
    (r'_max\b', r'_{\\max}'),      # maximum
    (r'_min\b', r'_{\\min}'),      # minimum
    (r'_total\b', r'_{\\mathrm{total}}'),
    (r'_ref\b', r'_{\\mathrm{ref}}'),
    (r'_th\b', r'_{\\mathrm{th}}'),
    (r'_sat\b', r'_{\\mathrm{sat}}'),
    (r'_avg\b', r'_{\\mathrm{avg}}'),
    (r'_in\b', r'_{\\mathrm{in}}'),
    (r'_out\b', r'_{\\mathrm{out}}'),
    (r'_crit\b', r'_{\\mathrm{crit}}'),
    (r'_opt\b', r'_{\\mathrm{opt}}'),
    (r'_init\b', r'_{\\mathrm{init}}'),
    (r'_jc\b', r'_{\\mathrm{jc}}'),    # junction-to-case
    (r'_hs\b', r'_{\\mathrm{hs}}'),    # heat sink
]

# Named operators / text phrases that appear in equations
# Replaced BEFORE general _/^ processing
NAMED_OPERATORS = [
    (' min:', r' \min '),
    (' max:', r' \max '),
    ('subject to:', r' \text{subject to:} '),
    ('subject to ', r' \text{subject to:} '),
    ('s.t.', r' \text{s.t.} '),
]


def unicode_to_latex(text):
    """Convert Unicode math symbols and sub/superscripts to LaTeX commands.

    Processing order (critical — each step builds on the previous):
      1. Unicode sub/superscript chars → LaTeX {}-wrapped
      2. Named operators (min:, subject to:) → LaTeX commands
      3. Unicode math symbols → LaTeX commands
      4. Named subscripts (_max, _min...) → proper LaTeX
      5. General ASCII _ → _{...} wrapping
      6. General ASCII ^ → ^{...} wrapping
      7. Consecutive subscript joining

    Returns a LaTeX string suitable for latex2word.
    """
    s = text

    # Step 1: Unicode sub/superscript characters
    for sub, latex in UNICODE_SUBSCRIPTS.items():
        s = s.replace(sub, latex)
    for sup, latex in UNICODE_SUPERSCRIPTS.items():
        s = s.replace(sup, latex)

    # Step 2: Named operators (before symbol conversion)
    for pattern, replacement in NAMED_OPERATORS:
        s = s.replace(pattern, replacement)

    # Step 3: Unicode math symbols (longest match first to avoid partial replacements)
    for uni, latex in sorted(UNICODE_MATH_SYMBOLS, key=lambda x: -len(x[0])):
        s = s.replace(uni, latex)

    # Step 4: Named subscripts (before general _ processing)
    for pattern, replacement in NAMED_SUBSCRIPTS:
        s = re.sub(pattern, replacement, s)

    # Step 5: General ASCII _ subscript wrapping
    # Multi-char: _text → _{text}
    s = re.sub(r'_([a-zA-Z]{2,})', r'_{\1}', s)
    # LaTeX commands: _\alpha → _{\alpha}
    s = re.sub(r'_(\\[a-zA-Z]+)', r'{\1}', s)
    # Single char: _X → _{X}
    s = re.sub(r'_([a-zA-Z])', r'_{\1}', s)
    # Digits: _123 → _{123}
    s = re.sub(r'_(\d+)', r'_{\1}', s)

    # Step 6: General ASCII ^ superscript wrapping
    s = re.sub(r'\^([a-zA-Z]{2,})', r'^{\1}', s)
    s = re.sub(r'\^(\\[a-zA-Z]+)', r'^{\1}', s)
    s = re.sub(r'\^([a-zA-Z])', r'^{\1}', s)
    s = re.sub(r'\^(-?\d+)', r'^{\1}', s)

    # Step 7: Fix consecutive subscripts (critical!)
    # V_{f}_{max} → V_{f,\max} (prevents invalid LaTeX double subscripts)
    s = re.sub(r'\}_\{', r',', s)

    # Final cleanup
    s = re.sub(r'\{(\{[^}]+\})\}', r'\1', s)  # remove doubled braces
    s = re.sub(r'\s+\{', '{', s)               # spaces before braces
    s = re.sub(r'\}\s+', '}', s)               # spaces after braces
    s = re.sub(r' +', ' ', s)                  # collapse spaces

    return s.strip()


# ---------------------------------------------------------------------------
# Equation line detection
# ---------------------------------------------------------------------------

def is_equation_line(line):
    """Detect if a stripped line is an equation (has math operators)
    or a text constraint that merely has a trailing equation number.

    Returns True if the line appears to be a mathematical equation.
    Used when splitting multi-line code blocks into individual equations.

    Examples:
      'ρ(u·∇)u = -∇p + μ∇²u - α(γ)u'   → True  (has =, Greek)
      'subject to: V_f / V_total ≤ V_f_max' → True  (has ≤)
      '0 ≤ γ_e ≤ 1'                      → True  (has ≤)
      '(governing equations (1)-(3))'    → False (text constraint)
    """
    if not line:
        return False
    return bool(re.search(r'[=≤≥×·]|\\[a-zA-Z]|[_^]', line))


def parse_equation_line(line):
    """Parse a single line from a code block as a potential equation.

    Args:
        line: A stripped line from a code block

    Returns:
        ('eq', latex_str, eq_num) for equations
        ('text_eq', text, eq_num) for text constraints with equation numbers
        (None, None, None) if the line should be skipped

    Equation number is extracted from trailing patterns like '    (4)' or '    (8-10)'.
    """
    line = line.strip()
    if not line:
        return None, None, None

    # Extract trailing equation number e.g. "    (4)" or "    (8-10)"
    eq_num = None
    m = re.search(r'\s*\((\d+(?:[–\-]\d+)?)\)\s*$', line)
    if m:
        eq_num = m.group(1)
        line = line[:m.start()].strip()

    if not line:
        # Only the number remained — text constraint
        return 'text_eq', '', eq_num

    if is_equation_line(line):
        latex = unicode_to_latex(line)
        return 'eq', latex, eq_num
    else:
        return 'text_eq', line, eq_num


def parse_equation_block(code_lines):
    """Process a multi-line code block into individual equation items.

    Each non-empty line is processed independently via parse_equation_line().
    Returns a list of ('eq', latex, num) or ('text_eq', text, num) tuples.
    """
    results = []
    for raw_line in code_lines:
        etype, content, eq_num = parse_equation_line(raw_line)
        if etype is not None:
            results.append((etype, content, eq_num))
    return results
