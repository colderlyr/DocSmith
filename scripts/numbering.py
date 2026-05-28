"""Heading and list numbering utilities. Pure functions, no side effects."""

import re


# ---- Roman numerals ----

def to_roman(n):
    """Convert integer to uppercase Roman numeral. 1→I, 4→IV, 9→IX."""
    vals = [(1000, 'M'), (900, 'CM'), (500, 'D'), (400, 'CD'),
            (100, 'C'), (90, 'XC'), (50, 'L'), (40, 'XL'),
            (10, 'X'), (9, 'IX'), (5, 'V'), (4, 'IV'), (1, 'I')]
    result = ''
    for v, r in vals:
        while n >= v:
            result += r
            n -= v
    return result


# ---- Letter numbering: 1→A, 2→B, ..., 26→Z, 27→AA ----

def to_letter(n, upper=True):
    """Convert integer to letter: 1→A/a, 2→B/b, ..., 27→AA/aa."""
    result = ''
    base = ord('A' if upper else 'a')
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result = chr(base + rem) + result
    return result


# ---- Chinese numerals ----

_CN_NUMS = ["", "一", "二", "三", "四", "五", "六", "七", "八", "九", "十"]


def to_cn_num(n):
    """Convert integer 1-10 to Chinese numeral; >10 returns str(n)."""
    return _CN_NUMS[n] if n < len(_CN_NUMS) else str(n)


# ---- Dot-fallback numbering: counters → "1.", "1.1", "1.1.1" ----

def dot_fallback(counters, level):
    """Generate '1.', '1.1', '1.1.1' style numbering from counters."""
    parts = [str(counters[i]) for i in range(level) if counters[i] > 0]
    return ".".join(parts) + " "


# ---- Heading numbering families ----

def _derive_h2_family(h1_family):
    """Derive the default H2 numbering from the H1 family."""
    mapping = {
        "1.": "1.1", "一、": "（一）", "I.": "A.", "i.": "a.",
        "(1)": "(a)", "1)": "a)", "（一）": "1."
    }
    return mapping.get(h1_family, "")


def _format_by_style(style, counters, level):
    """Format a counter for a specific numbering style.
    Returns empty string for unrecognized styles."""
    c = counters
    idx = level - 1

    if style == "(1)":
        return f"({c[idx]}) "
    elif style == "(a)":
        return f"({to_letter(c[idx], upper=False)}) "
    elif style == "1)":
        return f"{c[idx]}) "
    elif style == "a)":
        return f"{to_letter(c[idx], upper=False)}) "
    elif style == "i)":
        return f"{to_roman(c[idx]).lower()}) "
    elif style == "A.":
        return f"{to_letter(c[idx])}. "
    elif style == "I.":
        return f"{to_roman(c[idx])}. "
    elif re.match(r'^[\d.]+$', style):
        return dot_fallback(counters, level)
    return ""


def heading_prefix(level, cfg, ctx):
    """Generate auto-numbering prefix based on H1 numbering family.

    Supported families:
      "1."      → H1: "1.", H2: "1.1", H3: "1.1.1"
      "一、"     → H1: "一、", H2: "（一）"
      "I."      → H1: "I.", H2: "A.", H3: "1)", H4: "a)"
      "i."      → H1: "i.", H2: "a."
      "1)"      → H1: "1)", H2: "a)", H3: "i)"
      "(1)"     → H1: "(1)", H2: "(a)"
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
        return dot_fallback(c, level)
    elif family == "一、":
        if level == 1:
            return f"{to_cn_num(c[0])}、"
        elif level == 2:
            return f"（{to_cn_num(c[1])}）"
        else:
            return dot_fallback(c, level)
    elif family == "（一）":
        if level == 1:
            return f"（{to_cn_num(c[0])}）"
        elif level == 2:
            return f"{c[1]}. "
        else:
            return dot_fallback(c, level)
    elif family == "I.":
        if level == 1:
            return f"{to_roman(c[0])}. "
        elif level == 2:
            return f"{to_letter(c[1])}. "
        elif level == 3:
            return f"{c[2]}) "
        elif level == 4:
            return f"{to_letter(c[3], upper=False)}) "
        else:
            return dot_fallback(c, level)
    elif family == "i.":
        if level == 1:
            return f"{to_roman(c[0]).lower()}. "
        elif level == 2:
            return f"{to_letter(c[1], upper=False)}. "
        else:
            return dot_fallback(c, level)
    elif family == "1)":
        if level == 1:
            return f"{c[0]}) "
        elif level == 2:
            return f"{to_letter(c[1], upper=False)}) "
        elif level == 3:
            return f"{to_roman(c[2]).lower()}) "
        else:
            return dot_fallback(c, level)
    elif family in ("(1)", "(a)"):
        result = _format_by_style(family, c, level)
        if result:
            return result
        return dot_fallback(c, level)

    return dot_fallback(c, level)


# ---- Ordered list prefix formatting ----

def format_list_prefix(style, num):
    """Format ordered list prefix from style pattern and counter value.
    style is the markdown pattern (e.g. '1.', '1)', 'a.', '(a)')."""
    if style.endswith('.'):
        if style[0].isalpha():
            return f"{to_letter(num, style[0].isupper())}. "
        return f"{num}. "
    elif style.endswith(')'):
        if style[0].isalpha():
            return f"{to_letter(num, style[0].isupper())}) "
        return f"{num}) "
    elif style.startswith('('):
        inner = style.strip('()')
        if inner.isalpha():
            return f"({to_letter(num, inner.isupper())}) "
        return f"({num}) "
    return f"{num}. "
