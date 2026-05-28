"""Formula problem-text guards that depend on inline span extraction."""
from __future__ import annotations

import re

from .classification import _cjk_count, formula_text_looks_contaminated, normalize_math_text
from .inline_spans import split_inline_math_spans


def is_formula_problem_text(text: str) -> bool:
    """True only for short, dense formula-like text that should block QA."""
    t = normalize_math_text(text)
    if not formula_text_looks_contaminated(t):
        return False
    if len(t) > 180:
        return False
    if str(text or "").strip().endswith(("\u3002", "\uff1b", ";", ".")) and not re.search(r"[=<>]\s*$|\u2264\s*$|\u2265\s*$|\u2260\s*$|\u2248\s*$|\u2208\s*$|\u2209\s*$", t):
        return False
    if split_inline_math_spans(text):
        return False
    math_marks = len(re.findall(r"[=<>+\-*/%]|\u2264|\u2265|\u2260|\u2248|\u2208|\u2209", t))
    cjk = _cjk_count(t)
    if re.search(r"[=<>]|\u2264|\u2265|\u2260|\u2248|\u2208|\u2209", t[-3:]):
        return True
    return bool(math_marks >= 3 and cjk <= 20 and len(t) <= 150)
