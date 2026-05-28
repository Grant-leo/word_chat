"""Conservative inline math span extraction."""
from __future__ import annotations

import re
from typing import Dict

from .classification import (
    _cjk_count,
    _has_math_operator,
    _has_relation,
    formula_text_looks_contaminated,
    is_citation_text,
    is_formula_label,
    is_quantity_text,
    normalize_math_text,
)
from .models import FormulaSpan
from .patterns import CATEGORY_INLINE_MATH, DIGIT_RE, DOLLAR_INLINE_RE, INLINE_EQUATION_RE


def _valid_inline_math_candidate(text: str) -> bool:
    t = normalize_math_text(text)
    if not t or len(t) > 140:
        return False
    if is_citation_text(t) or is_formula_label(t) or is_quantity_text(t):
        return False
    if formula_text_looks_contaminated(t):
        return False
    if _cjk_count(t) > 0:
        return False
    for left, right in (("(", ")"), ("{", "}"), ("[", "]")):
        if t.count(left) != t.count(right):
            return False
    if re.search(r"[\u3002\uff0c\uff1b\uff1a,;]|\s[\u4e00-\u9fff]", str(text or "")):
        return False
    if not _has_relation(t):
        return False
    if _has_math_operator(t) or DIGIT_RE.search(t):
        return True
    parts = re.split(r"<=|>=|=|<|>|\u2264|\u2265|\u2248|\u2260", t, maxsplit=1)
    return len(parts) == 2 and all(re.search(r"[A-Za-z\u0370-\u03ff]", p) for p in parts)


def _valid_dollar_inline_math(text: str) -> bool:
    t = normalize_math_text(text)
    if not t or len(t) > 120:
        return False
    if is_citation_text(t) or is_formula_label(t) or is_quantity_text(t):
        return False
    if _cjk_count(t) > 0:
        return False
    for left, right in (("(", ")"), ("{", "}"), ("[", "]")):
        if t.count(left) != t.count(right):
            return False
    if re.search(r"\\[A-Za-z]+|[_^]|[=<>]|\u2264|\u2265|\u2248|\u2260|[+\-*/%]|\u2211|\u222b|\u221a", t):
        return True
    if re.fullmatch(r"[A-Za-z\u0370-\u03ff][A-Za-z0-9\u0370-\u03ff]*(?:\([A-Za-z0-9]+\))?", t):
        return True
    return False


def _trim_inline_candidate(text: str) -> str:
    t = str(text or "").strip()
    while t and t[-1] in ")]\uff09\uff3d":
        normalized = normalize_math_text(t)
        if normalized.count(")") <= normalized.count("(") and normalized.count("]") <= normalized.count("["):
            break
        t = t[:-1].rstrip()
    return t


def split_inline_math_spans(text: str) -> list[Dict[str, object]]:
    """Return conservative inline math spans from a normal prose paragraph."""
    raw = str(text or "")
    if not raw or len(raw) > 1200:
        return []
    if re.match(r"^\s*\$\$.+\$\$\s*$", raw, re.S):
        return []
    spans: list[FormulaSpan] = []

    def overlaps(start: int, end: int) -> bool:
        return any(not (end <= span.start or start >= span.end) for span in spans)

    for match in DOLLAR_INLINE_RE.finditer(raw):
        inner = match.group(1).strip()
        if not inner or overlaps(match.start(), match.end()) or not _valid_dollar_inline_math(inner):
            continue
        spans.append(
            FormulaSpan(
                match.start(),
                match.end(),
                inner,
                CATEGORY_INLINE_MATH,
                0.96,
                "dollar-delimited inline math",
                latex=inner,
            )
        )

    for match in INLINE_EQUATION_RE.finditer(raw):
        candidate = _trim_inline_candidate(match.group(1))
        start = match.start(1) + (len(match.group(1)) - len(match.group(1).lstrip()))
        end = start + len(candidate)
        if overlaps(start, end) or not _valid_inline_math_candidate(candidate):
            continue
        spans.append(
            FormulaSpan(
                start,
                end,
                candidate,
                CATEGORY_INLINE_MATH,
                0.84,
                "inline equation span",
            )
        )

    spans.sort(key=lambda span: span.start)
    return [span.to_dict() for span in spans]
