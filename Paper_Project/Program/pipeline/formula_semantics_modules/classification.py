"""Standalone formula text classification."""
from __future__ import annotations

import re

from .models import FormulaSemanticResult
from .patterns import (
    CATEGORY_CITATION,
    CATEGORY_CONTAMINATED,
    CATEGORY_DISPLAY_MATH,
    CATEGORY_FORMULA_LABEL,
    CATEGORY_QUANTITY_TEXT,
    CATEGORY_TEXT,
    CITATION_RE,
    CJK_RE,
    DIGIT_RE,
    FORMULA_LABEL_RE,
    LATEX_DELIMITED_RE,
    MATH_OP_RE,
    NARRATIVE_WORD_RE,
    RELATION_RE,
    SENTENCE_PUNCT_RE,
    TRANSLATION,
    UNIT_RE,
)


def normalize_math_text(text: str) -> str:
    return re.sub(r"\s+", " ", str(text or "").translate(TRANSLATION)).strip()


def _cjk_count(text: str) -> int:
    return len(CJK_RE.findall(text))


def _has_relation(text: str) -> bool:
    return bool(RELATION_RE.search(normalize_math_text(text)))


def _has_math_operator(text: str) -> bool:
    return bool(MATH_OP_RE.search(normalize_math_text(text)))


def is_citation_text(text: str) -> bool:
    return bool(CITATION_RE.match(str(text or "").strip()))


def is_formula_label(text: str) -> bool:
    return bool(FORMULA_LABEL_RE.match(str(text or "").strip()))


def is_quantity_text(text: str) -> bool:
    t = normalize_math_text(text)
    if not t or not DIGIT_RE.search(t):
        return False
    if _has_relation(t):
        return False
    if _has_math_operator(t) and not re.search(r"\d\s*-\s*\d", t):
        return False
    return bool(UNIT_RE.search(t))


def formula_text_looks_contaminated(text: str) -> bool:
    t = normalize_math_text(text)
    if not t:
        return False
    cjk = _cjk_count(t)
    if cjk < 4:
        return False
    has_formula_signal = bool(
        _has_relation(t)
        or re.search(r"[A-Za-z\u0370-\u03ff][A-Za-z0-9_\u0370-\u03ff]*\s*\(", t)
    )
    if not has_formula_signal or not DIGIT_RE.search(t):
        return False
    has_sentence_punct = bool(SENTENCE_PUNCT_RE.search(t))
    has_narrative_word = bool(NARRATIVE_WORD_RE.search(t))
    if len(t) > 30 and has_sentence_punct and has_narrative_word:
        return True
    if len(t) > 70 and cjk > 12 and has_sentence_punct:
        return True
    if len(t) > 90 and cjk > 18:
        return True
    return False


def formula_should_number(text: str) -> bool:
    t = normalize_math_text(text)
    if not t or formula_text_looks_contaminated(t):
        return False
    if not DIGIT_RE.search(t):
        return False
    return bool(_has_relation(t) and _has_math_operator(t))


def looks_like_formula_text(text: str) -> bool:
    t = normalize_math_text(text)
    if not t or len(t) > 180:
        return False
    if is_citation_text(t) or is_formula_label(t) or is_quantity_text(t):
        return False
    if LATEX_DELIMITED_RE.match(str(text or "").strip()):
        return True
    if formula_text_looks_contaminated(t):
        return False
    cjk = _cjk_count(t)
    if (cjk > 3 and SENTENCE_PUNCT_RE.search(t)) or cjk > 8:
        return False
    starts_continuation = bool(re.match(r"^[=<>]", t)) and bool(DIGIT_RE.search(t))
    if starts_continuation:
        return True
    if _has_relation(t) and (_has_math_operator(t) or DIGIT_RE.search(t)):
        return True
    return False


def classify_formula_text(text: str) -> FormulaSemanticResult:
    raw = str(text or "").strip()
    t = normalize_math_text(raw)
    if not t:
        return FormulaSemanticResult(CATEGORY_TEXT, 1.0, "empty")
    if is_citation_text(t):
        return FormulaSemanticResult(CATEGORY_CITATION, 0.98, "citation marker")
    if is_formula_label(t):
        return FormulaSemanticResult(CATEGORY_FORMULA_LABEL, 0.98, "standalone formula label")
    if formula_text_looks_contaminated(t):
        return FormulaSemanticResult(CATEGORY_CONTAMINATED, 0.92, "narrative text mixed with math operators")
    if is_quantity_text(t):
        return FormulaSemanticResult(CATEGORY_QUANTITY_TEXT, 0.86, "quantity/unit expression without equation relation")
    if LATEX_DELIMITED_RE.match(raw):
        return FormulaSemanticResult(CATEGORY_DISPLAY_MATH, 0.96, "latex delimiter", should_number=formula_should_number(t))
    if looks_like_formula_text(t):
        return FormulaSemanticResult(CATEGORY_DISPLAY_MATH, 0.88, "standalone equation", should_number=formula_should_number(t))
    return FormulaSemanticResult(CATEGORY_TEXT, 0.74, "plain text")
