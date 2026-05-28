"""Stable public entry point for formula semantic classification."""
from __future__ import annotations

from formula_semantics_modules.classification import (
    classify_formula_text,
    formula_should_number,
    formula_text_looks_contaminated,
    is_citation_text,
    is_formula_label,
    is_quantity_text,
    looks_like_formula_text,
    normalize_math_text,
)
from formula_semantics_modules.inline_spans import split_inline_math_spans
from formula_semantics_modules.models import FormulaSemanticResult, FormulaSpan
from formula_semantics_modules.patterns import (
    CATEGORY_CITATION,
    CATEGORY_CONTAMINATED,
    CATEGORY_DISPLAY_MATH,
    CATEGORY_FORMULA_LABEL,
    CATEGORY_INLINE_MATH,
    CATEGORY_QUANTITY_TEXT,
    CATEGORY_TEXT,
    CATEGORY_UNIT_TEXT,
)
from formula_semantics_modules.problem_detection import is_formula_problem_text

__all__ = [
    "CATEGORY_CITATION",
    "CATEGORY_CONTAMINATED",
    "CATEGORY_DISPLAY_MATH",
    "CATEGORY_FORMULA_LABEL",
    "CATEGORY_INLINE_MATH",
    "CATEGORY_QUANTITY_TEXT",
    "CATEGORY_TEXT",
    "CATEGORY_UNIT_TEXT",
    "FormulaSemanticResult",
    "FormulaSpan",
    "classify_formula_text",
    "formula_should_number",
    "formula_text_looks_contaminated",
    "is_citation_text",
    "is_formula_label",
    "is_formula_problem_text",
    "is_quantity_text",
    "looks_like_formula_text",
    "normalize_math_text",
    "split_inline_math_spans",
]
