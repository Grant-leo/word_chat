"""Compatibility exports for content formula extraction helpers."""
from __future__ import annotations

try:
    from content_parser_modules.formula_labels import (
        _math_text,
        _should_strip_trailing_formula_labels,
        _split_trailing_formula_labels,
        _strip_trailing_formula_labels,
        _strip_trailing_formula_labels_from_xml,
    )
    from content_parser_modules.formula_omml import extract_math
    from content_parser_modules.formula_repairs import repair_split_formula_layouts
    from content_parser_modules.formula_text_items import (
        _clean_formula_text,
        _default_clean_text_artifacts,
        _formula_item_from_text,
        _formula_problem_item_from_text,
        _formula_should_number,
        _latex_escape_text,
        _latex_from_formula_text,
        _looks_like_formula_text,
        _omml_text_looks_like_body,
        _rich_text_item_from_inline_formula_spans,
        set_clean_text_artifacts_func,
    )
except ImportError:  # pragma: no cover - package-style imports
    from .formula_labels import (
        _math_text,
        _should_strip_trailing_formula_labels,
        _split_trailing_formula_labels,
        _strip_trailing_formula_labels,
        _strip_trailing_formula_labels_from_xml,
    )
    from .formula_omml import extract_math
    from .formula_repairs import repair_split_formula_layouts
    from .formula_text_items import (
        _clean_formula_text,
        _default_clean_text_artifacts,
        _formula_item_from_text,
        _formula_problem_item_from_text,
        _formula_should_number,
        _latex_escape_text,
        _latex_from_formula_text,
        _looks_like_formula_text,
        _omml_text_looks_like_body,
        _rich_text_item_from_inline_formula_spans,
        set_clean_text_artifacts_func,
    )

__all__ = [
    "_clean_formula_text",
    "_default_clean_text_artifacts",
    "_formula_item_from_text",
    "_formula_problem_item_from_text",
    "_formula_should_number",
    "_latex_escape_text",
    "_latex_from_formula_text",
    "_looks_like_formula_text",
    "_math_text",
    "_omml_text_looks_like_body",
    "_rich_text_item_from_inline_formula_spans",
    "_should_strip_trailing_formula_labels",
    "_split_trailing_formula_labels",
    "_strip_trailing_formula_labels",
    "_strip_trailing_formula_labels_from_xml",
    "extract_math",
    "repair_split_formula_layouts",
    "set_clean_text_artifacts_func",
]
