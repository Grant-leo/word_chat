"""Repairs for formula layouts fragmented by source DOCX extraction."""
from __future__ import annotations

import re

try:
    from content_parser_modules.formula_repair_strategies import (
        _repair_fraction_sum_layout,
        _repair_labeled_inline_sum_missing_lower,
        _repair_labeled_sum_continuation,
        _repair_max_sum_layout,
        _repair_missing_sum_symbol_bounds,
        _repair_split_ratio_cluster,
        _repair_split_sum_bounds,
        _repair_split_sum_prefix,
    )
    from content_parser_modules.formula_repair_utils import _is_split_formula_fragment, _item_text
except ImportError:  # pragma: no cover - package-style imports
    from .formula_repair_strategies import (
        _repair_fraction_sum_layout,
        _repair_labeled_inline_sum_missing_lower,
        _repair_labeled_sum_continuation,
        _repair_max_sum_layout,
        _repair_missing_sum_symbol_bounds,
        _repair_split_ratio_cluster,
        _repair_split_sum_bounds,
        _repair_split_sum_prefix,
    )
    from .formula_repair_utils import _is_split_formula_fragment, _item_text

def repair_split_formula_layouts(paragraphs):
    """Repair formula layouts that were already fragmented in a source DOCX."""
    out = []
    i = 0
    while i < len(paragraphs):
        if _is_split_formula_fragment(paragraphs[i]):
            max_sum_repair = _repair_max_sum_layout(paragraphs, i)
            if max_sum_repair:
                repaired, next_i = max_sum_repair
                out.extend(repaired)
                i = next_i
                continue
            fraction_sum_repair = _repair_fraction_sum_layout(paragraphs, i)
            if fraction_sum_repair:
                repaired, next_i = fraction_sum_repair
                out.extend(repaired)
                i = next_i
                continue
            sum_repair = _repair_split_sum_bounds(paragraphs, i)
            if sum_repair:
                repaired, next_i = sum_repair
                out.extend(repaired)
                i = next_i
                continue
            missing_sum_repair = _repair_missing_sum_symbol_bounds(paragraphs, i)
            if missing_sum_repair:
                repaired, next_i = missing_sum_repair
                out.extend(repaired)
                i = next_i
                continue
            sum_prefix_repair = _repair_split_sum_prefix(paragraphs, i)
            if sum_prefix_repair:
                repaired, next_i = sum_prefix_repair
                out.extend(repaired)
                i = next_i
                continue
            ratio_repair = _repair_split_ratio_cluster(paragraphs, i)
            if ratio_repair:
                if out and re.fullmatch(r'[A-Za-z]\s*=\s*\d+', str(_item_text(out[-1]) or '').strip()):
                    out.pop()
                repaired, next_i = ratio_repair
                out.extend(repaired)
                i = next_i
                continue
        labeled_sum_repair = _repair_labeled_sum_continuation(paragraphs, i, out)
        if labeled_sum_repair:
            repaired, next_i = labeled_sum_repair
            out.extend(repaired)
            i = next_i
            continue
        inline_sum_repair = _repair_labeled_inline_sum_missing_lower(paragraphs[i], out)
        if inline_sum_repair:
            repaired, _ = inline_sum_repair
            out.extend(repaired)
            i += 1
            continue
        out.append(paragraphs[i])
        i += 1
    return out
