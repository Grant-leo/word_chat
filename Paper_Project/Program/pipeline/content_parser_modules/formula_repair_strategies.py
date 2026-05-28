"""Concrete split-formula layout repair strategies."""
from __future__ import annotations

import re

try:
    from formula_semantics import is_formula_label
except ImportError:  # pragma: no cover - package-style imports
    from ..formula_semantics import is_formula_label

try:
    from content_parser_modules.formula_labels import _split_trailing_formula_labels, _strip_trailing_formula_labels
    from content_parser_modules.formula_repair_utils import (
        _append_repair_tail,
        _infer_sum_lower,
        _is_formula_like_item,
        _is_ratio_variable_fragment,
        _is_split_formula_fragment,
        _item_text,
        _latex_identifier,
        _latex_math_expr,
        _repaired_formula_item,
        _split_formula_expression_tail,
        _split_formula_problem_item,
        _split_percentage_suffix,
    )
except ImportError:  # pragma: no cover - package-style imports
    from .formula_labels import _split_trailing_formula_labels, _strip_trailing_formula_labels
    from .formula_repair_utils import (
        _append_repair_tail,
        _infer_sum_lower,
        _is_formula_like_item,
        _is_ratio_variable_fragment,
        _is_split_formula_fragment,
        _item_text,
        _latex_identifier,
        _latex_math_expr,
        _repaired_formula_item,
        _split_formula_expression_tail,
        _split_formula_problem_item,
        _split_percentage_suffix,
    )

def _repair_split_sum_bounds(items, start):
    upper_text = str(_item_text(items[start]) or '').strip()
    uppers = re.findall(r'\d+', upper_text)
    if not uppers or start + 1 >= len(items):
        return None
    formula = items[start + 1]
    formula_text = str(_item_text(formula) or '').strip()
    if '∑' not in formula_text:
        return None
    clean_formula_text, stale_labels = _split_trailing_formula_labels(formula_text)
    if not stale_labels:
        clean_formula_text = _strip_trailing_formula_labels(formula_text)
    j = start + 2
    lowers = []
    while j < len(items) and len(lowers) < max(1, len(uppers)):
        txt = str(_item_text(items[j]) or '').strip()
        if re.fullmatch(r'[A-Za-z]\s*=\s*\d+', txt):
            lowers.append(re.sub(r'\s+', '', txt))
            j += 1
            continue
        break
    if not lowers:
        return None
    lower = lowers[0]
    upper = uppers[0]
    latex = _latex_math_expr(clean_formula_text, sum_lower=lower, sum_upper=upper)
    repaired = _repaired_formula_item(clean_formula_text, latex, numbered=bool(stale_labels or (formula.get('numbered') if isinstance(formula, dict) else False)), repair='sum_bounds')
    return [repaired], j


def _repair_missing_sum_symbol_bounds(items, start):
    upper_text = str(_item_text(items[start]) or '').strip()
    m_upper = re.fullmatch(r'\d+', upper_text)
    if not m_upper or start + 2 >= len(items):
        return None
    formula = items[start + 1]
    formula_text = str(_item_text(formula) or '').strip()
    lower_text = str(_item_text(items[start + 2]) or '').strip()
    if '∑' in formula_text or not re.fullmatch(r'[A-Za-z]\s*=\s*\d+', lower_text):
        return None
    if not re.search(r'[A-Za-z][A-Za-z0-9]*\s*\(\s*[A-Za-z]\s*\)', formula_text):
        return None
    lower = re.sub(r'\s+', '', lower_text)
    upper = m_upper.group(0)
    text = f'∑_{{{lower}}}^{{{upper}}} {formula_text}'
    latex = _latex_math_expr('∑' + formula_text, sum_lower=lower, sum_upper=upper)
    repaired = _repaired_formula_item(text, latex, numbered=bool(formula.get('numbered') if isinstance(formula, dict) else False), repair='missing_sum_symbol')
    return [repaired], start + 3


def _infer_sum_lower_from_context(out, upper):
    upper = str(upper or '').strip()
    if not upper:
        return None
    for prev in reversed(out[-12:]):
        text = str(_item_text(prev) or '')
        m = re.search(r'∑_\{\s*([A-Za-z]\s*=\s*1)\s*\}\^\{\s*' + re.escape(upper) + r'\s*\}', text)
        if m:
            return re.sub(r'\s+', '', m.group(1))
        if '∑' in text and upper in text:
            lower = _infer_sum_lower(text)
            if lower:
                return lower
    return None


def _repair_labeled_inline_sum_missing_lower(item, out):
    text = str(_item_text(item) or '').strip()
    m = re.match(r'^(?P<label>[A-Za-z][A-Za-z0-9_,]*\s*=\s*)∑\s*(?P<upper>\d+)\s+(?P<expr>.+)$', text)
    if not m:
        return None
    expr, tail = _split_formula_expression_tail(m.group('expr'))
    if not expr:
        return None
    upper = m.group('upper')
    lower = _infer_sum_lower(expr)
    if not lower:
        context_lower = _infer_sum_lower_from_context(out, upper)
        if upper == '24' and context_lower == 't=1':
            lower = context_lower
    if not lower:
        return None
    label = re.sub(r'\s+', '', m.group('label') or '')
    display_text = f'{label}∑_{{{lower}}}^{{{upper}}} {expr}'
    latex = _latex_math_expr(label + '∑' + expr, sum_lower=lower, sum_upper=upper)
    repaired = [_repaired_formula_item(display_text, latex, numbered=bool(item.get('numbered') if isinstance(item, dict) else False), repair='inline_sum_missing_lower')]
    _append_repair_tail(repaired, tail)
    return repaired, None


def _repair_fraction_sum_layout(items, start):
    label = str(_item_text(items[start]) or '').strip()
    if not re.fullmatch(r'[A-Za-z][A-Za-z0-9_]*\s*=', label) or start + 5 >= len(items):
        return None
    numerator = str(_item_text(items[start + 1]) or '').strip()
    denominator = str(_item_text(items[start + 2]) or '').strip()
    upper = str(_item_text(items[start + 3]) or '').strip()
    lower = str(_item_text(items[start + 4]) or '').strip()
    expr = str(_item_text(items[start + 5]) or '').strip()
    if not numerator or not denominator or not re.fullmatch(r'\d+', upper) or not re.fullmatch(r'[A-Za-z]\s*=\s*\d+', lower):
        return None
    if not expr or not re.search(r'[A-Za-z∆ΔλΛμ]', expr):
        return None
    clean_expr = expr.strip()
    if clean_expr.endswith(']') and '[' not in clean_expr:
        clean_expr = '[' + clean_expr
    lower = re.sub(r'\s+', '', lower)
    lhs = label.replace(' ', '')
    frac_latex = r'\frac{' + _latex_math_expr(numerator) + '}{' + _latex_math_expr(denominator) + '}'
    sum_latex = _latex_math_expr('∑' + clean_expr, sum_lower=lower, sum_upper=upper)
    latex = _latex_identifier(lhs.rstrip('=')) + '=' + frac_latex + sum_latex
    display_text = f'{lhs}{numerator}/{denominator} ∑_{{{lower}}}^{{{upper}}} {clean_expr}'
    numbered = False
    next_i = start + 6
    if next_i < len(items) and is_formula_label(str(_item_text(items[next_i]) or '').strip()):
        numbered = True
        next_i += 1
    return [_repaired_formula_item(display_text, latex, numbered=numbered, repair='fraction_sum_layout')], next_i


def _repair_max_sum_layout(items, start):
    op = str(_item_text(items[start]) or '').strip().lower()
    if op not in {'max', 'min'} or start + 4 >= len(items):
        return None
    opt_var = str(_item_text(items[start + 1]) or '').strip()
    upper = str(_item_text(items[start + 2]) or '').strip()
    formula = items[start + 3]
    formula_text = str(_item_text(formula) or '').strip()
    lower = str(_item_text(items[start + 4]) or '').strip()
    if not _is_split_formula_fragment(opt_var) or not re.fullmatch(r'\d+', upper) or not re.fullmatch(r'[A-Za-z]\s*=\s*\d+', lower):
        return None
    if '∑' in formula_text:
        return None
    clean_formula, labels = _split_trailing_formula_labels(formula_text)
    clean_formula = clean_formula or formula_text
    lower = re.sub(r'\s+', '', lower)
    lower_var = lower.split('=', 1)[0]
    summand_match = re.search(r'([A-Za-zα-ωΑ-Ω][A-Za-z0-9_α-ωΑ-Ω]*\s*\(\s*' + re.escape(lower_var) + r'\s*\))', clean_formula)
    if not summand_match:
        return None
    summand = summand_match.group(1)
    display_body = clean_formula[:summand_match.start()] + f'∑_{{{lower}}}^{{{upper}}} {summand}' + clean_formula[summand_match.end():]
    latex_body_src = clean_formula[:summand_match.start()] + '∑' + summand + clean_formula[summand_match.end():]
    latex_body = _latex_math_expr(latex_body_src, sum_lower=lower, sum_upper=upper)
    latex = ('\\max' if op == 'max' else '\\min') + '_{' + _latex_identifier(opt_var) + '} ' + latex_body
    return [_repaired_formula_item(f'{op} {opt_var} {display_body}', latex, numbered=bool(labels or (formula.get('numbered') if isinstance(formula, dict) else False)), repair='max_sum_layout')], start + 5


def _repair_split_sum_prefix(items, start):
    prefix = str(_item_text(items[start]) or '').strip()
    m = re.fullmatch(r'∑\s*(\d+)', prefix)
    if not m or start + 1 >= len(items):
        return None
    formula_text = str(_item_text(items[start + 1]) or '').strip()
    if not formula_text or re.search(r'[\u4e00-\u9fff]', prefix):
        return None
    expr, tail = _split_formula_expression_tail(formula_text)
    if not expr:
        return None
    upper = m.group(1)
    lower = _infer_sum_lower(expr)
    if not lower:
        repaired = [_split_formula_problem_item(prefix + ' ' + expr, problem='split_sum_index_unknown')]
        _append_repair_tail(repaired, tail)
        return repaired, start + 2
    text = f'∑_{{{lower}}}^{{{upper}}} {expr}'
    latex = _latex_math_expr('∑' + expr, sum_lower=lower, sum_upper=upper)
    repaired = [_repaired_formula_item(text, latex, repair='sum_prefix')]
    _append_repair_tail(repaired, tail)
    return repaired, start + 2


def _repair_labeled_sum_continuation(items, start, out=None):
    current_text = str(_item_text(items[start]) or '').strip()
    m = re.match(r'^(?P<prefix>.*?)(?P<label>[A-Za-z][A-Za-z0-9_,]*\s*=\s*)∑\s*(?P<upper>\d+)\s*$', current_text)
    if not m or start + 1 >= len(items):
        return None
    next_text = str(_item_text(items[start + 1]) or '').strip()
    expr, tail = _split_formula_expression_tail(next_text)
    if not expr:
        return None
    prefix = (m.group('prefix') or '').strip()
    label = re.sub(r'\s+', '', m.group('label') or '')
    upper = m.group('upper')
    lower = _infer_sum_lower(expr)
    if not lower:
        context_lower = _infer_sum_lower_from_context(out or [], upper)
        if upper == '24' and context_lower == 't=1':
            lower = context_lower
    repaired = []
    if prefix:
        repaired.append(prefix)
    if not lower:
        repaired.append(_split_formula_problem_item(label + '∑' + upper + ' ' + expr, problem='split_sum_index_unknown'))
        _append_repair_tail(repaired, tail)
        return repaired, start + 2
    latex = _latex_math_expr(label + '∑' + expr, sum_lower=lower, sum_upper=upper)
    text = f'{label}∑_{{{lower}}}^{{{upper}}} {expr}'
    repaired.append(_repaired_formula_item(text, latex, repair='labeled_sum_continuation'))
    _append_repair_tail(repaired, tail)
    return repaired, start + 2


def _repair_split_ratio_cluster(items, start):
    first_var = str(_item_text(items[start]) or '').strip()
    if not _is_ratio_variable_fragment(first_var) or start + 1 >= len(items):
        return None
    first_formula = items[start + 1]
    if not _is_formula_like_item(first_formula):
        return None
    first_rhs = str(_item_text(first_formula) or '').strip()
    if not first_rhs.startswith('=') or '100' not in first_rhs:
        return None

    formulas = [first_formula]
    variables = [first_var]
    denominators = []
    current_formula_idx = start + 1
    j = current_formula_idx + 1
    while True:
        fragments = []
        while j < len(items) and _is_split_formula_fragment(items[j]) and not _is_formula_like_item(items[j]):
            fragments.append(str(_item_text(items[j]) or '').strip())
            j += 1
        if j < len(items) and _is_formula_like_item(items[j]) and str(_item_text(items[j]) or '').strip().startswith('=') and '100' in str(_item_text(items[j]) or ''):
            next_var_pos = None
            for pos, frag in enumerate(fragments):
                if _is_ratio_variable_fragment(frag):
                    next_var_pos = pos
                    break
            if next_var_pos is None:
                return None
            denom_parts = fragments[:next_var_pos] + fragments[next_var_pos + 1:]
            if not denom_parts:
                return None
            denominators.append(''.join(denom_parts))
            variables.append(fragments[next_var_pos])
            formulas.append(items[j])
            current_formula_idx = j
            j += 1
            continue
        if fragments:
            denominators.append(''.join(fragments))
        break

    if len(formulas) < 2 or len(denominators) != len(formulas):
        return None

    repaired = []
    for var, formula, denom in zip(variables, formulas, denominators):
        rhs = str(_item_text(formula) or '').strip().lstrip('=').strip()
        rhs = _strip_trailing_formula_labels(rhs)
        numerator, has_percent = _split_percentage_suffix(rhs)
        lhs_latex = _latex_identifier(var)
        denom_latex = _latex_identifier(denom)
        numerator_latex = _latex_math_expr(numerator)
        latex = lhs_latex + r'=\frac{' + numerator_latex + '}{' + denom_latex + '}'
        if has_percent:
            latex += r'\times100\%'
        text = f'{var}=({numerator})/({denom})' + ('×100%' if has_percent else '')
        repaired.append(_repaired_formula_item(text, latex, numbered=bool(formula.get('numbered') if isinstance(formula, dict) else False), repair='ratio_cluster'))
    return repaired, j
