"""Shared helpers for split-formula layout repair."""
from __future__ import annotations

import re

try:
    from formula_semantics import classify_formula_text
except ImportError:  # pragma: no cover - package-style imports
    from ..formula_semantics import classify_formula_text

try:
    from content_parser_modules.formula_text_items import _clean_formula_text, _rich_text_item_from_inline_formula_spans
except ImportError:  # pragma: no cover - package-style imports
    from .formula_text_items import _clean_formula_text, _rich_text_item_from_inline_formula_spans

def _item_text(item):
    if isinstance(item, str):
        return item
    if isinstance(item, dict):
        return item.get('text') or item.get('code') or ''
    return ''


def _item_role(item):
    return item.get('role') if isinstance(item, dict) else 'text'


def _is_formula_like_item(item):
    return isinstance(item, dict) and (item.get('role') == 'formula' or item.get('latex') or item.get('xml') or item.get('math'))


def _is_split_formula_fragment(item):
    text = str(_item_text(item) or '').strip()
    if not text or len(text) > 18:
        return False
    if re.search(r'[\u4e00-\u9fff]', text):
        return False
    return bool(re.fullmatch(r'[A-Za-z0-9_\s∆ΔλΛµμ%+\-*/·×÷=<>≤≥≈∈∑().,α-ωΑ-Ω]+', text))


def _is_ratio_variable_fragment(text):
    t = re.sub(r'\s+', '', str(text or ''))
    return bool(re.fullmatch(r'[a-zα-ω][A-Za-z0-9_α-ωΑ-Ω]*', t))


def _latex_identifier(token):
    t = re.sub(r'\s+', '', str(token or ''))
    if not t:
        return ''
    greek = {
        '∆': r'\Delta',
        'Δ': r'\Delta',
        'λ': r'\lambda',
        'Λ': r'\Lambda',
        'μ': r'\mu',
        'α': r'\alpha',
        'β': r'\beta',
        'γ': r'\gamma',
    }
    if t in greek:
        return greek[t]
    m = re.fullmatch(r'([∆ΔλΛμ])([A-Za-z0-9]+)', t)
    if m:
        base, suffix = m.groups()
        command = greek.get(base, base)
        if base in ('∆', 'Δ'):
            return command + ' ' + _latex_identifier(suffix)
        return command + r'_{\mathrm{' + suffix + '}}'
    if re.fullmatch(r'[A-Za-z][A-Za-z0-9]*', t):
        if len(t) == 1:
            return t
        return t[0] + r'_{\mathrm{' + t[1:] + '}}'
    return t


def _latex_math_expr(text, sum_lower=None, sum_upper=None):
    s = str(text or '').strip()
    s = s.replace('−', '-').replace('－', '-').replace('＝', '=')
    s = s.replace('×', r'\times ').replace('·', r'\cdot ').replace('÷', r'\div ')
    s = s.replace('%', r'\%')
    if sum_lower and sum_upper:
        s = s.replace('∑', r'\sum_{' + sum_lower + '}^{' + str(sum_upper) + '}')
    else:
        s = s.replace('∑', r'\sum')
    s = re.sub(r'[∆ΔλΛμ][A-Za-z0-9]*', lambda m: _latex_identifier(m.group(0)), s)
    s = re.sub(r'(?<![\\{])\b[A-Za-z][A-Za-z0-9]*\b', lambda m: _latex_identifier(m.group(0)), s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s


def _repaired_formula_item(text, latex, numbered=False, repair='split_formula_layout'):
    semantic = classify_formula_text(text)
    return {
        'role': 'formula',
        'source': 'repaired_' + repair,
        'text': text,
        'latex': latex,
        'numbered': bool(numbered),
        'formula_semantics': semantic.to_dict(),
    }


def _split_formula_problem_item(text, problem='split_formula_layout'):
    clean = _clean_formula_text(text)
    semantic = classify_formula_text(clean)
    return {
        'role': 'formula_problem',
        'problem': problem,
        'source': 'repair',
        'text': clean,
        'formula_semantics': semantic.to_dict(),
    }


def _split_percentage_suffix(rhs):
    text = str(rhs or '').strip()
    m = re.match(r'^(.*?)(?:[×x*]\s*100\s*%|\\times\s*100\s*%)$', text)
    if m:
        return m.group(1).strip(), True
    return text, False


def _append_repair_tail(out, tail):
    text = str(tail or '').strip()
    if not text:
        return
    rich = _rich_text_item_from_inline_formula_spans(text)
    out.append(rich or text)


def _split_formula_expression_tail(text):
    s = str(text or '').strip()
    if not s:
        return '', ''
    m = re.search(r'[\u4e00-\u9fff]', s)
    if not m:
        return s, ''
    idx = m.start()
    if idx > 0 and s[idx - 1] in '（(':
        idx -= 1
    return s[:idx].strip().rstrip('，,。;；'), s[idx:].strip()


def _infer_sum_lower(expr):
    text = str(expr or '')
    candidates = []
    candidates.extend(re.findall(r'\b[A-Za-z][A-Za-z0-9]*\s*\(\s*([A-Za-z])\s*\)', text))
    candidates.extend(re.findall(r'\b[A-Za-z][A-Za-z0-9]*\s*_\s*([A-Za-z])\b', text))
    candidates = [c for c in candidates if re.fullmatch(r'[A-Za-z]', c or '')]
    if not candidates:
        return None
    unique = set(candidates)
    if len(unique) != 1:
        return None
    return candidates[0] + '=1'
