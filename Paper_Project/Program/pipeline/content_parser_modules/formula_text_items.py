"""Text formula classification and item builders for content parsing."""
from __future__ import annotations

import re

try:
    from formula_semantics import (
        classify_formula_text,
        formula_should_number as semantic_formula_should_number,
        looks_like_formula_text as semantic_looks_like_formula_text,
        split_inline_math_spans,
    )
except ImportError:  # pragma: no cover - package-style imports
    from ..formula_semantics import (
        classify_formula_text,
        formula_should_number as semantic_formula_should_number,
        looks_like_formula_text as semantic_looks_like_formula_text,
        split_inline_math_spans,
    )

try:
    from content_parser_modules.formula_labels import _split_trailing_formula_labels
except ImportError:  # pragma: no cover - package-style imports
    from .formula_labels import _split_trailing_formula_labels

def _default_clean_text_artifacts(text, preserve_newlines=False):
    t = str(text or '').replace('\u00a0', ' ')
    if preserve_newlines:
        lines = []
        for line in t.replace('\r\n', '\n').replace('\r', '\n').split('\n'):
            line = re.sub(r'[ \t]+', ' ', line).strip()
            if line:
                lines.append(line)
        return '\n'.join(lines).strip()
    return re.sub(r'\s+', ' ', t).strip()


_clean_text_artifacts = _default_clean_text_artifacts


def set_clean_text_artifacts_func(func):
    """Let the orchestration parser provide its artifact cleaner."""
    global _clean_text_artifacts
    if callable(func):
        _clean_text_artifacts = func


def _clean_formula_text(text):
    t = _clean_text_artifacts(text)
    if t.count('|') >= 3:
        t = t.replace('|', '')
    t = re.sub(r'\s+', ' ', t).strip()
    return t


def _looks_like_formula_text(text):
    """Detect standalone calculation/formula paragraphs.

    The rule is intentionally structural: formulas are short standalone lines
    with equality/calculation operators. Some thesis sources store formulas as
    plain text, including definition lines without numbers and continuation
    lines that start with "=".
    """
    return semantic_looks_like_formula_text(text)


def _latex_escape_text(text):
    return str(text or '').replace('\\', r'\backslash ').replace('{', r'\{').replace('}', r'\}')


def _latex_from_formula_text(text):
    t = str(text or '').strip()
    body, labels = _split_trailing_formula_labels(t)
    candidates = [body] if labels else []
    candidates.append(t)
    for candidate in candidates:
        candidate = str(candidate or '').strip()
        if candidate.startswith('$$') and candidate.endswith('$$'):
            return candidate[2:-2].strip()
        if candidate.startswith('$') and candidate.endswith('$'):
            return candidate[1:-1].strip()
    return ''


def _formula_should_number(text):
    if _omml_text_looks_like_body(text):
        return False
    return semantic_formula_should_number(text)


def _omml_text_looks_like_body(text):
    t = str(text or '').strip()
    if len(t) > 220:
        return True
    cjk = len(re.findall(r'[\u4e00-\u9fff]', t))
    if len(t) > 35 and cjk > 18 and re.search(r'(表明|显示|说明|分析|结果|选择|问题|模型|成本|指标)', t):
        return True
    if len(t) > 90 and cjk > 20 and re.search(r'[。；;，,\.]', t):
        return True
    return False


def _formula_item_from_text(text):
    clean = _clean_formula_text(text)
    semantic = classify_formula_text(clean)
    if semantic.category == 'CONTAMINATED':
        return _formula_problem_item_from_text(clean)
    item = {
        'role': 'formula',
        'source': 'text',
        'text': clean,
        'numbered': _formula_should_number(clean),
        'formula_semantics': semantic.to_dict(),
    }
    latex = _latex_from_formula_text(clean)
    if latex:
        item['source'] = 'latex'
        item['latex'] = latex
    return item


def _formula_problem_item_from_text(text):
    clean = _clean_formula_text(text)
    semantic = classify_formula_text(clean)
    return {
        'role': 'formula_problem',
        'problem': 'contaminated_formula_text',
        'source': 'text',
        'text': clean,
        'formula_semantics': semantic.to_dict(),
    }


def _rich_text_item_from_inline_formula_spans(text):
    spans = split_inline_math_spans(text)
    if not spans:
        return None
    runs = []
    math_entries = []
    pos = 0
    for span in spans:
        start = int(span.get('start') or 0)
        end = int(span.get('end') or start)
        if start > pos:
            runs.append({'type': 'text', 'text': text[pos:start]})
        formula_text = str(span.get('text') or '').strip()
        if not formula_text:
            pos = max(pos, end)
            continue
        entry = {
            'type': 'inline',
            'text': formula_text,
            'formula_semantics': span,
        }
        if span.get('latex'):
            entry['latex'] = span.get('latex')
        math_entries.append(entry)
        runs.append({'type': 'math', 'text': formula_text, 'math': [entry]})
        pos = max(pos, end)
    if pos < len(text):
        runs.append({'type': 'text', 'text': text[pos:]})
    if not math_entries:
        return None
    return {
        'role': 'rich_text',
        'text': text,
        'runs': runs,
        'math': math_entries,
    }

