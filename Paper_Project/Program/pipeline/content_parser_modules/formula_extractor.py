"""Formula extraction and repair helpers for content_parser.py."""
from lxml import etree
import re

try:
    from formula_semantics import (
        classify_formula_text,
        formula_should_number as semantic_formula_should_number,
        is_formula_label,
        is_formula_problem_text,
        looks_like_formula_text as semantic_looks_like_formula_text,
        split_inline_math_spans,
    )
except ImportError:  # pragma: no cover - package-style imports
    from ..formula_semantics import (
        classify_formula_text,
        formula_should_number as semantic_formula_should_number,
        is_formula_label,
        is_formula_problem_text,
        looks_like_formula_text as semantic_looks_like_formula_text,
        split_inline_math_spans,
    )


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


def _math_text(elem):
    """Extract plain text from a math OOXML element."""
    M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
    parts = []
    for t in elem.iter(f'{{{M}}}t'):
        if t.text:
            parts.append(t.text)
    return ''.join(parts)


_FORMULA_LABEL_BODY_RE = r'(?:\d+(?:\s*[-.]\s*\d+)?|[A-Za-z]\s*(?:[-.]\s*)?\d+(?:\s*[-.]\s*\d+)?)'
_FORMULA_TRAILING_LABEL_RE = re.compile(r'((?:\s*[\(\uff08]\s*' + _FORMULA_LABEL_BODY_RE + r'\s*[\)\uff09])+\s*)$')


def _split_trailing_formula_labels(text):
    t = str(text or '').strip()
    m = _FORMULA_TRAILING_LABEL_RE.search(t)
    if not m:
        return t, []
    labels = re.findall(r'[\(\uff08]\s*(' + _FORMULA_LABEL_BODY_RE + r')\s*[\)\uff09]', m.group(1))
    return t[:m.start()].strip(), labels


def _should_strip_trailing_formula_labels(text):
    body, labels = _split_trailing_formula_labels(text)
    if not body or not labels:
        return False
    # A single trailing "(1)" can be a function argument or superscript marker
    # after OOXML is flattened to text, e.g. f(1) or x^{(1)}. Treat it as an
    # equation number only when the preceding body has equation-like structure.
    if len(labels) == 1 and re.search(r'[A-Za-z\u0370-\u03ff]$', body):
        return False
    if len(labels) >= 2:
        return True
    return bool(re.search(r'[=＝≈≤≥<>]', body) and re.search(r'\d|[+*/×÷%·]', body))


def _strip_trailing_formula_labels(text):
    """Remove stale equation numbers copied from source documents."""
    if not _should_strip_trailing_formula_labels(text):
        return str(text or '').strip()
    body, _labels = _split_trailing_formula_labels(text)
    return body


def _strip_trailing_formula_labels_from_xml(xml):
    """Strip trailing formula labels from m:t nodes while preserving OMML."""
    try:
        root = etree.fromstring(str(xml or '').encode('utf-8'))
    except Exception:
        return xml, '', False
    M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
    nodes = [n for n in root.iter(f'{{{M}}}t') if n.text]
    original = ''.join(n.text or '' for n in nodes)
    stripped = _strip_trailing_formula_labels(original)
    if not original or stripped == original:
        return xml, original, False
    remove_chars = len(original) - len(stripped)
    for node in reversed(nodes):
        if remove_chars <= 0:
            break
        txt = node.text or ''
        if remove_chars >= len(txt):
            node.text = ''
            remove_chars -= len(txt)
        else:
            node.text = txt[:-remove_chars]
            remove_chars = 0
    return etree.tounicode(root, with_tail=False), stripped, True


def extract_math(para):
    """Extract OOXML math elements from a paragraph. Returns (text, math_list).
    math_list entries: {'type': 'inline'|'display', 'xml': escaped_xml_string, 'text': plain_text}
    Text is cleaned of formula garbling."""
    xml = para._element.xml
    if 'm:oMath' not in xml and 'oMathPara' not in xml:
        return para.text, []

    math_list = []
    root = para._element

    # Extract m:oMathPara (display formulas) — these ARE the paragraph
    for omp in root.findall('{http://schemas.openxmlformats.org/officeDocument/2006/math}oMathPara'):
        raw = etree.tounicode(omp, with_tail=False)
        math_list.append({'type': 'display', 'xml': raw, 'text': _math_text(omp)})

    # Extract m:oMath (inline formulas) — embedded in runs
    for om in root.iter('{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath'):
        # Skip if already inside an oMathPara
        parent_tag = om.getparent().tag.split('}')[-1] if '}' in om.getparent().tag else om.getparent().tag
        if parent_tag == 'oMathPara':
            continue
        raw = etree.tounicode(om, with_tail=False)
        math_list.append({'type': 'inline', 'xml': raw, 'text': _math_text(om)})

    # Reconstruct text without formula garbling: get text from runs, skip math-only runs
    text_parts = []
    for child in root:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag == 'r':
            # Skip runs that contain only math (no w:t)
            has_text = child.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t') is not None
            if has_text:
                text_parts.append(child.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t').text or '')
        elif tag == 'oMathPara' or tag == 'oMath':
            # Don't add math XML to text
            pass
        elif tag == 'pPr':
            pass
        else:
            # Other elements (like w:r with math only) — skip
            pass

    text = ''.join(text_parts).strip()
    return text, math_list


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
