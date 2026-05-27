"""Plain-text formula conversion runtime template fragment for generated build scripts."""
from __future__ import annotations

FORMULA_TEXT_RUNTIME = r'''
def chapter_number_from_heading(text):
    t = str(text or '').strip()
    m = re.match(r'^第(\d+)章', t)
    if m:
        return int(m.group(1))
    cn = {'一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8, '九': 9, '十': 10}
    m = re.match(r'^第([一二三四五六七八九十])章', t)
    if m:
        return cn.get(m.group(1))
    m = re.match(r'^(\d+)(?:\.|\s)', t)
    return int(m.group(1)) if m else None


def latex_escape_text(text):
    return str(text or '').replace('\\', r'\backslash ').replace('{', r'\{').replace('}', r'\}')


def latex_text_arg(text):
    return str(text or '').replace('\\', r'\backslash ').replace('{', r'\{').replace('}', r'\}')


def split_formula_number(text):
    t = str(text or '').strip()
    label_pattern = r'(?:\d+(?:\s*[-.]\s*\d+)?|[A-Za-z]\s*(?:[-.]\s*)?\d+(?:\s*[-.]\s*\d+)?)'
    m = re.search(r'((?:\s*[\(\uff08]\s*' + label_pattern + r'\s*[\)\uff09])+\s*)$', t)
    if not m:
        return t, ''
    labels = re.findall(r'[\(\uff08]\s*(' + label_pattern + r')\s*[\)\uff09]', m.group(1))
    body = t[:m.start()].strip()
    if not body or not labels:
        return t, ''
    if len(labels) == 1 and re.search(r'[A-Za-z\u0370-\u03ff]$', body):
        return t, ''
    if len(labels) == 1 and not (re.search(r'[=＝≈≤≥<>]', body) and re.search(r'\d|[+*/×÷%·]', body)):
        return t, ''
    label = re.sub(r'\s+', '', labels[-1]).replace('.', '-') if labels else ''
    return body, label


def formula_token_to_latex(token):
    if not token:
        return ''
    if re.fullmatch(r'[\u4e00-\u9fff]+', token):
        return r'\text{' + latex_text_arg(token) + '}'
    if re.fullmatch(r'[A-Za-z]+', token):
        return r'\mathrm{' + token + '}'
    mapping = {
        '×': r'\times',
        '÷': r'\div',
        '≤': r'\leq',
        '≥': r'\geq',
        '≈': r'\approx',
        '≒': r'\approx',
        '％': r'\%',
        '%': r'\%',
        '²': '^{2}',
        '³': '^{3}',
        '（': '(',
        '）': ')',
        '，': ',',
        '。': '.',
        '：': '=',
        '＝': '=',
        '＋': '+',
        '－': '-',
    }
    return mapping.get(token, token)


def expression_to_latex(text):
    s = str(text or '').strip()
    if not s:
        return ''
    tokens = re.findall(r'[\u4e00-\u9fff]+|[A-Za-z]+|\d+(?:,\d{3})*(?:\.\d+)?|[²³]|.', s)
    out = []
    for tok in tokens:
        if tok.isspace():
            continue
        out.append(formula_token_to_latex(tok))
    return ''.join(out)


def formula_colon_split(text):
    t = str(text or '').strip()
    for sep in ('：', ':'):
        if sep in t:
            left, right = t.split(sep, 1)
            if left.strip() and right.strip() and re.search(r'\d|[=＝+\-*/×÷]', right):
                return left.strip(), right.strip()
    return '', t


def latex_delimited_formula(text):
    clean = clean_formula_text(text)
    body, existing_label = split_formula_number(clean)
    candidates = [body] if existing_label else []
    candidates.append(clean)
    for candidate in candidates:
        candidate = str(candidate or '').strip()
        if candidate.startswith('$$') and candidate.endswith('$$'):
            return candidate[2:-2].strip(), existing_label
        if candidate.startswith('$') and candidate.endswith('$'):
            return candidate[1:-1].strip(), existing_label
    return '', existing_label


def text_formula_to_latex(text):
    clean = clean_formula_text(text)
    latex, existing_label = latex_delimited_formula(clean)
    if latex:
        return latex, existing_label
    body, existing_label = split_formula_number(clean)
    if not body:
        return '', existing_label
    for src, dst in ((r'\uff1d', '='), (r'\uff0b', '+'), (r'\uff0d', '-')):
        body = body.replace(src, dst)
    return expression_to_latex(body), existing_label


def formula_latex_from_text(text):
    t = clean_formula_text(text)
    if not t:
        return ''
    latex, _existing_label = latex_delimited_formula(t)
    if latex:
        return latex
    latex, _existing_label = text_formula_to_latex(t)
    return latex


def formula_has_number(text):
    _body, label = split_formula_number(text)
    return bool(label)
'''
