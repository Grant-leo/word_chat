"""Formula label cleanup helpers for content parsing."""
from __future__ import annotations

import re
from lxml import etree

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
