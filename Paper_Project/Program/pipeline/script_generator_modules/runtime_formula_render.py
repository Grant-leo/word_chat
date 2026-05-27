"""Formula rendering runtime template fragment for generated build scripts."""
from __future__ import annotations

FORMULA_RENDER_RUNTIME = r'''
def next_formula_label(chapter):
    ch = chapter or 0
    FORMULA_COUNTERS[ch] = FORMULA_COUNTERS.get(ch, 0) + 1
    return f'{ch}-{FORMULA_COUNTERS[ch]}' if ch else str(FORMULA_COUNTERS[ch])


def render_plain_formula(text, chapter=None):
    text = clean_formula_text(text)
    if not text:
        return None
    rules = DATA.get('rules') or {}
    if rules.get('formula_numbered') and not formula_has_number(text):
        text = text + '(' + next_formula_label(chapter) + ')'
    p = doc.add_paragraph()
    apply_paragraph_profile(p, profile('formula'), first_indent_override=0)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.keep_together = True
    r = p.add_run(text)
    apply_run_profile(r, profile('formula'), text)
    return p


def render_formula(item, chapter=None):
    if isinstance(item, str):
        item = {'text': item}
    latex = str(item.get('latex') or '').strip()
    text = clean_formula_text(item.get('text') or '')
    xml = item.get('xml')
    math_entries = item.get('math') or []
    if len(math_entries) > 1 and not latex and not xml:
        last = None
        for m in math_entries:
            sub = dict(item)
            sub['math'] = [m]
            sub['text'] = clean_formula_text(m.get('text') or text)
            last = render_formula(sub, chapter)
        return last
    if math_entries and not latex and not xml:
        first_math = math_entries[0] or {}
        latex = str(first_math.get('latex') or '').strip()
        xml = first_math.get('xml')
        text = clean_formula_text(text or first_math.get('text') or '')
    existing_label = ''
    if not latex and not xml:
        latex, existing_label = text_formula_to_latex(text)
    elif text:
        _body, existing_label = split_formula_number(text)
    if not xml and not latex:
        return render_plain_formula(text, chapter)
    rules = DATA.get('rules') or {}
    numbered = item.get('numbered')
    should_number = bool(existing_label) or (bool(rules.get('formula_numbered')) if numbered is None else bool(numbered))
    if latex and should_number and r'\tag' not in latex and r'\begin{equation}' not in latex and r'\begin{align}' not in latex:
        label = existing_label or next_formula_label(chapter)
        latex = latex + r'\tag{' + label + '}'
    p = doc.add_paragraph()
    apply_paragraph_profile(p, profile('formula'), first_indent_override=0)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    try:
        xml_str = ensure_omml_wps_compat(xml or latex_to_omath(latex, display=True))
        p._element.append(etree.fromstring(xml_str.encode('utf-8') if isinstance(xml_str, str) else xml_str))
        BUILD_STATS['content_formulas_rendered'] = BUILD_STATS.get('content_formulas_rendered', 0) + 1
        BUILD_STATS['display_formulas_rendered'] = BUILD_STATS.get('display_formulas_rendered', 0) + 1
    except Exception:
        r = p.add_run(text or latex)
        apply_run_profile(r, profile('formula'), text or latex)
    return p
'''