"""Reference and backmatter runtime template fragment for generated build scripts."""
from __future__ import annotations

REFERENCES_RUNTIME = r'''
def clean_ref_text(ref):
    text = re.sub(r'\s+', ' ', str(ref or '')).strip()
    if text.startswith('[') and ']' in text:
        prefix, rest = text.split(']', 1)
        if prefix[1:].isdigit():
            return rest.strip()
    parts = text.split(None, 1)
    if len(parts) == 2 and parts[0].strip('.、[]').isdigit():
        return parts[1].strip()
    return text


def split_refs_backmatter(refs):
    pure, ack, app = [], [], []
    mode = 'refs'
    for raw in refs or []:
        if isinstance(raw, dict):
            if mode == 'ack':
                ack.append(raw)
            elif mode == 'app':
                app.append(raw)
            continue
        text = str(raw or '').strip()
        if not text:
            continue
        if is_ack_heading(text):
            mode = 'ack'; continue
        if is_appendix_heading(text):
            mode = 'app'
            if normalize_label(text) != '附录':
                app.append(text)
            continue
        (pure if mode == 'refs' else ack if mode == 'ack' else app).append(text)
    return pure, ack, app


def add_reference_mixed_runs(p, text, prof):
    # Chinese parts use the role's CJK font; Latin/numeric punctuation uses Times New Roman.
    for seg in re.findall(r'[\u4e00-\u9fff]+|[^\u4e00-\u9fff]+', text):
        r = p.add_run(seg)
        if has_cjk(seg):
            apply_run_profile(r, prof, seg, force_latin='Times New Roman')
        else:
            p_latin = dict(prof); p_latin['font'] = 'Times New Roman'
            apply_run_profile(r, p_latin, seg, force_latin='Times New Roman')


def apply_reference_indent(p, prof, rule_hanging_cm=None):
    left = prof.get('left_indent_cm')
    hanging = prof.get('hanging_indent_cm')
    if hanging is None:
        hanging = rule_hanging_cm
    try:
        if left is not None:
            p.paragraph_format.left_indent = Cm(float(left or 0))
        elif hanging:
            p.paragraph_format.left_indent = Cm(float(hanging))
    except Exception:
        pass
    try:
        if hanging:
            p.paragraph_format.first_line_indent = Cm(-float(hanging))
    except Exception:
        pass


def render_reference_entries(refs):
    if not refs:
        return
    doc.add_page_break()
    add_text('参考文献', role='reference_heading', first_indent=False, outline_level=0)
    base_prof = profile('reference')
    hang_chars = (DATA.get('rules') or {}).get('reference_hanging_chars')
    try:
        hang_cm = float(hang_chars) * float(base_prof.get('size') or 12) * 0.0352778 if hang_chars else 0.0
    except Exception:
        hang_cm = 0.0
    for idx, raw in enumerate(refs, 1):
        if isinstance(raw, dict):
            continue
        text = clean_ref_text(raw)
        if not text:
            continue
        prof = profile('reference_english') if (DATA.get('rules') or {}).get('reference_english_left') and ascii_ratio(text[:120]) > 0.55 else base_prof
        p = doc.add_paragraph()
        apply_paragraph_profile(p, prof)
        apply_reference_indent(p, prof, hang_cm)
        p.paragraph_format.keep_together = True
        add_reference_mixed_runs(p, '[' + str(idx) + '] ' + text, prof)


def render_backmatter_section(title, paragraphs, code_sensitive=False):
    if not paragraphs:
        return
    doc.add_page_break()
    add_heading(title, 1)
    for item in paragraphs:
        if isinstance(item, dict):
            render_paragraph_item(item, code_sensitive=code_sensitive, chapter=None)
            continue
        text = str(item or '').strip()
        if not text:
            continue
        if code_sensitive and looks_like_code_line(text):
            add_code_block(text)
        else:
            add_text(text, role='body', first_indent=not code_sensitive)


def collect_structural_backmatter():
    ack, app = [], []
    mode = None
    for i, sec in enumerate(DATA.get('sections') or []):
        if is_front_section_index(i):
            continue
        h = (sec.get('heading') or '').strip()
        role = sec.get('role') or ''
        if role == 'acknowledgement' or is_ack_heading(h):
            mode = 'ack'
            for para in sec.get('paragraphs', []) or []:
                ack.append(para)
            continue
        if role == 'appendix' or is_appendix_heading(h):
            mode = 'app'
            if normalize_label(h) != '附录':
                app.append(h)
            for para in sec.get('paragraphs', []) or []:
                app.append(para)
            continue
        if mode in ('ack', 'app'):
            if sec.get('level') and h:
                (ack if mode == 'ack' else app).append(h)
            for para in sec.get('paragraphs', []) or []:
                (ack if mode == 'ack' else app).append(para)
    return ack, app


'''
