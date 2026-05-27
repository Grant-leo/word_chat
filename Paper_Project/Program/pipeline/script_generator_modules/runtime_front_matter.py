"""Front matter rendering runtime template fragment for generated build scripts."""
from __future__ import annotations

FRONT_MATTER_RUNTIME = r'''
def section_text(sec):
    out = []
    for para in sec.get('paragraphs', []) or []:
        if isinstance(para, str):
            out.append(para.strip())
        elif isinstance(para, dict) and para.get('text'):
            out.append(str(para.get('text')).strip())
    return '\n'.join(x for x in out if x)


def add_keywords(label, value, role):
    prof = profile(role)
    p = doc.add_paragraph()
    apply_paragraph_profile(p, prof, first_indent_override=0)
    r1 = p.add_run(label)
    p_bold = dict(prof); p_bold['bold'] = True
    apply_run_profile(r1, p_bold, label)
    r2 = p.add_run(value)
    p_norm = dict(prof); p_norm['bold'] = False
    apply_run_profile(r2, p_norm, value)
    return p


def add_blank_line(role='body'):
    return add_text('', role=role, first_indent=False)


def render_front_matter():
    front = DATA.get('front') or {}
    title_cn = DATA.get('title_cn') or ''
    if title_cn:
        add_text(title_cn, role='cn_title', first_indent=False)
    cn_abs = front.get('cn_abs')
    if cn_abs:
        add_text('摘 要', role='cn_abstract_heading', first_indent=False)
        cn_items = []
        for para in cn_abs.get('paragraphs', []) or []:
            if isinstance(para, dict) and (para.get('math') or para.get('role') == 'formula' or para.get('latex') or para.get('xml')):
                cn_items.append(para)
            else:
                text = para if isinstance(para, str) else para.get('text', '')
                if str(text).strip():
                    cn_items.append(str(text).strip())
        if DATA.get('rules', {}).get('cn_abstract_single_paragraph') and cn_items:
            plain_items = [x for x in cn_items if isinstance(x, str)]
            rich_items = [x for x in cn_items if not isinstance(x, str)]
            if plain_items:
                add_text(''.join(plain_items), role='cn_abstract_body', first_indent=True)
            for item in rich_items:
                add_rich_text_item(item, role='cn_abstract_body', first_indent=True)
        else:
            for item in cn_items:
                add_rich_text_item(item, role='cn_abstract_body', first_indent=True)
    cn_kw = front.get('cn_kw')
    if cn_kw and cn_abs:
        add_blank_line('cn_abstract_body')
    if cn_kw:
        val = section_text(cn_kw)
        if val:
            add_keywords('关键词：', val, 'cn_keywords')
    has_en = bool(front.get('en_title') or front.get('en_abs') or front.get('en_kw'))
    if has_en:
        doc.add_page_break()
    en_title = front.get('en_title') or ''
    if en_title and DATA.get('rules', {}).get('en_title_upper'):
        en_title = en_title.upper()
    if en_title:
        add_text(en_title, role='en_title', first_indent=False)
    en_abs = front.get('en_abs')
    if en_abs:
        add_text('ABSTRACT', role='en_abstract_heading', first_indent=False)
        for para in en_abs.get('paragraphs', []) or []:
            add_rich_text_item(para, role='en_abstract_body', first_indent=True)
    en_kw = front.get('en_kw')
    if en_kw:
        val = section_text(en_kw).replace('；', ';')
        if val:
            add_keywords('KEY WORDS: ', val, 'en_keywords')
    add_section_with_header('upperRoman', None)
    add_toc()
'''