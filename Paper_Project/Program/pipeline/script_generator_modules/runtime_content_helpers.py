"""Content cleanup and classification runtime helpers for generated build scripts."""
from __future__ import annotations

CONTENT_HELPERS_RUNTIME = r'''
def is_front_section_index(i):
    return i in set(DATA.get('front_indices') or [])


def is_reference_heading(h):
    h = str(h or '').strip()
    return h.startswith('参考文献') or bool(re.match(r'(?i)^references?$', h))


def is_ack_heading(h):
    text = str(h or '').strip()
    return bool(re.search(r'致\s*谢', text) or re.match(r'(?i)^acknowledg(?:e)?ments?\b|^acknowledgment\b', text))


def is_appendix_heading(h):
    text = str(h or '').strip()
    return bool(re.search(r'附\s*录', text) or re.match(r'(?i)^append(?:ix|ices)\b', text))


def is_backmatter_heading(h):
    return is_ack_heading(h) or is_appendix_heading(h)


def is_caption_heading(h):
    return bool(re.match(r'^(图|表)\s*\d+(?:[.-]\d+)?\s*', str(h or '').strip()))


def normalize_caption(text):
    t = str(text or '').strip()
    space = (DATA.get('rules') or {}).get('caption_number_space')
    if space is True:
        t = re.sub(r'^(图|表)\s*(\d+(?:[.-]\d+)?)\s*', r'\1 \2 ', t)
    elif space is False:
        t = re.sub(r'^(图|表)\s*(\d+(?:[.-]\d+)?)\s*', r'\1\2 ', t)
    else:
        t = re.sub(r'^(图|表)\s*(\d+(?:[.-]\d+)?)\s*', r'\1 \2 ', t)
    return t.strip()


def clean_markdown_links(text):
    def repl(m):
        label = (m.group(1) or '').strip()
        target = (m.group(2) or '').strip()
        return label or target
    return re.sub(r'\[([^\]]+)\]\(([^)]+)\)', repl, str(text or ''))


def is_noise_text(text):
    return str(text or '').strip() in {'复制', 'Copy', 'Plain Text', '纯文本'}


def clean_text_artifacts(text, preserve_newlines=False):
    t = clean_markdown_links(text).replace('\u00a0', ' ')
    if preserve_newlines:
        lines = []
        for line in t.replace('\r\n', '\n').replace('\r', '\n').split('\n'):
            s = re.sub(r'[ \t]+', ' ', line).strip()
            if not is_noise_text(s):
                lines.append(s)
        return '\n'.join(lines).strip()
    t = re.sub(r'\s+', ' ', t).strip()
    return '' if is_noise_text(t) else t


def clean_code_text(text):
    return clean_text_artifacts(text, preserve_newlines=True)


def clean_formula_text(text):
    t = clean_text_artifacts(text)
    if t.count('|') >= 3:
        t = t.replace('|', '')
    t = re.sub(r'\s+', ' ', t).strip()
    return t


def add_caption(text, role='figure_caption'):
    caption_text = normalize_caption(text)
    if role == 'table_caption':
        m = re.match(r'^表\s*(\d+)(?:[-.](\d+))?', caption_text)
        if m:
            ch = int(m.group(1) or 0)
            no = int(m.group(2) or 0)
            if no:
                TABLE_COUNTERS[ch] = max(TABLE_COUNTERS.get(ch, 0), no)
    p = add_text(caption_text, role=role, first_indent=False)
    if p is None:
        return None
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.keep_together = True
    if role == 'table_caption':
        p.paragraph_format.keep_with_next = True
    return p
'''
