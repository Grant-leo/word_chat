"""Body rendering runtime template fragment for generated build scripts."""
from __future__ import annotations

BODY_RUNTIME = r'''
def paragraph_item_has_image(item):
    if not isinstance(item, dict):
        return False
    if item.get('role') in ('image', 'figure') or item.get('image'):
        return True
    for cell in item.get('table_cell_items') or []:
        for nested in cell.get('items') or []:
            if paragraph_item_has_image(nested):
                return True
    return False


def paragraph_item_has_display_math(item):
    if not isinstance(item, dict):
        return False

    def iter_math_entries(value):
        if isinstance(value, dict):
            return [value]
        if isinstance(value, list):
            return value
        return []

    def math_entry_is_display(entry):
        if not isinstance(entry, dict):
            return False
        kind = str(entry.get('type') or '').strip().lower()
        return kind in ('display', 'block') or bool(entry.get('display') or entry.get('block'))

    for math_entry in iter_math_entries(item.get('math')):
        if math_entry_is_display(math_entry):
            return True
    for run in item.get('runs') or []:
        if not isinstance(run, dict):
            continue
        for math_entry in iter_math_entries(run.get('math')):
            if math_entry_is_display(math_entry):
                return True
    for cell in item.get('table_cell_items') or []:
        for nested in cell.get('items') or []:
            if paragraph_item_has_display_math(nested):
                return True
    return False


def looks_like_list_bridge_text(text):
    t = clean_text_artifacts(text).strip()
    if not t:
        return False
    if re.match(r'^(?:[-*+]|\u2022|\u00b7|\u25e6|\u2023)\s+\S', t):
        return True
    if re.match(r'^(?:\(?\d{1,3}\)?[.)、]|[A-Za-z][.)])\s+\S', t):
        return True
    if re.match(r'^[（(]?\d{1,3}[）)][、.]?\s*\S', t):
        return True
    if re.match(r'^[一二三四五六七八九十]+[、.)）]\s*\S', t):
        return True
    if re.match(r'^[①②③④⑤⑥⑦⑧⑨⑩]\s*\S', t):
        return True
    return False


def render_paragraph_item(item, code_sensitive=False, chapter=None):
    if isinstance(item, dict) and item.get('role') == 'rich_text':
        add_rich_text_runs(item, role='body', first_indent=True)
        return
    if isinstance(item, dict) and item.get('role') == 'formula_problem':
        add_text(item.get('text') or '', role='body', first_indent=True)
        return
    if isinstance(item, dict) and (item.get('role') == 'formula' or item.get('latex') or item.get('xml')):
        render_formula(item, chapter)
        return
    if isinstance(item, dict) and item.get('math'):
        math_items = item.get('math') or []
        text = str(item.get('text') or '').strip()
        if any((m.get('type') == 'display') for m in math_items) and not text:
            render_formula(item, chapter)
        else:
            add_rich_text_runs(item, role='body', first_indent=True)
        return
    if isinstance(item, dict) and item.get('role') == 'figure':
        render_image(item.get('image') or item.get('filename') or item.get('asset') or '', item.get('caption') or '')
        return
    if isinstance(item, dict) and (item.get('role') == 'image' or item.get('image')):
        render_image(item.get('image') or item.get('filename') or item.get('asset') or '')
        return
    if isinstance(item, dict) and item.get('table_rows'):
        rows = item.get('table_rows') or []
        if item.get('role') == 'code' or rows_look_like_code(rows):
            add_code_block(item.get('code') or code_text_from_rows(rows))
        else:
            render_table_item(item)
        return
    if isinstance(item, dict) and (item.get('role') == 'figure_caption'):
        add_caption(item.get('text') or '', 'figure_caption')
        return
    if isinstance(item, dict) and (item.get('role') == 'table_caption'):
        add_caption(item.get('text') or '', 'table_caption')
        return
    if isinstance(item, dict) and (item.get('code') or item.get('role') == 'code'):
        add_code_block(item.get('code') or item.get('text') or '')
        return
    text = str(item.get('text') if isinstance(item, dict) else item or '').strip()
    if not text:
        return
    if len(text) > 20 and any(k in text[:80] for k in ['完成后删除', '格式要求', '字体要求', '页眉页脚']):
        return
    if code_sensitive and looks_like_code_line(text):
        add_code_block(text)
    elif is_figure_caption_text(text):
        add_caption(text, 'figure_caption')
    elif is_table_caption_text(text):
        add_caption(text, 'table_caption')
    else:
        add_text(text, role='body', first_indent=True)


def is_landscape_table_item(item):
    return is_table_item(item) and bool(table_render_section_page_setup(item))


def landscape_table_section_signature(item):
    setup = table_render_section_page_setup(item)
    if not setup:
        return ()
    margins = setup.get('margins_twips') or {}
    if not isinstance(margins, dict):
        margins = {}
    return (
        str(setup.get('orientation') or '').strip().lower(),
        safe_positive_int(setup.get('page_width_twips')),
        safe_positive_int(setup.get('page_height_twips')),
        safe_positive_int(margins.get('left')),
        safe_positive_int(margins.get('right')),
        safe_positive_int(margins.get('top')),
        safe_positive_int(margins.get('bottom')),
        safe_positive_int(margins.get('header')),
        safe_positive_int(margins.get('footer')),
        safe_positive_int(margins.get('gutter')),
    )


def landscape_table_section_compatible(first_item, next_item):
    first_sig = landscape_table_section_signature(first_item)
    next_sig = landscape_table_section_signature(next_item)
    return bool(first_sig and next_sig and first_sig == next_sig)


def table_group_at(paragraphs, idx):
    if idx >= len(paragraphs):
        return None
    para = paragraphs[idx]
    nxt = paragraphs[idx + 1] if idx + 1 < len(paragraphs) else None
    if isinstance(para, dict) and para.get('role') == 'table_caption' and is_landscape_table_item(nxt):
        return {'caption_item': para, 'table_item': nxt, 'next_idx': idx + 2}
    if isinstance(para, str) and is_landscape_table_item(nxt) and looks_like_table_title(para):
        return {'title_text': para, 'table_item': nxt, 'next_idx': idx + 2}
    if is_landscape_table_item(para):
        return {'table_item': para, 'next_idx': idx + 1}
    return None


def landscape_table_bridge_text(item):
    if not isinstance(item, str):
        if not isinstance(item, dict):
            return ''
        role = str(item.get('role') or '').strip()
        if role in ('table_caption', 'figure_caption', 'figure', 'image', 'code', 'formula', 'formula_problem'):
            return ''
        if paragraph_item_has_image(item) or paragraph_item_has_display_math(item) or is_table_item(item) or item.get('code'):
            return ''
        text = clean_text_artifacts(item.get('text') or '').strip()
    else:
        text = clean_text_artifacts(item).strip()
    if not text:
        return ''
    if looks_like_table_title(text) or is_figure_caption_text(text) or is_table_caption_text(text):
        return ''
    if looks_like_list_bridge_text(text):
        return ''
    if looks_like_code_line(text):
        return ''
    return text


def landscape_table_bridge_run_at(paragraphs, idx, anchor_table_item=None):
    bridge_items = []
    total_len = 0
    pos = idx
    while pos < len(paragraphs) and len(bridge_items) < 3:
        if table_group_at(paragraphs, pos):
            return None
        bridge = paragraphs[pos]
        text = landscape_table_bridge_text(bridge)
        if not text or len(text) > 220:
            return None
        total_len += len(text)
        if total_len > 360:
            return None
        bridge_items.append(bridge)
        pos += 1
        next_group = table_group_at(paragraphs, pos)
        if next_group:
            if anchor_table_item is not None and not landscape_table_section_compatible(
                anchor_table_item, next_group.get('table_item') or {}
            ):
                return None
            return {'items': bridge_items, 'next_idx': pos}
    return None


def render_landscape_table_group(paragraphs, idx, current_chapter):
    first_group = table_group_at(paragraphs, idx)
    if not first_group:
        return idx
    started = begin_table_source_section(first_group.get('table_item'))
    try:
        while True:
            group = table_group_at(paragraphs, idx)
            if not group:
                break
            caption_item = group.get('caption_item')
            title_text = group.get('title_text')
            if isinstance(caption_item, dict):
                add_caption(caption_item.get('text') or '', 'table_caption')
            elif title_text:
                add_caption(next_table_caption(title_text, current_chapter), 'table_caption')
            render_table_from_item(group.get('table_item') or {})
            idx = group.get('next_idx') or (idx + 1)
            bridge_run = landscape_table_bridge_run_at(
                paragraphs,
                idx,
                anchor_table_item=group.get('table_item') or {},
            )
            if bridge_run:
                for bridge in bridge_run.get('items') or []:
                    render_paragraph_item(bridge, code_sensitive=False, chapter=current_chapter)
                idx = bridge_run.get('next_idx') or idx
                continue
            break
    finally:
        end_table_source_section(started)
    return idx


def render_body():
    add_section_with_header('decimal', 1)
    fig_no = 0
    current_chapter = None
    rendered_body_sections = 0
    for i, sec in enumerate(DATA.get('sections') or []):
        if is_front_section_index(i):
            continue
        h = (sec.get('heading') or '').strip()
        role = sec.get('role') or ''
        if is_reference_heading(h) or is_backmatter_heading(h) or role in ('references', 'acknowledgement', 'appendix'):
            continue
        if sec.get('page_break_before') and rendered_body_sections > 0:
            doc.add_page_break()
        rendered_body_sections += 1
        if is_caption_heading(h):
            add_caption(h, 'figure_caption' if is_figure_caption_text(h) else 'table_caption')
        elif h and h != '正文':
            add_heading(h, sec.get('level') or 1)
            if int(sec.get('level') or 1) == 1:
                current_chapter = chapter_number_from_heading(h) or current_chapter
        paragraphs = sec.get('paragraphs', []) or []
        has_inline_images = any(paragraph_item_has_image(x) for x in paragraphs)
        # New content_parser keeps images in the paragraph stream.  For old
        # content.json files, fall back to section-level images, but do not
        # invent a caption from the heading because that caused figure-title
        # mismatch.
        if not has_inline_images:
            for img in sec.get('images', []) or []:
                render_image(img, '')
        idx = 0
        while idx < len(paragraphs):
            para = paragraphs[idx]
            nxt = paragraphs[idx + 1] if idx + 1 < len(paragraphs) else None
            if table_group_at(paragraphs, idx):
                idx = render_landscape_table_group(paragraphs, idx, current_chapter)
                continue
            if isinstance(para, dict) and para.get('role') == 'table_caption' and is_table_item(nxt):
                started = begin_table_source_section(nxt)
                add_caption(para.get('text') or '', 'table_caption')
                render_table_from_item(nxt)
                end_table_source_section(started)
                idx += 2
                continue
            if isinstance(para, str) and is_table_item(nxt) and looks_like_table_title(para):
                started = begin_table_source_section(nxt)
                add_caption(next_table_caption(para, current_chapter), 'table_caption')
                render_table_from_item(nxt)
                end_table_source_section(started)
                idx += 2
                continue
            if is_table_item(para):
                prev = paragraphs[idx - 1] if idx > 0 else None
                has_caption = isinstance(prev, dict) and prev.get('role') == 'table_caption'
                if not has_caption and idx == 0 and h and looks_like_table_title(h):
                    add_caption(next_table_caption(h, current_chapter), 'table_caption')
            render_paragraph_item(para, code_sensitive=False, chapter=current_chapter)
            idx += 1
    pure_refs, ack_from_refs, app_from_refs = split_refs_backmatter(DATA.get('references') or [])
    ack_sections, app_sections = collect_structural_backmatter()
    render_reference_entries(pure_refs)
    render_backmatter_section('\u81f4  \u8c22', ack_sections or ack_from_refs, code_sensitive=False)
    render_backmatter_section('\u9644  \u5f55', app_sections or app_from_refs, code_sensitive=True)
'''
