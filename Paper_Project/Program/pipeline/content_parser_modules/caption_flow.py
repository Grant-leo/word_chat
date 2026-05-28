"""Caption detection and image-caption pairing for content_parser.py."""
import re


_REFERENTIAL_PROSE_START = (
    "展示",
    "显示",
    "给出",
    "给出了",
    "说明",
    "表明",
    "反映",
    "描述",
    "列出",
    "汇总",
    "呈现",
    "可见",
    "所示",
    "为",
    "是",
)


def _caption_tail(text, label_pattern):
    match = re.match(label_pattern, text)
    if not match:
        return ""
    return text[match.end():].strip(" \t:：.．、-—")


def _looks_like_referential_prose(tail):
    tail = str(tail or "").strip()
    return any(tail.startswith(word) for word in _REFERENTIAL_PROSE_START)


def is_figure_caption(text):
    text = str(text or '').strip()
    cn_tail = _caption_tail(text, r'^\u56fe\s*\d+(?:[.-]\d+)?')
    en_tail = _caption_tail(text, r'(?i)^(?:fig\.?|figure)\s*\d+(?:[.-]\d+)?')
    if _looks_like_referential_prose(cn_tail) or _looks_like_referential_prose(en_tail):
        return False
    return bool(
        re.match(r'^\u56fe\s*\d+(?:[.-]\d+)?\s*[^\d\s]', text)
        or re.match(r'(?i)^(?:fig\.?|figure)\s*\d+(?:[.-]\d+)?\s+[^\d\s]', text)
    )


def is_table_caption(text):
    text = str(text or '').strip()
    cn_tail = _caption_tail(text, r'^\u8868\s*\d+(?:[.-]\d+)?')
    en_tail = _caption_tail(text, r'(?i)^table\s*\d+(?:[.-]\d+)?')
    if _looks_like_referential_prose(cn_tail) or _looks_like_referential_prose(en_tail):
        return False
    return bool(
        re.match(r'^\u8868\s*\d+(?:[.-]\d+)?\s*[^\d\s]', text)
        or re.match(r'(?i)^table\s*\d+(?:[.-]\d+)?\s+[^\d\s]', text)
    )


def normalize_caption_spacing(text):
    text = str(text or '').strip()
    text = re.sub(r'(?i)^(fig\.?|figure)\s*(\d+(?:[.-]\d+)?)\s*', r'Fig. \2 ', text)
    text = re.sub(r'(?i)^table\s*(\d+(?:[.-]\d+)?)\s*', r'Table \1 ', text)
    return re.sub(r'^(\u56fe|\u8868)\s*(\d+(?:[.-]\d+)?)\s*', r'\1 \2 ', text).strip()


def _caption_kind(item):
    if isinstance(item, dict):
        role = item.get('role')
        if role in ('figure_caption', 'table_caption'):
            return role
        text = item.get('text') or ''
    else:
        text = str(item or '')
    if is_figure_caption(text):
        return 'figure_caption'
    if is_table_caption(text):
        return 'table_caption'
    return None


def _is_image_item(item):
    return isinstance(item, dict) and item.get('role') == 'image' and item.get('image')


def _caption_text(item):
    if isinstance(item, dict):
        return item.get('text') or ''
    return str(item or '')


def pair_figure_blocks(paragraphs):
    """Pair images with nearby figure captions while preserving all text."""
    out = []
    i = 0
    n = len(paragraphs or [])
    while i < n:
        item = paragraphs[i]
        if not _is_image_item(item):
            out.append(item)
            i += 1
            continue

        images = []
        while i < n and _is_image_item(paragraphs[i]):
            images.append(paragraphs[i])
            i += 1

        j = i
        held_text = []
        captions = []
        max_probe = min(n, i + max(8, len(images) * 3 + 3))
        while j < max_probe and len(captions) < len(images):
            nxt = paragraphs[j]
            if _is_image_item(nxt):
                break
            kind = _caption_kind(nxt)
            if kind == 'figure_caption':
                captions.append(nxt)
                j += 1
                continue
            if kind == 'table_caption':
                break
            held_text.append(nxt)
            j += 1

        if captions:
            for idx, img in enumerate(images):
                cap = captions[idx] if idx < len(captions) else None
                if cap is not None:
                    out.append({'role': 'figure', 'image': img.get('image'), 'caption': _caption_text(cap)})
                else:
                    out.append(img)
            if len(captions) > len(images):
                out.extend(captions[len(images):])
            out.extend(held_text)
            i = j
            continue

        out.extend(images)
    return out
