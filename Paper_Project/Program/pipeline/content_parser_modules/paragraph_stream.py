"""OOXML paragraph stream extraction helpers for content_parser.py."""
from lxml import etree

try:
    from formula_semantics import classify_formula_text, is_formula_problem_text
except ImportError:  # pragma: no cover - package-style imports
    from ..formula_semantics import classify_formula_text, is_formula_problem_text

try:
    from content_parser_modules.formula_extractor import (
        _formula_should_number,
        _math_text,
        _omml_text_looks_like_body,
        _strip_trailing_formula_labels,
        _strip_trailing_formula_labels_from_xml,
    )
    from content_parser_modules.image_extractor import images_from_run_ooxml
except ImportError:  # pragma: no cover - package-style imports
    from .formula_extractor import (
        _formula_should_number,
        _math_text,
        _omml_text_looks_like_body,
        _strip_trailing_formula_labels,
        _strip_trailing_formula_labels_from_xml,
    )
    from .image_extractor import images_from_run_ooxml


def local_name(element):
    return element.tag.split('}')[-1] if '}' in element.tag else element.tag


def run_text_preserve_breaks(run_elem):
    """Return visible text carried by a run, preserving explicit breaks."""
    parts = []
    for child in run_elem:
        name = local_name(child)
        if name == 't':
            parts.append(child.text or '')
        elif name in ('tab',):
            parts.append('\t')
        elif name in ('br', 'cr'):
            parts.append('\n')
    return ''.join(parts)


W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
_TRANSPARENT_CONTENT_CONTAINERS = {'customXml', 'smartTag'}
_ACCEPTED_REVISION_CONTAINERS = {'ins', 'moveTo'}
_DELETED_REVISION_CONTAINERS = {'del', 'moveFrom'}


def _sdt_content_children(sdt_elem):
    for part in list(sdt_elem):
        if local_name(part) == 'sdtContent':
            return list(part)
    return []


def visible_text_from_ooxml(element):
    """Return deterministic Word final-view text from a paragraph-like OOXML element."""
    parts = []

    def consume(node):
        name = local_name(node)
        if name in _DELETED_REVISION_CONTAINERS or name in ('delText', 'delInstrText'):
            return
        if name == 'txbxContent':
            return
        if name == 'sdt':
            for part in _sdt_content_children(node):
                consume(part)
            return
        if name == 't':
            parts.append(node.text or '')
            return
        if name == 'tab':
            parts.append('\t')
            return
        if name in ('br', 'cr'):
            parts.append('\n')
            return
        if name in ('instrText',):
            return
        for child in list(node):
            consume(child)

    consume(element)
    return ''.join(parts)


def paragraph_visible_text(para):
    return visible_text_from_ooxml(para._element)


def math_entry_from_ooxml(math_elem, math_type='inline'):
    raw = etree.tounicode(math_elem, with_tail=False)
    clean_xml, clean_text, had_label = _strip_trailing_formula_labels_from_xml(raw)
    if not clean_text:
        clean_text = _strip_trailing_formula_labels(_math_text(math_elem))
    entry = {'type': math_type, 'xml': clean_xml, 'text': clean_text}
    if clean_text:
        entry['formula_semantics'] = classify_formula_text(clean_text).to_dict()
    if had_label:
        entry['had_number_label'] = True
    return entry


def paragraph_stream_items(para, registry, notes=None):
    """Yield paragraph text/image/math items in true OOXML run order."""
    items = []
    buf = []
    seen_rids = set()
    notes = notes or {}

    def flush_text():
        text = ''.join(buf)
        buf.clear()
        if text.strip():
            items.append({'role': 'text', 'text': text})

    def append_math(math_elem, math_type='inline'):
        entry = math_entry_from_ooxml(math_elem, math_type)
        flush_text()
        semantic = classify_formula_text(entry.get('text') or '')
        if is_formula_problem_text(entry.get('text') or ''):
            items.append({
                'role': 'formula_problem',
                'problem': 'contaminated_formula_text',
                'source': 'omml',
                'text': entry.get('text') or '',
                'math': [entry],
                'formula_semantics': semantic.to_dict(),
            })
            return
        if _omml_text_looks_like_body(entry.get('text') or ''):
            items.append({'role': 'text', 'text': entry.get('text') or ''})
            return
        if math_type == 'display':
            items.append({
                'role': 'formula',
                'source': 'omml',
                'text': entry.get('text') or '',
                'math': [entry],
                'numbered': bool(entry.get('had_number_label')) or _formula_should_number(entry.get('text') or ''),
            })
        else:
            items.append({
                'role': 'math_inline',
                'source': 'omml',
                'text': entry.get('text') or '',
                'math': [entry],
            })

    def append_note_ref(note_elem, note_type):
        flush_text()
        note_id = str(note_elem.get(f'{{{W_NS}}}id') or '').strip()
        items.append({
            'role': 'note_ref',
            'note_type': note_type,
            'source_id': note_id,
            'text': (notes.get(note_type) or {}).get(note_id, ''),
        })

    def consume_run(run_elem):
        for part in run_elem:
            name = local_name(part)
            if name in ('drawing', 'pict'):
                flush_text()
                items.extend(images_from_run_ooxml(run_elem, para.part.rels, registry, seen_rids))
            elif name == 'oMath':
                append_math(part, 'inline')
            elif name in ('footnoteReference', 'endnoteReference'):
                append_note_ref(part, 'footnote' if name == 'footnoteReference' else 'endnote')
            elif name == 't':
                if part.text:
                    buf.append(part.text)
            elif name in ('tab',):
                buf.append('\t')
            elif name in ('br', 'cr'):
                buf.append('\n')
            elif name in _ACCEPTED_REVISION_CONTAINERS:
                consume_children(part)
            elif name in _DELETED_REVISION_CONTAINERS or name == 'delText':
                continue
            elif name in ('hyperlink', 'fldSimple'):
                consume_children(part)
            elif name == 'sdt':
                consume_inline_sdt(part)
            elif name in _TRANSPARENT_CONTENT_CONTAINERS:
                consume_children(part)
            elif name == 'oMathPara':
                append_math(part, 'display')

    def consume_inline_sdt(sdt_elem):
        for part in _sdt_content_children(sdt_elem):
            consume_child(part)

    def consume_children(parent_elem):
        for part in list(parent_elem):
            consume_child(part)

    def consume_child(child):
        name = local_name(child)
        if name == 'r':
            consume_run(child)
        elif name == 'sdt':
            consume_inline_sdt(child)
        elif name in ('hyperlink', 'fldSimple'):
            consume_children(child)
        elif name in _TRANSPARENT_CONTENT_CONTAINERS:
            consume_children(child)
        elif name in _ACCEPTED_REVISION_CONTAINERS:
            consume_children(child)
        elif name in _DELETED_REVISION_CONTAINERS:
            return
        elif name in ('oMath', 'oMathPara'):
            append_math(child, 'display' if name == 'oMathPara' else 'inline')
        elif name == 'p':
            consume_children(child)

    for child in para._element:
        consume_child(child)
    flush_text()
    return items


def append_stream_run_group(section, runs, append_text_or_code_func=None, in_appendix=False):
    if not runs:
        return
    text = ''.join(str(run.get('text') or '') for run in runs).strip()
    math_items = []
    note_items = []
    for run in runs:
        if run.get('type') == 'math':
            math_items.extend(run.get('math') or [])
        if run.get('type') == 'note_ref':
            note_items.append(run)
    if note_items:
        section['paragraphs'].append({
            'role': 'rich_text',
            'text': ''.join(str(run.get('text') or '') for run in runs if run.get('type') == 'text').strip(),
            'runs': runs,
            'notes': [
                {
                    'type': note.get('note_type') or 'footnote',
                    'source_id': note.get('source_id') or '',
                    'text': note.get('text') or '',
                }
                for note in note_items
            ],
        })
        return
    if not math_items:
        if append_text_or_code_func is not None:
            append_text_or_code_func(section, text, in_appendix=in_appendix)
        elif text:
            section.setdefault('paragraphs', []).append(text)
        return
    semantic = classify_formula_text(text)
    if is_formula_problem_text(text):
        section['paragraphs'].append({
            'role': 'formula_problem',
            'problem': 'contaminated_formula_text',
            'source': 'omml',
            'text': text,
            'math': math_items,
            'formula_semantics': semantic.to_dict(),
        })
        return
    non_math_text = ''.join(str(run.get('text') or '') for run in runs if run.get('type') != 'math').strip()
    if non_math_text:
        section['paragraphs'].append({
            'role': 'rich_text',
            'text': text,
            'runs': runs,
            'math': math_items,
        })
    else:
        section['paragraphs'].append({
            'role': 'formula',
            'source': 'omml',
            'text': text,
            'math': math_items,
            'numbered': any(item.get('had_number_label') for item in math_items) or _formula_should_number(text),
        })
