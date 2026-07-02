"""Formula and rich-text runtime template fragment for generated build scripts."""
from __future__ import annotations

FORMULA_RUNTIME = r'''
def _math_entry_list(value):
    if isinstance(value, dict):
        return [value]
    if isinstance(value, list):
        return [entry for entry in value if isinstance(entry, dict)]
    return []


def _math_entries_from_item(item):
    if not isinstance(item, dict):
        return []
    if item.get('math'):
        return _math_entry_list(item.get('math'))
    if item.get('latex') or item.get('xml'):
        return [item]
    return []


def ensure_omml_wps_compat(xml_str):
    """Ensure imported/source OMML has m:rPr for WPS compatibility."""
    try:
        root = etree.fromstring(xml_str.encode('utf-8') if isinstance(xml_str, str) else xml_str)
        m_ns = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
        for mr in root.iter('{%s}r' % m_ns):
            if mr.find('{%s}rPr' % m_ns) is None:
                mr.insert(0, etree.Element('{%s}rPr' % m_ns))
        return etree.tostring(root, encoding='unicode')
    except Exception:
        return xml_str


def append_inline_formula(p, item):
    if isinstance(item, str):
        item = {'text': item}
    latex = str(item.get('latex') or '').strip()
    text = clean_formula_text(item.get('text') or '')
    xml = item.get('xml')
    if not xml and not latex:
        latex, _existing_label = text_formula_to_latex(text)
    if not xml and not latex:
        r = p.add_run(text)
        apply_run_profile(r, profile('body'), text)
        return False
    try:
        xml_str = ensure_omml_wps_compat(xml or latex_to_omath(latex, display=False))
        elem = etree.fromstring(xml_str.encode('utf-8') if isinstance(xml_str, str) else xml_str)
        if elem.tag.endswith('}oMathPara') or elem.tag.endswith('oMathPara'):
            for child in list(elem):
                p._element.append(child)
        else:
            p._element.append(elem)
        BUILD_STATS['content_formulas_rendered'] = BUILD_STATS.get('content_formulas_rendered', 0) + 1
        BUILD_STATS['inline_formulas_rendered'] = BUILD_STATS.get('inline_formulas_rendered', 0) + 1
        return True
    except Exception:
        r = p.add_run(text or latex)
        apply_run_profile(r, profile('body'), text or latex)
        return False


def _rich_text_image_items(run):
    if not isinstance(run, dict):
        return []
    items = []
    for item in run.get('items') or []:
        if isinstance(item, dict):
            items.extend(_rich_text_image_items(item))
    kind = str(run.get('type') or '').strip()
    role = str(run.get('role') or '').strip()
    if (kind in ('image', 'figure') or role in ('image', 'figure') or run.get('image') or run.get('filename') or run.get('asset')):
        filename = run.get('image') or run.get('filename') or run.get('asset')
        if filename:
            items.append(run)
    return items


def _rich_text_direct_image_item(run):
    if not isinstance(run, dict):
        return False
    kind = str(run.get('type') or '').strip()
    role = str(run.get('role') or '').strip()
    return kind in ('image', 'figure') or role in ('image', 'figure') or bool(run.get('image') or run.get('filename') or run.get('asset'))


def append_inline_image_run(p, run, rendered_image_keys=None):
    wrote = False
    for image_item in _rich_text_image_items(run):
        filename = image_item.get('image') or image_item.get('filename') or image_item.get('asset') or ''
        if rendered_image_keys is not None and filename in rendered_image_keys:
            continue
        path = content_image_path(filename)
        if not path:
            continue
        try:
            r = p.add_run()
            width, height = fit_picture_dimensions(path, has_caption=False)
            r.add_picture(path, width=width, height=height)
            BUILD_STATS['content_images_rendered'] = BUILD_STATS.get('content_images_rendered', 0) + 1
            if rendered_image_keys is not None:
                rendered_image_keys.add(filename)
            wrote = True
        except Exception:
            continue
    return wrote


def _rich_text_item_math_entries(item):
    if not isinstance(item, dict):
        return []
    entries = []
    for nested in item.get('items') or []:
        if isinstance(nested, dict):
            entries.extend(_math_entries_from_item(nested))
            entries.extend(_rich_text_item_math_entries(nested))
    return entries


def add_rich_text_runs(item, role='body', first_indent=True, render_item_media=True):
    prof = profile(role)
    runs = item.get('runs') or []
    has_item_images = render_item_media and bool(_rich_text_image_items(item))
    item_math_entries = _rich_text_item_math_entries(item) if render_item_media else []
    if not runs:
        text = str(item.get('text') or '').strip()
        if text:
            runs = [{'type': 'text', 'text': text}]
        for m in _math_entries_from_item(item):
            runs.append({'type': 'math', 'text': m.get('text') or '', 'math': [m]})
    if not runs and not has_item_images and not item_math_entries:
        return None
    p = doc.add_paragraph()
    apply_paragraph_profile(p, prof, first_indent_override=(prof.get('first_indent_cm') if first_indent else 0))
    superscript = role == 'body'
    wrote = False
    rendered_image_keys = set()
    for run in runs:
        kind = run.get('type') or ('math' if run.get('math') else 'text')
        if kind == 'math':
            for m in _math_entries_from_item(run):
                wrote = append_inline_formula(p, m) or wrote
        elif _rich_text_direct_image_item(run) or (render_item_media and _rich_text_image_items(run)):
            image_run = run
            if _rich_text_direct_image_item(run) and not render_item_media:
                image_run = dict(run)
                image_run['items'] = []
            wrote = append_inline_image_run(p, image_run, rendered_image_keys) or wrote
        elif kind == 'note_ref':
            wrote = append_note_reference(p, run) or wrote
        else:
            text = str(run.get('text') or '')
            if text:
                add_text_runs(p, text, prof, superscript)
                wrote = True
    if has_item_images:
        wrote = append_inline_image_run(p, item, rendered_image_keys) or wrote
    for m in item_math_entries:
        wrote = append_inline_formula(p, m) or wrote
    if not wrote:
        try:
            p._element.getparent().remove(p._element)
        except Exception:
            pass
        return None
    return p


def add_rich_text_item(item, role='body', first_indent=True, chapter=None):
    """Render a content item that may contain plain text plus extracted math.

    `rich_text` items carry an ordered token stream, so inline formulas remain
    inside the current paragraph. Older content JSON files may only carry
    {"text": "...", "math": [...]}; those are still rendered as editable inline
    formulas after the text rather than being dropped.
    """
    if isinstance(item, str):
        return add_text(item, role=role, first_indent=first_indent)
    if not isinstance(item, dict):
        return None
    if item.get('role') == 'rich_text':
        return add_rich_text_runs(item, role=role, first_indent=first_indent)
    text = str(item.get('text') or '').strip()
    if item.get('role') == 'formula' or item.get('latex') or item.get('xml'):
        return render_formula(item, chapter)
    if item.get('math') and not item.get('latex') and not item.get('xml'):
        return add_rich_text_runs(item, role=role, first_indent=first_indent)
    if text:
        add_text(text, role=role, first_indent=first_indent)
    return None


'''
