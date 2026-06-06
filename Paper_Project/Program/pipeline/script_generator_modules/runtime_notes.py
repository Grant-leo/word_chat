"""Native footnote/endnote runtime template fragment."""
from __future__ import annotations

NOTES_RUNTIME = r'''
NOTE_DEFS = {'footnote': {}, 'endnote': {}}
NOTE_KEY_TO_ID = {'footnote': {}, 'endnote': {}}
NOTE_REF_COUNTS = {'footnote': 0, 'endnote': 0}


def reset_note_state():
    global NOTE_DEFS, NOTE_KEY_TO_ID, NOTE_REF_COUNTS
    NOTE_DEFS = {'footnote': {}, 'endnote': {}}
    NOTE_KEY_TO_ID = {'footnote': {}, 'endnote': {}}
    NOTE_REF_COUNTS = {'footnote': 0, 'endnote': 0}


def _note_key(note_type, source_id, text):
    sid = str(source_id or '').strip()
    body = str(text or '').strip()
    return sid or body


def _register_note(note_type, source_id, text):
    note_type = 'endnote' if note_type == 'endnote' else 'footnote'
    key = _note_key(note_type, source_id, text)
    if key in NOTE_KEY_TO_ID[note_type]:
        return NOTE_KEY_TO_ID[note_type][key]
    new_id = len(NOTE_DEFS[note_type]) + 1
    NOTE_KEY_TO_ID[note_type][key] = new_id
    NOTE_DEFS[note_type][new_id] = str(text or '').strip()
    return new_id


def append_note_reference(p, run_info):
    note_type = 'endnote' if (run_info.get('note_type') == 'endnote') else 'footnote'
    text = str(run_info.get('text') or '').strip()
    if not text:
        return False
    note_id = _register_note(note_type, run_info.get('source_id') or '', text)
    r = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    rStyle = OxmlElement('w:rStyle')
    rStyle.set(qn('w:val'), 'EndnoteReference' if note_type == 'endnote' else 'FootnoteReference')
    rPr.append(rStyle)
    r.append(rPr)
    ref = OxmlElement('w:endnoteReference' if note_type == 'endnote' else 'w:footnoteReference')
    ref.set(qn('w:id'), str(note_id))
    r.append(ref)
    p._p.append(r)
    NOTE_REF_COUNTS[note_type] = NOTE_REF_COUNTS.get(note_type, 0) + 1
    return True


def _note_text_paragraph(note_kind, note_id, text):
    W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    XML_NS = 'http://www.w3.org/XML/1998/namespace'
    tag = 'endnote' if note_kind == 'endnote' else 'footnote'
    ref_tag = 'endnoteRef' if note_kind == 'endnote' else 'footnoteRef'
    text_style = 'EndnoteText' if note_kind == 'endnote' else 'FootnoteText'
    ref_style = 'EndnoteReference' if note_kind == 'endnote' else 'FootnoteReference'
    note = etree.Element('{%s}%s' % (W, tag))
    note.set('{%s}id' % W, str(note_id))
    p = etree.SubElement(note, '{%s}p' % W)
    pPr = etree.SubElement(p, '{%s}pPr' % W)
    pStyle = etree.SubElement(pPr, '{%s}pStyle' % W)
    pStyle.set('{%s}val' % W, text_style)
    r1 = etree.SubElement(p, '{%s}r' % W)
    rPr = etree.SubElement(r1, '{%s}rPr' % W)
    rStyle = etree.SubElement(rPr, '{%s}rStyle' % W)
    rStyle.set('{%s}val' % W, ref_style)
    etree.SubElement(r1, '{%s}%s' % (W, ref_tag))
    r2 = etree.SubElement(p, '{%s}r' % W)
    t2 = etree.SubElement(r2, '{%s}t' % W)
    t2.set('{%s}space' % XML_NS, 'preserve')
    t2.text = ' ' + str(text or '')
    return note


def _separator_note(note_kind, note_id, separator_tag, note_type):
    W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    tag = 'endnote' if note_kind == 'endnote' else 'footnote'
    note = etree.Element('{%s}%s' % (W, tag))
    note.set('{%s}type' % W, note_type)
    note.set('{%s}id' % W, str(note_id))
    p = etree.SubElement(note, '{%s}p' % W)
    r = etree.SubElement(p, '{%s}r' % W)
    etree.SubElement(r, '{%s}%s' % (W, separator_tag))
    return note


def _notes_root(note_kind, definitions):
    W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    root_tag = 'endnotes' if note_kind == 'endnote' else 'footnotes'
    root = etree.Element('{%s}%s' % (W, root_tag), nsmap={'w': W})
    root.append(_separator_note(note_kind, -1, 'separator', 'separator'))
    root.append(_separator_note(note_kind, 0, 'continuationSeparator', 'continuationSeparator'))
    for note_id in sorted(definitions):
        root.append(_note_text_paragraph(note_kind, note_id, definitions[note_id]))
    return root


def _ensure_relationship(rels_path, rel_type, target):
    R = 'http://schemas.openxmlformats.org/package/2006/relationships'
    if os.path.exists(rels_path):
        root = etree.parse(rels_path).getroot()
    else:
        root = etree.Element('Relationships', nsmap={None: R})
    for rel in root.findall('{%s}Relationship' % R):
        if rel.get('Target') == target or rel.get('Type') == rel_type:
            return root
    used = set()
    for rel in root.findall('{%s}Relationship' % R):
        rid = str(rel.get('Id') or '')
        m = re.match(r'rId(\d+)$', rid)
        if m:
            used.add(int(m.group(1)))
    next_id = 1
    while next_id in used:
        next_id += 1
    rel = etree.SubElement(root, '{%s}Relationship' % R)
    rel.set('Id', 'rId' + str(next_id))
    rel.set('Type', rel_type)
    rel.set('Target', target)
    return root


def _ensure_content_type(content_types_path, part_name, content_type):
    if not os.path.exists(content_types_path):
        return None
    root = etree.parse(content_types_path).getroot()
    ns = root.tag.split('}')[0].strip('{') if root.tag.startswith('{') else ''
    override_tag = '{%s}Override' % ns if ns else 'Override'
    for item in root.findall(override_tag):
        if item.get('PartName') == part_name:
            return root
    item = etree.SubElement(root, override_tag)
    item.set('PartName', part_name)
    item.set('ContentType', content_type)
    return root


def _write_xml(path, root):
    if root is None:
        return
    os.makedirs(os.path.dirname(path), exist_ok=True)
    etree.ElementTree(root).write(path, encoding='UTF-8', xml_declaration=True, standalone=True)


def _zip_dir(src_dir, dst_docx):
    tmp_docx = dst_docx + '.notes.tmp'
    with zipfile.ZipFile(tmp_docx, 'w', zipfile.ZIP_DEFLATED) as zf:
        for folder, _dirs, files in os.walk(src_dir):
            for name in files:
                full = os.path.join(folder, name)
                rel = os.path.relpath(full, src_dir).replace(os.sep, '/')
                zf.write(full, rel)
    os.replace(tmp_docx, dst_docx)


def apply_note_parts_to_docx(docx_path):
    if not NOTE_DEFS.get('footnote') and not NOTE_DEFS.get('endnote'):
        return
    tmp = tempfile.mkdtemp(prefix='wordchat_notes_')
    try:
        with zipfile.ZipFile(docx_path) as zf:
            zf.extractall(tmp)
        word_dir = os.path.join(tmp, 'word')
        rels_path = os.path.join(word_dir, '_rels', 'document.xml.rels')
        ct_path = os.path.join(tmp, '[Content_Types].xml')
        if NOTE_DEFS.get('footnote'):
            _write_xml(os.path.join(word_dir, 'footnotes.xml'), _notes_root('footnote', NOTE_DEFS['footnote']))
            rels = _ensure_relationship(
                rels_path,
                'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes',
                'footnotes.xml',
            )
            _write_xml(rels_path, rels)
            ct = _ensure_content_type(
                ct_path,
                '/word/footnotes.xml',
                'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
            )
            _write_xml(ct_path, ct)
        if NOTE_DEFS.get('endnote'):
            _write_xml(os.path.join(word_dir, 'endnotes.xml'), _notes_root('endnote', NOTE_DEFS['endnote']))
            rels = _ensure_relationship(
                rels_path,
                'http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes',
                'endnotes.xml',
            )
            _write_xml(rels_path, rels)
            ct = _ensure_content_type(
                ct_path,
                '/word/endnotes.xml',
                'application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml',
            )
            _write_xml(ct_path, ct)
        _zip_dir(tmp, docx_path)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
'''
