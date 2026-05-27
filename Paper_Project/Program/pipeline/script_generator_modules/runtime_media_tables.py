"""Table, image, and code-block runtime template fragment for generated build scripts."""
from __future__ import annotations

MEDIA_TABLE_RUNTIME = r'''
def strip_heading_number(text):
    t = str(text or '').strip()
    t = re.sub(r'^第[一二三四五六七八九十百千万\d]+章\s*', '', t)
    t = re.sub(r'^\d+(?:\.\d+)*\s*', '', t)
    return t.strip()


def is_table_item(item):
    return isinstance(item, dict) and item.get('table_rows') and item.get('role') != 'code'


def is_code_table_item(item):
    return isinstance(item, dict) and item.get('table_rows') and (item.get('role') == 'code' or rows_look_like_code(item.get('table_rows') or []))


def looks_like_table_title(text):
    t = strip_heading_number(clean_text_artifacts(text))
    if not t or len(t) > 50:
        return False
    if re.match(r'^(图|表)\s*\d+', t) or re.match(r'^代码\s*\d+', t):
        return False
    if re.search(r'[。！？；;=<>]|如下|所示', t):
        return False
    return True


def next_table_caption(title, chapter=None):
    title = strip_heading_number(title)
    ch = chapter or 0
    TABLE_COUNTERS[ch] = TABLE_COUNTERS.get(ch, 0) + 1
    label = f'{ch}-{TABLE_COUNTERS[ch]}' if ch else str(TABLE_COUNTERS[ch])
    return f'表 {label} {title}'.strip()



def looks_like_code_line(text):
    t = str(text or '').strip()
    if not t or len(t) > 240:
        return False
    if re.match(r'^[A-Za-z0-9_.-]+[>#]', t):
        return True
    if re.match(r'^(interface|vlan|ip route|ip address|router|switchport|acl|rule|nat|dhcp|dns|ospf|bgp|display|show|ping|tracert|undo|quit|return|sysname|description|gateway|firewall|security-policy)\b', t, re.I):
        return True
    if re.match(r'^[a-z][a-z0-9_-]+\s+[-A-Za-z0-9_/.:]+', t) and any(ch in t for ch in ['/', '.', '-', '_']):
        return True
    return False




def rows_look_like_code(rows):
    flat = []
    for row in rows or []:
        for cell in row or []:
            for line in str(cell or '').splitlines():
                if line.strip():
                    flat.append(line.strip())
    if not flat:
        return False
    if len(rows or []) >= 4 and max((len(r) for r in rows or []), default=0) <= 2:
        code_hits = sum(1 for x in flat if looks_like_code_line(x))
        return code_hits >= max(2, len(flat) // 3)
    if max((len(r) for r in rows or []), default=0) == 1 and len(flat) >= 2:
        return sum(1 for x in flat if looks_like_code_line(x)) >= 2
    return False


def code_text_from_rows(rows):
    lines = []
    for row in rows or []:
        cells = [str(c or '').rstrip() for c in row]
        if len(cells) == 1:
            lines.append(cells[0])
        else:
            lines.append('    '.join(cells).rstrip())
    return clean_code_text('\n'.join(lines).rstrip())

def set_cell_margins(cell, top=80, start=120, bottom=80, end=120):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = tcPr.find(qn('w:tcMar'))
    if tcMar is not None:
        tcPr.remove(tcMar)
    tcMar = OxmlElement('w:tcMar')
    for side, val in [('top', top), ('start', start), ('bottom', bottom), ('end', end)]:
        el = OxmlElement('w:' + side)
        el.set(qn('w:w'), str(int(val)))
        el.set(qn('w:type'), 'dxa')
        tcMar.append(el)
    tcPr.append(tcMar)


def add_code_block(text):
    """Render code/configuration as a bordered block.

    A one-cell table is used instead of a normal paragraph so the output has a
    real solid frame, which is the conventional way to present command/config
    blocks in thesis appendices and network-design papers.  The detection of
    code remains semantic; no vendor, school, or fixed heading text is used.
    """
    text = clean_code_text(text)
    if not text:
        return None
    prof = profile('code')
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    try:
        _w = text_width_cm()
        table.columns[0].width = Cm(_w)
        table.rows[0].cells[0].width = Cm(_w)
    except Exception:
        pass
    cell = table.rows[0].cells[0]
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    set_cell_borders(cell,
                     top={'val': 'single', 'sz': '8', 'color': '000000'},
                     left={'val': 'single', 'sz': '8', 'color': '000000'},
                     bottom={'val': 'single', 'sz': '8', 'color': '000000'},
                     right={'val': 'single', 'sz': '8', 'color': '000000'})
    set_cell_margins(cell, top=80, start=120, bottom=80, end=120)
    cell.text = ''
    lines = text.splitlines() or ['']
    for i, line in enumerate(lines):
        p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
        apply_paragraph_profile(p, prof, first_indent_override=0)
        p.paragraph_format.left_indent = Cm(0)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(line)
        apply_run_profile(r, prof, line)
    # Add a tiny spacing paragraph after the code box so following text does not
    # visually touch the frame; it has no text and therefore cannot enter TOC.
    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_after = Pt(3)
    spacer.paragraph_format.line_spacing = 1
    return table


def apply_three_line_borders(table):
    rows = len(table.rows)
    if rows == 0:
        return
    for ri, row in enumerate(table.rows):
        for cell in row.cells:
            sides = {'top': 'nil', 'left': 'nil', 'bottom': 'nil', 'right': 'nil', 'insideH': 'nil', 'insideV': 'nil'}
            if ri == 0:
                sides['top'] = {'val': 'single', 'sz': '12', 'color': '000000'}
                sides['bottom'] = {'val': 'single', 'sz': '8', 'color': '000000'}
            if ri == rows - 1:
                sides['bottom'] = {'val': 'single', 'sz': '12', 'color': '000000'}
            set_cell_borders(cell, **sides)


def repeat_table_header(row):
    try:
        trPr = row._tr.get_or_add_trPr()
        tbl_header = trPr.find(qn('w:tblHeader'))
        if tbl_header is None:
            tbl_header = OxmlElement('w:tblHeader')
            trPr.append(tbl_header)
        tbl_header.set(qn('w:val'), 'true')
    except Exception:
        pass


def prevent_row_split(row):
    try:
        trPr = row._tr.get_or_add_trPr()
        cant = trPr.find(qn('w:cantSplit'))
        if cant is None:
            cant = OxmlElement('w:cantSplit')
            trPr.append(cant)
    except Exception:
        pass


def should_keep_table_together(rows):
    if not rows or len(rows) > 10:
        return False
    text_cells = [str(cell or '') for row in rows for cell in row]
    if any(len(cell) > 90 for cell in text_cells):
        return False
    estimated_lines = 0
    for row in rows:
        row_lines = 1
        for cell in row:
            parts = str(cell or '').split('\n') or ['']
            cell_lines = sum(max(1, math.ceil(len(part) / 14)) for part in parts)
            row_lines = max(row_lines, cell_lines)
        estimated_lines += row_lines
    return estimated_lines <= 18


def keep_table_together(table):
    for ri, row in enumerate(table.rows):
        for cell in row.cells:
            for p in cell.paragraphs:
                p.paragraph_format.keep_together = True
                if ri < len(table.rows) - 1:
                    p.paragraph_format.keep_with_next = True


def render_table(rows):
    if not rows:
        return
    ncols = max(len(r) for r in rows)
    table = doc.add_table(rows=len(rows), cols=ncols)
    BUILD_STATS['content_tables_rendered'] = BUILD_STATS.get('content_tables_rendered', 0) + 1
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for ri, row in enumerate(rows):
        if ri == 0:
            repeat_table_header(table.rows[ri])
        prevent_row_split(table.rows[ri])
        prof = profile('table_header' if ri == 0 else 'table_body')
        for ci in range(ncols):
            text = row[ci] if ci < len(row) else ''
            cell = table.rows[ri].cells[ci]
            cell.text = ''
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            parts = str(text or '').split('\n') or ['']
            for pi, part in enumerate(parts):
                p = cell.paragraphs[0] if pi == 0 else cell.add_paragraph()
                apply_paragraph_profile(p, prof, first_indent_override=0)
                r = p.add_run(part)
                apply_run_profile(r, prof, part)
    apply_three_line_borders(table)
    if should_keep_table_together(rows):
        keep_table_together(table)
    return table


def keep_paragraph_with_previous(p):
    # Word has no high-level python-docx flag for keep-with-previous.  We use
    # keepNext on the preceding image paragraph and keepLines on the caption so
    # a figure title is less likely to drift to the next page alone.
    try:
        pPr = p._element.get_or_add_pPr()
        keep = pPr.find(qn('w:keepLines'))
        if keep is None:
            keep = OxmlElement('w:keepLines')
            pPr.append(keep)
    except Exception:
        pass


def render_image(filename, caption=''):
    img_dir = DATA.get('images_dir') or ''
    candidates = []
    if os.path.isabs(img_dir):
        candidates.append(os.path.join(img_dir, filename))
    candidates += [
        os.path.join(BASE, img_dir, filename),
        os.path.abspath(os.path.join(os.getcwd(), img_dir, filename)),
        os.path.abspath(os.path.join(BASE, '..', img_dir, filename)),
        os.path.join(BASE, 'figures', filename),
    ]
    path = next((p for p in candidates if p and os.path.exists(p)), None)
    if not path:
        return None
    p = doc.add_paragraph()
    configure_picture_paragraph(p, keep_with_next=bool(caption))
    r = p.add_run()
    try:
        is_contained_fragment = picture_looks_like_low_res_text_fragment(path, caption)
        width, height = fit_picture_dimensions(path, has_caption=bool(caption))
        r.add_picture(path, width=width, height=height)
    except Exception:
        return None
    BUILD_STATS['content_images_rendered'] = BUILD_STATS.get('content_images_rendered', 0) + 1
    if is_contained_fragment:
        BUILD_STATS['content_image_fragments_contained'] = BUILD_STATS.get('content_image_fragments_contained', 0) + 1
    if caption:
        cap = add_caption(caption, 'figure_caption')
        cap.paragraph_format.keep_together = True
        keep_paragraph_with_previous(cap)
    return p


'''
