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


def estimate_table_row_lines(row):
    estimated = 1
    for cell in row or []:
        parts = str(cell or '').split('\n') or ['']
        cell_lines = sum(max(1, math.ceil(len(part) / 14)) for part in parts)
        estimated = max(estimated, cell_lines)
    return estimated


def should_prevent_row_split(row, is_header=False):
    if is_header:
        return True
    text_cells = [str(cell or '') for cell in row or []]
    if any(len(cell) > 420 for cell in text_cells):
        return False
    return estimate_table_row_lines(row) <= 12


def should_keep_table_together(rows):
    if not rows or len(rows) > 10:
        return False
    text_cells = [str(cell or '') for row in rows for cell in row]
    if any(len(cell) > 90 for cell in text_cells):
        return False
    estimated_lines = 0
    for row in rows:
        estimated_lines += estimate_table_row_lines(row)
    return estimated_lines <= 18


def keep_table_together(table):
    for ri, row in enumerate(table.rows):
        for cell in row.cells:
            for p in cell.paragraphs:
                p.paragraph_format.keep_together = True
                if ri < len(table.rows) - 1:
                    p.paragraph_format.keep_with_next = True


def content_image_path(filename):
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
    return next((p for p in candidates if p and os.path.exists(p)), None)


def fit_table_cell_picture_dimensions(path, ncols=1):
    try:
        px_width, px_height, image = picture_pixel_size(path)
        if image is None:
            return Inches(max(1.0, text_width_inches(1.0) / max(int(ncols or 1), 1) * 0.9)), None
        max_width = Inches(max(0.75, text_width_inches(1.0) / max(int(ncols or 1), 1) * 0.9))
        if px_width and px_height and (px_width < 240 or px_height < 80):
            width, height = image.scaled_dimensions()
            if width <= max_width:
                return width, height
        return image.scaled_dimensions(width=max_width)
    except Exception:
        return Inches(max(0.75, text_width_inches(1.0) / max(int(ncols or 1), 1) * 0.9)), None


def render_table_cell_image(cell, filename, ncols=1, force_new_paragraph=False):
    path = content_image_path(filename)
    if not path:
        return None
    p = cell.add_paragraph() if force_new_paragraph else (
        cell.paragraphs[0] if cell.paragraphs and not cell.paragraphs[0].text.strip() else cell.add_paragraph()
    )
    configure_picture_paragraph(p, keep_with_next=False)
    r = p.add_run()
    try:
        width, height = fit_table_cell_picture_dimensions(path, ncols=ncols)
        r.add_picture(path, width=width, height=height)
    except Exception:
        return None
    BUILD_STATS['content_images_rendered'] = BUILD_STATS.get('content_images_rendered', 0) + 1
    return p


def table_cell_media_map(cell_items):
    by_cell = {}
    for entry in cell_items or []:
        if not isinstance(entry, dict):
            continue
        try:
            key = (int(entry.get('row') or 0), int(entry.get('col') or 0))
        except Exception:
            continue
        by_cell.setdefault(key, []).extend(entry.get('items') or [])
    return by_cell


def normalize_table_col_widths(table_col_widths, ncols, max_total_twips=None):
    widths = []
    for value in table_col_widths or []:
        try:
            width = int(value or 0)
        except Exception:
            width = 0
        widths.append(max(0, width))
        if len(widths) >= ncols:
            break
    if len(widths) < ncols:
        widths.extend([0] * (ncols - len(widths)))
    positive_widths = [width for width in widths if width > 0]
    if not positive_widths:
        return []
    fallback_width = max(120, int(sum(positive_widths) / len(positive_widths)))
    widths = [width if width > 0 else fallback_width for width in widths]
    total = sum(widths)
    max_total = safe_positive_int(max_total_twips)
    if not max_total:
        try:
            max_total = int(text_width_cm() * 567)
        except Exception:
            max_total = 0
    if total > 0 and max_total > 0 and total > max_total:
        scale = max_total / float(total)
        min_width = 120
        if len(widths) * min_width > max_total:
            min_width = max(1, int(max_total / max(len(widths), 1)))
        widths = [max(min_width, int(width * scale)) for width in widths]
        overflow = sum(widths) - max_total
        while overflow > 0 and widths:
            idx = max(range(len(widths)), key=lambda i: widths[i])
            reducible = max(0, widths[idx] - 1)
            if reducible <= 0:
                break
            delta = min(reducible, overflow)
            widths[idx] -= delta
            overflow -= delta
    return widths


def set_fixed_table_layout(table):
    try:
        table.autofit = False
    except Exception:
        pass
    try:
        tblPr = table._tbl.tblPr
        layout = tblPr.find(qn('w:tblLayout'))
        if layout is None:
            layout = OxmlElement('w:tblLayout')
            insert_property_child(tblPr, layout, 'tblPr')
        layout.set(qn('w:type'), 'fixed')
    except Exception:
        pass


def set_table_grid_widths(table, widths):
    if not widths:
        return False
    set_fixed_table_layout(table)
    try:
        tblGrid = table._tbl.tblGrid
        if tblGrid is None:
            tblGrid = OxmlElement('w:tblGrid')
            table._tbl.insert(0, tblGrid)
        for old in list(tblGrid):
            tblGrid.remove(old)
        for width in widths:
            grid_col = OxmlElement('w:gridCol')
            grid_col.set(qn('w:w'), str(int(width or 0)))
            tblGrid.append(grid_col)
        return True
    except Exception:
        return False


def set_cell_width_twips(cell, width):
    if not width:
        return
    try:
        tcPr = cell._tc.get_or_add_tcPr()
        tcW = tcPr.find(qn('w:tcW'))
        if tcW is None:
            tcW = OxmlElement('w:tcW')
            tcPr.insert(0, tcW)
        tcW.set(qn('w:w'), str(int(width)))
        tcW.set(qn('w:type'), 'dxa')
    except Exception:
        pass


def normalize_margin_twips(margins):
    if not isinstance(margins, dict):
        return {}
    out = {}
    aliases = {'left': ('left', 'start'), 'right': ('right', 'end'), 'top': ('top',), 'bottom': ('bottom',)}
    for side, names in aliases.items():
        value = None
        for name in names:
            if name in margins:
                value = margins.get(name)
                break
        try:
            width = int(value or 0)
        except Exception:
            width = 0
        if width > 0:
            out[side] = width
    return out


def set_margin_container(parent, container_name, margins):
    margins = normalize_margin_twips(margins)
    if not margins:
        return False
    try:
        old = parent.find(qn('w:' + container_name))
        if old is not None:
            parent.remove(old)
        container = OxmlElement('w:' + container_name)
        for side in ('top', 'left', 'bottom', 'right'):
            if side not in margins:
                continue
            el = OxmlElement('w:' + side)
            el.set(qn('w:w'), str(int(margins[side])))
            el.set(qn('w:type'), 'dxa')
            container.append(el)
        insert_property_child(parent, container, property_parent_kind(container_name))
        return True
    except Exception:
        return False


def set_table_default_cell_margins(table, margins):
    try:
        tblPr = table._tbl.tblPr
        return set_margin_container(tblPr, 'tblCellMar', margins)
    except Exception:
        return False


def set_cell_margins_twips(cell, margins):
    try:
        tcPr = cell._tc.get_or_add_tcPr()
        return set_margin_container(tcPr, 'tcMar', margins)
    except Exception:
        return False


def normalize_border_spec(spec):
    if isinstance(spec, dict):
        val = str(spec.get('val') or spec.get('type') or '').strip()
        if not val:
            return {}
        out = {'val': val}
        for attr, aliases in (
            ('sz', ('sz', 'size')),
            ('color', ('color',)),
            ('space', ('space',)),
        ):
            value = None
            for name in aliases:
                if name in spec:
                    value = spec.get(name)
                    break
            if value is None:
                continue
            if attr in ('sz', 'space'):
                try:
                    value = max(0, int(value))
                except Exception:
                    pass
            value = str(value).strip()
            if value:
                out[attr] = value
        if val in ('nil', 'none'):
            out.setdefault('sz', '0')
        return out
    val = str(spec or '').strip()
    if not val:
        return {}
    if val in ('nil', 'none'):
        return {'val': 'nil', 'sz': '0'}
    return {'val': val}


def normalize_border_map(borders):
    if not isinstance(borders, dict):
        return {}
    out = {}
    allowed = ('top', 'left', 'bottom', 'right', 'insideH', 'insideV', 'tl2br', 'tr2bl', 'start', 'end')
    for side in allowed:
        if side not in borders:
            continue
        spec = normalize_border_spec(borders.get(side))
        if spec:
            out[side] = spec
    return out


def property_parent_kind(container_name):
    if container_name in ('tblBorders', 'tblCellMar', 'tblLayout'):
        return 'tblPr'
    if container_name in ('tcBorders', 'tcMar'):
        return 'tcPr'
    return ''


def property_order_index(local_name, kind):
    orders = {
        'tblPr': (
            'tblStyle', 'tblpPr', 'tblOverlap', 'bidiVisual',
            'tblStyleRowBandSize', 'tblStyleColBandSize', 'tblW', 'jc',
            'tblCellSpacing', 'tblInd', 'tblBorders', 'shd', 'tblLayout',
            'tblCellMar', 'tblLook', 'tblCaption', 'tblDescription', 'tblPrChange',
        ),
        'tcPr': (
            'cnfStyle', 'tcW', 'gridSpan', 'hMerge', 'vMerge', 'tcBorders',
            'shd', 'noWrap', 'tcMar', 'textDirection', 'tcFitText', 'vAlign',
            'hideMark', 'headers',
        ),
    }
    try:
        return orders.get(kind, ()).index(local_name)
    except ValueError:
        return 10_000


def insert_property_child(parent, child, kind):
    if not kind:
        parent.append(child)
        return
    child_name = str(child.tag).rsplit('}', 1)[-1]
    child_index = property_order_index(child_name, kind)
    for pos, existing in enumerate(list(parent)):
        existing_name = str(existing.tag).rsplit('}', 1)[-1]
        if property_order_index(existing_name, kind) > child_index:
            parent.insert(pos, child)
            return
    parent.append(child)


def border_specs_from_container(container):
    specs = {}
    if container is None:
        return specs
    for child in list(container):
        side = str(child.tag).rsplit('}', 1)[-1]
        if side not in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV', 'tl2br', 'tr2bl', 'start', 'end'):
            continue
        spec = {}
        for attr in ('val', 'sz', 'color', 'space'):
            value = child.get(qn('w:' + attr))
            if value is not None:
                spec[attr] = str(value)
        if spec:
            specs[side] = spec
    return specs


def set_border_container(parent, container_name, borders, merge_existing=False):
    borders = normalize_border_map(borders)
    if not borders:
        return False
    try:
        old = parent.find(qn('w:' + container_name))
        if merge_existing and old is not None:
            merged = border_specs_from_container(old)
            merged.update(borders)
            borders = normalize_border_map(merged)
        if old is not None:
            parent.remove(old)
        container = OxmlElement('w:' + container_name)
        for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV', 'tl2br', 'tr2bl', 'start', 'end'):
            spec = borders.get(side)
            if not spec:
                continue
            el = OxmlElement('w:' + side)
            for attr in ('val', 'sz', 'color', 'space'):
                if attr in spec:
                    el.set(qn('w:' + attr), str(spec[attr]))
            container.append(el)
        insert_property_child(parent, container, property_parent_kind(container_name))
        return True
    except Exception:
        return False


def set_table_borders(table, borders):
    try:
        tblPr = table._tbl.tblPr
        return set_border_container(tblPr, 'tblBorders', borders)
    except Exception:
        return False


def set_cell_borders_from_spec(cell, borders):
    try:
        tcPr = cell._tc.get_or_add_tcPr()
        return set_border_container(tcPr, 'tcBorders', borders, merge_existing=True)
    except Exception:
        return False


def apply_table_borders(table, table_borders):
    if set_table_borders(table, table_borders or {}):
        BUILD_STATS['content_table_borders_rendered'] = BUILD_STATS.get('content_table_borders_rendered', 0) + 1
        return True
    return False


def normalize_row_height_spec(spec):
    if isinstance(spec, dict):
        value = spec.get('val') if 'val' in spec else spec.get('height')
        rule = str(spec.get('rule') or spec.get('hRule') or '').strip()
    else:
        value = spec
        rule = ''
    try:
        val = int(value or 0)
    except Exception:
        val = 0
    if val <= 0:
        return None
    out = {'val': val}
    if rule in ('auto', 'exact', 'atLeast'):
        out['rule'] = rule
    return out


def apply_row_heights(table, row_heights):
    rendered = False
    for ri, spec in enumerate(row_heights or []):
        if ri >= len(table.rows):
            break
        height = normalize_row_height_spec(spec)
        if not height:
            continue
        try:
            trPr = table.rows[ri]._tr.get_or_add_trPr()
            tr_height = trPr.find(qn('w:trHeight'))
            if tr_height is None:
                tr_height = OxmlElement('w:trHeight')
                trPr.append(tr_height)
            tr_height.set(qn('w:val'), str(int(height['val'])))
            if height.get('rule'):
                tr_height.set(qn('w:hRule'), height['rule'])
            rendered = True
        except Exception:
            continue
    if rendered:
        BUILD_STATS['content_table_row_heights_rendered'] = BUILD_STATS.get('content_table_row_heights_rendered', 0) + 1
    return rendered


def apply_repeat_header_rows(table, repeat_rows=None):
    if repeat_rows is None:
        if len(table.rows):
            repeat_table_header(table.rows[0])
        return 0
    try:
        count = int(repeat_rows or 0)
    except Exception:
        count = 0
    count = max(0, min(count, len(table.rows)))
    for ri in range(count):
        repeat_table_header(table.rows[ri])
    if count:
        BUILD_STATS['content_table_repeat_header_rows_rendered'] = BUILD_STATS.get('content_table_repeat_header_rows_rendered', 0) + count
    return count


def apply_cell_overrides(table, overrides):
    rendered = 0
    border_rendered = 0
    align_map = {
        'top': WD_CELL_VERTICAL_ALIGNMENT.TOP,
        'center': WD_CELL_VERTICAL_ALIGNMENT.CENTER,
        'middle': WD_CELL_VERTICAL_ALIGNMENT.CENTER,
        'bottom': WD_CELL_VERTICAL_ALIGNMENT.BOTTOM,
    }
    for override in overrides or []:
        if not isinstance(override, dict):
            continue
        try:
            row = int(override.get('row') or 0)
            col = int(override.get('col') or 0)
        except Exception:
            continue
        if row < 0 or col < 0 or row >= len(table.rows) or col >= len(table.rows[row].cells):
            continue
        cell = table.rows[row].cells[col]
        changed = False
        align = str(override.get('v_align') or override.get('vertical_alignment') or '').strip().lower()
        if align in align_map:
            try:
                cell.vertical_alignment = align_map[align]
                changed = True
            except Exception:
                pass
        if set_cell_margins_twips(cell, override.get('margins_twips') or override.get('margins') or {}):
            changed = True
        if set_cell_borders_from_spec(cell, override.get('borders') or override.get('border_sides') or {}):
            changed = True
            border_rendered += 1
        if changed:
            rendered += 1
    if rendered:
        BUILD_STATS['content_table_cell_overrides_rendered'] = BUILD_STATS.get('content_table_cell_overrides_rendered', 0) + rendered
    if border_rendered:
        BUILD_STATS['content_table_cell_borders_rendered'] = BUILD_STATS.get('content_table_cell_borders_rendered', 0) + border_rendered
    return rendered


def apply_table_merges(table, table_merges):
    rendered = 0
    for merge in table_merges or []:
        if not isinstance(merge, dict):
            continue
        try:
            row = int(merge.get('row') or 0)
            col = int(merge.get('col') or 0)
            rowspan = max(1, int(merge.get('rowspan') or 1))
            colspan = max(1, int(merge.get('colspan') or 1))
            if rowspan <= 1 and colspan <= 1:
                continue
            end_row = row + rowspan - 1
            end_col = col + colspan - 1
            if row < 0 or col < 0 or end_row >= len(table.rows) or end_col >= len(table.rows[row].cells):
                continue
            table.cell(row, col).merge(table.cell(end_row, end_col))
            rendered += 1
        except Exception:
            continue
    if rendered:
        BUILD_STATS['content_table_merges_rendered'] = BUILD_STATS.get('content_table_merges_rendered', 0) + rendered
    return rendered


def media_after_paragraph_index(media):
    if not isinstance(media, dict) or 'after_paragraph_index' not in media:
        return None
    try:
        return max(0, int(media.get('after_paragraph_index') or 0))
    except Exception:
        return None


def media_replace_paragraph_index(media):
    if not isinstance(media, dict) or 'replace_paragraph_index' not in media:
        return None
    try:
        return max(0, int(media.get('replace_paragraph_index') or 0))
    except Exception:
        return None


def render_table_cell_rich_text(cell, item, prof, force_new_paragraph=False):
    if not isinstance(item, dict):
        return False
    runs = item.get('runs') or []
    if not runs and item.get('text'):
        runs = [{'type': 'text', 'text': item.get('text') or ''}]
    if not runs:
        return False
    p = cell.add_paragraph() if force_new_paragraph else (
        cell.paragraphs[0] if cell.paragraphs and not cell.paragraphs[0].text.strip() else cell.add_paragraph()
    )
    apply_paragraph_profile(p, prof, first_indent_override=0)
    wrote = False
    for run in runs:
        kind = run.get('type') or ('math' if run.get('math') else 'text')
        if kind == 'math':
            for m in run.get('math') or []:
                wrote = append_inline_formula(p, m) or wrote
        elif kind == 'note_ref':
            wrote = append_note_reference(p, run) or wrote
        else:
            text = str(run.get('text') or '')
            if text:
                add_text_runs(p, text, prof, False)
                wrote = True
    return wrote


def render_table_cell_media_item(cell, media, ncols, prof, force_new_paragraph=False):
    if not isinstance(media, dict):
        return False
    if media.get('role') == 'rich_text' or media.get('math'):
        return render_table_cell_rich_text(cell, media, prof, force_new_paragraph=force_new_paragraph)
    if media.get('role') == 'image' or media.get('image'):
        render_table_cell_image(
            cell,
            media.get('image') or media.get('filename') or media.get('asset') or '',
            ncols=ncols,
            force_new_paragraph=force_new_paragraph,
        )
        return True
    if media.get('table_rows'):
        render_table(
            media.get('table_rows') or [],
            media.get('table_cell_items') or [],
            media.get('table_merges') or [],
            media.get('table_col_widths_twips') or [],
            media.get('table_row_heights_twips') or [],
            media.get('table_repeat_header_rows'),
            media.get('table_cell_margins_twips') or {},
            media.get('table_cell_overrides') or [],
            media.get('table_borders') or {},
            container=cell,
            nested=True,
        )
        return True
    if media.get('role') == 'missing_image':
        p = cell.add_paragraph() if force_new_paragraph else (
            cell.paragraphs[0] if cell.paragraphs and not cell.paragraphs[0].text.strip() else cell.add_paragraph()
        )
        apply_paragraph_profile(p, prof, first_indent_override=0)
        r = p.add_run(media.get('text') or media.get('source') or 'missing image')
        apply_run_profile(r, prof, r.text)
        return True
    return False


def safe_positive_int(value):
    try:
        parsed = int(value or 0)
    except Exception:
        return 0
    return parsed if parsed > 0 else 0


def twips_to_cm(value):
    return float(value) / 567.0


def table_source_section_page_setup(item):
    if not isinstance(item, dict):
        return {}
    setup = item.get('source_section_page_setup') or {}
    if not isinstance(setup, dict):
        return {}
    width = safe_positive_int(setup.get('page_width_twips'))
    height = safe_positive_int(setup.get('page_height_twips'))
    orientation = str(setup.get('orientation') or '').strip().lower()
    if orientation != 'landscape' and not (width and height and width > height):
        return {}
    return setup


def table_source_section_text_width_twips(item):
    setup = table_source_section_page_setup(item)
    if not setup:
        return 0
    width = safe_positive_int(setup.get('page_width_twips'))
    if not width:
        return 0
    margins = setup.get('margins_twips') or {}
    left = safe_positive_int(margins.get('left')) if isinstance(margins, dict) else 0
    right = safe_positive_int(margins.get('right')) if isinstance(margins, dict) else 0
    available = width - left - right
    return available if available > 0 else width


def apply_source_section_page_setup(sec, setup):
    width = safe_positive_int(setup.get('page_width_twips'))
    height = safe_positive_int(setup.get('page_height_twips'))
    if width and height:
        sec.page_width = Cm(twips_to_cm(width))
        sec.page_height = Cm(twips_to_cm(height))
    margins = setup.get('margins_twips') or {}
    if isinstance(margins, dict):
        for attr, key in (
            ('top_margin', 'top'),
            ('bottom_margin', 'bottom'),
            ('left_margin', 'left'),
            ('right_margin', 'right'),
            ('header_distance', 'header'),
            ('footer_distance', 'footer'),
            ('gutter', 'gutter'),
        ):
            value = safe_positive_int(margins.get(key))
            if value:
                try:
                    setattr(sec, attr, Cm(twips_to_cm(value)))
                except Exception:
                    pass
    try:
        pg_sz = sec._sectPr.find(qn('w:pgSz'))
        if pg_sz is None:
            pg_sz = OxmlElement('w:pgSz')
            sec._sectPr.insert(0, pg_sz)
        pg_sz.set(qn('w:orient'), 'landscape')
    except Exception:
        pass


def begin_table_source_section(item):
    setup = table_source_section_page_setup(item)
    if not setup:
        return False
    remove_trailing_empty_body_paragraphs()
    sec = doc.add_section(WD_SECTION.NEW_PAGE)
    setup_section(sec)
    apply_header_footer(sec, 'decimal', None)
    apply_source_section_page_setup(sec, setup)
    BUILD_STATS['content_landscape_table_sections_rendered'] = BUILD_STATS.get('content_landscape_table_sections_rendered', 0) + 1
    return True


def end_table_source_section(started):
    if not started:
        return
    remove_trailing_empty_body_paragraphs()
    sec = doc.add_section(WD_SECTION.NEW_PAGE)
    setup_section(sec)
    apply_header_footer(sec, 'decimal', None)


def render_table_from_item(item, container=None, nested=False):
    return render_table(
        item.get('table_rows') or [],
        item.get('table_cell_items') or [],
        item.get('table_merges') or [],
        item.get('table_col_widths_twips') or [],
        item.get('table_row_heights_twips') or [],
        item.get('table_repeat_header_rows'),
        item.get('table_cell_margins_twips') or {},
        item.get('table_cell_overrides') or [],
        item.get('table_borders') or {},
        container=container,
        nested=nested,
        max_width_twips=table_source_section_text_width_twips(item) if not nested else None,
    )


def render_table_item(item):
    started = begin_table_source_section(item)
    try:
        return render_table_from_item(item)
    finally:
        end_table_source_section(started)


def render_table(rows, cell_items=None, table_merges=None, table_col_widths_twips=None, table_row_heights_twips=None, table_repeat_header_rows=None, table_cell_margins_twips=None, table_cell_overrides=None, table_borders=None, container=None, nested=False, max_width_twips=None):
    if not rows:
        return
    container = container or doc
    ncols = max(len(r) for r in rows)
    media_by_cell = table_cell_media_map(cell_items)
    explicit_borders = bool(normalize_border_map(table_borders))
    for override in table_cell_overrides or []:
        if isinstance(override, dict) and normalize_border_map(override.get('borders') or override.get('border_sides') or {}):
            explicit_borders = True
            break
    table = container.add_table(rows=len(rows), cols=ncols)
    BUILD_STATS['content_tables_rendered'] = BUILD_STATS.get('content_tables_rendered', 0) + 1
    if nested:
        BUILD_STATS['content_nested_tables_rendered'] = BUILD_STATS.get('content_nested_tables_rendered', 0) + 1
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    col_widths = normalize_table_col_widths(table_col_widths_twips, ncols, max_total_twips=max_width_twips)
    if col_widths and set_table_grid_widths(table, col_widths):
        BUILD_STATS['content_table_widths_rendered'] = BUILD_STATS.get('content_table_widths_rendered', 0) + 1
    if set_table_default_cell_margins(table, table_cell_margins_twips or {}):
        BUILD_STATS['content_table_cell_margins_rendered'] = BUILD_STATS.get('content_table_cell_margins_rendered', 0) + 1
    try:
        repeat_header_count = int(table_repeat_header_rows or 0)
    except Exception:
        repeat_header_count = 0
    if table_repeat_header_rows is None and len(rows):
        repeat_header_count = 1
    repeat_header_count = max(0, min(repeat_header_count, len(rows)))
    for ri, row in enumerate(rows):
        if should_prevent_row_split(row, is_header=ri < repeat_header_count):
            prevent_row_split(table.rows[ri])
        prof = profile('table_header' if ri == 0 else 'table_body')
        for ci in range(ncols):
            text = row[ci] if ci < len(row) else ''
            cell = table.rows[ri].cells[ci]
            if ci < len(col_widths):
                set_cell_width_twips(cell, col_widths[ci])
            cell.text = ''
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            parts = str(text or '').split('\n') if str(text or '') else []
            positioned_media = {}
            trailing_media = []
            replacement_media = {}
            for media in media_by_cell.get((ri, ci), []):
                replace_idx = media_replace_paragraph_index(media)
                if replace_idx is not None:
                    replacement_media.setdefault(replace_idx, []).append(media)
                    continue
                idx = media_after_paragraph_index(media)
                if idx is None:
                    trailing_media.append(media)
                else:
                    positioned_media.setdefault(idx, []).append(media)
            wrote_any = False
            for pos in range(len(parts) + 1):
                for media in positioned_media.get(pos, []):
                    if render_table_cell_media_item(cell, media, ncols, prof, force_new_paragraph=wrote_any):
                        wrote_any = True
                if pos >= len(parts):
                    continue
                part = parts[pos]
                replacements = replacement_media.get(pos) or []
                if replacements:
                    for media in replacements:
                        if render_table_cell_media_item(cell, media, ncols, prof, force_new_paragraph=wrote_any):
                            wrote_any = True
                else:
                    p = cell.paragraphs[0] if not wrote_any and cell.paragraphs else cell.add_paragraph()
                    apply_paragraph_profile(p, prof, first_indent_override=0)
                    r = p.add_run(part)
                    apply_run_profile(r, prof, part)
                    wrote_any = True
            for pos in sorted(key for key in positioned_media if key > len(parts)):
                for media in positioned_media.get(pos, []):
                    if render_table_cell_media_item(cell, media, ncols, prof, force_new_paragraph=wrote_any):
                        wrote_any = True
            for pos in sorted(key for key in replacement_media if key >= len(parts)):
                for media in replacement_media.get(pos, []):
                    if render_table_cell_media_item(cell, media, ncols, prof, force_new_paragraph=wrote_any):
                        wrote_any = True
            for media in trailing_media:
                if render_table_cell_media_item(cell, media, ncols, prof, force_new_paragraph=wrote_any):
                    wrote_any = True
            if not wrote_any:
                p = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
                apply_paragraph_profile(p, prof, first_indent_override=0)
    apply_repeat_header_rows(table, table_repeat_header_rows)
    apply_row_heights(table, table_row_heights_twips)
    apply_table_merges(table, table_merges)
    if explicit_borders:
        apply_table_borders(table, table_borders)
    else:
        apply_three_line_borders(table)
    apply_cell_overrides(table, table_cell_overrides)
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
    path = content_image_path(filename)
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
