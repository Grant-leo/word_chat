"""Cover reconstruction runtime template fragment for generated build scripts."""
from __future__ import annotations

COVER_RUNTIME = r'''
def normalize_label(text):
    return re.sub(r'[\s：:]+', '', str(text or ''))


def para_text_from_cover_el(el):
    if el.get('type') in ('para', 'empty', 'image'):
        return ''.join(r.get('t', '') for r in el.get('r', []))
    if el.get('type') == 'table':
        parts = []
        for row in el.get('rows', []):
            for cell in row:
                for pp in cell.get('p', []):
                    parts.append(''.join(r.get('t', '') for r in pp.get('r', [])))
        return ''.join(parts)
    return ''


def is_template_placeholder_text(text):
    compact = re.sub(r'\s+', ' ', str(text or '')).strip()
    if not compact:
        return False
    placeholder_patterns = [
        r'^\(Insert\b.*\)$',
        r'^\(E\.g\.\s*X{2,}(?:\s+X{2,})*,?\s*PhD\)$',
        r'\((?:Insert|student|date of|title of|name of)\b[^)]*\)',
        r'^\s*(?:报名序号|论文编号|学号|学生姓名|姓名|学院|专业|班级|指导教师|指导老师|日期)\s*[:：]\s*\[[^\]]+\]\s*$',
        r'\[[^\]\n]*(?:报名|序号|姓名|学号|学院|专业|班级|题目|指导|教师|日期|编码|待填|请输入|XX|XXX)[^\]\n]*\]',
        r'(?:待填写|待补全|请输入)',
        r'X{3,}(?:\s+X{3,})+',
    ]
    return any(re.search(pat, compact, re.I) for pat in placeholder_patterns)


def apply_cover_run(run, rd):
    font = rd.get('fn') or rd.get('fe') or '宋体'
    east = rd.get('fe') or font
    latin = rd.get('fn') or ('Times New Roman' if east in CJK_FONTS else east)
    set_run_fonts(run, east, latin)
    if rd.get('sz'):
        run.font.size = Pt(float(rd.get('sz')))
    run.bold = bool(rd.get('b', False))


def apply_cover_paragraph_format(p, el):
    am = {'left': 'LEFT', 'center': 'CENTER', 'right': 'RIGHT', 'both': 'JUSTIFY', 'distribute': 'DISTRIBUTE'}
    if el.get('al'):
        p.alignment = ALIGN.get(am.get(el.get('al'), 'LEFT'), WD_ALIGN_PARAGRAPH.LEFT)
    pf = p.paragraph_format
    line = el.get('ls_val')
    rule = el.get('ls_rule')
    if line:
        try:
            n = int(line)
            pf.line_spacing = Pt(n / 20.0) if rule in ('exact', 'atLeast') else n / 240.0
        except Exception:
            pass
    if el.get('sp_before'):
        try: pf.space_before = Pt(int(el.get('sp_before')) / 20.0)
        except Exception: pass
    if el.get('sp_after'):
        try: pf.space_after = Pt(int(el.get('sp_after')) / 20.0)
        except Exception: pass
    if el.get('fl_indent'):
        try: pf.first_line_indent = Pt(int(el.get('fl_indent')) / 20.0)
        except Exception: pass


def asset_path(name):
    if not name:
        return None
    bases = []
    if DATA.get('assets_dir'):
        bases.append(DATA.get('assets_dir'))
    bases.extend([os.path.join(BASE, 'assets'), BASE, os.getcwd()])
    for b in bases:
        p = os.path.join(b, name) if not os.path.isabs(name) else name
        if os.path.exists(p):
            return p
    return None


def image_width_from_extent(extent, default_inches=1.2):
    if not extent:
        return Inches(default_inches)
    try:
        cx = int(extent.get('cx') or 0)
        if cx > 0:
            return Inches(cx / 914400.0)
    except Exception:
        pass
    return Inches(default_inches)


def image_dimensions_from_extent(extent, default_inches=1.2):
    if not extent:
        return Inches(default_inches), None
    try:
        cx = int(extent.get('cx') or 0)
        cy = int(extent.get('cy') or 0)
        width = Inches(cx / 914400.0) if cx > 0 else Inches(default_inches)
        height = Inches(cy / 914400.0) if cy > 0 else None
        return width, height
    except Exception:
        return Inches(default_inches), None


def add_asset_picture(run, rd, default_inches=1.2):
    path = asset_path(rd.get('asset') or rd.get('image'))
    if not path:
        return False
    try:
        width, height = image_dimensions_from_extent(rd.get('extent'), default_inches)
        if height is not None:
            run.add_picture(path, width=width, height=height)
        else:
            run.add_picture(path, width=width)
        return True
    except Exception:
        return False


def render_cover_para(el):
    p = doc.add_paragraph()
    apply_cover_paragraph_format(p, el)
    for rd in el.get('r', []):
        rr = p.add_run(rd.get('t', ''))
        apply_cover_run(rr, rd)
    return p


def render_cover_image(el):
    p = doc.add_paragraph()
    apply_cover_paragraph_format(p, el)
    extent = el.get('extent') or {}
    try:
        cy = int(extent.get('cy') or 0)
        if cy > 0:
            p.paragraph_format.line_spacing = Pt((cy / 12700.0) + 4.0)
            p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.AT_LEAST
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
    except Exception:
        pass
    if not el.get('al'):
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run()
    add_asset_picture(r, el, default_inches=1.35)
    return p


def set_cell_borders(cell, **sides):
    tcPr = cell._tc.get_or_add_tcPr()
    old = tcPr.find(qn('w:tcBorders'))
    if old is not None:
        tcPr.remove(old)
    tcB = OxmlElement('w:tcBorders')
    for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        val = sides.get(side, 'nil')
        el = OxmlElement('w:' + side)
        if isinstance(val, dict):
            el.set(qn('w:val'), val.get('val', 'single'))
            el.set(qn('w:sz'), str(val.get('sz', '8')))
            el.set(qn('w:color'), val.get('color', '000000'))
        else:
            el.set(qn('w:val'), val)
            el.set(qn('w:sz'), '0' if val in ('nil', 'none') else '8')
            el.set(qn('w:color'), '000000')
        el.set(qn('w:space'), '0')
        tcB.append(el)
    tcPr.append(tcB)


def set_table_indent(table, twips=0):
    tblPr = table._tbl.tblPr
    tblInd = tblPr.find(qn('w:tblInd'))
    if tblInd is None:
        tblInd = OxmlElement('w:tblInd')
        tblPr.append(tblInd)
    tblInd.set(qn('w:w'), str(int(twips)))
    tblInd.set(qn('w:type'), 'dxa')




def _set_or_remove_attr(el, name, value):
    attr = qn('w:' + name)
    if value is None:
        if attr in el.attrib:
            del el.attrib[attr]
    else:
        el.set(attr, str(value))


def set_table_width(table, spec):
    if not spec:
        return
    tblPr = table._tbl.tblPr
    tblW = tblPr.find(qn('w:tblW'))
    if tblW is None:
        tblW = OxmlElement('w:tblW')
        tblPr.insert(0, tblW)
    for k, v in spec.items():
        _set_or_remove_attr(tblW, k, v)


def set_table_alignment_from_jc(table, jc):
    if jc in ('left', 'start'):
        table.alignment = WD_TABLE_ALIGNMENT.LEFT
    elif jc == 'right':
        table.alignment = WD_TABLE_ALIGNMENT.RIGHT
    elif jc == 'center':
        table.alignment = WD_TABLE_ALIGNMENT.CENTER


def set_table_layout(table, layout_type):
    if not layout_type:
        return
    tblPr = table._tbl.tblPr
    layout = tblPr.find(qn('w:tblLayout'))
    if layout is None:
        layout = OxmlElement('w:tblLayout')
        tblPr.append(layout)
    layout.set(qn('w:type'), layout_type)


def set_table_cell_margins(table, margins):
    if not margins:
        return
    tblPr = table._tbl.tblPr
    cellMar = tblPr.find(qn('w:tblCellMar'))
    if cellMar is not None:
        tblPr.remove(cellMar)
    cellMar = OxmlElement('w:tblCellMar')
    for side, attrs in margins.items():
        el = OxmlElement('w:' + side)
        for k, v in (attrs or {}).items():
            _set_or_remove_attr(el, k, v)
        cellMar.append(el)
    tblPr.append(cellMar)


def set_table_grid(table, grid_cols):
    if not grid_cols:
        return
    tblGrid = table._tbl.tblGrid
    if tblGrid is None:
        tblGrid = OxmlElement('w:tblGrid')
        table._tbl.insert(0, tblGrid)
    for i, w in enumerate(grid_cols):
        if i < len(tblGrid.gridCol_lst):
            gc = tblGrid.gridCol_lst[i]
        else:
            gc = OxmlElement('w:gridCol')
            tblGrid.append(gc)
        gc.set(qn('w:w'), str(w))


def apply_cover_table_props(table, el):
    props = el.get('tblPr') or {}
    if props:
        set_table_width(table, props.get('tblW'))
        if props.get('tblInd'):
            try:
                set_table_indent(table, int(props['tblInd'].get('w') or 0))
            except Exception:
                pass
        set_table_alignment_from_jc(table, props.get('jc'))
        layout = props.get('tblLayout') or {}
        set_table_layout(table, layout.get('type'))
        set_table_cell_margins(table, props.get('cellMar'))
        set_table_grid(table, props.get('grid_cols'))


def set_cover_cell_margins(cell, margins):
    if not margins:
        return
    tcPr = cell._tc.get_or_add_tcPr()
    old = tcPr.find(qn('w:tcMar'))
    if old is not None:
        tcPr.remove(old)
    mar = OxmlElement('w:tcMar')
    for side, attrs in margins.items():
        el = OxmlElement('w:' + side)
        for k, v in (attrs or {}).items():
            _set_or_remove_attr(el, k, v)
        mar.append(el)
    tcPr.append(mar)


def apply_cell_props(cell, cell_data):
    tcPr_data = cell_data.get('tcPr') or {}
    tcW = tcPr_data.get('tcW') or {}
    if tcW:
        tcPr = cell._tc.get_or_add_tcPr()
        old = tcPr.find(qn('w:tcW'))
        if old is None:
            old = OxmlElement('w:tcW')
            tcPr.insert(0, old)
        for k, v in tcW.items():
            _set_or_remove_attr(old, k, v)
    if tcPr_data.get('tcMar'):
        set_cover_cell_margins(cell, tcPr_data.get('tcMar'))
    valign = tcPr_data.get('vAlign')
    if valign == 'top':
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    elif valign == 'bottom':
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.BOTTOM
    else:
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def set_cell_no_wrap(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    if tcPr.find(qn('w:noWrap')) is None:
        tcPr.append(OxmlElement('w:noWrap'))


def apply_row_props(row, row_props):
    if not row_props:
        return
    trPr = row._tr.get_or_add_trPr()
    height = row_props.get('height') or {}
    if height:
        old = trPr.find(qn('w:trHeight'))
        if old is None:
            old = OxmlElement('w:trHeight')
            trPr.append(old)
        for k, v in height.items():
            _set_or_remove_attr(old, k, v)
    if row_props.get('cantSplit'):
        if trPr.find(qn('w:cantSplit')) is None:
            trPr.append(OxmlElement('w:cantSplit'))


def cover_table_sample_value(el, label_suffix='题目'):
    for row in el.get('rows') or []:
        if len(row) < 2:
            continue
        label = normalize_label(first_cell_label(row))
        if label.endswith(label_suffix):
            return ''.join(r.get('t', '') for pp in row[1].get('p', []) for r in pp.get('r', []))
    return ''


def cover_table_value_cell(row):
    if not row or len(row) < 2:
        return None
    return row[1]


def cover_text_units(text):
    units = 0.0
    for ch in str(text or ''):
        if ch.isspace():
            continue
        units += 0.5 if ch.isascii() else 1.0
    return units


def cover_title_capacity_chars(el, label_suffix='题目'):
    for row in el.get('rows') or []:
        label = normalize_label(first_cell_label(row))
        if not label.endswith(label_suffix):
            continue
        cell = cover_table_value_cell(row)
        if not cell:
            continue
        width_dxa = 0
        try:
            width_dxa = int(cell.get('w') or 0)
        except Exception:
            width_dxa = 0
        try:
            tcw = (cell.get('tcPr') or {}).get('tcW') or {}
            width_dxa = max(width_dxa, int(tcw.get('w') or 0))
        except Exception:
            pass
        font_size = 16.0
        sizes = []
        for pp in cell.get('p') or []:
            for rd in pp.get('r') or []:
                try:
                    if rd.get('sz'):
                        sizes.append(float(rd.get('sz')))
                except Exception:
                    pass
        if sizes:
            font_size = max(sizes)
        if width_dxa <= 0 or font_size <= 0:
            return 20.0
        width_pt = width_dxa / 20.0
        # CJK cover titles are close to one em per character.  A small
        # reserve avoids treating text that barely fits as one-line content.
        return max(1.0, width_pt / (font_size * 1.05))
    return 20.0


def estimate_cover_title_lines(text, el):
    capacity = cover_title_capacity_chars(el)
    return max(1, int(math.ceil(cover_text_units(text) / max(capacity, 1.0))))


def is_cover_empty_paragraph(el):
    """True for structurally empty cover paragraphs, including section markers."""
    if el.get('type') != 'empty':
        return False
    return not ''.join(r.get('t', '') for r in el.get('r', [])).strip()


def is_elastic_cover_empty(el):
    """Empty spacer that can be removed without deleting a section break."""
    return is_cover_empty_paragraph(el) and not el.get('section_break_after')


def compute_cover_skip_indices(cover):
    """Remove template spacer paragraphs only when replacement content is longer.

    This implements the template instruction 'delete one return above/below the table
    if the title uses two lines' without checking any school-specific text.  It
    compares the sample title in the template with the actual title and removes
    only structurally empty spacer paragraphs adjacent to the cover info table.
    """
    info_idx = next((i for i, el in enumerate(cover) if el.get('role') == 'cover_info_table'), None)
    if info_idx is None:
        return set()
    info_el = cover[info_idx]
    actual = str((DATA.get('cover_info') or {}).get('paper_title') or DATA.get('title_cn') or '').strip()
    sample = cover_table_sample_value(info_el, '题目')
    if not actual or not sample:
        return set()
    extra_lines = max(0, estimate_cover_title_lines(actual, info_el) - estimate_cover_title_lines(sample, info_el))
    skip = set()
    j = info_idx - 1
    before = []
    while j >= 0 and is_elastic_cover_empty(cover[j]):
        before.append(j); j -= 1
    # Pre-table blank paragraphs are elastic vertical budget.  Generated
    # content must keep the full cover on page one, so these can be removed.
    for idx in before:
        skip.add(idx)
    j = info_idx + 1
    after = []
    while j < len(cover) and is_elastic_cover_empty(cover[j]):
        after.append(j); j += 1
    # Keep a bounded gap before the next visible paragraph (normally a
    # committee/signature line).  Keeping every template spacer can push that
    # paragraph to page two, while removing all of them visually attaches it
    # to the final table row.
    keep_after = max(1, min(2, len(after) - extra_lines))
    after_remove = max(0, len(after) - keep_after)
    for idx in after[:after_remove]:
        skip.add(idx)
    # Also delete trailing blank spacer paragraphs immediately before an
    # empty section-break carrier. Keep the marker itself so the render loop
    # can still create the intended next-page section.
    for marker_idx, marker_el in enumerate(cover):
        if cover_element_is_empty_section_marker(marker_el):
            k = marker_idx - 1
            while k >= 0 and is_elastic_cover_empty(cover[k]):
                skip.add(k)
                k -= 1
    return skip


def cover_element_is_empty_section_marker(el):
    # A section break can be stored on a visually empty paragraph.  Render the
    # break, but never render that paragraph itself; otherwise it becomes a
    # blank page between declaration/front-matter sections.
    return bool(el.get('section_break_after')) and is_cover_empty_paragraph(el)

def first_cell_label(row):
    if not row:
        return ''
    cell = row[0]
    return ''.join(r.get('t', '') for pp in cell.get('p', []) for r in pp.get('r', []))


def is_code_like_cover_table(el):
    if el.get('role') == 'cover_code_table':
        return True
    rows = el.get('rows') or []
    labels = [normalize_label(first_cell_label(row)) for row in rows if row]
    return bool(labels) and len(rows) <= 2 and all(x.endswith('编码') for x in labels if x)


def render_cover_table(el):
    rows = el.get('rows', [])
    if not rows:
        return None
    ncols = max(len(r) for r in rows)
    table = doc.add_table(rows=len(rows), cols=ncols)
    apply_cover_table_props(table, el)
    code_like_table = is_code_like_cover_table(el)
    if code_like_table and ncols >= 2:
        table.autofit = False
        try:
            table.columns[0].width = Cm(2.45)
            table.columns[1].width = Cm(1.75)
        except Exception:
            pass
    # Fallback only when the extractor did not provide table-level properties.
    if not (el.get('tblPr') or {}).get('jc'):
        table.alignment = WD_TABLE_ALIGNMENT.LEFT if code_like_table else WD_TABLE_ALIGNMENT.CENTER
    if code_like_table and not (el.get('tblPr') or {}).get('tblInd'):
        set_table_indent(table, 0)
    cover_info = DATA.get('cover_info') or {}
    label_map = {
        '学校编码': cover_info.get('school_code', '') or cover_info.get('degree_code', ''),
        '学位编码': cover_info.get('degree_code', '') or cover_info.get('school_code', ''),
        '论文题目': cover_info.get('paper_title', ''),
        '学生姓名': cover_info.get('student_name', ''),
        '学号': cover_info.get('student_id', ''),
        '所属学院': cover_info.get('college', ''),
        '专业班级': cover_info.get('class_name', ''),
        '指导老师': cover_info.get('advisor', ''),
        '指导教师': cover_info.get('advisor', ''),
    }
    norm_map = {normalize_label(k): v for k, v in label_map.items() if v}
    row_props = (el.get('tblPr') or {}).get('rows') or []
    for ri, row in enumerate(rows):
        if ri < len(table.rows):
            apply_row_props(table.rows[ri], row_props[ri] if ri < len(row_props) else {})
        row_label = ''
        if row and row[0].get('p'):
            row_label = ''.join(r.get('t', '') for r in row[0]['p'][0].get('r', []))
        row_key = normalize_label(row_label)
        row_value = norm_map.get(row_key, '')
        if not row_value:
            for k, v in norm_map.items():
                if k and (k in row_key or row_key in k):
                    row_value = v; break
        force_left = code_like_table or row_key.endswith('编码')
        for ci in range(ncols):
            cell = table.rows[ri].cells[ci]
            cell.text = ''
            cell_data = row[ci] if ci < len(row) else {'p': []}
            apply_cell_props(cell, cell_data)
            if code_like_table:
                set_cell_no_wrap(cell)
            if cell_data.get('w') and not (cell_data.get('tcPr') or {}).get('tcW'):
                try: cell.width = Cm(float(cell_data.get('w')) / 567.0)
                except Exception: pass
            paras = cell_data.get('p') or [{'r': []}]
            for pi, pp in enumerate(paras):
                p = cell.paragraphs[0] if pi == 0 else cell.add_paragraph()
                apply_cover_paragraph_format(p, pp)
                if force_left:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                use_value = (ci == 1 and row_value)
                if use_value:
                    rd = (pp.get('r') or [{}])[0]
                    rr = p.add_run(row_value)
                    apply_cover_run(rr, rd)
                    if code_like_table and not rd.get('sz'):
                        rr.font.size = Pt(10.5)
                    continue
                for rd in pp.get('r', []) or [{}]:
                    rr = p.add_run(rd.get('t', ''))
                    apply_cover_run(rr, rd)
                    if code_like_table and not rd.get('sz'):
                        rr.font.size = Pt(10.5)
                    if rd.get('asset') or rd.get('image'):
                        add_asset_picture(rr, rd)
            borders = cell_data.get('borders') or {}
            if borders:
                set_cell_borders(cell, **borders)
    return table


def render_cover_and_declarations():
    setup_section(doc.sections[0])
    clear_header_footer(doc.sections[0])
    cover = DATA.get('cover') or []
    if not cover:
        return add_section_with_header('upperRoman', 1)
    front_started = False
    skip_indices = compute_cover_skip_indices(cover)
    for idx, el in enumerate(cover):
        # A paragraph whose only purpose is to carry a section break must not
        # be rendered as a blank page. Treat it as a structural marker.
        if idx in skip_indices:
            continue
        if not el.get('section_break_after') and is_template_placeholder_text(para_text_from_cover_el(el)):
            continue
        should_render = not cover_element_is_empty_section_marker(el)
        if should_render:
            if el.get('type') in ('para', 'empty'):
                render_cover_para(el)
            elif el.get('type') == 'table':
                render_cover_table(el)
            elif el.get('type') == 'image':
                render_cover_image(el)
        if el.get('section_break_after'):
            if not front_started:
                front_started = True
                add_section_with_header('upperRoman', 1)
            else:
                add_section_with_header('upperRoman', None)
    if not front_started:
        add_section_with_header('upperRoman', 1)
    return doc.sections[-1]
'''
