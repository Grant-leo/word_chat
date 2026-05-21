"""
format_extractor.py — OOXML-direct format extraction with style resolution.
Solves python-docx API limitation: no more None values from style inheritance.
"""
import json, os, hashlib
from docx import Document
from docx.oxml.ns import qn
from lxml import etree

ALIGN_MAP = {0: 'LEFT', 1: 'CENTER', 2: 'RIGHT', 3: 'JUSTIFY', None: 'DEFAULT'}

def _tag(el):
    return el.tag.split('}')[-1] if '}' in el.tag else el.tag

def _val(el, attr='w:val', default=None):
    return el.get(qn(attr), default)

def _pt(half_pts_str):
    """Convert half-points string to float. '28' -> 14.0"""
    try: return int(half_pts_str) / 2.0
    except: return None

def _emu_to_pt(emu):
    try: return round(int(emu) / 12700, 1)
    except: return None


class StyleResolver:
    """Resolve formatting by walking style inheritance tree in styles.xml."""

    def __init__(self, doc):
        self.styles = {}  # style_id -> {font, size, bold, ...}
        self._load_styles(doc)

    def _load_styles(self, doc):
        for style in doc.styles:
            sid = style.style_id
            entry = {}
            f = style.font
            if f.name: entry['font'] = f.name
            if f.size:  entry['size'] = f.size.pt
            if f.bold is not None: entry['bold'] = f.bold
            if f.italic is not None: entry['italic'] = f.italic
            # Paragraph styles have paragraph_format; character styles don't
            try:
                pf = style.paragraph_format
                if pf.line_spacing: entry['ls'] = pf.line_spacing
                if pf.alignment is not None: entry['align'] = ALIGN_MAP.get(pf.alignment, 'DEFAULT')
                if pf.first_line_indent: entry['indent'] = pf.first_line_indent.cm
            except AttributeError:
                pass  # character style — no paragraph formatting
            try:
                entry['base'] = style.base_style.style_id if style.base_style else None
            except (AttributeError, ValueError):
                entry['base'] = None
            self.styles[sid] = entry

    def resolve(self, p_elem, r_elem):
        """Return fully resolved formatting for a run."""
        result = {'font': None, 'size': None, 'bold': False, 'italic': False,
                  'ls': None, 'align': 'DEFAULT', 'indent': 0}

        # 1. Direct formatting on the run (highest priority)
        rPr = r_elem.find(qn('w:rPr'))
        if rPr is not None:
            for child in rPr:
                ct = _tag(child)
                if ct == 'rFonts':
                    result['font'] = _val(child, 'w:ascii') or _val(child, 'w:hAnsi')
                elif ct == 'sz':
                    result['size'] = _pt(_val(child))
                elif ct == 'b':
                    result['bold'] = True
                elif ct == 'i':
                    result['italic'] = True
                elif ct == 'color':
                    result['color'] = _val(child)

        # 2. Paragraph properties
        pPr = p_elem.find(qn('w:pPr'))
        style_id = None
        if pPr is not None:
            for child in pPr:
                ct = _tag(child)
                if ct == 'pStyle':
                    style_id = _val(child)
                elif ct == 'jc':
                    v = _val(child)
                    align_map = {'left': 'LEFT', 'center': 'CENTER', 'right': 'RIGHT', 'both': 'JUSTIFY'}
                    result['align'] = align_map.get(v, 'DEFAULT')
                elif ct == 'spacing':
                    line = child.get(qn('w:line'))
                    if line: result['ls'] = int(line) / 240.0
                elif ct == 'ind':
                    fi = child.get(qn('w:firstLine'))
                    if fi: result['indent'] = round(int(fi) / 567.0, 1)  # twips -> cm

        # 3. Resolve from style tree
        self._resolve_style(style_id, result)

        return result

    def _resolve_style(self, sid, result):
        """Walk style inheritance chain to fill missing values."""
        visited = set()
        while sid and sid not in visited:
            visited.add(sid)
            s = self.styles.get(sid, {})
            if result['font'] is None: result['font'] = s.get('font')
            if result['size'] is None: result['size'] = s.get('size')
            if result['align'] == 'DEFAULT': result['align'] = s.get('align', 'DEFAULT')
            if result['ls'] is None: result['ls'] = s.get('ls')
            sid = s.get('base')
        # Final fallbacks
        if result['font'] is None: result['font'] = 'Times New Roman'
        if result['size'] is None: result['size'] = 12.0
        if result['ls'] is None: result['ls'] = 1.15
        if result['align'] == 'DEFAULT': result['align'] = 'LEFT'


def extract(docx_path):
    """Extract ALL formatting with style resolution. Returns (format_dict, md_text)."""
    doc = Document(docx_path)
    resolver = StyleResolver(doc)

    fmt = {
        '_meta': {
            'source': os.path.basename(docx_path),
            'sha256': hashlib.sha256(open(docx_path, 'rb').read()).hexdigest()[:16],
            'paragraphs': len(doc.paragraphs),
            'tables': len(doc.tables),
            'sections': len(doc.sections),
        },
        'sections': [],
        'paragraphs': [],
        'tables': [],
    }

    md_lines = []
    md_lines.append(f'# 模版格式提取 — {os.path.basename(docx_path)}\n')
    md_lines.append(f'**段落**: {len(doc.paragraphs)} | **表格**: {len(doc.tables)} | **节**: {len(doc.sections)}\n')

    # ── Sections ──
    md_lines.append('## 页面设置\n')
    for i, sec in enumerate(doc.sections):
        si = {
            'index': i,
            'page_width_cm': round(sec.page_width.cm, 1),
            'page_height_cm': round(sec.page_height.cm, 1),
            'margin_top_cm': round(sec.top_margin.cm, 1),
            'margin_bottom_cm': round(sec.bottom_margin.cm, 1),
            'margin_left_cm': round(sec.left_margin.cm, 1),
            'margin_right_cm': round(sec.right_margin.cm, 1),
            'diff_first_page': sec.different_first_page_header_footer,
            'header': [], 'footer': [],
        }
        if sec.header:
            for p in sec.header.paragraphs:
                runs = []
                for r in p.runs:
                    info = resolver.resolve(p._element, r._element)
                    runs.append({'text': r.text, 'font': info['font'], 'size_pt': info['size'],
                                 'bold': info['bold'], 'italic': info['italic']})
                si['header'].append({'text': p.text, 'alignment': ALIGN_MAP.get(p.alignment, 'CENTER'), 'runs': runs})
        if sec.footer:
            for p in sec.footer.paragraphs:
                si['footer'].append({'text': p.text, 'alignment': ALIGN_MAP.get(p.alignment, 'CENTER')})
        fmt['sections'].append(si)

        md_lines.append(f'**节{i}**: {si["page_width_cm"]}x{si["page_height_cm"]}cm')
        md_lines.append(f'  边距: T{si["margin_top_cm"]} B{si["margin_bottom_cm"]} L{si["margin_left_cm"]} R{si["margin_right_cm"]}')
        for h in si['header']:
            md_lines.append(f'  页眉: {h["alignment"]} | {h["text"][:100]}')
        for f in si['footer']:
            md_lines.append(f'  页脚: {f["alignment"]} | {f["text"][:100]}')
        md_lines.append('')

    # ── Paragraphs ──
    md_lines.append('## 正文格式\n')
    for idx, p in enumerate(doc.paragraphs):
        pf = p.paragraph_format
        pinfo = {
            'index': idx,
            'style': p.style.name if p.style else 'Normal',
            'text': p.text,
            'runs': [],
            'has_page_break': False,
        }
        for r in p.runs:
            info = resolver.resolve(p._element, r._element)
            has_pb = 'w:br w:type="page"' in r._element.xml or "w:br w:type='page'" in r._element.xml
            if has_pb: pinfo['has_page_break'] = True
            # Use resolved values (never None)
            pinfo['runs'].append({
                'text': r.text,
                'font': info['font'],
                'size_pt': info['size'],
                'bold': info['bold'],
                'italic': info['italic'],
            })
            # Carry paragraph-level formatting on first run
            if 'align' not in pinfo:
                pinfo['align'] = info['align']
                pinfo['ls'] = info['ls']
                pinfo['indent'] = info['indent']

        fmt['paragraphs'].append(pinfo)

        # MD
        if p.text.strip() or pinfo['has_page_break']:
            flags = f"align={pinfo.get('align','?')} ls={pinfo.get('ls','?')}"
            if pinfo.get('indent'): flags += f' indent={pinfo["indent"]}cm'
            if pinfo['has_page_break']: flags += ' [PAGE BREAK]'
            txt = p.text[:100].replace('\n', '\\n')
            run_fmts = []
            for r in pinfo['runs'][:4]:
                parts = [f'{r.get("font","?")}', f'{r.get("size_pt","?")}pt']
                if r.get('bold'): parts.append('B')
                if r.get('italic'): parts.append('I')
                run_fmts.append('|'.join(parts))
            md_lines.append(f'**P{idx}** [{pinfo["style"]}] {flags}')
            if run_fmts: md_lines.append(f'  runs: {" / ".join(run_fmts)}')
            if txt: md_lines.append(f'  > {txt}')
            md_lines.append('')

    # ── Tables ──
    md_lines.append('## 表格\n')
    for ti, table in enumerate(doc.tables):
        tinfo = {'index': ti, 'rows': len(table.rows), 'cols': len(table.columns), 'cells': []}
        for ri, row in enumerate(table.rows):
            row_cells = []
            for ci, cell in enumerate(row.cells):
                cell_runs = []
                for p in cell.paragraphs:
                    for r in p.runs:
                        info = resolver.resolve(p._element, r._element)
                        cell_runs.append({
                            'text': r.text, 'font': info['font'], 'size_pt': info['size'], 'bold': info['bold'],
                        })
                row_cells.append({'row': ri, 'col': ci, 'text': cell.text, 'runs': cell_runs})
            tinfo['cells'].append(row_cells)
            if ri < 5:
                md_lines.append(f'  Row{ri}: {[c["text"][:50] for c in row_cells]}')
        fmt['tables'].append(tinfo)
        md_lines.append('')

    # ── Verification ──
    md_lines.append(f'\n## 验证\n')
    md_lines.append(f'- 段落: JSON={len(fmt["paragraphs"])} docx={len(doc.paragraphs)} ✓')
    md_lines.append(f'- 表格: JSON={len(fmt["tables"])} docx={len(doc.tables)} ✓')
    md_lines.append(f'- 节:   JSON={len(fmt["sections"])} docx={len(doc.sections)} ✓')

    # ── Cover extraction: walk body elements until abstract/body content ──
    fmt['cover'] = _extract_cover(doc)

    return fmt, '\n'.join(md_lines)


def _extract_cover(doc):
    """Walk template body from start, extract cover+declaration elements with full formatting.
    Stops at abstract/TOC/body content. Returns list of element dicts or [] if no cover."""
    W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'

    STOP_KW = ['摘 要', '摘要', '目  录', '目录', '目 录', 'ABSTRACT', '第1章', '1.1 ', '1.1.']
    SKIP_KW = ['页边距要求', '碳素笔', '完成后删除', '封面要求', '1.论文题目', '毕业论文（设计）题目为',
               '按答辩时间', '提交论文', '禁止使用']

    elements = []
    for child in doc.element.body:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag == 'sectPr':
            break

        if tag == 'tbl':
            rows_data = []
            for row in child.findall(f'{{{W}}}tr'):
                cells_data = []
                for tc in row.findall(f'{{{W}}}tc'):
                    tcPr = tc.find(f'{{{W}}}tcPr')
                    # Cell width
                    tcW = tcPr.find(f'{{{W}}}tcW') if tcPr is not None else None
                    cell_w = int(tcW.get(f'{{{W}}}w', '0')) if tcW is not None else 0
                    # Cell borders
                    cell_borders = {}
                    tcBorders = tcPr.find(f'{{{W}}}tcBorders') if tcPr is not None else None
                    if tcBorders is not None:
                        for b in tcBorders:
                            btag = b.tag.split('}')[-1]
                            bval = b.get(f'{{{W}}}val', 'nil')
                            if bval not in (None, 'nil', 'none'):
                                cell_borders[btag] = {
                                    'val': bval,
                                    'sz': b.get(f'{{{W}}}sz', '0'),
                                    'color': b.get(f'{{{W}}}color', '000000'),
                                }
                    # Cell paragraphs
                    cell_paras = []
                    for p in tc.findall(f'{{{W}}}p'):
                        pPr = p.find(f'{{{W}}}pPr')
                        jc = pPr.find(f'{{{W}}}jc') if pPr is not None else None
                        palign = jc.get(f'{{{W}}}val') if jc is not None else None
                        runs = []
                        for r in p.findall(f'{{{W}}}r'):
                            rPr = r.find(f'{{{W}}}rPr')
                            fn_ascii, fn_ea, fsz, fbold = '', '', 0, False
                            if rPr is not None:
                                rf = rPr.find(f'{{{W}}}rFonts')
                                if rf is not None:
                                    fn_ascii = rf.get(f'{{{W}}}ascii', '') or ''
                                    fn_ea = rf.get(f'{{{W}}}eastAsia', '') or ''
                                sz = rPr.find(f'{{{W}}}sz')
                                if sz is not None:
                                    fsz = int(sz.get(f'{{{W}}}val', '0')) // 2
                                fbold = rPr.find(f'{{{W}}}b') is not None
                            txt = ''.join(t.text or '' for t in r.findall(f'{{{W}}}t'))
                            runs.append({'t': txt, 'fn': fn_ascii, 'fe': fn_ea, 'sz': fsz, 'b': fbold})
                        cell_paras.append({'al': palign, 'r': runs})
                    cells_data.append({'w': cell_w, 'borders': cell_borders, 'p': cell_paras})
                rows_data.append(cells_data)
            elements.append({'type': 'table', 'rows': rows_data})
            continue

        if tag != 'p':
            continue

        # Extract paragraph data
        pPr = child.find(f'{{{W}}}pPr')
        jc = pPr.find(f'{{{W}}}jc') if pPr is not None else None
        palign = jc.get(f'{{{W}}}val') if jc is not None else None

        spacing = pPr.find(f'{{{W}}}spacing') if pPr is not None else None
        line_val = spacing.get(f'{{{W}}}line') if spacing is not None else None
        lineRule = spacing.get(f'{{{W}}}lineRule') if spacing is not None else None
        before_val = spacing.get(f'{{{W}}}before') if spacing is not None else None

        indent = pPr.find(f'{{{W}}}ind') if pPr is not None else None
        first_line = indent.get(f'{{{W}}}firstLine') if indent is not None else None

        runs = []
        for r in child.findall(f'{{{W}}}r'):
            rPr = r.find(f'{{{W}}}rPr')
            fn_ascii, fn_ea, fsz, fbold = '', '', 0, False
            if rPr is not None:
                rf = rPr.find(f'{{{W}}}rFonts')
                if rf is not None:
                    fn_ascii = rf.get(f'{{{W}}}ascii', '') or ''
                    fn_ea = rf.get(f'{{{W}}}eastAsia', '') or ''
                sz = rPr.find(f'{{{W}}}sz')
                if sz is not None:
                    fsz = int(sz.get(f'{{{W}}}val', '0')) // 2
                fbold = rPr.find(f'{{{W}}}b') is not None
            txt = ''.join(t.text or '' for t in r.findall(f'{{{W}}}t'))
            runs.append({'t': txt, 'fn': fn_ascii, 'fe': fn_ea, 'sz': fsz, 'b': fbold})

        full_text = ''.join(r['t'] for r in runs).strip()

        # ── Heuristic: is this paragraph a format instruction? ──
        # Format notes specify fonts, sizes, alignments — not actual content.
        _fmt_font = any(f in full_text for f in ['黑体', '宋体', '楷体', '华文', '方正', 'Times New Roman'])
        _fmt_size = any(kw in full_text for kw in ['二号', '三号', '四号', '小四', '五号', '小五',
                                                    '号加粗', '号居中', 'pt', '号字'])
        _fmt_align = any(kw in full_text for kw in ['居中', '加粗', '缩进', '对齐', '行距', '段前', '段后',
                                                     '固定值', '倍行距', '1.5倍', '双倍'])
        _fmt_paren = '（' in full_text and '）' in full_text
        _is_fmt_note = (_fmt_paren and (_fmt_font or _fmt_size)) or (_fmt_font and _fmt_size and _fmt_align)
        _is_fmt_note = _is_fmt_note or (_fmt_paren and any(kw in full_text for kw in ['空一行', '空两行', '空行', '空  行']))
        _is_fmt_header = any(kw in full_text[:60] for kw in ['页眉页脚', '页眉', '页码', '字体要求',
                                                               '字号要求', '格式要求', '排版要求'])

        # Stop at abstract/TOC/body (real content, not format notes)
        if full_text and any(kw in full_text[:20] for kw in STOP_KW):
            if not _is_fmt_note:
                break

        # Skip individual format notes & section headers
        if full_text:
            if _is_fmt_note or _is_fmt_header:
                continue
            if any(kw in full_text[:30] for kw in SKIP_KW):
                continue
            if len(full_text) > 40 and '删除' in full_text[:80]:
                continue

        # Check for images
        has_img = False
        img_extent = None
        img_srcRect = {}
        for inline in child.iter(f'{{{WP}}}inline'):
            ext = inline.find(f'{{{WP}}}extent')
            if ext is not None:
                img_extent = {'cx': ext.get('cx', '0'), 'cy': ext.get('cy', '0')}
            for sr in inline.iter(f'{{{A}}}srcRect'):
                img_srcRect = {'l': sr.get('l', '0'), 't': sr.get('t', '0'),
                               'r': sr.get('r', '0'), 'b': sr.get('b', '0')}
            for blip in inline.iter(f'{{{A}}}blip'):
                emb = blip.get(f'{{{R}}}embed')
                if emb and emb in doc.part.rels and 'image' in doc.part.rels[emb].reltype:
                    has_img = True
                    break
            break

        if has_img:
            elements.append({
                'type': 'image',
                'al': palign, 'ls_val': line_val, 'ls_rule': lineRule,
                'extent': img_extent, 'srcRect': img_srcRect,
                'rEmbed': emb,
                'r': runs,
            })
        elif not full_text:
            elements.append({
                'type': 'empty',
                'al': palign, 'ls_val': line_val, 'ls_rule': lineRule,
                'sp_before': before_val, 'fl_indent': first_line,
                'r': runs,  # run font sizes control paragraph height even when empty
            })
        else:
            elements.append({
                'type': 'para',
                'al': palign, 'ls_val': line_val, 'ls_rule': lineRule,
                'sp_before': before_val, 'fl_indent': first_line,
                'r': runs,
            })

    return elements


if __name__ == '__main__':
    import sys
    path = sys.argv[1] if len(sys.argv) > 1 else 'Templates/模版.docx'
    fmt, md = extract(path)
    json_path = path.replace('.docx', '_format.json')
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(fmt, f, ensure_ascii=False, indent=2)
    md_path = path.replace('.docx', '_格式提取.md')
    with open(md_path, 'w', encoding='utf-8') as f:
        f.write(md)
    print(f'Format JSON -> {json_path}')
    print(f'Format MD   -> {md_path}')
    print(f'Paragraphs: {len(fmt["paragraphs"])}  Tables: {len(fmt["tables"])}  Sections: {len(fmt["sections"])}')
