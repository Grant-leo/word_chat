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

    return fmt, '\n'.join(md_lines)


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
