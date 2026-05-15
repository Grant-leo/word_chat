"""
script_generator.py
架构 = Acta 脚本骨架 (OOXML helpers + 三线表 + 交叉引用 + 分页)
参数 = 模板文字说明(优先) + format.json OOXML提取(补充)
内容 = content.json
"""
import json, os, re
from collections import Counter

# Chinese字号 -> pt  (ordered: check 小X号 before X号 to avoid false match)
SIZE_PATTERNS = [
    ('小初号', 36), ('初号', 42), ('小一号', 24), ('一号', 26),
    ('小二号', 18), ('二号', 22), ('小三号', 15), ('三号', 16),
    ('小四号', 12), ('四号', 14), ('小五号', 9), ('五号', 10.5),
    ('小初', 36), ('小一', 24), ('小二', 18), ('小三', 15),
    ('小四', 12), ('小五', 9),
]

def _parse_text_instructions(fmt):
    """Parse explicit format instructions from template text.
    These are the author's stated rules — higher priority than OOXML stats."""
    rules = {}  # key -> {size, font, bold, align, indent, ls, ...}

    all_text = '\n'.join(p.get('text', '') for p in fmt['paragraphs'])

    def _get_size(text, default=None):
        # SIZE_PATTERNS is ordered: "小三号" before "三号", "小四" before "四号"
        for name, pt in SIZE_PATTERNS:
            if name in text: return pt
        return default

    def _get_font(text, default=None):
        if '宋体' in text: return '宋体'
        if '黑体' in text: return '黑体'
        if '楷体' in text: return '楷体'
        if '微软雅黑' in text: return '微软雅黑'
        if 'Times New Roman' in text or '新罗马' in text: return 'Times New Roman'
        if '英文用Times New Roman' in text: return 'Times New Roman'
        return default

    def _get_align(text, default=None):
        if '居中' in text: return 'CENTER'
        if '左顶格' in text or '左对齐' in text: return 'LEFT'
        if '右对齐' in text: return 'RIGHT'
        if '两端对齐' in text: return 'JUSTIFY'
        return default

    # ── Heading levels ──
    for level_name, level_key in [('一级标题', 'h1'), ('二级标题', 'h2'), ('三级标题', 'h3')]:
        pattern = re.compile(rf'{level_name}[^。）]*(?:。|；|）)')
        matches = pattern.findall(all_text)
        # Prefer body format over TOC: body instructions are longer and contain '加粗'
        body_matches = [m for m in matches if '加粗' in m and '目录' not in m]
        use = body_matches if body_matches else matches
        # Take the longest match (most detailed instruction)
        combined = max(use, key=len) if use else ''
        if combined:
            rules[level_key] = {
                'size': _get_size(combined, 15 if level_key == 'h1' else 14 if level_key == 'h2' else 12),
                'font': _get_font(combined, 'Times New Roman'),
                'bold': '加粗' in combined,
                'align': _get_align(combined, 'CENTER' if level_key == 'h1' else 'LEFT'),
            }

    # ── Cover page — extract text from template tables, ZERO hardcoding ──
    cover_labels = []  # label texts for info table rows
    cover_title_text = ''
    cover_title_size = 36
    cover_title_font = '黑体'
    cover_label_font = '楷体'
    cover_label_size = 14

    # Get cover title: prefer the run with largest font in the paragraph
    for p in fmt['paragraphs'][:10]:
        txt = p['text'].strip()
        if not txt or len(txt) > 30:
            continue
        best_run = None
        for r in p.get('runs', []):
            if r.get('size_pt') and r['size_pt'] >= 22:
                if best_run is None or r['size_pt'] > best_run['size_pt']:
                    best_run = r
        if best_run:
            cover_title_text = txt
            cover_title_size = best_run['size_pt']
            cover_title_font = best_run.get('font') or '黑体'
            break

    # Get cover table labels from template's first two tables
    if len(fmt['tables']) >= 2:
        t0 = fmt['tables'][0]
        t1 = fmt['tables'][1]
        # Extract label font/size from FIRST cell of FIRST column only
        if t0.get('cells') and t0['cells'][0]:
            for r in t0['cells'][0][0].get('runs', []):
                if r.get('font'): cover_label_font = r['font']
                if r.get('size_pt'): cover_label_size = r['size_pt']
                break
        # Extract info labels from table1's first column (col 0 only)
        for row in t1.get('cells', []):
            for c in row:
                txt = c['text'].strip()
                # Exclude rows that are pure formatting notes
                if txt and not any(kw in txt for kw in ['行高', '居中', '1.5', '楷体四']):
                    cover_labels.append(txt)
                    break

    rules['cover'] = {
        'title_text': cover_title_text,
        'title_size': cover_title_size,
        'title_font': cover_title_font,
        'label_font': cover_label_font,
        'label_size': cover_label_size,
        'labels': cover_labels,
    }

    # ── Body text ──
    body_pattern = re.compile(r'正文[^。）]*(?:。|；|）)')
    body_matches = body_pattern.findall(all_text)
    # Prefer matches that specify BOTH line spacing AND indent/alignment
    best = [m for m in body_matches if '行距' in m and ('缩进' in m or '对齐' in m)]
    # Fall back to any match with line spacing
    if not best:
        best = [m for m in body_matches if '行距' in m]
    body_text = best[0] if best else (body_matches[0] if body_matches else '')
    if body_text:
        rules['body'] = {
            'size': _get_size(body_text, 12),
            'font': _get_font(body_text, 'Times New Roman'),
            'ls': 1.5 if '1.5倍' in body_text else (2.0 if '双倍' in body_text else None),
            'align': _get_align(body_text, 'JUSTIFY'),
            'indent': 0.74 if '缩进' in body_text and '2字符' in body_text else 0,
        }

    # ── Specific labels ──
    for label, key, default_size in [
        ('Abstract', 'abstract_label', 14),
        ('Key words', 'keywords_label', 14),
        ('References', 'references_label', 15),
        ('Contents', 'contents_label', 15),
        ('Acknowledgements', 'ack_label', 15),
    ]:
        pattern = re.compile(rf'{label}[^。）]*(?:。|；|）)')
        matches = pattern.findall(all_text)
        combined = ' '.join(matches)
        if combined:
            rules[key] = {
                'size': _get_size(combined, default_size),
                'bold': '加粗' in combined,
                'align': _get_align(combined, 'LEFT'),
            }

    # ── Chinese abstract page ──
    if '单独成页' in all_text:
        rules['chinese_abstract_page'] = True
    if '中文摘要内容必须与英文摘要完全对应' in all_text:
        rules['chinese_abstract_symmetric'] = True

    # ── Keywords ──
    kw_pattern = re.compile(r'关键词[^。）]*(?:。|；|）)')
    kw_matches = kw_pattern.findall(all_text)
    kw_text = ' '.join(kw_matches)
    if kw_text:
        rules['keywords'] = {
            'size': _get_size(kw_text, 12),
            'ls': 1.5 if '1.5倍' in kw_text else None,
            'separator': '分号' if '分号' in kw_text else 'comma',
        }

    # ── Appendix ──
    app_pattern = re.compile(r'Appendix[^。）]*(?:。|；|）)')
    app_matches = app_pattern.findall(all_text)
    app_text = ' '.join(app_matches)
    if app_text:
        rules['appendix'] = {
            'size': _get_size(app_text, 15),
            'bold': '加粗' in app_text,
            'align': _get_align(app_text, 'LEFT'),
        }

    # ── References count ──
    count_match = re.search(r'(\d+)篇', all_text)
    if count_match:
        rules['ref_count'] = int(count_match.group(1))
    eng_match = re.search(r'(\d+)篇英文', all_text)
    if eng_match:
        rules['ref_english_count'] = int(eng_match.group(1))

    # ── Cover table ──
    if '行高' in all_text and '0.9' in all_text:
        rules['cover_table_row_height'] = 0.9  # cm
    if '楷体' in all_text:
        rules['cover_font'] = '楷体'

    return rules


def _q(text):
    """Escape for Python single-quoted string."""
    return text.replace('\\', '\\\\').replace("'", "\\'").replace('\n', '\\n').replace('\r', '')


def _extract_params(fmt):
    """Extract ALL format parameters from format.json. No hardcoded defaults for key values."""
    P = {}  # params dict

    # ── Page ──
    s0 = fmt['sections'][0] if fmt['sections'] else {}
    P['page_w'] = s0.get('page_width_cm', 21.0)
    P['page_h'] = s0.get('page_height_cm', 29.7)
    P['mt'] = s0.get('margin_top_cm', 2.54)
    P['mb'] = s0.get('margin_bottom_cm', 2.54)
    P['ml'] = s0.get('margin_left_cm', 2.54)
    P['mr'] = s0.get('margin_right_cm', 2.54)

    # ── Header ──
    hdr = None
    if s0.get('header'):
        h0 = s0['header'][0]
        r0 = h0['runs'][0] if h0.get('runs') else {}
        hdr_text = h0.get('text', '')
        # Strip Chinese formatting notes like "（新罗马字体，五号加粗居中）"
        hdr_text = re.sub(r'（[^）]*[号字体新罗马粗细斜居左右顶格][^）]*）', '', hdr_text).strip()
        hdr = {
            'text': _q(hdr_text[:120]),
            'align': h0.get('alignment', 'RIGHT'),
            'font': r0.get('font') or 'Times New Roman',
            'size': r0.get('size_pt', 9),
            'bold': r0.get('bold', False),
            'italic': r0.get('italic', False),
        }
    P['header'] = hdr

    # ── Body text: sample from actual content paragraphs (not formatting notes) ──
    samples = []
    for p in fmt['paragraphs']:
        txt = p.get('text', '').strip()
        if len(txt) < 30:
            continue
        if txt.startswith('（') and any(k in txt[:60] for k in ['号', '行距', '缩进', '对齐', '空一行', '备注']):
            continue
        for r in p.get('runs', []):
            if r.get('size_pt') and r.get('font') and not r.get('bold'):
                samples.append({
                    'size': round(r['size_pt'], 1),
                    'font': r['font'],
                    'ls': p.get('line_spacing_val', 1.5),
                    'align': p.get('alignment', 'JUSTIFY'),
                    'indent': round(p.get('first_indent_cm') or 0, 1),
                })
                break

    if samples:
        P['body_size'] = Counter(s['size'] for s in samples).most_common(1)[0][0]
        P['body_font'] = Counter(s['font'] for s in samples).most_common(1)[0][0]
        P['body_ls']     = Counter(s['ls'] for s in samples).most_common(1)[0][0]
        raw_align = Counter(s['align'] for s in samples).most_common(1)[0][0]
        P['body_align']  = 'JUSTIFY' if raw_align == 'DEFAULT' else raw_align
        P['body_indent'] = Counter(s['indent'] for s in samples).most_common(1)[0][0]
    else:
        P['body_size'] = 12; P['body_font'] = 'Times New Roman'
        P['body_ls'] = 1.5; P['body_align'] = 'JUSTIFY'; P['body_indent'] = 0

    # ── CJK font: detect from template paragraphs that contain Chinese characters ──
    cjk_fonts = []
    for p in fmt['paragraphs']:
        txt = p.get('text', '').strip()
        if len(txt) < 20:
            continue
        if not any('一' <= c <= '鿿' for c in txt[:500]):
            continue
        for r in p.get('runs', []):
            fn = r.get('font', '')
            if fn and fn not in ('Times New Roman', 'Arial', 'Calibri', 'Consolas', 'Cambria'):
                cjk_fonts.append(fn)
                break
    P['cjk_font'] = Counter(cjk_fonts).most_common(1)[0][0] if cjk_fonts else '宋体'

    # ── Headings: bold sizes that appear >= 3 times ──
    bold_counts = Counter()
    bold_aligns = {}
    for p in fmt['paragraphs']:
        txt = p.get('text', '').strip()
        if not txt or len(txt) > 200:
            continue
        for r in p.get('runs', []):
            if r.get('bold') and r.get('size_pt') and r.get('size_pt') >= 12:
                sz = round(r['size_pt'], 1)
                bold_counts[sz] += 1
                if sz not in bold_aligns or bold_aligns[sz] == 'DEFAULT':
                    raw = p.get('alignment', 'LEFT')
                    bold_aligns[sz] = 'LEFT' if raw == 'DEFAULT' else raw
                break

    # Sort by SIZE descending (not frequency), so h1 > h2 > h3
    h_sizes = sorted([sz for sz, cnt in bold_counts.most_common() if cnt >= 3], reverse=True)
    P['h_levels'] = []
    for i, sz in enumerate(h_sizes[:3]):
        P['h_levels'].append({
            'level': i + 1,
            'size': sz,
            'align': bold_aligns.get(sz, 'CENTER' if i == 0 else 'LEFT'),
            'space_before': [12, 8, 6][i],
        })

    # ── Reference format: check if [N] style exists ──
    P['has_ref_bookmarks'] = any('bookmark' in str(p).lower() for p in fmt['paragraphs'])
    P['has_tables'] = any('Table' in (p.get('text', '') or '') for p in fmt['paragraphs'])

    # ── OVERRIDE with explicit text instructions (higher priority) ──
    text_rules = _parse_text_instructions(fmt)
    P['_text_rules'] = text_rules  # store for reference

    if 'h1' in text_rules:
        if P['h_levels']:
            P['h_levels'][0] = {**P['h_levels'][0], **text_rules['h1']}
        else:
            P['h_levels'].append({**text_rules['h1'], 'level': 1, 'space_before': 12})
    if 'h2' in text_rules:
        if len(P['h_levels']) > 1:
            P['h_levels'][1] = {**P['h_levels'][1], **text_rules['h2']}
        else:
            P['h_levels'].append({**text_rules['h2'], 'level': 2, 'space_before': 8})
    if 'h3' in text_rules:
        if len(P['h_levels']) > 2:
            P['h_levels'][2] = {**P['h_levels'][2], **text_rules['h3']}
        else:
            P['h_levels'].append({**text_rules['h3'], 'level': 3, 'space_before': 6})

    if 'body' in text_rules:
        b = text_rules['body']
        if b.get('size'): P['body_size'] = b['size']
        if b.get('font'): P['body_font'] = b['font']  # use text instruction font
        if b.get('ls'): P['body_ls'] = b['ls']
        if b.get('align'): P['body_align'] = b['align']
        if b.get('indent') is not None: P['body_indent'] = b['indent']

    return P


# ═══════════════════════════════════════════════════════════════
#  CODE GENERATION  (Acta architecture + format params + content)
# ═══════════════════════════════════════════════════════════════
def generate(format_json_path, content_json_path, output_dir, output_docx_name='最终论文.docx'):
    output_py_path = os.path.join(output_dir, 'build_generated.py')
    with open(format_json_path, encoding='utf-8') as f:
        fmt = json.load(f)
    with open(content_json_path, encoding='utf-8') as f:
        cnt = json.load(f)

    P = _extract_params(fmt)
    has_refs = len(cnt.get('references', [])) > 0
    has_images = any(s.get('images') for s in cnt.get('sections', []))
    img_dir = os.path.abspath(cnt['_meta'].get('images_dir', 'Inputs/figures')).replace('\\', '/')

    # ── Derived parameters (ALL from format.json, zero hardcoding) ──
    D = {}  # derived
    D['footer_size'] = round(P['body_size'] * 0.85, 1)
    D['ref_size'] = max(P['body_size'] - 2, 8)
    D['caption_size'] = max(P['body_size'] - 2, 8)
    D['ref_sep_size'] = max(P['body_size'] - 3, 7)
    # Usable page height in pt: (page_h - top - bottom) / 0.0352778, +0.5cm tolerance
    D['usable_pt'] = (P['page_h'] - P['mt'] - P['mb']) / 0.0352778 + 0.5
    # Image width: 55% of text width, capped at 4.2 inches (safe for A4)
    text_w = P['page_w'] - P['ml'] - P['mr']
    D['img_width'] = min(round(text_w * 0.55, 1), 4.2)
    # Reference hanging indent (cm)
    D['ref_indent'] = 1.27 if P['body_indent'] < 0.5 else round(P['body_indent'] + 0.5, 1)
    D['table_cell_size'] = max(P['body_size'] - 3, 8)
    D['footer_font'] = P['header']['font'] if P['header'] else P['body_font']

    L = []
    def l(s=''): L.append(s)

    # ═══ HEADER ═══
    l('"""')
    l(f'build_generated.py')
    l(f'模版: {fmt["_meta"]["source"]}  内容: {cnt["_meta"]["source"]}')
    l(f'运行: python build_generated.py')
    l('"""')
    l('from docx import Document')
    l('from docx.shared import Pt, Inches, Cm, RGBColor, Emu')
    l('from docx.enum.text import WD_ALIGN_PARAGRAPH')
    l('from docx.oxml.ns import qn')
    l('from docx.oxml import OxmlElement')
    if has_refs:
        l('from docx.opc.constants import RELATIONSHIP_TYPE as RT')
    l('import os')
    l('')
    l("BASE = os.path.dirname(os.path.abspath(__file__))")
    l(f"OUT = os.path.join(BASE, '{os.path.basename(output_docx_name)}')")
    l('')
    l('doc = Document()')
    l('')

    # ═══ PAGE SETUP ═══
    l('# ── Page setup ──')
    l('for sec in doc.sections:')
    l(f'    sec.page_width   = Cm({P["page_w"]})')
    l(f'    sec.page_height  = Cm({P["page_h"]})')
    l(f'    sec.top_margin    = Cm({P["mt"]})')
    l(f'    sec.bottom_margin = Cm({P["mb"]})')
    l(f'    sec.left_margin   = Cm({P["ml"]})')
    l(f'    sec.right_margin  = Cm({P["mr"]})')
    l('')
    l('# Footer: PAGE field code')
    l("sec = doc.sections[0]")
    l("footer = sec.footer; footer.is_linked_to_previous = False")
    l("fp = footer.paragraphs[0]; fp.alignment = WD_ALIGN_PARAGRAPH.CENTER")
    l(f"rf = fp.add_run(); rf.font.size = Pt({D['footer_size']}); rf.font.name = '{D['footer_font']}'")
    l('for tag, attrs in [("w:fldChar", {qn("w:fldCharType"): "begin"}),')
    l('                   ("w:instrText", {}),')
    l('                   ("w:fldChar", {qn("w:fldCharType"): "end"})]:')
    l('    el = OxmlElement(tag)')
    l('    for k, v in attrs.items(): el.set(k, v)')
    l('    if tag == "w:instrText":')
    l('        el.set(qn("xml:space"), "preserve"); el.text = " PAGE "')
    l('    rf._element.append(el)')
    l('')

    # ═══ HEADER ═══
    if P['header']:
        h = P['header']
        l('# Running header')
        l("hdr = sec.header; hdr.is_linked_to_previous = False")
        l(f"hp = hdr.paragraphs[0]; hp.alignment = WD_ALIGN_PARAGRAPH.{h['align']}")
        l(f"r = hp.add_run('{h['text']}')")
        l(f"r.font.size = Pt({h['size']}); r.font.name = '{h['font']}'")
        l(f"r.bold = {h['bold']}; r.italic = {h['italic']}")
        l('')

    # ═══ DEFAULT STYLE ═══
    l('# Default paragraph style')
    l('style = doc.styles["Normal"]')
    l(f"style.font.name = '{P['body_font']}'")
    l(f'style.font.size = Pt({P["body_size"]})')
    l(f'style.paragraph_format.line_spacing = {P["body_ls"]}')
    l('style.paragraph_format.space_after  = Pt(0)')
    l('style.paragraph_format.space_before = Pt(0)')
    l('')

    # ═══ HELPERS ═══
    l('# ═══════════════════ HELPERS ═══════════════════')
    l('')

    ls = P['body_ls']
    align = P['body_align']
    indent_pt = round(P['body_indent'] * 28.35) if P['body_indent'] > 0 else 0

    cjk = P.get('cjk_font', '宋体')
    need_eastAsia = (P['body_font'] != cjk)

    # body()
    l('def body(text, first_indent=True, comment=None):')
    l('    p = doc.add_paragraph()')
    l(f'    p.alignment = WD_ALIGN_PARAGRAPH.{align}')
    l(f'    pf = p.paragraph_format; pf.line_spacing = {ls}')
    if indent_pt:
        l('    if first_indent:')
        l(f'        pf.first_line_indent = Pt({indent_pt})')
    l('    r = p.add_run(text)')
    l(f"    r.font.name = '{P['body_font']}'; r.font.size = Pt({P['body_size']})")
    if need_eastAsia:
        l(f'    rp = r._element.get_or_add_rPr()')
        l(f'    rf = rp.find(qn("w:rFonts"))')
        l(f'    if rf is None: rf = OxmlElement("w:rFonts"); rp.insert(0, rf)')
        l(f'    rf.set(qn("w:eastAsia"), "{cjk}")')
    l('    if comment:')
    l('        _cc.add(p, comment)')
    l('    return p')
    l('')

    # headings
    for hl in P['h_levels']:
        n = hl['level']
        l(f'def heading{n}(text, comment=None):')
        l('    p = doc.add_paragraph()')
        l(f'    p.alignment = WD_ALIGN_PARAGRAPH.{hl["align"]}')
        l(f'    pf = p.paragraph_format; pf.line_spacing = {ls}')
        l(f'    pf.space_before = Pt({hl["space_before"]})')
        l('    r = p.add_run(text); r.bold = True')
        l(f"    r.font.name = '{P['body_font']}'; r.font.size = Pt({hl['size']})")
        if need_eastAsia:
            l(f'    rp = r._element.get_or_add_rPr()')
            l(f'    rf = rp.find(qn("w:rFonts"))')
            l(f'    if rf is None: rf = OxmlElement("w:rFonts"); rp.insert(0, rf)')
            l(f'    rf.set(qn("w:eastAsia"), "{cjk}")')
        l('    if comment:')
        l('        _cc.add(p, comment)')
        l('    return p')
        l('')

    # ═══ TOC (Table of Contents) ═══
    l('# ── Table of Contents ──')
    l('def insert_toc(doc, title="目录"):')
    l('    p = doc.add_paragraph()')
    l('    p.alignment = WD_ALIGN_PARAGRAPH.CENTER')
    l('    r = p.add_run(title)')
    l('    r.bold = True')
    l(f"    r.font.size = Pt({P['h_levels'][0]['size'] if P['h_levels'] else 15})")
    l(f"    r.font.name = '{P['body_font']}'")
    l('    p2 = doc.add_paragraph()')
    l('    r2 = p2.add_run()')
    l('    r2.font.size = Pt(10)  # small hint text')
    l('    r2.font.name = "Times New Roman"')
    l('    r2.font.color.rgb = RGBColor(128,128,128)')
    l('    r2.add_text("（在 Word/WPS 中右键此处 → 更新域 → 更新整个目录）")')
    l('    # TOC field code')
    l('    tp = doc.add_paragraph()')
    l('    tr = tp.add_run()')
    l('    fld_begin = OxmlElement("w:fldChar")')
    l('    fld_begin.set(qn("w:fldCharType"), "begin")')
    l('    tr._element.append(fld_begin)')
    l('    instr = OxmlElement("w:instrText")')
    l('    instr.set(qn("xml:space"), "preserve")')
    l("    instr.text = ' TOC \\\\o \"1-3\" \\\\h \\\\z \\\\u '")
    l('    tr._element.append(instr)')
    l('    fld_sep = OxmlElement("w:fldChar")')
    l('    fld_sep.set(qn("w:fldCharType"), "separate")')
    l('    tr._element.append(fld_sep)')
    l('    tr2 = tp.add_run("（在 Word 中右键更新目录）")')
    l('    tr2.font.size = Pt(10)')
    l('    tr2.font.color.rgb = RGBColor(128,128,128)')
    l('    fld_end = OxmlElement("w:fldChar")')
    l('    fld_end.set(qn("w:fldCharType"), "end")')
    l('    tr2._element.append(fld_end)')
    l('')
    l('# Uncomment the next line to insert TOC after cover page:')
    l('# insert_toc(doc)')
    l('')

    # ═══ CROSS-REFERENCES (Acta architecture) ═══
    if has_refs:
        l('# ── Cross-reference system ──')
        l('ref_bookmarks = {}')
        l('')
        l('def B_ref(text, ref_nums):')
        l('    p = doc.add_paragraph()')
        l(f'    p.alignment = WD_ALIGN_PARAGRAPH.{align}')
        l(f'    pf = p.paragraph_format; pf.line_spacing = {ls}')
        l('    r = p.add_run(text)')
        l(f"    r.font.name = '{P['body_font']}'; r.font.size = Pt({P['body_size']})")
        if need_eastAsia:
            l(f'    rp = r._element.get_or_add_rPr()')
            l(f'    rf = rp.find(qn("w:rFonts"))')
            l(f'    if rf is None: rf = OxmlElement("w:rFonts"); rp.insert(0, rf)')
            l(f'    rf.set(qn("w:eastAsia"), "{cjk}")')
        l('    for i, n in enumerate(ref_nums):')
        l("        if i > 0:")
        l(f"            s = p.add_run(', '); s.font.superscript = True; s.font.size = Pt({D['ref_sep_size']})")
        l("        _add_ref(p, n, f'[{{n}}]')")
        l('    return p')
        l('')
        l('def _add_ref(paragraph, ref_num, display_text):')
        l('    bm = f"_Ref{{ref_num}}"; ref_bookmarks[ref_num] = bm')
        l('    hl = paragraph._element.makeelement(qn("w:hyperlink"), {')
        l('        qn("w:anchor"): bm, qn("w:history"): "1"})')
        l('    nr = paragraph._element.makeelement(qn("w:r"), {})')
        l('    rpr = paragraph._element.makeelement(qn("w:rPr"), {})')
        l('    for tag, attrs in [("w:vertAlign", {qn("w:val"): "superscript"}),')
        l('                        ("w:sz", {qn("w:val"): "18"}),')
        l('                        ("w:color", {qn("w:val"): "0000FF"})]:')
        l('        e = OxmlElement(tag)')
        l('        for k, v in attrs.items(): e.set(k, v)')
        l('        rpr.append(e)')
        l('    nr.append(rpr)')
        l('    t = OxmlElement("w:t"); t.set(qn("xml:space"), "preserve")')
        l('    t.text = display_text; nr.append(t); hl.append(nr)')
        l('    paragraph._element.append(hl)')
        l('')

    # ═══ THREE-LINE TABLE ═══
    if P['has_tables']:
        l('# ── Three-line table ──')
        l('def _rm_tbl_borders(table):')
        l('    tbl = table._tbl; tblPr = tbl.tblPr')
        l('    if tblPr is None:')
        l('        tblPr = OxmlElement("w:tblPr"); tbl.insert(0, tblPr)')
        l('    for old in tblPr.findall(qn("w:tblBorders")): tblPr.remove(old)')
        l('')
        l('def _set_tc(cell, top=None, bottom=None):')
        l('    tcPr = cell._tc.get_or_add_tcPr()')
        l('    for old in tcPr.findall(qn("w:tcBorders")): tcPr.remove(old)')
        l('    tcB = OxmlElement("w:tcBorders"); tcPr.append(tcB)')
        l('    def _b(pos, val="nil", sz="0"):')
        l('        b = OxmlElement(f"w:{{pos}}");')
        l('        b.set(qn("w:val"), val); b.set(qn("w:sz"), str(sz))')
        l('        b.set(qn("w:space"), "0"); b.set(qn("w:color"), "000000")')
        l('        tcB.append(b)')
        l('    _b("top", *(top or ("nil","0"))); _b("bottom", *(bottom or ("nil","0")))')
        l('    _b("left","nil"); _b("right","nil")')
        l('')
        l('def three_line_table(table):')
        l('    _rm_tbl_borders(table); n = len(table.rows)')
        l('    if n == 0: return')
        l('    for ri, row in enumerate(table.rows):')
        l('        for cell in row.cells:')
        l('            if ri == 0 and n == 1:')
        l('                _set_tc(cell, top=("single","12"), bottom=("single","12"))')
        l('            elif ri == 0:')
        l('                _set_tc(cell, top=("single","12"), bottom=("single","4"))')
        l('            elif ri == 1:')
        l('                _set_tc(cell, top=("single","4"))')
        l('            elif ri == n - 1:')
        l('                _set_tc(cell, bottom=("single","12"))')
        l('            else: _set_tc(cell)')
        l('')
        l(f'def C(table, row, col, text, bold=False, size={D["table_cell_size"]}):')
        l('    cell = table.rows[row].cells[col]; cell.text = ""')
        l('    r = cell.paragraphs[0].add_run(text); r.bold = bold')
        l(f"    r.font.size = Pt(size); r.font.name = '{P['body_font']}'")
        if need_eastAsia:
            l(f'    rp = r._element.get_or_add_rPr()')
            l(f'    rf = rp.find(qn("w:rFonts"))')
            l(f'    if rf is None: rf = OxmlElement("w:rFonts"); rp.insert(0, rf)')
            l(f'    rf.set(qn("w:eastAsia"), "{cjk}")')
        l('    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER')
        l('')

    # ═══ IMAGE INSERT ═══
    if has_images:
        l('# ── Image insert ──')
        l(f"FIG_DIR = r'{img_dir}'")
        l('')
        l('def insert_figure(path, width_inches, fig_num, caption_text):')
        l('    p = doc.add_paragraph()')
        l('    p.alignment = WD_ALIGN_PARAGRAPH.CENTER')
        l('    if os.path.exists(path):')
        l('        doc.add_picture(path, width=Inches(width_inches))')
        l('        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER')
        l('        cap = doc.add_paragraph()')
        l('        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER')
        l(f'        r = cap.add_run(f"Fig. {{fig_num}}. ")')
        l(f'        r.font.size = Pt({D["caption_size"]}); r.font.name = "{P["body_font"]}"; r.bold = True')
        if need_eastAsia:
            l(f'        rp = r._element.get_or_add_rPr()')
            l(f'        rf = rp.find(qn("w:rFonts"))')
            l(f'        if rf is None: rf = OxmlElement("w:rFonts"); rp.insert(0, rf)')
            l(f'        rf.set(qn("w:eastAsia"), "{cjk}")')
        l(f'        r2 = cap.add_run(caption_text)')
        l(f'        r2.font.size = Pt({D["caption_size"]}); r2.font.name = "{P["body_font"]}"; r2.italic = True')
        if need_eastAsia:
            l(f'        rp = r2._element.get_or_add_rPr()')
            l(f'        rf = rp.find(qn("w:rFonts"))')
            l(f'        if rf is None: rf = OxmlElement("w:rFonts"); rp.insert(0, rf)')
            l(f'        rf.set(qn("w:eastAsia"), "{cjk}")')
        l('')

    # ═══ PAGINATION ═══
    # Font metrics: avg_char_width / font_size ratio (standard typographic constants)
    _FONT_RATIOS = {'Times New Roman': 0.42, 'Arial': 0.45, 'Consolas': 0.60,
                    '宋体': 1.0, '黑体': 1.0, '楷体': 1.0, '微软雅黑': 0.95}
    _ratio = next((v for k, v in _FONT_RATIOS.items() if k in str(P['body_font'])), 0.42)
    _cjk_ratio = next((v for k, v in _FONT_RATIOS.items() if k in str(P.get('cjk_font', '宋体'))), 1.0)
    text_w_pt = (P['page_w'] - P['ml'] - P['mr']) / 2.54 * 72
    avg_char_w = P['body_size'] * _ratio
    cpl = int(text_w_pt / avg_char_w)
    cpl_cjk = int(text_w_pt / (P['body_size'] * _cjk_ratio))
    usable_cm = P['page_h'] - P['mt'] - P['mb']  # exact theoretical, no fudge needed

    l(f'# ── A4 pagination (cpl={cpl}, cpl_cjk={cpl_cjk}, usable={usable_cm:.1f}cm) ──')
    l("M_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'")
    l('def _et(el): return el.tag.split("}")[-1]')
    l('def _is_cjk(c):')
    l("    return '\\u4e00' <= c <= '\\u9fff' or '\\u3400' <= c <= '\\u4dbf' or '\\uf900' <= c <= '\\ufaff'")
    l('def _formula_h(el, sz):')
    l('    """Estimate formula height from OOXML math structure."""')
    l("    n_frac = len(el.findall(f'.//{{{M_NS}}}f'))")
    l("    n_nary = len(el.findall(f'.//{{{M_NS}}}nary'))")
    l("    n_rad  = len(el.findall(f'.//{{{M_NS}}}rad'))")
    l("    n_d    = len(el.findall(f'.//{{{M_NS}}}d'))")
    l('    return sz * 1.5 * (1.0')
    l('        + min(n_frac, 4) * 0.8')
    l('        + min(n_nary, 4) * 0.7')
    l('        + min(n_rad,  4) * 0.6')
    l('        + min(n_d,    4) * 0.4)')
    l('def _ph(pe):')
    l('    txt=""; sz=12.0; ih=0; fh=0')
    l('    for re in pe:')
    l('        tag=_et(re)')
    l('        if tag=="oMathPara":')
    l('            fh=max(fh, _formula_h(re, sz))')
    l('        elif tag=="oMath":')
    l('            fh=max(fh, _formula_h(re, sz) * 0.65)')
    l('        elif tag!="r":')
    l('            continue')
    l('        else:')
    l('            for c in re:')
    l('                ct=_et(c)')
    l('                if ct=="t": txt+=(c.text or "")')
    l('                elif ct=="rPr":')
    l('                    for p in c:')
    l('                        if _et(p)=="sz": sz=max(sz,float(p.get(qn("w:val"),"24"))/2.0)')
    l('                elif ct=="drawing":')
    l('                    for il in c:')
    l('                        for ex in il:')
    l('                            if _et(ex)=="extent": ih=max(ih,int(ex.get("cy","0")))')
    l('    if fh>0:')
    l('        return max(fh, sz*4.5)')
    l('    if ih>0: return ih/12700.0')
    l(f'    if not txt.strip(): return sz*{P["body_ls"]}')
    l(f'    has_cjk = any(_is_cjk(c) for c in txt)')
    l(f'    use_cpl = {cpl_cjk} if has_cjk else {cpl}')
    l(f'    lines = (len(txt)+use_cpl//2)//use_cpl')
    l(f'    return max(1,lines)*sz*{P["body_ls"]}')
    l('def _th(te):')
    l('    n=sum(1 for r in te if _et(r)=="tr")')
    l(f'    return {P["body_size"]*2.5:.0f}+(n-1)*{P["body_size"]*1.8:.0f} if n>1 else {P["body_size"]*2.5:.0f}')
    l('def _ispb(pe):')
    l('    for r in pe:')
    l('        if _et(r)=="r":')
    l('            for br in r:')
    l('                if _et(br)=="br" and br.get(qn("w:type"))=="page": return True')
    l('    return False')
    l(f'# Pagination: chars_per_line computed from page geometry')
    l('def paginate_a4(doc):')
    # Acta-tested formula: theoretical + margin for estimation error
    l(f'    body=doc.element.body; usable={usable_cm}/0.0352778; acc=0.0; ins=[]')
    l('    for idx,child in enumerate(list(body)):')
    l('        tag=_et(child)')
    l('        if tag=="sectPr" or tag not in("p","tbl"): continue')
    l('        if tag=="p" and _ispb(child): acc=0.0; continue')
    l('        h=_th(child) if tag=="tbl" else _ph(child)')
    l('        if acc>0 and acc+h>usable: ins.append(idx); acc=h')
    l('        else: acc+=h')
    l('    for pos in reversed(ins):')
    l('        pb=OxmlElement("w:p"); r=OxmlElement("w:r"); br=OxmlElement("w:br")')
    l('        br.set(qn("w:type"),"page"); r.append(br); pb.append(r)')
    l('        body.insert(pos,pb)')
    l('    return len(ins)')
    l('')

    # ── Always copy comment_utils.py (used by body/heading helpers) ──
    import shutil
    pipeline_dir = os.path.dirname(os.path.abspath(__file__))
    comment_src = os.path.join(pipeline_dir, 'comment_utils.py')
    if os.path.exists(comment_src):
        shutil.copy2(comment_src, os.path.join(output_dir, 'comment_utils.py'))
    l('')
    l('import sys as _sys; _sys.path.insert(0, BASE)')
    l('from comment_utils import CommentCollector')
    l('')
    l('# ── Comment collector (optional: use comment="text" in body/heading) ──')
    l('_cc = CommentCollector()')
    l('')

    # ── Formula support ──
    has_formulas = any(
        isinstance(p, dict) and p.get('math')
        for s in cnt.get('sections', [])
        for p in s.get('paragraphs', [])
    )
    M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
    if has_formulas:
        l('')
        l('from lxml import etree')
        l('')
        l('# ── Formula utilities (reusable, permanent) ──')
        l(f"M = '{M}'")
        l('')
        l('def _strip_math_ns(elem):')
        l('    """Remove redundant xmlns declarations from math element so WPS can parse it."""')
        l('    KEEP = {')
        l('        "http://schemas.openxmlformats.org/officeDocument/2006/math",')
        l('        "http://schemas.openxmlformats.org/wordprocessingml/2006/main",')
        l('    }')
        l('    for key in list(elem.attrib):')
        l('        if key.startswith("xmlns:") and elem.attrib[key] not in KEEP:')
        l('            del elem.attrib[key]')
        l('    for child in elem:')
        l('        _strip_math_ns(child)')
        l('')
        l('def formula_text(xml_str):')
        l('    """Extract plain text from OOXML math formula."""')
        l('    parts = []')
        l('    for t in etree.fromstring(xml_str).iter(f"{{{M}}}t"):')
        l('        if t.text: parts.append(t.text)')
        l("    return ''.join(parts)")
        l('')
        l('def formula_remove(xml_str, target_text):')
        l('    """Remove math child elements containing target_text. Returns modified XML string."""')
        l('    root = etree.fromstring(xml_str)')
        l('    omath = root.find(f".//{{{M}}}oMath")')
        l('    if omath is not None:')
        l('        for child in list(omath):')
        l('            child_text = "".join(t.text or "" for t in child.iter(f"{{{M}}}t"))')
        l('            if target_text in child_text:')
        l('                omath.remove(child)')
        l('    return etree.tounicode(root, with_tail=False)')
        l('')
        l('def formula_replace(xml_str, old_text, new_text):')
        l('    """Replace text in all m:t elements. Returns modified XML string."""')
        l('    root = etree.fromstring(xml_str)')
        l('    for t in root.iter(f"{{{M}}}t"):')
        l('        if t.text and old_text in t.text:')
        l('            t.text = t.text.replace(old_text, new_text)')
        l('    return etree.tounicode(root, with_tail=False)')
        l('')
        l('def formula_build_matrix(cells, cols=2, brackets="[]"):')
        l('    """Build WPS-compliant matrix OOXML. cells=list of text strings, cols=column count.')
        l('    brackets: \"[]\", \"()\", \"||\", \"\"(none). Returns XML string ready for body_with_formula."""')
        l('    left, right = brackets if len(brackets) == 2 else (brackets, brackets)')
        l('    omath = etree.Element(f"{{{M}}}oMath")')
        l('    d = etree.SubElement(omath, f"{{{M}}}d")')
        l('    dPr = etree.SubElement(d, f"{{{M}}}dPr")')
        l('    if left:')
        l('        beg = etree.SubElement(dPr, f"{{{M}}}begChr")')
        l('        beg.set(f"{{{M}}}val", left)')
        l('    if right:')
        l('        end = etree.SubElement(dPr, f"{{{M}}}endChr")')
        l('        end.set(f"{{{M}}}val", right)')
        l('    m = etree.SubElement(d, f"{{{M}}}m")')
        l('    mPr = etree.SubElement(m, f"{{{M}}}mPr")')
        l('    mcs = etree.SubElement(mPr, f"{{{M}}}mcs")')
        l('    mc = etree.SubElement(mcs, f"{{{M}}}mc")')
        l('    mcPr = etree.SubElement(mc, f"{{{M}}}mcPr")')
        l('    count = etree.SubElement(mcPr, f"{{{M}}}count")')
        l('    count.set(f"{{{M}}}val", str(cols))')
        l('    etree.SubElement(mc, f"{{{M}}}mcPr")')
        l('    rows = len(cells) // cols')
        l('    idx = 0')
        l('    for _ in range(rows):')
        l('        mr = etree.SubElement(m, f"{{{M}}}mr")')
        l('        for _ in range(cols):')
        l('            e = etree.SubElement(mr, f"{{{M}}}e")')
        l('            r = etree.SubElement(e, f"{{{M}}}r")')
        l('            etree.SubElement(r, f"{{{M}}}rPr")  # WPS requires this, even if empty')
        l('            t = etree.SubElement(r, f"{{{M}}}t")')
        l('            t.text = cells[idx] if idx < len(cells) else ""')
        l('            idx += 1')
        l('    return etree.tounicode(omath, with_tail=False)')
        l('')

        # ── Copy latex_omath.py to output ──
        latex_src = os.path.join(pipeline_dir, 'latex_omath.py')
        if os.path.exists(latex_src):
            shutil.copy2(latex_src, os.path.join(output_dir, 'latex_omath.py'))
        l('')
        l('from latex_omath import latex_to_omath, body_latex, formula_text_from_omath')
        l('')
        l('# ── Formula display helper ──')
        l('def body_with_formula(text, math_xml_list):')
        l('    p = doc.add_paragraph()')
        l(f'    p.alignment = WD_ALIGN_PARAGRAPH.{align}')
        l(f'    pf = p.paragraph_format; pf.line_spacing = {ls}')
        if indent_pt:
            l('    pf.first_line_indent = Pt(' + str(indent_pt) + ')')
        l('    if text.strip():')
        l('        r = p.add_run(text)')
        l(f"        r.font.name = '{P['body_font']}'; r.font.size = Pt({P['body_size']})")
        if need_eastAsia:
            l('        rp = r._element.get_or_add_rPr()')
            l('        rf = rp.find(qn("w:rFonts"))')
            l('        if rf is None: rf = OxmlElement("w:rFonts"); rp.insert(0, rf)')
            l(f'        rf.set(qn("w:eastAsia"), "{cjk}")')
        l('    for xml_str in math_xml_list:')
        l('        math_el = etree.fromstring(xml_str)')
        l('        _strip_math_ns(math_el)  # remove redundant xmlns for WPS compatibility')
        l('        p._element.append(math_el)')
        l('    return p')
        l('')

    # ═══ CONTENT ═══
    l('# ═══════════════════ CONTENT ═══════════════════')
    l('')

    # ── Cover page — ALL from template body analysis ──
    cover = P.get('_text_rules', {}).get('cover')
    bf = P['body_font']
    if cover and cover.get('labels'):
        cv = cover
        labels = cv['labels']
        ti = cnt.get('title_info', {})
        paper_title = _q(ti.get('title_cn', 'Paper Title'))

        # Read template docx to count exact empty-paragraph gaps between cover elements
        template_path = os.path.join(TEMPLATE_DIR if 'TEMPLATE_DIR' in dir() else 'Templates',
                                     fmt['_meta']['source'])
        from docx import Document as _Doc
        try:
            _tpl = _Doc(template_path)
            _body = _tpl.element.body
            _gaps = []  # consecutive empty paragraph counts between non-empty elements
            _empty_run = 0
            for _ch in _body:
                _tag = _ch.tag.split('}')[-1]
                if _tag == 'p':
                    _txt = ''
                    for _r in _ch:
                        for _t in _r:
                            if _t.tag.split('}')[-1] == 't':
                                _txt += (_t.text or '')
                    if not _txt.strip():
                        _empty_run += 1
                    else:
                        if _empty_run > 0:
                            _gaps.append(_empty_run)
                        _empty_run = 0
                elif _tag == 'tbl':
                    if _empty_run > 0:
                        _gaps.append(_empty_run)
                    _empty_run = 0
            if _empty_run > 0:
                _gaps.append(_empty_run)
        except Exception:
            _gaps = [6, 4, 2, 0, 1]  # fallback from template analysis

        # _gaps are the empty paragraph counts between cover elements in order:
        # [before_title, after_title, between_tables, after_info_table, before_date]
        g = _gaps + [1] * (5 - len(_gaps))  # pad if needed
        g0, g1, g2, g3, g4 = g[0], g[1], g[2], g[3], g[4]

        l('# ═══════════════════ COVER ═══════════════════')
        l('from docx.oxml import OxmlElement as _OE')
        l('from docx.oxml.ns import qn as _qn')
        l(f'for _ in range({g0}): doc.add_paragraph()')
        ct_text = _q(cv.get('title_text', ''))
        if ct_text:
            l(f'p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER')
            l(f"r = p.add_run('{ct_text}'); r.bold = True")
            l(f'r.font.size = Pt({cv["title_size"]}); r.font.name = "{cv["title_font"]}"')
            # Set east-asia font so WPS renders CJK correctly
            cv_font_name = cv["title_font"]
            l(f'if "{cv_font_name}" != "Times New Roman":')
            l(f'    rp = r._element.get_or_add_rPr()')
            l(f'    rf = rp.find(_qn("w:rFonts"))')
            l(f'    if rf is None: rf = _OE("w:rFonts"); rp.insert(0, rf)')
            l(f'    rf.set(_qn("w:eastAsia"), "{cv_font_name}")')
        l(f'for _ in range({g1}): doc.add_paragraph()')
        l(f't_title = doc.add_table(rows=1, cols=2)')
        l(f"r = t_title.rows[0].cells[0].paragraphs[0].add_run('{_q(labels[0])}')")
        l(f'r.font.size = Pt({cv["label_size"]}); r.font.name = "{cv["label_font"]}"; r.bold = True')
        l(f'if "{cv["label_font"]}" != "Times New Roman":')
        l(f'    rp = r._element.get_or_add_rPr()')
        l(f'    rf = rp.find(_qn("w:rFonts"))')
        l(f'    if rf is None: rf = _OE("w:rFonts"); rp.insert(0, rf)')
        l(f'    rf.set(_qn("w:eastAsia"), "{cv["label_font"]}")')
        l(f"r = t_title.rows[0].cells[1].paragraphs[0].add_run('{paper_title}')")
        l(f'r.font.size = Pt({cv["label_size"]}); r.font.name = "{bf}"')
        if need_eastAsia:
            l(f'rp = r._element.get_or_add_rPr()')
            l(f'rf = rp.find(_qn("w:rFonts"))')
            l(f'if rf is None: rf = _OE("w:rFonts"); rp.insert(0, rf)')
            l(f'rf.set(_qn("w:eastAsia"), "{cjk}")')
        l('for t in [t_title]:')
        l('    for row in t.rows:')
        l('        tr = row._tr; trPr = tr.find(_qn("w:trPr"))')
        l('        if trPr is None: trPr = _OE("w:trPr"); tr.insert(0, trPr)')
        l('        th = _OE("w:trHeight"); th.set(_qn("w:val"), "510")')
        l('        th.set(_qn("w:hRule"), "atLeast"); trPr.append(th)')
        l(f'for _ in range({g2}): doc.add_paragraph()')
        l(f't_info = doc.add_table(rows={len(labels)-1}, cols=2)')
        l(f'info_labels = [')
        for lb in labels[1:]:
            l(f"    '{_q(lb)}',")
        l(']')
        l('for i, label in enumerate(info_labels):')
        l(f'    r = t_info.rows[i].cells[0].paragraphs[0].add_run(label)')
        l(f'    r.font.size = Pt({cv["label_size"]}); r.font.name = "{cv["label_font"]}"; r.bold = True')
        l(f'    if "{cv["label_font"]}" != "Times New Roman":')
        l(f'        rp = r._element.get_or_add_rPr()')
        l(f'        rf = rp.find(_qn("w:rFonts"))')
        l(f'        if rf is None: rf = _OE("w:rFonts"); rp.insert(0, rf)')
        l(f'        rf.set(_qn("w:eastAsia"), "{cv["label_font"]}")')
        l('    tr = t_info.rows[i]._tr; trPr = tr.find(_qn("w:trPr"))')
        l('    if trPr is None: trPr = _OE("w:trPr"); tr.insert(0, trPr)')
        l('    th = _OE("w:trHeight"); th.set(_qn("w:val"), "510")')
        l('    th.set(_qn("w:hRule"), "atLeast"); trPr.append(th)')
        l(f'for _ in range({g3}): doc.add_paragraph()')
        l(f'p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER')
        l(f"r = p.add_run('（完成时间按照答辩时间填写）'); r.bold = True")
        l(f'r.font.size = Pt({cv["label_size"]}); r.font.name = "{bf}"')
        if need_eastAsia:
            l(f'rp = r._element.get_or_add_rPr()')
            l(f'rf = rp.find(_qn("w:rFonts"))')
            l(f'if rf is None: rf = _OE("w:rFonts"); rp.insert(0, rf)')
            l(f'rf.set(_qn("w:eastAsia"), "{cjk}")')
        l(f'for _ in range({g4}): doc.add_paragraph()')
        l(f'p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER')
        l(f"r = p.add_run('年  月  日')")
        l(f'r.font.size = Pt({cv["label_size"]}); r.font.name = "{bf}"')
        if need_eastAsia:
            l(f'rp = r._element.get_or_add_rPr()')
            l(f'rf = rp.find(_qn("w:rFonts"))')
            l(f'if rf is None: rf = _OE("w:rFonts"); rp.insert(0, rf)')
            l(f'rf.set(_qn("w:eastAsia"), "{cjk}")')
        l('')
        l('doc.add_page_break()')
        l('')

    # ── Title ──
    ti = cnt.get('title_info', {})
    title_text = ti.get('title_cn', '')
    if title_text:
        l(f'# Title')
        l('p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER')
        l(f"r = p.add_run('{_q(title_text)}'); r.bold = True")
        if P['h_levels']:
            l(f'r.font.size = Pt({P["h_levels"][0]["size"]})')
        l(f"r.font.name = '{P['body_font']}'")
        if need_eastAsia:
            l(f'rp = r._element.get_or_add_rPr()')
            l(f'rf = rp.find(qn("w:rFonts"))')
            l(f'if rf is None: rf = OxmlElement("w:rFonts"); rp.insert(0, rf)')
            l(f'rf.set(qn("w:eastAsia"), "{cjk}")')
        l('')
    l('')

    # Sections
    fig_num = 0
    for sec in cnt.get('sections', []):
        h = sec.get('heading', '').strip()
        lv = sec.get('level', 0)
        if h:
            safe = _q(h)
            if lv >= 1 and lv <= 3:
                l(f'heading{lv}("{safe}")')
            else:
                l(f'body("{safe}", first_indent=False)')
            l('')

        for img in sec.get('images', []):
            fig_num += 1
            cap = _q(sec.get('heading', 'Figure')[:80])
            l(f"insert_figure(os.path.join(FIG_DIR, '{img}'), {D['img_width']}, {fig_num}, '{cap}')")
            l('')

        for para in sec.get('paragraphs', []):
            # Check for formula paragraphs (dict with 'math' key)
            if isinstance(para, dict) and para.get('math'):
                txt = _q(para.get('text', ''))
                # Add formula text + LaTeX annotation comments
                for m in para['math']:
                    if 'text' in m and m['text']:
                        l(f'# {m["text"]}')
                    if 'latex' in m and m['latex']:
                        l(f'# latex: {_q(m["latex"])}')
                l(f'body_with_formula("{txt}", [')
                for m in para['math']:
                    if 'latex' in m and m['latex']:
                        l(f"    latex_to_omath(r'{_q(m['latex'])}'),")
                    else:
                        l(f"    '{_q(m['xml'])}',")
                l('])')
                continue
            # Regular text paragraph
            p = (para if isinstance(para, str) else para.get('text', '')).strip()
            if not p or len(p) < 5:
                continue
            if any(k in p[:100] for k in ['号', '行距', '缩进', '对齐', '空一行', '备注', '按答辩时间', '小四', '四号', '小三']):
                continue
            printable = sum(1 for c in p if c.isprintable() or c in '\n\r\t ')
            if printable / max(len(p), 1) < 0.7:
                continue
            if len(p) < 20 and not any(c.isascii() and c.isalpha() for c in p):
                continue
            l(f"body('{_q(p)}')")
        l('')

    # ═══ REFERENCES ═══
    refs = cnt.get('references', [])
    if refs:
        l('# ── References with bookmarks ──')
        if P['h_levels']:
            l(f'p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.{P["h_levels"][0]["align"]}')
            l('r = p.add_run("References"); r.bold = True')
            l(f'r.font.size = Pt({P["h_levels"][0]["size"]})')
            l(f"r.font.name = '{P['body_font']}'")
            if need_eastAsia:
                l(f'rp = r._element.get_or_add_rPr()')
                l(f'rf = rp.find(qn("w:rFonts"))')
                l(f'if rf is None: rf = OxmlElement("w:rFonts"); rp.insert(0, rf)')
                l(f'rf.set(qn("w:eastAsia"), "{cjk}")')
        l('')
        l('refs = [')
        for i, ref in enumerate(refs):
            l(f"    ({i+1}, '{_q(ref)}'),")
        l(']')
        l('')
        l('for num, ref_text in refs:')
        l('    p = doc.add_paragraph()')
        l(f'    p.alignment = WD_ALIGN_PARAGRAPH.{align}')
        l(f'    p.paragraph_format.left_indent = Cm({D["ref_indent"]})')
        l(f'    p.paragraph_format.first_line_indent = Cm(-{D["ref_indent"]})')
        l('    bm = f"_Ref{num}"')
        l('    bk = OxmlElement("w:bookmarkStart")')
        l('    bk.set(qn("w:id"), str(num)); bk.set(qn("w:name"), bm)')
        l('    p._element.append(bk)')
        # Strip existing [N] prefix from ref text to avoid double numbering
        l('    import re')
        l('    clean_ref = re.sub(r"^\[\d+\]\s*", "", ref_text)')
        l(f'    r = p.add_run(f"[{{num}}] "); r.font.size = Pt({D["ref_size"]})')
        l(f"    r.font.name = '{P['body_font']}'")
        if need_eastAsia:
            l(f'    rp = r._element.get_or_add_rPr()')
            l(f'    rf = rp.find(qn("w:rFonts"))')
            l(f'    if rf is None: rf = OxmlElement("w:rFonts"); rp.insert(0, rf)')
            l(f'    rf.set(qn("w:eastAsia"), "{cjk}")')
        l(f'    r = p.add_run(clean_ref); r.font.size = Pt({D["ref_size"]})')
        l(f"    r.font.name = '{P['body_font']}'")
        if need_eastAsia:
            l(f'    rp = r._element.get_or_add_rPr()')
            l(f'    rf = rp.find(qn("w:rFonts"))')
            l(f'    if rf is None: rf = OxmlElement("w:rFonts"); rp.insert(0, rf)')
            l(f'    rf.set(qn("w:eastAsia"), "{cjk}")')
        l('    be = OxmlElement("w:bookmarkEnd")')
        l('    be.set(qn("w:id"), str(num)); p._element.append(be)')
        l('')

    # ═══ PAGINATE + SAVE ═══
    l('# ── Paginate + Save ──')
    l('n = paginate_a4(doc)')
    l('doc.save(OUT)')
    l('_cc.save(OUT)  # inject comments into saved docx')
    l("print(f'Saved: {OUT}  pages: ~{n+1}')")

    code = '\n'.join(L)
    os.makedirs(os.path.dirname(output_py_path), exist_ok=True)
    with open(output_py_path, 'w', encoding='utf-8') as f:
        f.write(code)
    return len(code)
