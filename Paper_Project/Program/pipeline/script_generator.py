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
    # Template text often combines all levels: "一级标题 黑体三号加粗、二级标题 黑体四号加粗 三级标题 黑体小四加粗"
    # Split by level markers to isolate each level's spec
    _h_specs = re.split(r'(?=一[级二]标题|二[级二]标题|三[级二]标题)', all_text)
    for level_name, level_key in [('一级标题', 'h1'), ('二级标题', 'h2'), ('三级标题', 'h3')]:
        # Find spec chunks that start with this level name
        _chunks = [c for c in _h_specs if c.strip().startswith(level_name)]
        if not _chunks:
            # Fallback: old regex approach
            pattern = re.compile(rf'{level_name}[^。）]*(?:。|；|）)')
            _chunks = pattern.findall(all_text)
        # Prefer chunks with font/size info, exclude TOC
        _good = [c for c in _chunks if '加粗' in c and '目录' not in c and _get_size(c) is not None]
        _use = _good if _good else _chunks
        combined = max(_use, key=len) if _use else ''
        if combined:
            rules[level_key] = {
                'size': _get_size(combined, 15 if level_key == 'h1' else 14 if level_key == 'h2' else 12),
                'font': _get_font(combined, '黑体'),
                'bold': '加粗' in combined,
                'align': _get_align(combined, 'CENTER'),
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
    # Template text may have: "正文：黑体三号加粗…固定值28磅" or "正文宋体小四…"
    # Search broadly for any text containing "正文" with format specs
    body_pattern = re.compile(r'正文[^。\n]{0,120}')
    body_matches = body_pattern.findall(all_text)
    # Find matches with line spacing info
    _with_ls = [m for m in body_matches if '行距' in m or '固定值' in m]
    # Also search for "固定值28磅" anywhere (even without "正文" prefix)
    if not _with_ls:
        _ls_anywhere = re.findall(r'[^。\n]{0,60}固定值\s*\d+[^。\n]{0,60}', all_text)
        if _ls_anywhere:
            body_text = _ls_anywhere[0]
        else:
            body_text = body_matches[0] if body_matches else ''
    else:
        body_text = _with_ls[0]

    if body_text:
        # Detect line spacing from "固定值28磅" or "1.5倍" or "双倍"
        _ls_fixed = re.search(r'固定值\s*(\d+(?:\.\d+)?)\s*磅', body_text)
        if _ls_fixed:
            _ls_val = f'Pt({float(_ls_fixed.group(1))})'
        elif '1.5倍' in body_text:
            _ls_val = 1.5
        elif '双倍' in body_text:
            _ls_val = 2.0
        else:
            _ls_val = None
        # Only extract font/size if the match is about body format, not headings
        # Check if this is a heading instruction misidentified as body
        _is_heading_spec = any(kw in body_text for kw in ['一级标题', '二级标题', '三级标题', '标题', '目录'])
        if _is_heading_spec:
            # The match is about heading specs within body — don't extract body font from it
            _body_size = 12
            _body_font = '宋体'
            _body_align = 'JUSTIFY'
            _body_indent = 0.74
        else:
            _body_size = _get_size(body_text, 12)
            _body_font = _get_font(body_text, '宋体')
            _body_align = _get_align(body_text, 'JUSTIFY')
            _body_indent = 0.74 if '缩进' in body_text and ('2字符' in body_text or '2个汉字' in body_text) else 0
        rules['body'] = {
            'size': _body_size,
            'font': _body_font,
            'ls': _ls_val,
            'ls_fixed_pt': float(_ls_fixed.group(1)) if _ls_fixed else None,
            'align': _body_align,
            'indent': _body_indent,
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
        _hdr_font = r0.get('font')
        if not _hdr_font or _hdr_font == 'Times New Roman':
            _hdr_font = '宋体'  # CJK default for Chinese templates
        hdr = {
            'text': _q(hdr_text[:120]),
            'align': h0.get('alignment', 'RIGHT'),
            'font': _hdr_font,
            'size': r0.get('size_pt', 9),
            'bold': r0.get('bold', False),
            'italic': r0.get('italic', False),
        }
    P['header'] = hdr

    # ── Body text: sample from long paragraphs (real content, not cover/headings/declarations) ──
    samples = []
    for p in fmt['paragraphs']:
        txt = p.get('text', '').strip()
        # Must be long enough to be real body content
        if len(txt) < 100:
            continue
        # Skip format notes that start with （
        if txt.startswith('（') or txt.startswith('('):
            continue
        # Skip known cover/declaration content
        if any(kw in txt[:30] for kw in ['本人郑重', '本科生毕业', '学位评定', '本人在导师', '原创性声明', '版权使用']):
            continue
        # Must contain Chinese text (periods, commas, or CJK characters)
        has_cjk = any('一' <= c <= '鿿' for c in txt[:200])
        if not has_cjk:
            continue
        for r in p.get('runs', []):
            if r.get('size_pt') and r.get('font'):
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
        P['body_size'] = 12; P['body_font'] = '宋体'
        P['body_ls'] = None; P['body_align'] = 'JUSTIFY'; P['body_indent'] = 0.74

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
        if b.get('font'): P['body_font'] = b['font']
        if b.get('ls'): P['body_ls'] = b['ls']
        if b.get('ls_fixed_pt'): P['body_ls_fixed_pt'] = b['ls_fixed_pt']
        if b.get('align'): P['body_align'] = b['align']
        if b.get('indent') is not None: P['body_indent'] = b['indent']

    # ── Sanity: if body values still look like heading, use defaults ──
    if P.get('body_size', 0) >= 16 and P.get('body_font') == '黑体':
        P['body_size'] = 12
        P['body_font'] = '宋体'
        P['body_align'] = 'JUSTIFY'
        P['body_indent'] = 0.74
    if P.get('body_align') not in ('LEFT', 'RIGHT', 'CENTER', 'JUSTIFY', 'DISTRIBUTE'):
        P['body_align'] = 'JUSTIFY'
    if P.get('body_align') == 'CENTER':
        P['body_align'] = 'JUSTIFY'

    # ── English abstract format: read from format.json P39-P43 ──
    P['eng_abs'] = {}
    for p in fmt['paragraphs']:
        txt = p.get('text', '').strip()
        # P39: ABSTRACT heading
        if 'ABSTRACT' in txt and len(txt) < 30:
            for r in p.get('runs', []):
                if r.get('size_pt') and r.get('size_pt') >= 14:
                    P['eng_abs']['heading_font'] = r.get('font', 'Times New Roman')
                    P['eng_abs']['heading_size'] = r['size_pt']
                    P['eng_abs']['heading_align'] = p.get('alignment', 'CENTER')
                    P['eng_abs']['heading_bold'] = r.get('bold', True)
                    break
        # P40: English abstract body
        if len(txt) > 100 and all(ord(c) < 128 or c == ' ' for c in txt[:50]):
            for r in p.get('runs', []):
                if r.get('size_pt'):
                    P['eng_abs']['body_font'] = r.get('font', 'Times New Roman')
                    P['eng_abs']['body_size'] = r['size_pt']
                    break
            P['eng_abs']['body_align'] = p.get('alignment', 'JUSTIFY')
            P['eng_abs']['body_ls'] = p.get('line_spacing_val', 1.5)
            P['eng_abs']['body_indent'] = p.get('first_indent_cm', 0.9)
            break
    if not P['eng_abs']:
        P['eng_abs'] = {'heading_font': 'Times New Roman', 'heading_size': 16,
                        'body_font': 'Times New Roman', 'body_size': 12,
                        'body_align': 'JUSTIFY', 'body_ls': 1.5, 'body_indent': 0.9}

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
    _hdr = P.get('header') or {}
    D['footer_size'] = _hdr.get('size', 9) if _hdr else 9
    D['footer_font'] = _hdr.get('font', '宋体') if _hdr else '宋体'
    D['ref_size'] = P['body_size']  # same as body text per template
    D['caption_size'] = max(P['body_size'] - 2, 8)
    D['ref_sep_size'] = max(P['body_size'] - 3, 7)
    D['usable_pt'] = (P['page_h'] - P['mt'] - P['mb']) / 0.0352778 + 0.5
    text_w = P['page_w'] - P['ml'] - P['mr']
    D['img_width'] = min(round(text_w * 0.55, 1), 4.2)
    # Reference hanging indent: standard 1.27cm (0.5in) in Chinese academic papers
    D['ref_indent'] = 1.27
    D['table_cell_size'] = max(P['body_size'] - 3, 8)

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
    _ftr_font = D.get('footer_font', '宋体')
    _ftr_size = D.get('footer_size', 9)
    l(f"rf = fp.add_run(); rf.font.size = Pt({_ftr_size}); rf.font.name = '{_ftr_font}'")
    l('rp = rf._element.get_or_add_rPr()')
    l('rf2 = rp.find(qn("w:rFonts"))')
    l('if rf2 is None: rf2 = OxmlElement("w:rFonts"); rp.insert(0, rf2)')
    l(f'rf2.set(qn("w:eastAsia"), "{_ftr_font}"); rf2.set(qn("w:hint"), "eastAsia")')
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
        _hdr_align = h['align'] if h['align'] in ('LEFT','CENTER','RIGHT','JUSTIFY','DISTRIBUTE') else 'CENTER'
        l(f"hp = hdr.paragraphs[0]; hp.alignment = WD_ALIGN_PARAGRAPH.{_hdr_align}")
        _hdr_font = h.get('font') or '宋体'
        _hdr_size = h.get('size') or 9
        l(f"r = hp.add_run('{h['text']}')")
        l(f"r.font.size = Pt({_hdr_size}); r.font.name = '{_hdr_font}'")
        l('rp = r._element.get_or_add_rPr()')
        l('rf = rp.find(qn("w:rFonts"))')
        l('if rf is None: rf = OxmlElement("w:rFonts"); rp.insert(0, rf)')
        l(f'rf.set(qn("w:eastAsia"), "{_hdr_font}"); rf.set(qn("w:hint"), "eastAsia")')
        l(f"r.bold = {h['bold']}; r.italic = {h['italic']}")
        l('')

    # ═══ DEFAULT STYLE ═══
    l('# Default paragraph style')
    l('style = doc.styles["Normal"]')
    l(f"style.font.name = '{P['body_font']}'")
    _ns = fmt.get('normal_style', {})
    _ns_size = _ns.get('font_size_pt') or 10
    _ns_ls = _ns.get('line_spacing') or 1.0
    l(f'style.font.size = Pt({_ns_size})  # from template Normal style')
    l(f'style.paragraph_format.line_spacing = {_ns_ls}  # from template Normal style')
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
    _body_ls = P.get('body_ls_fixed_pt')
    if _body_ls:
        _ls_code = f'Pt({_body_ls})'
    elif isinstance(P.get('body_ls'), str) and P['body_ls'].startswith('Pt('):
        _ls_code = P['body_ls']
    elif P.get('body_ls'):
        _ls_code = str(P['body_ls'])
    else:
        _ls_code = 'Pt(28.0)'
    _body_indent_pt = round(P['body_indent'] * 28.35) if P['body_indent'] > 0 else 24
    _body_font = P.get('body_font', '宋体')
    _body_size = P.get('body_size', 12)
    _body_align = P.get('body_align', 'JUSTIFY')
    _cjk = P.get('cjk_font', '宋体')
    l('def body(text, first_indent=True, comment=None):')
    l('    p = doc.add_paragraph()')
    l(f'    p.alignment = WD_ALIGN_PARAGRAPH.{_body_align}')
    l(f'    pf = p.paragraph_format; pf.line_spacing = {_ls_code}')
    l('    pf.space_after = Pt(0)')
    l('    pf.space_before = Pt(0)')
    l('    if first_indent:')
    l(f'        pf.first_line_indent = Pt({_body_indent_pt})')
    l('    r = p.add_run(text)')
    l(f"    r.font.name = '{_body_font}'; r.font.size = Pt({_body_size})")
    l('    rp = r._element.get_or_add_rPr()')
    l('    rf = rp.find(qn("w:rFonts"))')
    l('    if rf is None: rf = OxmlElement("w:rFonts"); rp.insert(0, rf)')
    l(f'    rf.set(qn("w:eastAsia"), "{_cjk}"); rf.set(qn("w:hint"), "eastAsia")')
    l('    if comment:')
    l('        _cc.add(p, comment)')
    l('    return p')
    l('')

    # headings — with English detection for auto TNR
    # sizes: H1=三号16pt, H2=四号14pt, H3=小四12pt
    _h_sizes = {1: 16, 2: 14, 3: 12}
    for hl in P['h_levels']:
        n = hl['level']
        _hsz = _h_sizes.get(n, hl.get('size', 12))
        l(f'def heading{n}(text, comment=None):')
        l('    p = doc.add_paragraph()')
        l(f'    p.alignment = WD_ALIGN_PARAGRAPH.CENTER')
        l(f'    pf = p.paragraph_format; pf.line_spacing = {_ls_code}')
        l(f'    pf.space_before = Pt({hl["space_before"]})')
        l('    r = p.add_run(text); r.bold = True')
        l('    eng = sum(1 for c in text if c.isascii() and c.isalpha()) / max(len(text),1)')
        l('    fn = "Times New Roman" if eng > 0.5 else "黑体"')
        l(f'    r.font.name = fn; r.font.size = Pt({_hsz})')
        l('    rp = r._element.get_or_add_rPr()')
        l('    rf = rp.find(qn("w:rFonts"))')
        l('    if rf is None: rf = OxmlElement("w:rFonts"); rp.insert(0, rf)')
        l('    rf.set(qn("w:eastAsia"), fn); rf.set(qn("w:hint"), "eastAsia")')
        l('    if comment:')
        l('        _cc.add(p, comment)')
        l('    return p')
        l('')
        l('')
        _ea = P['eng_abs']
        l('# English body text — format from template English abstract section')
        l('def english_body(text):')
        l('    p = doc.add_paragraph()')
        l(f"    p.alignment = WD_ALIGN_PARAGRAPH.{_ea['body_align']}")
        l(f'    pf = p.paragraph_format; pf.line_spacing = {_ea["body_ls"]}')
        l('    pf.space_after = Pt(0)')
        l('    pf.space_before = Pt(0)')
        l(f'    pf.first_line_indent = Cm({_ea["body_indent"]})')
        l('    r = p.add_run(text)')
        l(f"    r.font.name = '{_ea['body_font']}'; r.font.size = Pt({_ea['body_size']})")
        l('    rp = r._element.get_or_add_rPr()')
        l('    rf = rp.find(qn("w:rFonts"))')
        l('    if rf is None: rf = OxmlElement("w:rFonts"); rp.insert(0, rf)')
        l(f'    rf.set(qn("w:eastAsia"), "{_ea["body_font"]}"); rf.set(qn("w:hint"), "eastAsia")')
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
            l(f'    rf.set(qn("w:eastAsia"), "{cjk}"); rf.set(qn("w:hint"), "eastAsia")')
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
            l(f'    rf.set(qn("w:eastAsia"), "{cjk}"); rf.set(qn("w:hint"), "eastAsia")')
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
        _cap_size = P.get('body_size', 12)
        _cap_font = P.get('body_font', '宋体')
        _cap_cjk = P.get('cjk_font', '宋体')
        l('        r = cap.add_run(f"图{fig_num} ")')
        l(f'        r.font.size = Pt({_cap_size}); r.font.name = "{_cap_font}"; r.bold = True')
        l('        rp = r._element.get_or_add_rPr()')
        l('        rf = rp.find(qn("w:rFonts"))')
        l('        if rf is None: rf = OxmlElement("w:rFonts"); rp.insert(0, rf)')
        l(f'        rf.set(qn("w:eastAsia"), "{_cap_cjk}"); rf.set(qn("w:hint"), "eastAsia")')
        l('        r2 = cap.add_run(caption_text)')
        l(f'        r2.font.size = Pt({_cap_size}); r2.font.name = "{_cap_font}"; r2.italic = True')
        l('        rp = r2._element.get_or_add_rPr()')
        l('        rf = rp.find(qn("w:rFonts"))')
        l('        if rf is None: rf = OxmlElement("w:rFonts"); rp.insert(0, rf)')
        l(f'        rf.set(qn("w:eastAsia"), "{_cap_cjk}"); rf.set(qn("w:hint"), "eastAsia")')
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
    l('# Paginate disabled — Word handles page breaks naturally')
    l('def paginate_a4(doc):')
    l('    return 0  # no-op: let Word paginate naturally')
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
            l(f'        rf.set(qn("w:eastAsia"), "{cjk}"); rf.set(qn("w:hint"), "eastAsia")')
        l('    for xml_str in math_xml_list:')
        l('        math_el = etree.fromstring(xml_str)')
        l('        _strip_math_ns(math_el)  # remove redundant xmlns for WPS compatibility')
        l('        p._element.append(math_el)')
        l('    return p')
        l('')

    # ═══ SyncTeX source tracking ═══
    l('')
    l('# ── SyncTeX source tracking ──')
    shutil.copy2(os.path.join(pipeline_dir, 'sync_tracker.py'), os.path.join(output_dir, 'sync_tracker.py'))
    l('from sync_tracker import SyncTracker')
    l('st = SyncTracker(doc)')
    l('if os.environ.get("DOCX_SYNC_DISABLE", "0") != "1":')
    l('    body = st.track(body)')
    l('    heading1 = st.track(heading1)')
    l('    heading2 = st.track(heading2)')
    l('    heading3 = st.track(heading3)')
    if has_images:
        l('    insert_figure = st.track_multi(insert_figure, count=2)')
    l('')
    l('')

    # ═══ CONTENT ═══
    l('# ═══════════════════ CONTENT ═══════════════════')
    l('')
    # -- Cover page -- use pre-extracted cover elements from format.json --
    cover_elements = fmt.get('cover', [])
    cover_info = cnt.get('cover_info', {})
    _content_map = {}
    if cover_info:
        _ci = cover_info
        _LABELS = {
            'school_code': '学校编码：', 'paper_title': '论文题目：',
            'student_name': '学生姓名：', 'student_id': '学    号：',
            'college': '所属学院：', 'class_name': '专业班级：',
            'advisor': '指导老师：',
        }
        for _k, _label_tmpl in _LABELS.items():
            if _k in _ci:
                _content_map[_label_tmpl] = _ci[_k]

    if cover_elements:
        l('')
        l('# ' + '='*20 + ' COVER ' + '='*20)
        l('from docx import Document as _Doc')
        l('import io as _io, os as _os')
        _tp = os.path.abspath(os.path.join('Templates', fmt['_meta']['source'])).replace(chr(92), chr(47))
        l(f"_tpl = _Doc(r'{_tp}')")
        l('_tpl_rels = _tpl.part.rels')
        l('from PIL import Image as _PILImg')
        l('from docx.enum.text import WD_LINE_SPACING')
        l('')
        _els_json = json.dumps(cover_elements, ensure_ascii=False)
        _els_json = _els_json.replace('true', 'True').replace('false', 'False').replace('null', 'None')
        l('_cover_els = ' + _els_json)
        l('')
        l('_cmap = {')
        for _k, _v in sorted(_content_map.items(), key=lambda x: -len(x[0])):
            l(f"    '{_q(_k)}': '{_q(_v)}',")
        l('}')
        l('')
        l('_img_seq = 0')
        l('_img_dir = _os.path.join(BASE, "cover_images")')
        l('_os.makedirs(_img_dir, exist_ok=True)')
        l('')
        l('for _el in _cover_els:')
        l('    _etyp = _el.get("type","")')
        l('')
        l('    if _etyp in ("empty","para"):')
        l('        p = doc.add_paragraph()')
        l('        pf = p.paragraph_format')
        l('        _al = _el.get("al")')
        l('        _am = {"left":"LEFT","center":"CENTER","right":"RIGHT","both":"JUSTIFY","distribute":"DISTRIBUTE"}')
        l('        if _al: setattr(p, "alignment", getattr(WD_ALIGN_PARAGRAPH, _am.get(_al, "LEFT")))')
        l('        _ls_val = _el.get("ls_val")')
        l('        _ls_rule = _el.get("ls_rule")')
        l('        if _ls_val:')
        l('            _ls_n = int(_ls_val)')
        l('            if _ls_rule in ("exact",):')
        l('                pf.line_spacing = Pt(_ls_n / 20)')
        l('                pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY')
        l('            elif _ls_rule in ("atLeast",):')
        l('                pf.line_spacing = Pt(_ls_n / 20)')
        l('                pf.line_spacing_rule = WD_LINE_SPACING.AT_LEAST')
        l('            else:')
        l('                pf.line_spacing = _ls_n / 240')
        l('        _sp_before = _el.get("sp_before")')
        l('        if _sp_before: pf.space_before = Pt(int(_sp_before) / 20)')
        l('        _fl_indent = _el.get("fl_indent")')
        l('        if _fl_indent:')
        l('            _fli = int(_fl_indent)')
        l('            if _fli > 0: pf.first_line_indent = Pt(_fli / 20)')
        l('        _ft = "".join(r.get("t","") for r in _el.get("r",[]))')
        l('        for _r in _el.get("r",[]):')
        l('            _rt = _r.get("t","")')
        l('            _fn = _r.get("fn") or _r.get("fe") or ""')
        l('            _fea = _r.get("fe") or ""')
        l('            _fsz = _r.get("sz") or 0')
        l('            _fb = _r.get("b", False)')
        l('            if _etyp == "empty" and _fsz <= 0 and not _rt: continue')
        l('            rr = p.add_run(_rt)')
        l('            if _fsz > 0: rr.font.size = Pt(_fsz)')
        l('            rr.font.name = _fn')
        l('            if _fb: rr.bold = True')
        l('            if _fea:')
        l('                rp = rr._element.get_or_add_rPr()')
        l('                rf = rp.find(qn("w:rFonts"))')
        l('                if rf is None: rf = OxmlElement("w:rFonts"); rp.insert(0, rf)')
        l('                rf.set(qn("w:eastAsia"), _fea); rf.set(qn("w:hint"), "eastAsia")')
        l('        if "学位评定委员会" in _ft:')
        l('            _r = p.add_run()')
        l('            _br = OxmlElement("w:br"); _br.set(qn("w:type"), "page")')
        l('            _r._element.append(_br)')
        l('')
        l('    elif _etyp =="image":')
        l('        _ext = _el.get("extent",{})')
        l('        _cx = max(0.5, int(_ext.get("cx","0")) / 914400) if _ext else 5')
        l('        _src = _el.get("srcRect",{})')
        l('        _rEmbed = _el.get("rEmbed","")')
        l('        _fp = _os.path.join(_img_dir, f"cover_img_{_img_seq:02d}.png")')
        l('        if not _os.path.exists(_fp):')
        l('            if _rEmbed and _rEmbed in _tpl_rels:')
        l('                _rel = _tpl_rels[_rEmbed]')
        l('                if "image" in _rel.reltype:')
        l('                    _blob = _rel.target_part.blob')
        l('                    if _src:')
        l('                        _pi = _PILImg.open(_io.BytesIO(_blob))')
        l('                        _pic_w, _pic_h = _pi.size')
        l('                        _l = int(_pic_w*int(_src.get("l","0"))/100000)')
        l('                        _t = int(_pic_h*int(_src.get("t","0"))/100000)')
        l('                        _r = int(_pic_w*int(_src.get("r","0"))/100000)')
        l('                        _b = int(_pic_h*int(_src.get("b","0"))/100000)')
        l('                        _pi.crop((_l,_t,_pic_w-_r,_pic_h-_b)).save(_fp,"PNG")')
        l('                    else:')
        l('                        with open(_fp,"wb") as _f: _f.write(_blob)')
        l('        if _os.path.exists(_fp):')
        l('            doc.add_picture(_fp, width=Inches(_cx))')
        l('            pf = doc.paragraphs[-1].paragraph_format')
        l('            _al = _el.get("al")')
        l('            _am = {"left":"LEFT","center":"CENTER","right":"RIGHT","both":"JUSTIFY","distribute":"DISTRIBUTE"}')
        l('            if _al: doc.paragraphs[-1].alignment = getattr(WD_ALIGN_PARAGRAPH, _am.get(_al, "CENTER"))')
        l('            _ls_val = _el.get("ls_val")')
        l('            _ls_rule = _el.get("ls_rule")')
        l('            if _ls_val:')
        l('                _ls_n = int(_ls_val)')
        l('                if _ls_rule in ("exact",):')
        l('                    pf.line_spacing = Pt(_ls_n / 20)')
        l('                    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY')
        l('                elif _ls_rule in ("atLeast",):')
        l('                    pf.line_spacing = Pt(_ls_n / 20)')
        l('                    pf.line_spacing_rule = WD_LINE_SPACING.AT_LEAST')
        l('                else:')
        l('                    pf.line_spacing = _ls_n / 240')
        l('            for _r in _el.get("r",[]):')
        l('                _rt = _r.get("t","")')
        l('                if not _rt: continue')
        l('                _fn = _r.get("fn") or _r.get("fe") or ""')
        l('                _fea = _r.get("fe") or ""')
        l('                _fsz = _r.get("sz") or 0')
        l('                rr = doc.paragraphs[-1].add_run(_rt)')
        l('                if _fsz > 0: rr.font.size = Pt(_fsz)')
        l('                if _fn: rr.font.name = _fn')
        l('                if _fea:')
        l('                    rp = rr._element.get_or_add_rPr()')
        l('                    rf = rp.find(qn("w:rFonts"))')
        l('                    if rf is None: rf = OxmlElement("w:rFonts"); rp.insert(0, rf)')
        l('                    rf.set(qn("w:eastAsia"), _fea); rf.set(qn("w:hint"), "eastAsia")')
        l('        _img_seq += 1')
        l('')
        l('    elif _etyp =="table":')
        l('        _rows = _el.get("rows",[])')
        l('        _nrows = len(_rows)')
        l('        _ncols = max(len(r) for r in _rows) if _rows else 2')
        l('        t = doc.add_table(rows=_nrows, cols=_ncols)')
        l('        for _ri, _row in enumerate(_rows):')
        l('            _row_label = ""')
        l('            if len(_row) > 0:')
        l('                _p0_l = _row[0].get("p",[{}])[0] if _row[0].get("p") else {}')
        l('                _runs_l = _p0_l.get("r",[])')
        l('                _row_label = "".join(r.get("t","") for r in _runs_l).strip()')
        l('            _row_val = _cmap.get(_row_label, "")')
        l('            if not _row_val:')
        l('                for _ck, _cv in _cmap.items():')
        l('                    if _ck in _row_label or _row_label in _ck:')
        l('                        _row_val = _cv; break')
        l('            for _ci, _cell in enumerate(_row):')
        l('                _paras = _cell.get("p",[])')
        l('                if not _paras: continue')
        l('                _p0 = _paras[0]')
        l('                _runs0 = _p0.get("r",[])')
        l('                if not _runs0: continue')
        l('                _cell_label = "".join(r.get("t","") for r in _runs0).strip()')
        l('                if _ci == 1 and _row_val:')
        l('                    _use = _row_val')
        l('                else:')
        l('                    _use = _cell_label')
        l('                _r0 = _runs0[0]')
        l('                _fn2 = _r0.get("fn") or _r0.get("fe") or ""')
        l('                _fea2 = _r0.get("fe") or ""')
        l('                _fsz2 = _r0.get("sz") or 0')
        l('                _fb2 = _r0.get("b", False)')
        l('                _al2 = _p0.get("al")')
        l('                _am2 = {"left":"LEFT","center":"CENTER","right":"RIGHT","both":"JUSTIFY"}')
        l('                cell = t.rows[_ri].cells[_ci]')
        l('                cell.text = ""')
        l('                if _al2: cell.paragraphs[0].alignment = getattr(WD_ALIGN_PARAGRAPH, _am2.get(_al2, "CENTER"))')
        l('                r = cell.paragraphs[0].add_run(_use)')
        l('                r.font.name = _fn2 if _fn2 else "宋体"')
        l('                if _fsz2 > 0: r.font.size = Pt(_fsz2)')
        l('                if _fb2: r.bold = True')
        l('                _ea = _fea2 if _fea2 else "宋体"')
        l('                rp = r._element.get_or_add_rPr()')
        l('                rf = rp.find(qn("w:rFonts"))')
        l('                if rf is None: rf = OxmlElement("w:rFonts"); rp.insert(0, rf)')
        l('                rf.set(qn("w:eastAsia"), _ea); rf.set(qn("w:hint"), "eastAsia")')
        l('                _borders = _cell.get("borders",{})')
        l('                if _borders:')
        l('                    tcPr = cell._tc.get_or_add_tcPr()')
        l('                    for old in tcPr.findall(qn("w:tcBorders")): tcPr.remove(old)')
        l('                    tcB = OxmlElement("w:tcBorders"); tcPr.append(tcB)')
        l('                    for _bp, _bv in _borders.items():')
        l('                        if isinstance(_bv, dict):')
        l('                            b = OxmlElement(f"w:{_bp}")')
        l('                            b.set(qn("w:val"), _bv.get("val","single"))')
        l('                            b.set(qn("w:sz"), _bv.get("sz","4"))')
        l('                            b.set(qn("w:space"), "0")')
        l('                            b.set(qn("w:color"), _bv.get("color","000000"))')
        l('                            tcB.append(b)')
        l('                _cw = _cell.get("w",0)')
        l('                if _cw: cell.width = Cm(round(_cw/567,1))')
        l('')
        l('# Page break after cover+declarations')
        l('doc.add_page_break()')
        l('')
    else:
        l('# (no cover elements found in template)')
        l('')

    # -- Title --
    ti = cnt.get('title_info', {})
    # Sections
    fig_num = 0
    for sec in cnt.get('sections', []):
        h = sec.get('heading', '').strip()
        lv = sec.get('level', 0)
        if h:
            safe = _q(h)
            # English abstract/keywords headings should be 16pt (heading1 size)
            _use_lv = lv
            if lv == 2 and h in ('Abstract', 'KEYWORDS:', 'KEY WORDS:'):
                _use_lv = 1
            if _use_lv >= 1 and _use_lv <= 3:
                l(f'heading{_use_lv}("{safe}")')
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
            # Use english_body for English sections (Abstract, KEYWORDS)
            _h_eng = sum(1 for c in h if c.isascii() and c.isalpha())
            _is_eng_section = _h_eng > len(h) * 0.5 or h in ('Abstract', 'KEYWORDS:', 'KEY WORDS:')
            _body_fn = 'english_body' if _is_eng_section else 'body'
            l(f"{_body_fn}('{_q(p)}')")
        l('')

    # ═══ REFERENCES ═══
    refs = cnt.get('references', [])
    # Separate actual refs from 致谢/附录 — detect boundaries
    _xie_start = None; _app_start = None
    for _i, _r in enumerate(refs):
        if '致谢' in _r or '致  谢' in _r:
            _xie_start = _i
        if _xie_start is not None and _app_start is None and ('附录' in _r or '附  录' in _r):
            _app_start = _i
    _pure_refs = refs[:_xie_start] if _xie_start else refs
    _xie_items = refs[_xie_start:_app_start] if _xie_start else []
    _app_items = refs[_app_start:] if _app_start else []

    if _pure_refs:
        l('# ── References ──')
        l('doc.add_page_break()')
        l("heading1('参考文献')")
        l('')
        l('refs = [')
        for i, ref in enumerate(_pure_refs):
            l(f"    ({i+1}, '{_q(ref)}'),")
        l(']')
        l('')
        l('for num, ref_text in refs:')
        l('    p = doc.add_paragraph()')
        l('    p.alignment = WD_ALIGN_PARAGRAPH.LEFT')
        l(f'    p.paragraph_format.left_indent = Cm({D["ref_indent"]})')
        l(f'    p.paragraph_format.first_line_indent = Cm(-{D["ref_indent"]})')
        l('    bm = f"_Ref{num}"')
        l('    bk = OxmlElement("w:bookmarkStart")')
        l('    bk.set(qn("w:id"), str(num)); bk.set(qn("w:name"), bm)')
        l('    p._element.append(bk)')
        l('    import re')
        l('    clean_ref = re.sub(r"^\[\d+\]\s*", "", ref_text)')
        _ref_font = P.get('body_font', '宋体')
        _ref_size = D.get('ref_size', 12)
        l(f'    r = p.add_run("[" + str(num) + "] "); r.font.size = Pt({_ref_size})')
        l(f"    r.font.name = '{_ref_font}'")
        l('    rp = r._element.get_or_add_rPr()')
        l('    rf = rp.find(qn("w:rFonts"))')
        l('    if rf is None: rf = OxmlElement("w:rFonts"); rp.insert(0, rf)')
        _ref_cjk = P.get('cjk_font', '宋体')
        l(f'    rf.set(qn("w:eastAsia"), "{_ref_cjk}"); rf.set(qn("w:hint"), "eastAsia")')
        l(f'    r = p.add_run(clean_ref); r.font.size = Pt({_ref_size})')
        l(f"    r.font.name = '{_ref_font}'")
        l('    rp = r._element.get_or_add_rPr()')
        l('    rf = rp.find(qn("w:rFonts"))')
        l('    if rf is None: rf = OxmlElement("w:rFonts"); rp.insert(0, rf)')
        l(f'    rf.set(qn("w:eastAsia"), "{_ref_cjk}"); rf.set(qn("w:hint"), "eastAsia")')
        l('    be = OxmlElement("w:bookmarkEnd")')
        l('    be.set(qn("w:id"), str(num)); p._element.append(be)')
        l('')

    # ── 致谢 ──
    if _xie_items:
        l('# ── 致谢 ──')
        l('doc.add_page_break()')
        l("heading1('致  谢')")
        for _xi in _xie_items[1:]:  # skip heading row
            l(f"body('{_q(_xi)}')")
        l('')

    # ── 附录 ──
    if _app_items:
        l('# ── 附录 ──')
        l('doc.add_page_break()')
        l("heading1('附  录')")
        for _ai in _app_items[1:]:  # skip heading row
            l(f"body('{_q(_ai)}', first_indent=False)")
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
