"""
script_generator.py — role-driven DOCX build script generator.

输入: format.json + content.json
输出: build_generated.py；运行该脚本生成最终论文 docx。

修复点：
- 封面按模板元素重建，支持校徽/图片 assets，学位/学校编码行左对齐。
- 中文摘要、英文摘要、目录、正文使用独立 role；英文摘要另起一页，正文在目录后另起一节。
- 标题设置 outline level，编号和标题之间自动补一个空格。
- 正文 [N] / [1,2] / [1-3] 引用自动上标。
- 表格渲染为三线表，保留单元格换行。
- 图/表题走 caption role，居中并规范“图 1 标题”的空格。
- 代码块保留换行和空格，使用等宽字体。
- 参考文献单条一个段落，悬挂缩进，中文宋体/英文 Times New Roman 混排。
"""
from __future__ import annotations

import json
import os
import re
from typing import Any, Dict, List, Optional


def _is_cjk(text: str) -> bool:
    return any('\u4e00' <= c <= '\u9fff' for c in str(text or ''))


def _ascii_ratio(text: str) -> float:
    text = str(text or '')
    if not text:
        return 0.0
    return sum(1 for c in text if c.isascii() and c.isalpha()) / max(len(text), 1)


def _first_run(p: Dict[str, Any]) -> Dict[str, Any]:
    runs = [r for r in (p.get('runs') or []) if r.get('text', '').strip() or r.get('size_pt')]
    if not runs:
        return (p.get('runs') or [{}])[0] if p.get('runs') else {}

    def score(r: Dict[str, Any]) -> float:
        txt = r.get('text', '') or ''
        font = r.get('font', '') or ''
        size = float(r.get('size_pt') or 0)
        return size + (5 if _is_cjk(txt) else 0) + (3 if font and font not in ('Arial', 'Times New Roman', 'Calibri') else 0)

    return max(runs, key=score)


def _normalize_profile(prof: Optional[Dict[str, Any]], fallback: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    fallback = fallback or {}
    p = dict(fallback)
    if prof:
        p.update({k: v for k, v in prof.items() if v is not None})
    p.setdefault('font', '宋体')
    p.setdefault('size', 12)
    p.setdefault('bold', False)
    p.setdefault('italic', False)
    p.setdefault('align', 'LEFT')
    if p.get('align') == 'DEFAULT':
        p['align'] = fallback.get('align', 'LEFT')
    p.setdefault('line_spacing_val', 1.5)
    p.setdefault('line_spacing_fixed_pt', None)
    p.setdefault('space_before_pt', 0)
    p.setdefault('space_after_pt', 0)
    p.setdefault('first_indent_cm', 0)
    try:
        p['size'] = float(p.get('size') or fallback.get('size') or 12)
    except Exception:
        p['size'] = 12.0
    for k in ('space_before_pt', 'space_after_pt', 'first_indent_cm'):
        try:
            p[k] = float(p.get(k) or 0)
        except Exception:
            p[k] = 0.0
    try:
        if p.get('line_spacing_fixed_pt') is not None:
            p['line_spacing_fixed_pt'] = float(p.get('line_spacing_fixed_pt'))
    except Exception:
        p['line_spacing_fixed_pt'] = None
    try:
        if p.get('line_spacing_val') is not None:
            p['line_spacing_val'] = float(p.get('line_spacing_val'))
    except Exception:
        p['line_spacing_val'] = fallback.get('line_spacing_val', 1.5)
    if p.get('line_spacing_val') and p.get('line_spacing_val') > 20 and not p.get('line_spacing_fixed_pt'):
        p['line_spacing_val'] = fallback.get('line_spacing_val', 1.5)
    return p


def _profile_from_para(p: Dict[str, Any], fallback: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    r = _first_run(p)
    prof = {
        'font': r.get('font') or (fallback or {}).get('font') or '宋体',
        'size': r.get('size_pt') or (fallback or {}).get('size') or 12,
        'bold': bool(r.get('bold', (fallback or {}).get('bold', False))),
        'italic': bool(r.get('italic', (fallback or {}).get('italic', False))),
        'align': p.get('alignment') or p.get('align') or (fallback or {}).get('align') or 'LEFT',
        'line_spacing_val': p.get('line_spacing_val') if p.get('line_spacing_val') is not None else p.get('ls', (fallback or {}).get('line_spacing_val')),
        'line_spacing_rule': p.get('line_spacing_rule') or (fallback or {}).get('line_spacing_rule'),
        'line_spacing_fixed_pt': p.get('line_spacing_fixed_pt') or (fallback or {}).get('line_spacing_fixed_pt'),
        'space_before_pt': p.get('space_before_pt') if p.get('space_before_pt') is not None else (fallback or {}).get('space_before_pt', 0),
        'space_after_pt': p.get('space_after_pt') if p.get('space_after_pt') is not None else (fallback or {}).get('space_after_pt', 0),
        'first_indent_cm': p.get('first_indent_cm') if p.get('first_indent_cm') is not None else p.get('indent', (fallback or {}).get('first_indent_cm', 0)),
    }
    return _normalize_profile(prof, fallback)


def _infer_style_profiles(fmt: Dict[str, Any]) -> Dict[str, Dict[str, Any]]:
    existing = {k: _normalize_profile(v) for k, v in (fmt.get('style_profiles') or {}).items()}
    paras = fmt.get('paragraphs') or []
    profiles: Dict[str, Dict[str, Any]] = {}

    def put(role: str, p: Dict[str, Any]) -> None:
        if p and role not in profiles:
            profiles[role] = _profile_from_para(p)

    for p in paras:
        txt = (p.get('text') or '').strip()
        if not txt:
            continue
        compact = re.sub(r'\s+', '', txt)
        up = txt.upper()
        if ('论文' in txt and '题目' in txt and len(txt) < 100 and ('居中' in txt or p.get('style') == '论文题目')):
            put('cn_title', p)
        if re.match(r'^摘\s*要(?:[（(]|$)', txt) and len(txt) < 40:
            put('cn_abstract_heading', p)
        if txt.startswith('摘要是') or ('中文摘要' in txt and len(txt) > 60):
            put('cn_abstract_body', p)
        if txt.startswith('关键词'):
            put('cn_keywords', p)
        if '英文题目' in txt and ('Times' in txt or 'Roman' in txt):
            put('en_title', p)
        if up.startswith('ABSTRACT') and len(txt) < 40:
            put('en_abstract_heading', p)
        if len(txt) > 80 and _ascii_ratio(txt[:160]) > 0.55:
            put('en_abstract_body', p)
        if up.startswith('KEY WORD') or up.startswith('KEYWORDS'):
            put('en_keywords', p)
        if compact in ('目录', '目目录') or compact.startswith('目录'):
            put('toc_title', p)
        if ('一级标题' in txt or re.match(r'^第[一二三四五六七八九十\d]+章\s+', txt)) and len(txt) < 100:
            put('h1', p)
        if ('二级标题' in txt or re.match(r'^\d+\.\d+\s+', txt)) and len(txt) < 100:
            put('h2', p)
        if ('三级标题' in txt or re.match(r'^\d+\.\d+\.\d+\s+', txt)) and len(txt) < 100:
            put('h3', p)
        if re.match(r'^(图|表)\s*\d+', txt) and len(txt) < 80:
            put('figure_caption' if txt.startswith('图') else 'table_caption', p)
        if '参考文献' in txt and len(txt) < 40:
            put('reference_heading', p)

    def heading_score(p: Dict[str, Any], level: int) -> float:
        txt = (p.get('text') or '').strip()
        r = _first_run(p)
        font = r.get('font') or ''
        size = float(r.get('size_pt') or 0)
        score = size
        if p.get('style', '').lower().startswith('heading'):
            score += 20
        if '\t' in txt or re.search(r'\s\d+$', txt):
            score -= 18
        if font and font not in ('Arial', 'Times New Roman', 'Calibri'):
            score += 8
        if level == 1 and re.match(r'^第[一二三四五六七八九十\d]+章\s+', txt):
            score += 10
        if level == 2 and re.match(r'^\d+\.\d+\s+', txt):
            score += 10
        if level == 3 and re.match(r'^\d+\.\d+\.\d+\s+', txt):
            score += 10
        return score

    for role, level, pat in [
        ('h1', 1, r'^第[一二三四五六七八九十\d]+章\s+'),
        ('h2', 2, r'^\d+\.\d+\s+'),
        ('h3', 3, r'^\d+\.\d+\.\d+\s+'),
    ]:
        cands = [p for p in paras if re.match(pat, (p.get('text') or '').strip()) and len((p.get('text') or '').strip()) < 100]
        if cands:
            profiles[role] = _profile_from_para(max(cands, key=lambda x: heading_score(x, level)))
            profiles[role]['bold'] = True
            profiles[role]['first_indent_cm'] = 0
            if level == 1:
                profiles[role]['align'] = 'CENTER'

    body_cands = []
    for p in paras:
        txt = (p.get('text') or '').strip()
        if len(txt) < 80:
            continue
        if any(k in txt[:80] for k in ['本人郑重声明', '本人在导师', '原创性声明', '版权使用', '格式要求', '字体要求', '行距：', '字号', '页眉页脚']):
            continue
        if _is_cjk(txt):
            body_cands.append(p)
    if body_cands:
        put('body', body_cands[0])

    for role, prof in existing.items():
        profiles.setdefault(role, prof)

    body = _normalize_profile(profiles.get('body') or {'font': '宋体', 'size': 12, 'align': 'JUSTIFY', 'line_spacing_fixed_pt': 28, 'first_indent_cm': 0.74})
    profiles['body'] = body
    profiles.setdefault('h1', _normalize_profile({'font': '黑体', 'size': 16, 'bold': True, 'align': 'CENTER', 'line_spacing_fixed_pt': body.get('line_spacing_fixed_pt'), 'first_indent_cm': 0}, body))
    profiles.setdefault('h2', _normalize_profile({'font': '黑体', 'size': 14, 'bold': True, 'align': 'LEFT', 'line_spacing_fixed_pt': body.get('line_spacing_fixed_pt'), 'first_indent_cm': 0}, body))
    profiles.setdefault('h3', _normalize_profile({'font': '黑体', 'size': 12, 'bold': True, 'align': 'LEFT', 'line_spacing_fixed_pt': body.get('line_spacing_fixed_pt'), 'first_indent_cm': 0}, body))
    profiles.setdefault('cn_title', profiles['h1'])
    profiles.setdefault('cn_abstract_heading', profiles['h1'])
    profiles.setdefault('cn_abstract_body', body)
    profiles.setdefault('cn_keywords', _normalize_profile({'first_indent_cm': 0, 'bold': False}, body))
    profiles.setdefault('en_title', _normalize_profile({'font': 'Times New Roman', 'bold': True, 'align': 'CENTER'}, profiles['h1']))
    profiles.setdefault('en_abstract_heading', _normalize_profile({'font': 'Times New Roman', 'size': 16, 'bold': True, 'align': 'CENTER', 'first_indent_cm': 0}, profiles['h1']))
    profiles.setdefault('en_abstract_body', _normalize_profile({'font': 'Times New Roman', 'line_spacing_fixed_pt': None, 'line_spacing_val': 1.5, 'first_indent_cm': 0.9, 'align': 'JUSTIFY'}, body))
    profiles.setdefault('en_keywords', _normalize_profile({'font': 'Times New Roman', 'bold': False, 'first_indent_cm': 0, 'align': 'LEFT'}, profiles['en_abstract_body']))
    profiles.setdefault('toc_title', profiles['h1'])
    profiles.setdefault('figure_caption', _normalize_profile({'font': '宋体', 'size': 10.5, 'align': 'CENTER', 'first_indent_cm': 0, 'space_before_pt': 6, 'space_after_pt': 6, 'line_spacing_fixed_pt': 28}, body))
    profiles.setdefault('table_caption', dict(profiles['figure_caption']))
    profiles.setdefault('code', _normalize_profile({'font': 'Consolas', 'size': 10.5, 'align': 'LEFT', 'first_indent_cm': 0, 'line_spacing_fixed_pt': None, 'line_spacing_val': 1.0}, body))
    profiles.setdefault('reference', _normalize_profile({'font': '宋体', 'size': 12, 'align': 'JUSTIFY', 'first_indent_cm': 0, 'space_before_pt': 6, 'space_after_pt': 6, 'line_spacing_fixed_pt': 28}, body))
    profiles.setdefault('reference_heading', profiles['h1'])
    return profiles


def _extract_page_and_header(fmt: Dict[str, Any]) -> Dict[str, Any]:
    sections = fmt.get('sections') or []
    s0 = sections[0] if sections else {}
    page = {
        'page_w': s0.get('page_width_cm', 21.0),
        'page_h': s0.get('page_height_cm', 29.7),
        'mt': s0.get('margin_top_cm', 2.54),
        'mb': s0.get('margin_bottom_cm', 2.54),
        'ml': s0.get('margin_left_cm', 2.54),
        'mr': s0.get('margin_right_cm', 2.54),
        'header': None,
    }
    for sec in sections:
        for h in sec.get('header', []) or []:
            text = (h.get('text') or '').strip()
            if not text:
                continue
            run = next((r for r in h.get('runs', []) if r.get('text', '').strip() or r.get('size_pt')), {})
            page['header'] = {
                'text': text,
                'align': h.get('alignment') if h.get('alignment') != 'DEFAULT' else 'CENTER',
                'font': run.get('font') or '宋体',
                'size': run.get('size_pt') or 9,
                'bold': bool(run.get('bold', False)),
                'italic': bool(run.get('italic', False)),
            }
            return page
    return page


def _section_role(sec: Dict[str, Any]) -> str:
    role = (sec.get('role') or '').strip()
    if role:
        return role
    h = (sec.get('heading') or '').strip()
    compact = re.sub(r'[\s：:]+', '', h).lower()
    if compact in ('摘要', '中文摘要'):
        return 'cn_abstract'
    if h.startswith('关键词') or compact in ('关键词', '关键字'):
        return 'cn_keywords'
    if compact == 'abstract':
        return 'en_abstract'
    if h.upper().replace(' ', '').startswith('KEYWORDS') or re.match(r'(?i)^key\s*words?', h):
        return 'en_keywords'
    if h.startswith('参考文献') or re.match(r'(?i)^references?$', h):
        return 'references'
    if re.search(r'致\s*谢', h):
        return 'acknowledgement'
    if re.search(r'附\s*录', h):
        return 'appendix'
    return 'body'


def _front_matter_sections(cnt: Dict[str, Any]) -> Dict[str, Any]:
    result: Dict[str, Any] = {'cn_abs': None, 'cn_kw': None, 'en_title': '', 'en_abs': None, 'en_kw': None, 'front_indices': set()}
    sections = cnt.get('sections') or []
    for idx, sec in enumerate(sections):
        role = _section_role(sec)
        h = (sec.get('heading') or '').strip()
        if role == 'cn_abstract':
            result['cn_abs'] = sec; result['front_indices'].add(idx)
        elif role == 'cn_keywords':
            result['cn_kw'] = sec; result['front_indices'].add(idx)
        elif role == 'en_abstract':
            result['en_abs'] = sec; result['front_indices'].add(idx)
        elif role == 'en_keywords':
            result['en_kw'] = sec; result['front_indices'].add(idx)
        elif sec.get('level') == 1 and _ascii_ratio(h) > 0.55 and not re.match(r'(?i)^chapter\s+\d+', h):
            if not result['en_title']:
                result['en_title'] = h; result['front_indices'].add(idx)
    return result


RUNTIME_TEMPLATE = r'''
# -*- coding: utf-8 -*-
"""
build_generated.py — generated by role-driven script_generator.py.
运行: python build_generated.py
"""
import json
import os
import re

from docx import Document
from docx.shared import Pt, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

DATA = json.loads(__DATA_BLOB__)
BASE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(BASE, __OUT_DOCX__)

doc = Document()

ALIGN = {
    'LEFT': WD_ALIGN_PARAGRAPH.LEFT,
    'CENTER': WD_ALIGN_PARAGRAPH.CENTER,
    'RIGHT': WD_ALIGN_PARAGRAPH.RIGHT,
    'JUSTIFY': WD_ALIGN_PARAGRAPH.JUSTIFY,
    'DISTRIBUTE': WD_ALIGN_PARAGRAPH.DISTRIBUTE,
    'DEFAULT': WD_ALIGN_PARAGRAPH.LEFT,
}

CJK_FONTS = {'宋体', '黑体', '楷体', '微软雅黑', '仿宋', '华文宋体', '华文中宋'}


def has_cjk(text):
    return any('\u4e00' <= c <= '\u9fff' for c in str(text or ''))


def ascii_ratio(text):
    text = str(text or '')
    return sum(1 for c in text if c.isascii() and c.isalpha()) / max(len(text), 1)


def set_run_fonts(run, east_asia='宋体', latin=None):
    latin = latin or ('Times New Roman' if east_asia in CJK_FONTS else east_asia)
    run.font.name = latin
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:ascii'), latin)
    rFonts.set(qn('w:hAnsi'), latin)
    rFonts.set(qn('w:eastAsia'), east_asia)
    rFonts.set(qn('w:hint'), 'eastAsia')


def profile(name):
    return DATA['profiles'].get(name) or DATA['profiles']['body']


def apply_paragraph_profile(p, prof, outline_level=None, first_indent_override=None):
    p.alignment = ALIGN.get(prof.get('align') or 'LEFT', WD_ALIGN_PARAGRAPH.LEFT)
    pf = p.paragraph_format
    fixed = prof.get('line_spacing_fixed_pt')
    if fixed:
        pf.line_spacing = Pt(float(fixed))
    elif prof.get('line_spacing_val'):
        pf.line_spacing = float(prof.get('line_spacing_val'))
    if prof.get('space_before_pt') is not None:
        pf.space_before = Pt(float(prof.get('space_before_pt') or 0))
    if prof.get('space_after_pt') is not None:
        pf.space_after = Pt(float(prof.get('space_after_pt') or 0))
    indent = first_indent_override if first_indent_override is not None else prof.get('first_indent_cm')
    if indent:
        pf.first_line_indent = Cm(float(indent))
    else:
        pf.first_line_indent = Cm(0)
    if outline_level is not None:
        pPr = p._element.get_or_add_pPr()
        old = pPr.find(qn('w:outlineLvl'))
        if old is not None:
            pPr.remove(old)
        ol = OxmlElement('w:outlineLvl')
        ol.set(qn('w:val'), str(int(outline_level)))
        pPr.append(ol)


def apply_run_profile(run, prof, text='', superscript=False, force_latin=None):
    font = prof.get('font') or '宋体'
    latin = force_latin
    if not latin:
        if font in CJK_FONTS:
            latin = 'Times New Roman'
        else:
            latin = font
    east_asia = font if font in CJK_FONTS else font
    if ascii_ratio(text) > 0.75 and font in CJK_FONTS:
        east_asia = font
        latin = 'Times New Roman'
    set_run_fonts(run, east_asia, latin)
    run.font.size = Pt(float(prof.get('size') or 12))
    run.bold = bool(prof.get('bold', False))
    run.italic = bool(prof.get('italic', False))
    run.font.superscript = bool(superscript)


def citation_pattern():
    return re.compile(r'(\[\d+(?:\s*[-,，、]\s*\d+)*\])')


def add_text_runs(p, text, prof, superscript_citations=False):
    text = str(text or '')
    pos = 0
    pat = citation_pattern() if superscript_citations else None
    matches = list(pat.finditer(text)) if pat else []
    if not matches:
        r = p.add_run(text)
        apply_run_profile(r, prof, text)
        return
    for m in matches:
        if m.start() > pos:
            seg = text[pos:m.start()]
            r = p.add_run(seg)
            apply_run_profile(r, prof, seg)
        r = p.add_run(m.group(1))
        apply_run_profile(r, prof, m.group(1), superscript=True)
        pos = m.end()
    if pos < len(text):
        seg = text[pos:]
        r = p.add_run(seg)
        apply_run_profile(r, prof, seg)


def add_text(text, role='body', first_indent=True, outline_level=None, bold_prefix=None):
    prof = profile(role)
    p = doc.add_paragraph()
    apply_paragraph_profile(p, prof, outline_level=outline_level, first_indent_override=(prof.get('first_indent_cm') if first_indent else 0))
    superscript = role == 'body'
    text = str(text or '')
    if bold_prefix and text.startswith(bold_prefix):
        r1 = p.add_run(bold_prefix)
        p1 = dict(prof); p1['bold'] = True
        apply_run_profile(r1, p1, bold_prefix)
        add_text_runs(p, text[len(bold_prefix):], prof, superscript)
    else:
        add_text_runs(p, text, prof, superscript)
    return p


def normalize_heading_spacing(text):
    t = str(text or '').strip()
    t = re.sub(r'^(第[一二三四五六七八九十百千万\d]+章)\s*(\S)', r'\1 \2', t)
    t = re.sub(r'^(\d+(?:\.\d+)*)(?![.\d])\s*(\S)', r'\1 \2', t)
    return t


def add_heading(text, level):
    level = max(1, min(int(level or 1), 3))
    return add_text(normalize_heading_spacing(text), role='h' + str(level), first_indent=False, outline_level=level - 1)


def setup_section(sec):
    page = DATA['page']
    sec.page_width = Cm(float(page.get('page_w') or 21.0))
    sec.page_height = Cm(float(page.get('page_h') or 29.7))
    sec.top_margin = Cm(float(page.get('mt') or 2.54))
    sec.bottom_margin = Cm(float(page.get('mb') or 2.54))
    sec.left_margin = Cm(float(page.get('ml') or 2.54))
    sec.right_margin = Cm(float(page.get('mr') or 2.54))


def add_page_field(paragraph):
    r = paragraph.add_run()
    begin = OxmlElement('w:fldChar'); begin.set(qn('w:fldCharType'), 'begin')
    instr = OxmlElement('w:instrText'); instr.set(qn('xml:space'), 'preserve'); instr.text = ' PAGE '
    end = OxmlElement('w:fldChar'); end.set(qn('w:fldCharType'), 'end')
    r._element.append(begin); r._element.append(instr); r._element.append(end)
    return r


def set_page_numbering(sec, fmt='decimal', start=1):
    sectPr = sec._sectPr
    pg = sectPr.find(qn('w:pgNumType'))
    if pg is None:
        pg = OxmlElement('w:pgNumType')
        sectPr.append(pg)
    pg.set(qn('w:start'), str(start))
    pg.set(qn('w:fmt'), fmt)


def apply_header_footer(sec, page_fmt='decimal', start=1):
    hdr = DATA['page'].get('header') or {}
    sec.header.is_linked_to_previous = False
    sec.footer.is_linked_to_previous = False
    hp = sec.header.paragraphs[0]
    hp.text = ''
    hp.alignment = ALIGN.get(hdr.get('align', 'CENTER'), WD_ALIGN_PARAGRAPH.CENTER)
    if hdr.get('text'):
        r = hp.add_run(hdr['text'])
        apply_run_profile(r, {'font': hdr.get('font') or '宋体', 'size': hdr.get('size') or 9, 'bold': hdr.get('bold', False), 'italic': hdr.get('italic', False)}, hdr['text'])
        pPr = hp._element.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'), 'single'); bottom.set(qn('w:sz'), '4'); bottom.set(qn('w:space'), '1'); bottom.set(qn('w:color'), 'auto')
        pBdr.append(bottom); pPr.append(pBdr)
    fp = sec.footer.paragraphs[0]
    fp.text = ''
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = add_page_field(fp)
    apply_run_profile(r, {'font': hdr.get('font') or '宋体', 'size': hdr.get('size') or 9}, '')
    set_page_numbering(sec, page_fmt, start)


def clear_header_footer(sec):
    sec.header.is_linked_to_previous = False
    sec.footer.is_linked_to_previous = False
    for part in (sec.header, sec.footer):
        for p in part.paragraphs:
            p.text = ''


def add_section_with_header(page_fmt='decimal', start=1):
    sec = doc.add_section(WD_SECTION.NEW_PAGE)
    setup_section(sec)
    apply_header_footer(sec, page_fmt, start)
    return sec


def enable_update_fields_on_open():
    settings = doc.settings._element
    upd = settings.find(qn('w:updateFields'))
    if upd is None:
        upd = OxmlElement('w:updateFields')
        settings.append(upd)
    upd.set(qn('w:val'), 'true')


def remove_initial_empty_paragraph():
    if len(doc.paragraphs) == 1 and not doc.paragraphs[0].text.strip():
        el = doc.paragraphs[0]._element
        el.getparent().remove(el)


def force_cover_headerless():
    if not doc.sections:
        return
    sec = doc.sections[0]
    sec.different_first_page_header_footer = True
    for part in (sec.header, sec.first_page_header, sec.footer, sec.first_page_footer):
        part.is_linked_to_previous = False
        for p in part.paragraphs:
            p.text = ''


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


def add_asset_picture(run, rd, default_inches=1.2):
    path = asset_path(rd.get('asset') or rd.get('image'))
    if not path:
        return False
    try:
        run.add_picture(path, width=image_width_from_extent(rd.get('extent'), default_inches))
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


def render_school_or_degree_code_line(el, first_key):
    """学校编码/学位编码不要按两列表格渲染，直接渲染为左侧一行。"""
    cover_info = DATA.get('cover_info') or {}

    code_value = ''
    if '学校编码' in first_key:
        code_value = cover_info.get('school_code', '') or cover_info.get('degree_code', '')
        label = '学校编码：'
    else:
        code_value = cover_info.get('degree_code', '') or cover_info.get('school_code', '')
        label = '学位编码：'

    if not code_value:
        try:
            code_value = ''.join(
                r.get('t', '')
                for r in el.get('rows', [])[0][1].get('p', [])[0].get('r', [])
            ).strip()
        except Exception:
            code_value = ''

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.left_indent = Cm(0)
    p.paragraph_format.first_line_indent = Cm(0)

    rd_label = {}
    rd_value = {}
    try:
        rd_label = el['rows'][0][0]['p'][0]['r'][0]
    except Exception:
        pass
    try:
        rd_value = el['rows'][0][1]['p'][0]['r'][0]
    except Exception:
        pass

    r1 = p.add_run(label)
    apply_cover_run(r1, rd_label)

    r2 = p.add_run('    ' + code_value)
    apply_cover_run(r2, rd_value or rd_label)

    return p


def render_cover_table(el):
    rows = el.get('rows', [])
    if not rows:
        return None

    first_key = ''
    if rows and rows[0] and rows[0][0].get('p'):
        first_key = normalize_label(
            ''.join(r.get('t', '') for r in rows[0][0]['p'][0].get('r', []))
        )

    if '学校编码' in first_key or '学位编码' in first_key:
        return render_school_or_degree_code_line(el, first_key)

    ncols = max(len(r) for r in rows)
    table = doc.add_table(rows=len(rows), cols=ncols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
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
    for ri, row in enumerate(rows):
        row_label = ''
        if row and row[0].get('p'):
            row_label = ''.join(r.get('t', '') for r in row[0]['p'][0].get('r', []))
        row_key = normalize_label(row_label)
        row_value = norm_map.get(row_key, '')
        if not row_value:
            for k, v in norm_map.items():
                if k and (k in row_key or row_key in k):
                    row_value = v; break
        force_left = ('学位编码' in row_key or '学校编码' in row_key)
        for ci in range(ncols):
            cell = table.rows[ri].cells[ci]
            cell.text = ''
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            cell_data = row[ci] if ci < len(row) else {'p': []}
            if cell_data.get('w'):
                try: cell.width = Cm(float(cell_data.get('w')) / 567.0)
                except Exception: pass
            paras = cell_data.get('p') or [{'r': []}]
            for pi, pp in enumerate(paras):
                p = cell.paragraphs[0] if pi == 0 else cell.add_paragraph()
                apply_cover_paragraph_format(p, pp)
                if force_left:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                raw = ''.join(r.get('t', '') for r in pp.get('r', []))
                use_value = (ci == 1 and row_value)
                if use_value:
                    rd = (pp.get('r') or [{}])[0]
                    rr = p.add_run(row_value)
                    apply_cover_run(rr, rd)
                    continue
                for rd in pp.get('r', []) or [{}]:
                    rr = p.add_run(rd.get('t', ''))
                    apply_cover_run(rr, rd)
                    if rd.get('asset') or rd.get('image'):
                        add_asset_picture(rr, rd)
            borders = cell_data.get('borders') or {}
            if borders:
                set_cell_borders(cell, **borders)
    return table


def is_empty_cover_element(el):
    return el.get('type') == 'empty' and not para_text_from_cover_el(el).strip()


def next_nonempty_cover_text(cover, idx):
    for j in range(idx + 1, len(cover)):
        t = para_text_from_cover_el(cover[j]).strip()
        if t:
            return t
    return ''


def render_cover_and_declarations():
    setup_section(doc.sections[0])
    clear_header_footer(doc.sections[0])
    cover = DATA.get('cover') or []
    if not cover:
        return add_section_with_header('upperRoman', 1)
    front_started = False
    empty_run = 0
    for idx, el in enumerate(cover):
        text = para_text_from_cover_el(el)

        if is_empty_cover_element(el):
            empty_run += 1
            next_text = next_nonempty_cover_text(cover, idx)
            if '论文题目' in next_text and empty_run > 1:
                continue
            if '学位评定委员会' in next_text and empty_run > 1:
                continue
        else:
            empty_run = 0

        if (not front_started) and ('原创性声明' in text or '版权使用授权书' in text):
            front_started = True
            add_section_with_header('upperRoman', 1)

        if el.get('type') in ('para', 'empty'):
            render_cover_para(el)
        elif el.get('type') == 'table':
            render_cover_table(el)
        elif el.get('type') == 'image':
            render_cover_image(el)
    if not front_started:
        add_section_with_header('upperRoman', 1)
    return doc.sections[-1]


def section_text(sec):
    out = []
    for para in sec.get('paragraphs', []) or []:
        if isinstance(para, str):
            out.append(para.strip())
        elif isinstance(para, dict) and para.get('text'):
            out.append(str(para.get('text')).strip())
    return '\n'.join(x for x in out if x)


def add_keywords(label, value, role):
    prof = profile(role)
    p = doc.add_paragraph()
    apply_paragraph_profile(p, prof, first_indent_override=0)
    r1 = p.add_run(label)
    p_bold = dict(prof); p_bold['bold'] = True
    apply_run_profile(r1, p_bold, label)
    r2 = p.add_run(value)
    p_norm = dict(prof); p_norm['bold'] = False
    apply_run_profile(r2, p_norm, value)
    return p


def add_blank_line(role='body'):
    return add_text('', role=role, first_indent=False)


def add_toc():
    enable_update_fields_on_open()
    add_text('目录', role='toc_title', first_indent=False)
    p = doc.add_paragraph()
    r = p.add_run()
    begin = OxmlElement('w:fldChar'); begin.set(qn('w:fldCharType'), 'begin'); begin.set(qn('w:dirty'), 'true')
    instr = OxmlElement('w:instrText'); instr.set(qn('xml:space'), 'preserve'); instr.text = ' TOC \\o "1-3" \\h \\z \\u '
    sep = OxmlElement('w:fldChar'); sep.set(qn('w:fldCharType'), 'separate')
    end = OxmlElement('w:fldChar'); end.set(qn('w:fldCharType'), 'end')
    r._element.append(begin); r._element.append(instr); r._element.append(sep); r._element.append(end)
    return p


def render_front_matter():
    front = DATA.get('front') or {}
    title_cn = DATA.get('title_cn') or ''
    if title_cn:
        add_text(title_cn, role='cn_title', first_indent=False)
    cn_abs = front.get('cn_abs')
    if cn_abs:
        add_text('摘 要', role='cn_abstract_heading', first_indent=False)
        for para in cn_abs.get('paragraphs', []) or []:
            text = para if isinstance(para, str) else para.get('text', '')
            if str(text).strip():
                add_text(str(text).strip(), role='cn_abstract_body', first_indent=True)
    cn_kw = front.get('cn_kw')
    if cn_kw and cn_abs:
        add_blank_line('cn_abstract_body')
    if cn_kw:
        val = section_text(cn_kw)
        if val:
            add_keywords('关键词：', val, 'cn_keywords')
    has_en = bool(front.get('en_title') or front.get('en_abs') or front.get('en_kw'))
    if has_en:
        doc.add_page_break()
    en_title = front.get('en_title') or ''
    if en_title:
        add_text(en_title, role='en_title', first_indent=False)
    en_abs = front.get('en_abs')
    if en_abs:
        add_text('ABSTRACT', role='en_abstract_heading', first_indent=False)
        for para in en_abs.get('paragraphs', []) or []:
            text = para if isinstance(para, str) else para.get('text', '')
            if str(text).strip():
                add_text(str(text).strip(), role='en_abstract_body', first_indent=True)
    en_kw = front.get('en_kw')
    if en_kw:
        val = section_text(en_kw).replace('；', ';')
        if val:
            add_keywords('KEY WORDS: ', val, 'en_keywords')
    add_section_with_header('upperRoman', 1)
    add_toc()


def is_front_section_index(i):
    return i in set(DATA.get('front_indices') or [])


def is_reference_heading(h):
    h = str(h or '').strip()
    return h.startswith('参考文献') or bool(re.match(r'(?i)^references?$', h))


def is_ack_heading(h):
    return bool(re.search(r'致\s*谢', str(h or '')))


def is_appendix_heading(h):
    return bool(re.search(r'附\s*录', str(h or '')))


def is_backmatter_heading(h):
    return is_ack_heading(h) or is_appendix_heading(h)


def normalize_caption(text):
    t = str(text or '').strip()
    t = re.sub(r'^(图|表)\s*(\d+(?:[.-]\d+)?)\s*', r'\1 \2 ', t)
    return t.strip()


def add_caption(text, role='figure_caption'):
    return add_text(normalize_caption(text), role=role, first_indent=False)


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


def add_code_block(text):
    prof = profile('code')
    p = doc.add_paragraph()
    apply_paragraph_profile(p, prof, first_indent_override=0)
    p.paragraph_format.left_indent = Cm(0.74)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    try:
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    except Exception:
        pass
    r = p.add_run(str(text).rstrip())
    apply_run_profile(r, prof, text, force_latin=prof.get('font') or 'Consolas')
    return p


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


def render_table(rows):
    if not rows:
        return
    ncols = max(len(r) for r in rows)
    table = doc.add_table(rows=len(rows), cols=ncols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    prof = profile('body')
    for ri, row in enumerate(rows):
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
    return table


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
        return
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run()
    try:
        r.add_picture(path, width=Inches(4.2))
    except Exception:
        return
    if caption:
        add_caption(caption)


def clean_ref_text(ref):
    text = re.sub(r'\s+', ' ', str(ref or '')).strip()
    if text.startswith('[') and ']' in text:
        prefix, rest = text.split(']', 1)
        if prefix[1:].isdigit():
            return rest.strip()
    parts = text.split(None, 1)
    if len(parts) == 2 and parts[0].strip('.、[]').isdigit():
        return parts[1].strip()
    return text


def split_refs_backmatter(refs):
    pure, ack, app = [], [], []
    mode = 'refs'
    for raw in refs or []:
        text = str(raw or '').strip()
        if not text:
            continue
        if is_ack_heading(text):
            mode = 'ack'; continue
        if is_appendix_heading(text):
            mode = 'app'
            if normalize_label(text) != '附录':
                app.append(text)
            continue
        (pure if mode == 'refs' else ack if mode == 'ack' else app).append(text)
    return pure, ack, app


def add_reference_mixed_runs(p, text, prof):
    # Chinese parts use the role's CJK font; Latin/numeric punctuation uses Times New Roman.
    for seg in re.findall(r'[\u4e00-\u9fff]+|[^\u4e00-\u9fff]+', text):
        r = p.add_run(seg)
        if has_cjk(seg):
            apply_run_profile(r, prof, seg, force_latin='Times New Roman')
        else:
            p_latin = dict(prof); p_latin['font'] = 'Times New Roman'
            apply_run_profile(r, p_latin, seg, force_latin='Times New Roman')


def render_reference_entries(refs):
    if not refs:
        return
    doc.add_page_break()
    add_text('参考文献', role='reference_heading', first_indent=False, outline_level=0)
    prof = profile('reference')
    for idx, raw in enumerate(refs, 1):
        text = clean_ref_text(raw)
        if not text:
            continue
        p = doc.add_paragraph()
        apply_paragraph_profile(p, prof, first_indent_override=0)
        p.paragraph_format.left_indent = Cm(0.74)
        p.paragraph_format.first_line_indent = Cm(-0.74)
        p.paragraph_format.keep_together = True
        add_reference_mixed_runs(p, '[' + str(idx) + '] ' + text, prof)


def render_backmatter_section(title, paragraphs, code_sensitive=False):
    if not paragraphs:
        return
    doc.add_page_break()
    add_heading(title, 1)
    for item in paragraphs:
        if isinstance(item, dict) and (item.get('code') or item.get('role') == 'code'):
            add_code_block(item.get('code') or item.get('text') or '')
            continue
        text = str(item.get('text') if isinstance(item, dict) else item or '').strip()
        if not text:
            continue
        if code_sensitive and looks_like_code_line(text):
            add_code_block(text)
        else:
            add_text(text, role='body', first_indent=not code_sensitive)


def collect_structural_backmatter():
    ack, app = [], []
    mode = None
    for i, sec in enumerate(DATA.get('sections') or []):
        if is_front_section_index(i):
            continue
        h = (sec.get('heading') or '').strip()
        role = sec.get('role') or ''
        if role == 'acknowledgement' or is_ack_heading(h):
            mode = 'ack'; continue
        if role == 'appendix' or is_appendix_heading(h):
            mode = 'app'
            if normalize_label(h) != '附录':
                app.append(h)
            continue
        if mode in ('ack', 'app'):
            if sec.get('level') and h:
                (ack if mode == 'ack' else app).append(h)
            for para in sec.get('paragraphs', []) or []:
                (ack if mode == 'ack' else app).append(para)
    return ack, app


def render_paragraph_item(item, code_sensitive=False):
    if isinstance(item, dict) and item.get('table_rows'):
        render_table(item.get('table_rows') or [])
        return
    if isinstance(item, dict) and (item.get('role') == 'figure_caption'):
        add_caption(item.get('text') or '', 'figure_caption')
        return
    if isinstance(item, dict) and (item.get('role') == 'table_caption'):
        add_caption(item.get('text') or '', 'table_caption')
        return
    if isinstance(item, dict) and (item.get('code') or item.get('role') == 'code'):
        add_code_block(item.get('code') or item.get('text') or '')
        return
    text = str(item.get('text') if isinstance(item, dict) else item or '').strip()
    if not text:
        return
    if len(text) > 20 and any(k in text[:80] for k in ['完成后删除', '格式要求', '字体要求', '页眉页脚']):
        return
    if code_sensitive and looks_like_code_line(text):
        add_code_block(text)
    elif re.match(r'^图\s*\d+', text):
        add_caption(text, 'figure_caption')
    elif re.match(r'^表\s*\d+', text):
        add_caption(text, 'table_caption')
    else:
        add_text(text, role='body', first_indent=True)


def render_body():
    add_section_with_header('decimal', 1)
    fig_no = 0
    for i, sec in enumerate(DATA.get('sections') or []):
        if is_front_section_index(i):
            continue
        h = (sec.get('heading') or '').strip()
        role = sec.get('role') or ''
        if is_reference_heading(h) or is_backmatter_heading(h) or role in ('references', 'acknowledgement', 'appendix'):
            continue
        if h and h != '正文':
            add_heading(h, sec.get('level') or 1)
        for img in sec.get('images', []) or []:
            fig_no += 1
            render_image(img, '图 ' + str(fig_no) + ' ' + re.sub(r'^第[一二三四五六七八九十\d]+章\s*', '', h))
        for para in sec.get('paragraphs', []) or []:
            render_paragraph_item(para, code_sensitive=False)
    pure_refs, ack_from_refs, app_from_refs = split_refs_backmatter(DATA.get('references') or [])
    ack_sections, app_sections = collect_structural_backmatter()
    render_reference_entries(pure_refs)
    render_backmatter_section('致  谢', ack_sections or ack_from_refs, code_sensitive=False)
    render_backmatter_section('附  录', app_sections or app_from_refs, code_sensitive=True)


def main():
    setup_section(doc.sections[0])
    clear_header_footer(doc.sections[0])
    remove_initial_empty_paragraph()
    render_cover_and_declarations()
    render_front_matter()
    render_body()
    force_cover_headerless()
    doc.save(OUT)
    print('Saved:', OUT)


if __name__ == '__main__':
    main()
'''


def generate(format_json_path: str, content_json_path: str, output_dir: str, output_docx_name: str = '最终论文.docx') -> int:
    os.makedirs(output_dir, exist_ok=True)
    with open(format_json_path, 'r', encoding='utf-8') as f:
        fmt = json.load(f)
    with open(content_json_path, 'r', encoding='utf-8') as f:
        cnt = json.load(f)

    profiles = _infer_style_profiles(fmt)
    page = _extract_page_and_header(fmt)
    front = _front_matter_sections(cnt)
    cover_info = cnt.get('cover_info', {}) or {}
    title_cn = cover_info.get('paper_title') or cnt.get('title_info', {}).get('title_cn') or ''

    data_blob = {
        'fmt_meta': fmt.get('_meta', {}),
        'content_meta': cnt.get('_meta', {}),
        'page': page,
        'profiles': profiles,
        'cover': fmt.get('cover', []),
        'cover_info': cover_info,
        'title_cn': title_cn,
        'sections': cnt.get('sections', []),
        'references': cnt.get('references', []),
        'front_indices': sorted(front['front_indices']),
        'front': {k: v for k, v in front.items() if k != 'front_indices'},
        'images_dir': cnt.get('_meta', {}).get('images_dir') or '',
        'assets_dir': fmt.get('_meta', {}).get('assets_dir') or '',
    }
    blob = json.dumps(data_blob, ensure_ascii=False, indent=2)
    code = RUNTIME_TEMPLATE.replace('__DATA_BLOB__', repr(blob)).replace('__OUT_DOCX__', repr(os.path.basename(output_docx_name)))
    out_py = os.path.join(output_dir, 'build_generated.py')
    with open(out_py, 'w', encoding='utf-8') as f:
        f.write(code)
    return len(code)


if __name__ == '__main__':
    import sys
    if len(sys.argv) < 4:
        print('Usage: python script_generator.py format.json content.json output_dir [output.docx]')
        raise SystemExit(2)
    out_name = sys.argv[4] if len(sys.argv) > 4 else '最终论文.docx'
    n = generate(sys.argv[1], sys.argv[2], sys.argv[3], out_name)
    print(f'Generated build script, {n} chars')
