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
import shutil
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


def _first_text_run(p: Dict[str, Any]) -> Dict[str, Any]:
    for r in p.get('runs') or []:
        if str(r.get('text') or '').strip():
            return r
    return _first_run(p)


def _profile_from_para_first_text(p: Dict[str, Any], fallback: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    r = _first_text_run(p)
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
            profiles.setdefault('en_title', _profile_from_para_first_text(p))
        if up.startswith('ABSTRACT') and len(txt) < 40:
            profiles.setdefault('en_abstract_heading', _profile_from_para_first_text(p))
        if len(txt) > 80 and _ascii_ratio(txt[:160]) > 0.55:
            put('en_abstract_body', p)
        if up.startswith('KEY WORD') or up.startswith('KEYWORDS'):
            profiles.setdefault('en_keywords', _profile_from_para_first_text(p))
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
            run = next((r for r in h.get('runs', []) if str(r.get('text', '')).strip()), next((r for r in h.get('runs', []) if r.get('size_pt')), {}))
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


def _infer_template_rules(fmt: Dict[str, Any]) -> Dict[str, Any]:
    """Infer layout rules from template instruction text, not from fixed school names."""
    texts = '\n'.join(str(p.get('text') or '') for p in fmt.get('paragraphs') or [])
    return {
        'cn_abstract_single_paragraph': bool(re.search(r'中文摘要[^\n。；;]{0,80}不分自然段|不分自然段[^\n。；;]{0,80}中文摘要', texts)),
        'en_title_upper': bool(re.search(r'英文题目[^\n。；;]{0,120}(大写字母|大写)', texts)),
    }


RUNTIME_TEMPLATE = r'''
# -*- coding: utf-8 -*-
"""
build_generated.py — generated by role-driven script_generator.py.
运行: python build_generated.py
"""
import json
import os
import re
import shutil
import subprocess
import tempfile

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
    # New sections created by python-docx copy the previous section's
    # w:start.  If this continuation section should not restart numbering,
    # remove the copied attribute instead of leaving another start=1 behind.
    if start is None:
        start_attr = qn('w:start')
        if start_attr in pg.attrib:
            del pg.attrib[start_attr]
    else:
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


def _paragraph_is_empty_for_cleanup(p):
    if p.text.strip():
        return False
    xml = p._element.xml
    if '<w:drawing' in xml or '<w:pict' in xml:
        return False
    pPr = p._element.find(qn('w:pPr'))
    if pPr is not None and pPr.find(qn('w:sectPr')) is not None:
        return False
    return True


def remove_trailing_empty_body_paragraphs(limit=12):
    removed = 0
    while len(doc.paragraphs) > 1 and removed < limit and _paragraph_is_empty_for_cleanup(doc.paragraphs[-1]):
        el = doc.paragraphs[-1]._element
        el.getparent().remove(el)
        removed += 1


def add_section_with_header(page_fmt='decimal', start=1):
    remove_trailing_empty_body_paragraphs()
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


def set_cell_margins(cell, margins):
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
        set_cell_margins(cell, tcPr_data.get('tcMar'))
    valign = tcPr_data.get('vAlign')
    if valign == 'top':
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    elif valign == 'bottom':
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.BOTTOM
    else:
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


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


def is_elastic_cover_empty(el):
    if el.get('type') != 'empty':
        return False
    if el.get('section_break_after'):
        return False
    return not ''.join(r.get('t', '') for r in el.get('r', [])).strip()


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
    actual = str((DATA.get('cover_info') or {}).get('paper_title') or DATA.get('title_cn') or '').strip()
    sample = cover_table_sample_value(cover[info_idx], '题目')
    if not actual or not sample or len(actual) <= len(sample):
        return set()
    skip = set()
    j = info_idx - 1
    before = []
    while j >= 0 and is_elastic_cover_empty(cover[j]):
        before.append(j); j -= 1
    # For a longer replacement title, all directly adjacent blank spacer
    # paragraphs around the info table are elastic. They exist for the sample
    # cover only, not as mandatory content. Removing all of them prevents the
    # info table and committee line from spilling onto the next page.
    for idx in before:
        skip.add(idx)
    j = info_idx + 1
    after = []
    while j < len(cover) and is_elastic_cover_empty(cover[j]):
        after.append(j); j += 1
    for idx in after:
        skip.add(idx)
    # Also delete trailing blank spacer paragraphs immediately before an
    # empty section-break carrier. Otherwise the carrier occupies a whole
    # Roman-numbered blank page before the next section content.
    for marker_idx, marker_el in enumerate(cover):
        if cover_element_is_empty_section_marker(marker_el):
            skip.add(marker_idx)
            k = marker_idx - 1
            while k >= 0 and is_elastic_cover_empty(cover[k]):
                skip.add(k)
                k -= 1
    return skip


def cover_element_is_empty_section_marker(el):
    return bool(el.get('section_break_after')) and el.get('type') == 'empty' and is_elastic_cover_empty(el)

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


def collect_toc_entries():
    entries = []
    for i, sec in enumerate(DATA.get('sections') or []):
        if is_front_section_index(i):
            continue
        h = (sec.get('heading') or '').strip()
        role = sec.get('role') or ''
        if not h or h == '正文':
            continue
        if is_reference_heading(h) or is_backmatter_heading(h) or is_caption_heading(h) or role in ('references', 'acknowledgement', 'appendix'):
            continue
        level = max(1, min(int(sec.get('level') or 1), 3))
        entries.append({'level': level, 'text': normalize_heading_spacing(h)})
    if DATA.get('references'):
        entries.append({'level': 1, 'text': '参考文献'})
    ack_sections, app_sections = collect_structural_backmatter()
    pure_refs, ack_from_refs, app_from_refs = split_refs_backmatter(DATA.get('references') or [])
    if ack_sections or ack_from_refs:
        entries.append({'level': 1, 'text': '致  谢'})
    if app_sections or app_from_refs:
        entries.append({'level': 1, 'text': '附  录'})
    return entries


def add_toc_line(text, level, page_text=''):
    prof = profile('body')
    p = doc.add_paragraph()
    apply_paragraph_profile(p, prof, first_indent_override=0)
    p.paragraph_format.left_indent = Cm(0.0 if level == 1 else 0.74 if level == 2 else 1.48)
    tabs = p.paragraph_format.tab_stops
    tabs.add_tab_stop(Cm(15.2))
    r = p.add_run(text)
    line_prof = dict(prof)
    if level == 1:
        line_prof['bold'] = True
    apply_run_profile(r, line_prof, text)
    p.add_run('\t')
    r2 = p.add_run(str(page_text))
    apply_run_profile(r2, prof, str(page_text))
    return p


def add_toc():
    # Keep the document update flag for Word users, but generate visible TOC
    # entries immediately so headless conversion does not produce a blank TOC.
    enable_update_fields_on_open()
    add_text('目录', role='toc_title', first_indent=False)
    page_no = 1
    for entry in collect_toc_entries():
        add_toc_line(entry['text'], entry['level'], page_no if entry['level'] == 1 else '')
        if entry['level'] == 1:
            page_no += 1


def render_front_matter():
    front = DATA.get('front') or {}
    title_cn = DATA.get('title_cn') or ''
    if title_cn:
        add_text(title_cn, role='cn_title', first_indent=False)
    cn_abs = front.get('cn_abs')
    if cn_abs:
        add_text('摘 要', role='cn_abstract_heading', first_indent=False)
        cn_items = []
        for para in cn_abs.get('paragraphs', []) or []:
            text = para if isinstance(para, str) else para.get('text', '')
            if str(text).strip():
                cn_items.append(str(text).strip())
        if DATA.get('rules', {}).get('cn_abstract_single_paragraph') and cn_items:
            add_text(''.join(cn_items), role='cn_abstract_body', first_indent=True)
        else:
            for text in cn_items:
                add_text(text, role='cn_abstract_body', first_indent=True)
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
    if en_title and DATA.get('rules', {}).get('en_title_upper'):
        en_title = en_title.upper()
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
    add_section_with_header('upperRoman', None)
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


def is_caption_heading(h):
    return bool(re.match(r'^(图|表)\s*\d+(?:[.-]\d+)?\s*', str(h or '').strip()))


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
    return '\n'.join(lines).rstrip()

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
        rows = item.get('table_rows') or []
        if item.get('role') == 'code' or rows_look_like_code(rows):
            add_code_block(item.get('code') or code_text_from_rows(rows))
        else:
            render_table(rows)
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
        if is_caption_heading(h):
            add_caption(h, 'figure_caption' if str(h).strip().startswith('图') else 'table_caption')
        elif h and h != '正文':
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




def _norm_for_pdf_match(text):
    return re.sub(r'\s+', '', str(text or '')).lower()


def _extract_pdf_pages(pdf_path):
    exe = shutil.which('pdftotext')
    if not exe:
        return []
    with tempfile.TemporaryDirectory() as td:
        txt = os.path.join(td, 'out.txt')
        cmd = [exe, '-layout', pdf_path, txt]
        subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=60)
        if not os.path.exists(txt):
            return []
        data = open(txt, 'r', encoding='utf-8', errors='ignore').read()
    return data.split('\f')


def _make_pdf_for_pagination(docx_path):
    soffice = shutil.which('libreoffice') or shutil.which('soffice')
    if not soffice:
        return None
    td = tempfile.mkdtemp(prefix='toc_pages_')
    cmd = [soffice, '--headless', '--convert-to', 'pdf', '--outdir', td, docx_path]
    subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=90)
    base = os.path.splitext(os.path.basename(docx_path))[0] + '.pdf'
    pdf = os.path.join(td, base)
    return pdf if os.path.exists(pdf) else None


def _infer_heading_pages_from_pdf():
    """Best-effort pagination pass for the static TOC.

    python-docx has no layout engine.  When LibreOffice + pdftotext are
    available, we render once, locate headings in the rendered pages, then
    rewrite the visible TOC page numbers.  If the tools are unavailable, the
    document still contains visible TOC entries and Word can update fields on
    open if desired.
    """
    try:
        pdf = _make_pdf_for_pagination(OUT)
        if not pdf:
            return {}
        pages = _extract_pdf_pages(pdf)
        if not pages:
            return {}
        norm_pages = [_norm_for_pdf_match(p) for p in pages]
        entries = collect_toc_entries()
        if not entries:
            return {}
        first = _norm_for_pdf_match(entries[0]['text'])
        toc_last = 0
        for i, text in enumerate(norm_pages):
            if '目录' in pages[i] or '目录' in text:
                toc_last = i
        body_start = None
        for i in range(toc_last + 1, len(norm_pages)):
            if first and first in norm_pages[i]:
                body_start = i
                break
        if body_start is None:
            return {}
        page_map = {}
        for ent in entries:
            key = _norm_for_pdf_match(ent['text'])
            if not key:
                continue
            for i in range(body_start, len(norm_pages)):
                if key in norm_pages[i]:
                    page_map[_norm_for_pdf_match(ent['text'])] = i - body_start + 1
                    break
        return page_map
    except Exception:
        return {}


def _rewrite_static_toc_pages(page_map):
    if not page_map:
        return False
    try:
        d = Document(OUT)
        in_toc = False
        changed = False
        for p in d.paragraphs:
            txt = p.text.strip()
            if txt == '目录':
                in_toc = True
                continue
            if not in_toc:
                continue
            pPr = p._element.find(qn('w:pPr'))
            if pPr is not None and pPr.find(qn('w:sectPr')) is not None:
                break
            if '\t' not in p.text:
                continue
            label = p.text.split('\t', 1)[0].strip()
            key = _norm_for_pdf_match(label)
            if key not in page_map:
                continue
            # Preserve paragraph formatting; rebuild simple runs only.
            for r in list(p.runs):
                r._element.getparent().remove(r._element)
            prof = profile('body')
            r1 = p.add_run(label)
            apply_run_profile(r1, prof, label)
            p.add_run('\t')
            r2 = p.add_run(str(page_map[key]))
            apply_run_profile(r2, prof, str(page_map[key]))
            changed = True
        if changed:
            d.save(OUT)
        return changed
    except Exception:
        return False


def update_static_toc_pages():
    page_map = _infer_heading_pages_from_pdf()
    _rewrite_static_toc_pages(page_map)

def main():
    setup_section(doc.sections[0])
    clear_header_footer(doc.sections[0])
    remove_initial_empty_paragraph()
    render_cover_and_declarations()
    render_front_matter()
    render_body()
    force_cover_headerless()
    doc.save(OUT)
    update_static_toc_pages()
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
    rules = _infer_template_rules(fmt)

    src_assets = fmt.get('_meta', {}).get('assets_dir') or ''
    runtime_assets = src_assets
    if src_assets and os.path.isdir(src_assets):
        dst_assets = os.path.join(output_dir, 'assets')
        if os.path.abspath(src_assets) != os.path.abspath(dst_assets):
            shutil.rmtree(dst_assets, ignore_errors=True)
            shutil.copytree(src_assets, dst_assets)
        runtime_assets = 'assets'

    src_images = cnt.get('_meta', {}).get('images_dir') or ''
    runtime_images = src_images
    if src_images and os.path.isdir(src_images):
        dst_images = os.path.join(output_dir, 'figures')
        if os.path.abspath(src_images) != os.path.abspath(dst_images):
            shutil.rmtree(dst_images, ignore_errors=True)
            shutil.copytree(src_images, dst_images)
        runtime_images = 'figures'

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
        'images_dir': runtime_images,
        'assets_dir': runtime_assets,
        'rules': rules,
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
