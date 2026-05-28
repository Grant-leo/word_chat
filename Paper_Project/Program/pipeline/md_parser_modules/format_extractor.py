"""Markdown format-spec extraction helpers."""
from __future__ import annotations

import hashlib
import os
import re
from typing import Any, Dict, Tuple


DEFAULT_PAGE = {
    'page_width_cm': 21.0, 'page_height_cm': 29.7,
    'margin_top_cm': 2.54, 'margin_bottom_cm': 2.54,
    'margin_left_cm': 2.54, 'margin_right_cm': 2.54,
}

DEFAULT_FORMAT_TEXT = (
    "一级标题：黑体，小三号(15pt)，加粗，居中，段前12pt。\n"
    "二级标题：黑体，四号(14pt)，加粗，左对齐，段前8pt。\n"
    "三级标题：黑体，小四号(12pt)，加粗，左对齐，段前6pt。\n"
    "正文：Times New Roman，小四号(12pt)，两端对齐，首行缩进2字符(21pt)，1.5倍行距。\n"
    "中文字体使用宋体。\n"
    "Abstract：Times New Roman，小三号(15pt)，加粗，左对齐。\n"
    "Key words：Times New Roman，小四号(12pt)，加粗，左对齐。"
)


def _parse_page_geometry(text: str) -> Dict[str, float]:
    """Extract page dimensions from Chinese description like 'A4，上2.5cm，下2.4cm...'."""
    geo: Dict[str, float] = {}
    m = re.search(r'(\d+\.?\d*)\s*[x×]\s*(\d+\.?\d*)\s*cm', text)
    if m:
        geo['page_width_cm'] = float(m.group(1))
        geo['page_height_cm'] = float(m.group(2))
    for key, pat in [('margin_top_cm', r'上\s*(\d+\.?\d*)\s*cm'),
                     ('margin_bottom_cm', r'下\s*(\d+\.?\d*)\s*cm'),
                     ('margin_left_cm', r'左\s*(\d+\.?\d*)\s*cm'),
                     ('margin_right_cm', r'右\s*(\d+\.?\d*)\s*cm')]:
        m = re.search(pat, text)
        if m:
            geo[key] = float(m.group(1))
    return geo


def _parse_yaml_frontmatter(raw: str) -> Tuple[Dict[str, Any], int]:
    """Parse simple YAML-like frontmatter between --- markers."""
    if not raw.startswith('---'):
        return {}, 0
    end = raw.find('---', 3)
    if end == -1:
        return {}, 0
    config: Dict[str, Any] = {}
    for line in raw[3:end].strip().split('\n'):
        line = line.strip()
        if ':' in line and not line.startswith('#'):
            key, _, val = line.partition(':')
            key = key.strip()
            val = val.strip().strip('"').strip("'")
            if val.lower() in ('true', 'yes', 'on'):
                parsed: Any = True
            elif val.lower() in ('false', 'no', 'off'):
                parsed = False
            else:
                parsed = val
                try:
                    parsed = int(val) if '.' not in val else float(val)
                except ValueError:
                    pass
            config[key] = parsed
    return config, end + 3


def _find_format_section(text: str, start_pos: int) -> Tuple[str | None, int]:
    """Find natural-language format description section in MD.
    Returns (format_lines_text, end_position) or (None, start_pos)."""
    m = re.search(r'^#{1,3}\s+[格式排版要求说明].*$', text[start_pos:], re.MULTILINE)
    if not m:
        return None, start_pos
    sec_start = start_pos + m.start()
    rest = text[sec_start:].split('\n')
    collected = [rest[0]]
    body_start = 1
    for i, line in enumerate(rest[1:], 1):
        stripped = line.strip()
        if stripped == '---':
            body_start = i + 1
            break
        if re.match(r'^#{1,3}\s+', stripped) and not re.match(r'^#{1,3}\s+[格式排版要求说明]', stripped):
            break
        collected.append(line)
        body_start = i + 1
    fmt_text = '\n'.join(collected)
    body_pos = sec_start + sum(len(l) + 1 for l in rest[:body_start])
    return fmt_text, body_pos


def _build_format_dict(md_path: str, fmt_text: str, page_override: Dict[str, Any] | None = None, header_override: Dict[str, Any] | None = None) -> Dict[str, Any]:
    """Build a format.json-compatible dict from format description text."""
    geo = _parse_page_geometry(fmt_text)
    page = {**DEFAULT_PAGE, **geo}
    if page_override:
        page.update(page_override)

    paragraphs = []
    for line in fmt_text.split('\n'):
        line = line.strip()
        if not line or line.startswith('#'):
            continue
        paragraphs.append({
            'index': len(paragraphs),
            'style': 'Normal',
            'text': line,
            'runs': [{'text': line, 'font': 'Times New Roman', 'size_pt': 12,
                      'bold': False, 'italic': False}],
            'has_page_break': False,
            'align': 'JUSTIFY', 'ls': 1.5, 'indent': 0,
        })

    header_override = header_override or {}
    header_text = str(header_override.get('text') or '').strip()
    header_font = header_override.get('font') or 'Times New Roman'
    header_size = float(header_override.get('size') or 10.5)
    header_bold = bool(header_override.get('bold', False))
    header_align = header_override.get('align') or 'CENTER'
    header = []
    if header_text:
        header = [{
            'text': header_text,
            'alignment': header_align,
            'runs': [{'text': header_text, 'font': header_font,
                      'size_pt': header_size, 'bold': header_bold, 'italic': False}],
        }]

    section = {
        'index': 0,
        **page,
        'diff_first_page': False,
        'header': header,
        'footer': [],
    }

    return {
        '_meta': {
            'source': os.path.basename(md_path),
            'sha256': hashlib.sha256(open(md_path, 'rb').read()).hexdigest()[:16],
            'paragraphs': len(paragraphs),
            'tables': 0,
            'sections': 1,
        },
        'sections': [section],
        'paragraphs': paragraphs,
        'tables': [],
    }


def extract_format(md_path: str) -> Tuple[Dict[str, Any], str]:
    """Extract format information from MD file.
    Returns (format_dict, md_text) — same signature as format_extractor.extract().
    """
    with open(md_path, 'r', encoding='utf-8') as f:
        raw = f.read()

    yaml_config, body_pos = _parse_yaml_frontmatter(raw)
    fmt_text = None
    page_override: Dict[str, Any] = {}
    header_override: Dict[str, Any] = {}

    if yaml_config:
        parts = []
        body_sz = yaml_config.get('body_size', 12)
        body_font = yaml_config.get('body_font', 'Times New Roman')
        cjk_font = yaml_config.get('body_cjk_font', '宋体')
        body_align = yaml_config.get('body_align', 'JUSTIFY')
        align_cn = '两端对齐' if body_align == 'JUSTIFY' else ('居中' if body_align == 'CENTER' else '左对齐')
        parts.append(f"正文：{body_font}，{body_sz}pt，{align_cn}，1.5倍行距。")
        parts.append(f"中文字体使用{cjk_font}。")

        for i, key in enumerate(['heading1', 'heading2', 'heading3']):
            sz = yaml_config.get(f'{key}_size')
            if sz:
                font = yaml_config.get(f'{key}_font', '黑体')
                align = yaml_config.get(f'{key}_align', 'CENTER' if i == 0 else 'LEFT')
                align_cn = '居中' if align == 'CENTER' else '左对齐'
                parts.append(f"{'一二三'[i]}级标题：{font}，{sz}pt，加粗，{align_cn}。")

        if yaml_config.get('abstract_size'):
            parts.append(f"Abstract：Times New Roman，{yaml_config['abstract_size']}pt，加粗，左对齐。")
        if yaml_config.get('keywords_size'):
            parts.append(f"Key words：Times New Roman，{yaml_config['keywords_size']}pt，加粗，左对齐。")

        fmt_text = '\n'.join(parts)

        for k in ['page_width_cm', 'page_height_cm', 'margin_top_cm',
                  'margin_bottom_cm', 'margin_left_cm', 'margin_right_cm']:
            if k in yaml_config:
                page_override[k] = float(yaml_config[k])
        if yaml_config.get('header_text'):
            header_override = {
                'text': yaml_config.get('header_text'),
                'font': yaml_config.get('header_font', 'Times New Roman'),
                'size': yaml_config.get('header_size', 10.5),
                'bold': yaml_config.get('header_bold', False),
                'align': yaml_config.get('header_align', 'CENTER'),
            }

    if not fmt_text:
        fmt_text, body_pos = _find_format_section(raw, body_pos)
    if not fmt_text:
        fmt_text = DEFAULT_FORMAT_TEXT

    fmt_dict = _build_format_dict(md_path, fmt_text, page_override, header_override)
    return fmt_dict, raw

