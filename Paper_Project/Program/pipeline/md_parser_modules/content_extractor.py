"""Markdown content extraction orchestration."""
from __future__ import annotations

import hashlib
import os
import re

try:
    from path_safety import ensure_safe_output_dir, safe_rmtree_generated_child
    from md_parser_modules.content_helpers import (
        _RE_BACKMATTER_HEADING,
        _RE_REF_HEADING,
        _classify_markdown_heading_role,
        _detect_title,
        _parse_markdown_table,
        _parse_paragraph_items,
        _skip_format_section,
        _strip_md_formatting,
    )
except ImportError:  # pragma: no cover - package-style imports
    from ..path_safety import ensure_safe_output_dir, safe_rmtree_generated_child
    from .content_helpers import (
        _RE_BACKMATTER_HEADING,
        _RE_REF_HEADING,
        _classify_markdown_heading_role,
        _detect_title,
        _parse_markdown_table,
        _parse_paragraph_items,
        _skip_format_section,
        _strip_md_formatting,
    )

def _default_output_dir():
    return os.path.abspath(os.path.join(os.getcwd(), 'Outputs', '_md_parser_extract'))


def extract_content(md_path, output_dir=None):
    """Extract content from MD file into content.json-compatible dict.
    Returns dict with same structure as content_parser.extract().
    """
    with open(md_path, 'r', encoding='utf-8') as f:
        raw = f.read()

    lines = raw.split('\n')
    lines = _skip_format_section(lines)

    base = os.path.splitext(os.path.basename(md_path))[0]
    base_dir = os.path.dirname(os.path.abspath(md_path))
    output_dir = ensure_safe_output_dir(output_dir or _default_output_dir())
    content_dir = os.path.join(output_dir, base)
    fig_dir = os.path.join(content_dir, 'figures')
    safe_rmtree_generated_child(fig_dir, output_dir, allowed_names={"figures"})
    os.makedirs(fig_dir, exist_ok=True)

    # Detect title
    title, title_idx = _detect_title(lines)

    content = {
        '_meta': {
            'source': os.path.basename(md_path),
            'sha256': hashlib.sha256(open(md_path, 'rb').read()).hexdigest()[:16],
            'paragraphs': 0,
            'tables_count': 0,
        },
        'title_info': {'title_cn': title} if title else {},
        'sections': [],
        'references': [],
    }

    sections = []
    current_section = None
    ref_section = None
    collected_references = []
    para_lines = []
    all_images = []
    missing_images = []
    total_paras = 0
    total_tables = 0

    def _flush_paragraphs():
        nonlocal total_paras
        if current_section is None:
            return
        if para_lines:
            text_block = '\n'.join(para_lines).strip()
            if text_block:
                # Split by blank lines into paragraphs
                blocks = re.split(r'\n\s*\n', text_block)
                for block in blocks:
                    block = block.strip()
                    if not block:
                        continue
                    items, imgs, missing = _parse_paragraph_items(block, fig_dir, f'{base}_img', base_dir=base_dir)
                    current_section['images'].extend(imgs)
                    all_images.extend(imgs)
                    missing_images.extend(missing)
                    for para in items:
                        if not para:
                            continue
                        current_section['paragraphs'].append(para)
                        total_paras += 1
            para_lines.clear()

    def _flush_section():
        _flush_paragraphs()
        if current_section is None:
            return
        if current_section['paragraphs'] or current_section['images']:
            sections.append(current_section)

    if title:
        current_section = {
            'heading': '正文',
            'level': 1,
            'role': 'body',
            'paragraphs': [],
            'images': [],
        }

    # Parse line by line
    i = title_idx + 1 if title else 0
    while i < len(lines):
        line = lines[i]
        m = re.match(r'^(#{1,3})\s+(.+)', line)
        if m:
            _flush_section()
            level = len(m.group(1))
            heading = m.group(2).strip()

            if _RE_REF_HEADING.match(heading):
                ref_section = {'heading': heading, 'entries': []}
                current_section = None
                i += 1
                continue
            if ref_section is not None and _RE_BACKMATTER_HEADING.match(heading):
                if ref_section.get('entries'):
                    collected_references.extend(ref_section['entries'])
                ref_section = None

            current_section = {
                'heading': heading,
                'level': level,
                'role': _classify_markdown_heading_role(heading),
                'paragraphs': [],
                'images': [],
            }
            i += 1
            continue

        # References section
        if ref_section is not None:
            stripped = line.strip()
            if stripped and not stripped.startswith('#'):
                ref_section['entries'].append(_strip_md_formatting(stripped))
            i += 1
            continue

        # Content
        if current_section is not None:
            fence = re.match(r'^\s*(```|~~~)\s*([\w.+-]*)\s*$', line)
            if fence:
                _flush_paragraphs()
                marker = fence.group(1)
                language = fence.group(2).strip()
                code_lines = []
                i += 1
                while i < len(lines) and not lines[i].strip().startswith(marker):
                    code_lines.append(lines[i])
                    i += 1
                if i < len(lines):
                    i += 1
                current_section['paragraphs'].append({
                    'role': 'code',
                    'language': language,
                    'code': '\n'.join(code_lines).rstrip('\n'),
                })
                total_paras += 1
                continue

            table_rows, next_i = _parse_markdown_table(lines, i)
            if table_rows:
                _flush_paragraphs()
                current_section['paragraphs'].append({'role': 'table', 'table_rows': table_rows})
                total_paras += 1
                total_tables += 1
                i = next_i
                continue

            para_lines.append(line)
        i += 1

    _flush_section()

    content['sections'] = sections
    content['_meta']['paragraphs'] = total_paras
    content['_meta']['tables_count'] = total_tables
    content['_meta']['images_extracted'] = len(all_images)
    content['_meta']['images_dir'] = fig_dir
    content['_meta']['missing_images'] = missing_images

    if ref_section and ref_section['entries']:
        collected_references.extend(ref_section['entries'])
    if collected_references:
        content['references'] = collected_references

    return content


