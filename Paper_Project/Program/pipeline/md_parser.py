"""
md_parser.py — Parse Markdown files into format.json and content.json.
Supports optional YAML frontmatter or natural-language # 格式说明 for format specs.
"""
import os, re, shutil, hashlib, json

# ═══════════════════════════════════════════════════════════════
#  FORMAT EXTRACTION
# ═══════════════════════════════════════════════════════════════

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

def _parse_page_geometry(text):
    """Extract page dimensions from Chinese description like 'A4，上2.5cm，下2.4cm...'."""
    geo = {}
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


def _parse_yaml_frontmatter(raw):
    """Parse simple YAML-like frontmatter between --- markers."""
    if not raw.startswith('---'):
        return {}, 0
    end = raw.find('---', 3)
    if end == -1:
        return {}, 0
    config = {}
    for line in raw[3:end].strip().split('\n'):
        line = line.strip()
        if ':' in line and not line.startswith('#'):
            key, _, val = line.partition(':')
            key = key.strip()
            val = val.strip().strip('"').strip("'")
            if val.lower() in ('true', 'yes', 'on'):
                val = True
            elif val.lower() in ('false', 'no', 'off'):
                val = False
            else:
                try: val = int(val) if '.' not in val else float(val)
                except ValueError: pass
            config[key] = val
    return config, end + 3


def _find_format_section(text, start_pos):
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


def _build_format_dict(md_path, fmt_text, page_override=None, header_override=None):
    """Build a format.json-compatible dict from format description text."""
    geo = _parse_page_geometry(fmt_text)
    page = {**DEFAULT_PAGE, **geo}
    if page_override:
        page.update(page_override)

    # Put each non-empty, non-heading format line as a paragraph
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


def extract_format(md_path):
    """Extract format information from MD file.
    Returns (format_dict, md_text) — same signature as format_extractor.extract().
    """
    with open(md_path, 'r', encoding='utf-8') as f:
        raw = f.read()

    yaml_config, body_pos = _parse_yaml_frontmatter(raw)
    fmt_text = None
    page_override = {}
    header_override = {}

    if yaml_config:
        # Build format description from YAML config (use Latin parens to avoid regex issues)
        parts = []
        body_sz = yaml_config.get('body_size', 12)
        body_font = yaml_config.get('body_font', 'Times New Roman')
        cjk_font = yaml_config.get('body_cjk_font', '宋体')
        body_align = yaml_config.get('body_align', 'JUSTIFY')
        align_cn = '两端对齐' if body_align == 'JUSTIFY' else ('居中' if body_align == 'CENTER' else '左对齐')
        parts.append(
            f"正文：{body_font}，{body_sz}pt，{align_cn}，1.5倍行距。")
        parts.append(f"中文字体使用{cjk_font}。")

        for i, key in enumerate(['heading1', 'heading2', 'heading3']):
            sz = yaml_config.get(f'{key}_size')
            if sz:
                font = yaml_config.get(f'{key}_font', '黑体')
                align = yaml_config.get(f'{key}_align', 'CENTER' if i == 0 else 'LEFT')
                align_cn = '居中' if align == 'CENTER' else '左对齐'
                parts.append(
                    f"{'一二三'[i]}级标题：{font}，{sz}pt，加粗，{align_cn}。")

        if yaml_config.get('abstract_size'):
            parts.append(f"Abstract：Times New Roman，{yaml_config['abstract_size']}pt，加粗，左对齐。")
        if yaml_config.get('keywords_size'):
            parts.append(f"Key words：Times New Roman，{yaml_config['keywords_size']}pt，加粗，左对齐。")

        fmt_text = '\n'.join(parts)

        # Extract page geometry from YAML
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


# ═══════════════════════════════════════════════════════════════
#  CONTENT EXTRACTION
# ═══════════════════════════════════════════════════════════════

# Patterns for detecting reference sections
_RE_REF_HEADING = re.compile(r'(?i)^references?\b|^参考文献|^引用文献')


def _is_format_section_heading(line):
    return bool(re.match(r'^#{1,3}\s+[格式排版要求说明]', line))


def _skip_format_section(lines):
    """Skip YAML frontmatter and # 格式 section. Returns remaining lines."""
    # Skip YAML frontmatter
    if lines and lines[0].strip() == '---':
        for i in range(1, len(lines)):
            if lines[i].strip() == '---':
                lines = lines[i + 1:]
                break
    # Skip format description section
    while lines and not lines[0].strip():
        lines = lines[1:]
    if lines and _is_format_section_heading(lines[0]):
        for i in range(1, len(lines)):
            stripped = lines[i].strip()
            if stripped == '---' or (re.match(r'^#{1,3}\s+', stripped)
                                     and not _is_format_section_heading(stripped)):
                lines = lines[i + (1 if stripped == '---' else 0):]
                break
        else:
            lines = []  # entire file was format section
    return lines


def _detect_title(lines):
    """Find first # heading as document title."""
    for i, line in enumerate(lines):
        m = re.match(r'^#\s+(.+)', line)
        if m:
            return m.group(1).strip(), i
    return '', 0


def _process_inline_math(text):
    """Split text by $...$ spans. Returns list of (text, is_math) tuples."""
    parts = []
    pos = 0
    while pos < len(text):
        dollar = text.find('$', pos)
        if dollar == -1:
            parts.append((text[pos:], False))
            break
        if dollar > 0 and text[dollar - 1] == '\\':
            # Escaped dollar
            parts.append((text[pos:dollar - 1] + '$', False))
            pos = dollar + 1
            continue
        # Check for $$ (display math) — handled by caller
        if dollar + 1 < len(text) and text[dollar + 1] == '$':
            parts.append((text[pos:dollar], False))
            end = text.find('$$', dollar + 2)
            if end == -1:
                parts.append((text[dollar + 2:], True))
                pos = len(text)
            else:
                parts.append((text[dollar + 2:end], True))
                pos = end + 2
            continue
        # Single $ — inline math
        parts.append((text[pos:dollar], False))
        end = text.find('$', dollar + 1)
        if end == -1:
            parts.append((text[dollar + 1:], True))
            pos = len(text)
        else:
            parts.append((text[dollar + 1:end], True))
            pos = end + 1
    return parts


def _extract_images_from_text(text, fig_dir, prefix, base_dir=''):
    """Extract image references like ![...](path) from text.
    Copies images to fig_dir. Returns (clean_text, image_filenames)."""
    imgs = []
    clean = text

    for m in re.finditer(r'!\[.*?\]\((.+?)\)', text):
        src = m.group(1).strip().strip('"').strip("'")
        if re.match(r'^[a-z]+://', src, re.I):
            continue
        fname = os.path.basename(src)
        if not fname:
            continue
        name, ext = os.path.splitext(fname)
        if not ext:
            ext = '.png'
        # Generate unique name
        existing = [f for f in os.listdir(fig_dir) if f.startswith(prefix)]
        seq = len(existing) + 1
        dest = os.path.join(fig_dir, f'{prefix}_{seq:03d}{ext}')
        candidates = [src]
        if not os.path.isabs(src):
            if base_dir:
                candidates.insert(0, os.path.join(base_dir, src))
            candidates.append(os.path.abspath(src))
        src_path = next((p for p in candidates if os.path.exists(p)), None)
        if src_path:
            shutil.copy2(src_path, dest)
            imgs.append(os.path.basename(dest))

    # Remove image syntax from text
    clean = re.sub(r'!\[.*?\]\(.+?\)', '', clean).strip()
    return clean, imgs


def _looks_like_formula_text(text):
    t = str(text or '').strip()
    if not t or len(t) > 180 or t.endswith(('。', '！', '？')):
        return False
    if not re.search(r'[=＝≈≤≥<>]', t) or not re.search(r'\d', t):
        return False
    return len(re.findall(r'[=＝+\-*/×÷%≈≤≥<>]', t)) >= 2


def _latex_escape_text(text):
    return str(text or '').replace('\\', r'\backslash ').replace('{', r'\{').replace('}', r'\}')


def _latex_from_formula_text(text):
    return r'\text{' + _latex_escape_text(str(text or '').strip()) + '}'


def _parse_paragraph(text, fig_dir, prefix, base_dir=''):
    """Parse a paragraph text into content.json paragraph entry.
    Detects inline $...$ math, images, and formatting marks."""
    # Extract images first
    text, images = _extract_images_from_text(text, fig_dir, prefix, base_dir=base_dir)

    # Check for standalone display math (entire paragraph is $$...$$)
    stripped = text.strip()
    if stripped.startswith('$$') and stripped.endswith('$$'):
        latex = stripped[2:-2].strip()
        return {'role': 'formula', 'text': '', 'latex': latex, 'math': [{'type': 'display', 'latex': latex}]}, images

    # Check for inline math
    if '$' not in text:
        # Strip markdown formatting
        clean = _strip_md_formatting(text)
        if _looks_like_formula_text(clean):
            return {'role': 'formula', 'text': clean, 'latex': _latex_from_formula_text(clean)}, images
        return (clean if clean else None), images

    # Process inline math spans
    parts = _process_inline_math(text)
    plain_parts = []
    math_items = []
    for content, is_math in parts:
        if is_math:
            math_items.append({'type': 'inline', 'latex': content.strip()})
        else:
            clean = _strip_md_formatting(content)
            if clean:
                plain_parts.append(clean)

    plain_text = ' '.join(plain_parts).strip()
    if not plain_text:
        plain_text = ''

    if math_items:
        return {'text': plain_text, 'math': math_items}, images
    else:
        return (plain_text if plain_text else None), images


def _strip_md_formatting(text):
    """Strip markdown formatting: **bold**, *italic*, `code`, [links](url), > quote, etc."""
    # Remove images
    text = re.sub(r'!\[.*?\]\(.+?\)', '', text)
    # Remove links (keep text)
    text = re.sub(r'\[([^\]]*)\]\(.+?\)', r'\1', text)
    # Remove bold/italic markers
    text = re.sub(r'\*\*(.+?)\*\*', r'\1', text)
    text = re.sub(r'__(.+?)__', r'\1', text)
    text = re.sub(r'\*(.+?)\*', r'\1', text)
    text = re.sub(r'_(.+?)_', r'\1', text)
    # Remove inline code
    text = re.sub(r'`([^`]+)`', r'\1', text)
    # Remove blockquote marker
    text = re.sub(r'^>\s?', '', text, flags=re.MULTILINE)
    # Remove list markers
    text = re.sub(r'^[\s]*[-*+]\s+', '', text, flags=re.MULTILINE)
    text = re.sub(r'^[\s]*\d+[.)]\s+', '', text, flags=re.MULTILINE)
    # Collapse whitespace
    text = re.sub(r'\s+', ' ', text).strip()
    return text


def extract_content(md_path, output_dir='Inputs'):
    """Extract content from MD file into content.json-compatible dict.
    Returns dict with same structure as content_parser.extract().
    """
    with open(md_path, 'r', encoding='utf-8') as f:
        raw = f.read()

    lines = raw.split('\n')
    lines = _skip_format_section(lines)

    base = os.path.splitext(os.path.basename(md_path))[0]
    base_dir = os.path.dirname(os.path.abspath(md_path))
    content_dir = os.path.join(output_dir, base)
    fig_dir = os.path.join(content_dir, 'figures')
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
    para_lines = []
    all_images = []
    total_paras = 0

    def _flush_section():
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
                    para, imgs = _parse_paragraph(block, fig_dir, f'{base}_img', base_dir=base_dir)
                    current_section['images'].extend(imgs)
                    all_images.extend(imgs)
                    if para:
                        current_section['paragraphs'].append(para)
                        total_paras += 1
            para_lines.clear()
        if current_section['paragraphs'] or current_section['images']:
            sections.append(current_section)

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

            current_section = {
                'heading': heading,
                'level': level,
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
            para_lines.append(line)
        i += 1

    _flush_section()

    content['sections'] = sections
    content['_meta']['paragraphs'] = total_paras
    content['_meta']['images_extracted'] = len(all_images)
    content['_meta']['images_dir'] = fig_dir

    if ref_section and ref_section['entries']:
        content['references'] = ref_section['entries']

    return content


if __name__ == '__main__':
    import sys
    path = sys.argv[1] if len(sys.argv) > 1 else 'Inputs/test.md'

    # Test format extraction
    fmt, md_text = extract_format(path)
    with open(path.replace('.md', '_format.json'), 'w', encoding='utf-8') as f:
        json.dump(fmt, f, ensure_ascii=False, indent=2)
    print(f'Format JSON → {path.replace(".md", "_format.json")}')

    # Test content extraction
    cnt = extract_content(path)
    with open(path.replace('.md', '_content.json'), 'w', encoding='utf-8') as f:
        json.dump(cnt, f, ensure_ascii=False, indent=2)
    print(f'Content JSON → {path.replace(".md", "_content.json")}')
