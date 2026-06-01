"""Markdown content parsing helpers."""
from __future__ import annotations

import os
import re
import shutil
from typing import Any, Dict, List, Tuple
from urllib.parse import unquote


_RE_REF_HEADING = re.compile(r'(?i)^references?\b|^参考文献|^引用文献')
_RE_BACKMATTER_HEADING = re.compile(r'(?i)^append(?:ix|ices)\b|^acknowledg(?:e)?ments?\b|^acknowledgment\b|^附\s*录|^致\s*谢')


def _strip_bom_prefix(line: str) -> str:
    return str(line or "").lstrip("\ufeff")


def _classify_markdown_heading_role(heading: str) -> str:
    text = str(heading or "").strip()
    compact = re.sub(r"\s+", "", text).lower()
    if compact in {"摘要", "中文摘要"}:
        return "cn_abstract"
    if compact in {"关键词", "关键字"}:
        return "cn_keywords"
    if compact in {"abstract", "englishabstract"}:
        return "en_abstract"
    if compact in {"keywords", "keyword", "keywords:", "keywords："}:
        return "en_keywords"
    if re.match(r"(?i)^append(?:ix|ices)\b", text) or re.search(r"附\s*录", text):
        return "appendix"
    if re.match(r"(?i)^acknowledg(?:e)?ments?\b|^acknowledgment\b", text) or re.search(r"致\s*谢", text):
        return "acknowledgement"
    return ""


def _markdown_heading_text(line: str) -> str:
    match = re.match(r'^#{1,3}\s+(.+?)\s*#*\s*$', _strip_bom_prefix(line).strip())
    return match.group(1).strip() if match else ''


def _is_format_section_heading(line: str) -> bool:
    text = re.sub(r"\s+", "", _markdown_heading_text(line))
    if not text:
        return False
    explicit_names = {"格式", "排版", "格式说明", "格式要求", "排版说明", "排版要求", "格式排版", "要求说明"}
    if text in explicit_names:
        return True
    return any(token in text for token in ("格式说明", "格式要求", "排版说明", "排版要求"))


def _looks_like_format_rule_line(line: str) -> bool:
    text = str(line or '').strip()
    if not text or text == '---' or re.match(r'^#{1,3}\s+', text):
        return False
    lower = text.lower()
    score = 0
    if any(font.lower() in lower for font in (
        "times new roman", "arial", "calibri", "宋体", "黑体", "楷体", "仿宋", "微软雅黑"
    )):
        score += 1
    if re.search(
        r'正文|标题|摘要|关键词|字体|字号|行距|页边距|缩进|居中|对齐|加粗|'
        r'\bbody\b|\bheading\b|\bfont\b|\babstract\b|\bkeywords?\b|'
        r'\bmargin\b|\bspacing\b|\bindent\b|\bjustif(?:y|ied)\b|\bcenter\b|\bbold\b',
        text,
        re.I,
    ):
        score += 1
    if re.search(r'\d+(?:\.\d+)?\s*(?:pt|cm|mm|号|字符|倍)', text, re.I):
        score += 1
    elif re.search(r'\d+(?:\.\d+)?', text) and ("times new roman" in lower or "font" in lower):
        score += 1
    if re.search(r'[:：]', text) and score:
        score += 1
    return score >= 2


def _format_section_end(lines: List[str]) -> int | None:
    if not lines or not re.match(r'^#{1,3}\s+', _strip_bom_prefix(lines[0]).strip()):
        return None

    explicit_heading = _is_format_section_heading(lines[0])
    seen_rule = False
    seen_nonblank = False
    for i in range(1, len(lines)):
        stripped = _strip_bom_prefix(lines[i]).strip()
        if not stripped:
            continue
        if stripped == '---':
            return i + 1 if explicit_heading or seen_rule else None
        if re.match(r'^#{1,3}\s+', stripped):
            return i if explicit_heading else None

        seen_nonblank = True
        if _looks_like_format_rule_line(stripped):
            seen_rule = True
            continue
        if not explicit_heading:
            return None

    if explicit_heading:
        return len(lines)
    if seen_nonblank and seen_rule:
        return len(lines)
    return None


def _skip_format_section(lines: List[str]) -> List[str]:
    """Skip YAML frontmatter and # 格式 section. Returns remaining lines."""
    if lines and _strip_bom_prefix(lines[0]).strip() == '---':
        for i in range(1, len(lines)):
            if _strip_bom_prefix(lines[i]).strip() == '---':
                lines = lines[i + 1:]
                break
    while lines and not _strip_bom_prefix(lines[0]).strip():
        lines = lines[1:]
    end = _format_section_end(lines)
    if end is not None:
        lines = lines[end:]
    return lines


def _detect_title(lines: List[str]) -> Tuple[str, int]:
    """Find the document title and return the index where title syntax ends."""
    for i, line in enumerate(lines):
        stripped = _strip_bom_prefix(line).strip()
        m = re.match(r'^#\s+(.+?)\s*#*\s*$', stripped)
        if m:
            return m.group(1).strip(), i
        if stripped and i + 1 < len(lines):
            underline = str(lines[i + 1] or "").strip()
            if re.fullmatch(r'=+\s*', underline):
                return stripped, i + 1
    return '', 0


def _title_info_from_title(title: str) -> Dict[str, str]:
    title = str(title or "").strip()
    if not title:
        return {}
    if re.search(r"[\u3400-\u9fff]", title):
        return {"title_cn": title}
    return {"title_en": title}


def _normalize_reference_label(label: str) -> str:
    return re.sub(r'\s+', ' ', str(label or '').strip()).lower()


def _parse_markdown_reference_definition(line: str) -> Tuple[str, str] | None:
    match = re.match(r'^\s{0,3}\[([^\]]+)\]:\s*(.+?)\s*$', str(line or ''))
    if not match:
        return None
    label = _normalize_reference_label(match.group(1))
    rest = match.group(2).strip()
    if not label or not rest:
        return None
    if rest.startswith('<'):
        end = rest.find('>')
        target = rest[1:end].strip() if end != -1 else rest[1:].strip()
    else:
        target = rest.split()[0].strip()
    if not target:
        return None
    return label, target


def _extract_markdown_reference_definitions(lines: List[str]) -> Tuple[List[str], Dict[str, str]]:
    filtered: List[str] = []
    references: Dict[str, str] = {}
    fence_marker = ''
    for line in lines:
        fence = re.match(r'^\s*(```|~~~)', str(line or ''))
        if fence:
            marker = fence.group(1)
            if not fence_marker:
                fence_marker = marker
            elif marker == fence_marker:
                fence_marker = ''
            filtered.append(line)
            continue
        if fence_marker:
            filtered.append(line)
            continue
        parsed = _parse_markdown_reference_definition(line)
        if parsed:
            label, target = parsed
            references[label] = target
            continue
        filtered.append(line)
    return filtered, references


def _process_inline_math(text: str) -> List[Tuple[str, bool]]:
    """Split text by $...$ spans. Returns list of (text, is_math) tuples."""
    parts: List[Tuple[str, bool]] = []
    pos = 0
    while pos < len(text):
        dollar = text.find('$', pos)
        if dollar == -1:
            parts.append((text[pos:], False))
            break
        if dollar > 0 and text[dollar - 1] == '\\':
            parts.append((text[pos:dollar - 1] + '$', False))
            pos = dollar + 1
            continue
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
        parts.append((text[pos:dollar], False))
        end = text.find('$', dollar + 1)
        if end == -1:
            parts.append((text[dollar + 1:], True))
            pos = len(text)
        else:
            parts.append((text[dollar + 1:end], True))
            pos = end + 1
    return parts


def _is_remote_image_source(src: str) -> bool:
    return bool(re.match(r'^[a-z][a-z0-9+.-]*://', str(src or ''), re.I))


def _local_image_filesystem_source(src: str) -> str:
    raw = str(src or '').strip()
    end = len(raw)
    for marker in ('?', '#'):
        idx = raw.find(marker)
        if idx != -1:
            end = min(end, idx)
    local = unquote(raw[:end].strip())
    if os.sep != '\\':
        local = local.replace('\\', os.sep)
    return local


def _split_image_tokens_from_text(text: str, fig_dir: str, prefix: str, base_dir: str = '', image_refs: Dict[str, str] | None = None) -> Tuple[List[Dict[str, Any]], List[str], List[Dict[str, str]]]:
    """Split Markdown image references into ordered text/image tokens."""
    imgs: List[str] = []
    missing: List[Dict[str, str]] = []
    tokens: List[Dict[str, Any]] = []
    pos = 0
    refs = image_refs or {}

    for m in re.finditer(r'!\[([^\]]*)\](?:\((.+?)\)|\[([^\]]*)\])?', text):
        if m.start() > pos:
            tokens.append({'type': 'text', 'text': text[pos:m.start()]})
        alt = m.group(1).strip()
        inline_src = m.group(2)
        ref_label = m.group(3)
        if inline_src is None:
            label = _normalize_reference_label(ref_label if ref_label is not None and ref_label.strip() else alt)
            src = refs.get(label, '')
            if not src:
                source = f'[{label}]' if label else ''
                missing.append({'source': source, 'alt': alt, 'reason': 'reference_not_found'})
                tokens.append({'type': 'missing_image', 'source': source, 'alt': alt, 'reason': 'reference_not_found'})
                pos = m.end()
                continue
        else:
            src = inline_src
        src = src.strip().strip('"').strip("'")
        if src.startswith("<") and src.endswith(">"):
            src = src[1:-1].strip()
        source_for_report = unquote(src)
        if _is_remote_image_source(src) or _is_remote_image_source(source_for_report):
            missing.append({'source': source_for_report, 'alt': alt, 'reason': 'remote'})
            tokens.append({'type': 'missing_image', 'source': source_for_report, 'alt': alt, 'reason': 'remote'})
            pos = m.end()
            continue
        local_src = _local_image_filesystem_source(src)
        fname = os.path.basename(local_src)
        if not fname:
            missing.append({'source': source_for_report, 'alt': alt, 'reason': 'empty_filename'})
            tokens.append({'type': 'missing_image', 'source': source_for_report, 'alt': alt, 'reason': 'empty_filename'})
            pos = m.end()
            continue
        name, ext = os.path.splitext(fname)
        if not ext:
            ext = '.png'
        existing = [f for f in os.listdir(fig_dir) if f.startswith(prefix)]
        seq = len(existing) + 1
        dest = os.path.join(fig_dir, f'{prefix}_{seq:03d}{ext}')
        candidates = [local_src]
        if not os.path.isabs(local_src):
            if base_dir:
                candidates.insert(0, os.path.join(base_dir, local_src))
            candidates.append(os.path.abspath(local_src))
        src_path = next((p for p in candidates if os.path.exists(p)), None)
        if src_path:
            shutil.copy2(src_path, dest)
            copied = os.path.basename(dest)
            imgs.append(copied)
            tokens.append({'type': 'image', 'image': copied})
        else:
            missing.append({'source': source_for_report, 'alt': alt, 'reason': 'not_found'})
            tokens.append({'type': 'missing_image', 'source': source_for_report, 'alt': alt, 'reason': 'not_found'})
        pos = m.end()

    if pos < len(text):
        tokens.append({'type': 'text', 'text': text[pos:]})
    if not tokens:
        tokens.append({'type': 'text', 'text': text})
    return tokens, imgs, missing


def _looks_like_formula_text(text: str) -> bool:
    t = str(text or '').strip()
    if not t or len(t) > 180 or t.endswith(('。', '！', '？')):
        return False
    if not re.search(r'[=＝≈≤≥<>]', t) or not re.search(r'\d', t):
        return False
    return len(re.findall(r'[=＝+\-*/×÷%≈≤≥<>]', t)) >= 2


def _latex_escape_text(text: str) -> str:
    return str(text or '').replace('\\', r'\backslash ').replace('{', r'\{').replace('}', r'\}')


def _latex_from_formula_text(text: str) -> str:
    return r'\text{' + _latex_escape_text(str(text or '').strip()) + '}'


def _parse_text_paragraph(text: str) -> Any:
    """Parse a paragraph text fragment into a content.json paragraph entry."""
    stripped = text.strip()
    if stripped.startswith('$$') and stripped.endswith('$$'):
        latex = stripped[2:-2].strip()
        return {'role': 'formula', 'text': '', 'latex': latex, 'math': [{'type': 'display', 'latex': latex}]}

    if '$' not in text:
        clean = _strip_md_formatting(text)
        if _looks_like_formula_text(clean):
            return {'role': 'formula', 'text': clean, 'latex': _latex_from_formula_text(clean)}
        return clean if clean else None

    parts = _process_inline_math(text)
    plain_parts = []
    math_items = []
    runs = []
    for content, is_math in parts:
        if is_math:
            latex = content.strip()
            item = {'type': 'inline', 'latex': latex, 'text': latex}
            math_items.append(item)
            runs.append({'type': 'math', 'text': latex, 'math': [item]})
        else:
            clean = _strip_md_formatting(content, preserve_edges=True)
            if clean.strip():
                plain_parts.append(clean)
                runs.append({'type': 'text', 'text': clean})

    plain_text = ' '.join(plain_parts).strip()
    if not plain_text:
        plain_text = ''

    if math_items:
        display_text = ''.join(str(r.get('text') or '') for r in runs).strip()
        return {'role': 'rich_text', 'text': display_text or plain_text, 'runs': runs, 'math': math_items}
    return plain_text if plain_text else None


def _parse_paragraph_items(text: str, fig_dir: str, prefix: str, base_dir: str = '', image_refs: Dict[str, str] | None = None) -> Tuple[List[Any], List[str], List[Dict[str, str]]]:
    """Parse a Markdown paragraph into ordered content items and image names."""
    tokens, images, missing = _split_image_tokens_from_text(text, fig_dir, prefix, base_dir=base_dir, image_refs=image_refs)
    items = []
    for tok in tokens:
        if tok.get('type') == 'image':
            items.append({'role': 'image', 'image': tok.get('image')})
            continue
        if tok.get('type') == 'missing_image':
            label = tok.get('alt') or tok.get('source') or 'missing image'
            items.append({'role': 'missing_image', 'text': label, 'source': tok.get('source'), 'reason': tok.get('reason')})
            continue
        fragment = tok.get('text') or ''
        if not fragment.strip():
            continue
        para = _parse_text_paragraph(fragment)
        if para:
            items.append(para)
    return items, images, missing


def _strip_md_formatting(text: str, preserve_edges: bool = False) -> str:
    """Strip markdown formatting: **bold**, *italic*, `code`, [links](url), > quote, etc."""
    raw = str(text or '')
    has_leading = bool(re.match(r'^\s+', raw))
    has_trailing = bool(re.search(r'\s+$', raw))
    text = raw
    text = re.sub(r'!\[.*?\]\(.+?\)', '', text)
    text = re.sub(r'\[([^\]]*)\]\(.+?\)', r'\1', text)
    text = re.sub(r'\*\*(.+?)\*\*', r'\1', text)
    text = re.sub(r'__(.+?)__', r'\1', text)
    text = re.sub(r'\*(.+?)\*', r'\1', text)
    text = re.sub(r'_(.+?)_', r'\1', text)
    text = re.sub(r'`([^`]+)`', r'\1', text)
    text = re.sub(r'^>\s?', '', text, flags=re.MULTILINE)
    text = re.sub(r'^[\s]*[-*+]\s+', '', text, flags=re.MULTILINE)
    text = re.sub(r'^[\s]*\d+[.)]\s+', '', text, flags=re.MULTILINE)
    text = re.sub(r'\s+', ' ', text)
    if preserve_edges:
        core = text.strip()
        if not core:
            return ''
        return (' ' if has_leading else '') + core + (' ' if has_trailing else '')
    return text.strip()


def _is_markdown_table_separator(line: str) -> bool:
    """Return True for a Markdown table separator row such as | --- | :---: |."""
    text = str(line or '').strip()
    if '|' not in text:
        return False
    text = text.strip('|').strip()
    if not text:
        return False
    cells = [c.strip() for c in text.split('|')]
    return bool(cells) and all(re.fullmatch(r':?-{3,}:?', c or '') for c in cells)


def _split_markdown_table_row(line: str) -> List[str]:
    text = str(line or '').strip()
    if '|' not in text:
        return []
    text = text.strip()
    if text.startswith('|'):
        text = text[1:]
    if text.endswith('|'):
        text = text[:-1]
    return [_strip_md_formatting(c.strip()) for c in text.split('|')]


def _parse_markdown_table(lines: List[str], start: int) -> Tuple[List[List[str]], int]:
    """Parse a GitHub-style Markdown table. Returns (rows, next_index)."""
    if start + 1 >= len(lines):
        return [], start
    header = _split_markdown_table_row(lines[start])
    if not header or not _is_markdown_table_separator(lines[start + 1]):
        return [], start

    rows = [header]
    i = start + 2
    while i < len(lines):
        stripped = lines[i].strip()
        if not stripped or stripped.startswith('#') or '|' not in stripped:
            break
        row = _split_markdown_table_row(lines[i])
        if row:
            if len(row) < len(header):
                row += [''] * (len(header) - len(row))
            elif len(row) > len(header):
                row = row[:len(header) - 1] + [' | '.join(row[len(header) - 1:])]
            rows.append(row)
        i += 1
    return rows if len(rows) > 1 else [], i
