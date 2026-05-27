"""Front-matter extraction helpers for content_parser.py."""
import re

try:
    from content_parser_modules.heading_detector import (
        ascii_alpha_ratio,
        classify_section_role,
    )
    from content_parser_modules.placeholders import (
        extract_labeled_title,
        is_unfilled_placeholder_text,
    )
    from content_parser_modules.style import heading_level_from_style
except ImportError:  # pragma: no cover - package-style imports
    from .heading_detector import (
        ascii_alpha_ratio,
        classify_section_role,
    )
    from .placeholders import (
        extract_labeled_title,
        is_unfilled_placeholder_text,
    )
    from .style import heading_level_from_style


_ABSTRACT_PREFIXES = ('\u6458\u8981', '\u5173\u952e\u8bcd')
_PAREN_OPEN = '\uff08'
_PAREN_CLOSE = '\uff09'
_YEAR_PREFIX = '\u5e74'
_UNDERGRAD_PREFIX = '\u672c\u79d1'

_COVER_LABEL_MAP = {
    '\u5b66\u6821\u7f16\u7801': 'school_code',
    '\u5b66\u4f4d\u7f16\u7801': 'degree_code',
    '\u8bba\u6587\u9898\u76ee': 'paper_title',
    '\u5b66\u751f\u59d3\u540d': 'student_name',
    '\u5b66\u53f7': 'student_id',
    '\u5b66    \u53f7': 'student_id',
    '\u6240\u5c5e\u5b66\u9662': 'college',
    '\u4e13\u4e1a\u73ed\u7ea7': 'class_name',
    '\u6307\u5bfc\u8001\u5e08': 'advisor',
    '\u6307\u5bfc\u6559\u5e08': 'advisor',
}


def _default_clean_text(text):
    return re.sub(r'\s+', ' ', str(text or '')).strip()


def _max_run_font_size_pt(paragraph):
    max_size = 0
    for run in paragraph.runs:
        if run.font.size:
            max_size = max(max_size, run.font.size.pt)
    return max_size


def _extract_cover_info_from_tables(tables):
    cover_info = {}
    for table in tables[:5]:
        for row in table.rows:
            if len(row.cells) < 2:
                continue
            label = row.cells[0].text.strip()
            value = row.cells[1].text.strip()
            if not label or not value:
                continue
            if is_unfilled_placeholder_text(value):
                continue
            for keyword, key in _COVER_LABEL_MAP.items():
                if keyword in label:
                    cover_info[key] = value
                    break
    return cover_info


def _clean_title_candidate(text):
    return str(text or '').split(_PAREN_OPEN)[0].strip().replace('\n', ' ').replace('\r', '')


def _looks_like_structural_heading(text):
    clean = str(text or '').strip()
    cjk_num = r'\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341\u767e\u5343\u4e07'
    if re.match(r'^\d+(?:\.\d+)*\s+', clean):
        return True
    if re.match(r'^\d+(?:\.\d+)+\s*\S+', clean):
        return True
    if re.match(r'^\d+[\u3001\uff0e.]\s*\S+', clean):
        return True
    if re.match(r'^第[' + cjk_num + r'\d]+[章节篇部分]\s*\S*', clean):
        return True
    if re.match(r'^[' + cjk_num + r']+[\u3001\uff0e.]\s*\S+', clean):
        return True
    if re.match(r'^[\uff08(][' + cjk_num + r'\d]+[\uff09)]\s*\S+', clean):
        return True
    return False


def _looks_like_title_candidate(text, paragraph, max_size):
    clean = _clean_title_candidate(text)
    if not clean or len(clean) < 10 or len(clean) > 120:
        return False
    if clean.startswith(_YEAR_PREFIX) or clean.startswith(_UNDERGRAD_PREFIX):
        return False
    if _looks_like_structural_heading(clean):
        return False
    if classify_section_role(clean, 1) in {'cn_abstract', 'cn_keywords', 'en_abstract', 'en_keywords', 'references', 'acknowledgement', 'appendix'}:
        return False
    if max_size >= 14:
        return True
    return bool(heading_level_from_style(paragraph))


def _looks_like_plain_front_title(text):
    clean = _clean_title_candidate(text)
    if not clean or len(clean) < 10 or len(clean) > 120:
        return False
    if clean.startswith(_YEAR_PREFIX) or clean.startswith(_UNDERGRAD_PREFIX):
        return False
    if _looks_like_structural_heading(clean):
        return False
    if classify_section_role(clean, 1) in {'cn_abstract', 'cn_keywords', 'en_abstract', 'en_keywords', 'references', 'acknowledgement', 'appendix'}:
        return False
    if re.search(r'(\u5b66\u6821|\u5b66\u9662|\u5927\u5b66|\u59d4\u5458\u4f1a|\u58f0\u660e|\u6388\u6743|\u76ee\u5f55|\u53c2\u8003\u6587\u732e)', clean) and len(clean) < 24:
        return False
    return True


def extract_front_matter(doc, clean_text_func=None):
    """Extract title/cover fields and return the first body paragraph index."""
    clean_text = clean_text_func or _default_clean_text
    title_info = {}
    cover_info = {}
    text_start = 0

    best_title = ('', 0, 0)  # (text, size, paragraph index)
    plain_title = ('', 0)  # (text, paragraph index)
    saw_front_marker = False
    for index, paragraph in enumerate(doc.paragraphs[:30]):
        text = paragraph.text.strip()
        if not text:
            continue
        if classify_section_role(text, 1) in {'cn_abstract', 'cn_keywords', 'en_abstract', 'en_keywords'}:
            saw_front_marker = True
            break
        if len(text) < 10:
            continue
        labeled_title = extract_labeled_title(text)
        if labeled_title:
            title_info['title_cn'] = labeled_title
            cover_info['paper_title'] = labeled_title
            text_start = max(text_start, index + 1)
            continue
        if is_unfilled_placeholder_text(text):
            continue
        if text.startswith(_PAREN_OPEN) and text.endswith(_PAREN_CLOSE):
            continue

        max_size = _max_run_font_size_pt(paragraph)
        if _looks_like_title_candidate(text, paragraph, max_size):
            clean = _clean_title_candidate(text)
            score = max(max_size, 13 if heading_level_from_style(paragraph) else 0)
            if score > best_title[1]:
                best_title = (clean, score, index)
        elif not plain_title[0] and _looks_like_plain_front_title(text):
            plain_title = (_clean_title_candidate(text), index)

    if not best_title[0] and plain_title[0] and saw_front_marker:
        best_title = (plain_title[0], 1, plain_title[1])

    if best_title[0]:
        title_info['title_cn'] = best_title[0]
        text_start = best_title[2] + 1
        end = min(text_start + 6, len(doc.paragraphs))
        for index, paragraph in enumerate(doc.paragraphs[text_start:end], start=text_start):
            text = clean_text(paragraph.text)
            if not text:
                continue
            if re.match(r'(?i)^(abstract|key\s*words?)\b', text) or text.startswith(_ABSTRACT_PREFIXES):
                break
            if classify_section_role(text, 1) in {'references', 'acknowledgement', 'appendix'}:
                continue
            if ascii_alpha_ratio(text) > 0.55 and not re.match(r'^\d+(?:\.\d+)*\s+', text):
                title_info['title_en'] = text
                text_start = index + 1
                break

    table_cover_info = _extract_cover_info_from_tables(doc.tables)
    if table_cover_info:
        cover_info.update(table_cover_info)

    # Cover tables are the most reliable source for the paper title.
    if cover_info.get('paper_title'):
        title_info['title_cn'] = cover_info['paper_title']

    result = {
        'text_start': text_start,
        'title_info': title_info,
    }
    if cover_info:
        result['cover_info'] = cover_info
    return result
