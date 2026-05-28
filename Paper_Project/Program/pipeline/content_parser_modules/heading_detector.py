"""Heading detection and section role helpers for content_parser.py."""
import re

try:
    from content_parser_modules.style import (
        heading_level_from_style as _heading_level_from_style,
        looks_like_heading_style as _looks_like_heading_style,
    )
except ImportError:  # pragma: no cover - package-style imports
    from .style import (
        heading_level_from_style as _heading_level_from_style,
        looks_like_heading_style as _looks_like_heading_style,
    )


_CJK_NUMERALS = r'\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341\u767e\u5343'
_COMMON_COUNT_WORDS = set('\u79cd\u4e2a\u5929\u5e74\u6708\u65e5\u5428\u9879\u7ec4\u7c7b\u573a')


def _numbered_line_looks_like_body_sentence(text: str) -> bool:
    t = str(text or '').strip()
    if len(t) > 140:
        return True
    if len(t) > 70 and re.search(r'[\u3002\uff1b\uff0c,;]', t):
        return True
    if len(t) > 35 and re.search(r'[\u3002\uff1b;]\s*$', t):
        return True
    return False


def detect_heading_level(para):
    """Detect heading level using OOXML-direct size + heuristics."""
    if not para.runs:
        return 0
    text = para.text.strip()
    if not text:
        return 0
    # Figure/table captions are captions, not outline headings or TOC entries.
    if re.match(r'^(\u56fe|\u8868)\s*\d+(?:[.-]\d+)?\s*', text):
        return 0
    if re.fullmatch(r'[\d\s.,\uff0c\uff0e\u3002]+', text) or re.fullmatch(r'\d+\s*[.\uff0e]\s*\d+', text):
        return 0

    label_patterns = [
        (r'(?i)^(Abstract\s*:?)', 2),
        (r'(?i)^(Key\s*words?\s*:?)', 2),
        (r'^(\u6458\u8981\s*[\uff1a:]?)', 2),
        (r'^(\u5173\u952e\u8bcd\s*[\uff1a:]?)', 2),
    ]
    for pat, lvl in label_patterns:
        if re.match(pat, text):
            return lvl

    if re.match(r'^\u7b2c[\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341\d]+\u7ae0', text):
        return 1
    m_count_heading = re.match(r'^\d+\s+([\u4e00-\u9fff])', text)
    if m_count_heading and m_count_heading.group(1) in _COMMON_COUNT_WORDS:
        return 2 if _looks_like_heading_style(para) else 0
    if re.match(r'^(?:Chapter\s*)?\d+\s+[\u4e00-\u9fffA-Za-z]', text) and not re.match(r'^\d+\.\d+', text):
        if _numbered_line_looks_like_body_sentence(text) and not _looks_like_heading_style(para):
            return 0
        return 1
    if len(text) <= 80 and re.match(r'^[' + _CJK_NUMERALS + r']+[\u3001\uff0e.]\s*\S+', text):
        return 1
    if len(text) <= 80 and re.match(r'^\d+[\u3001\uff0e]\s*\S+', text):
        return 1
    if len(text) <= 80 and re.match(r'^[\uff08(][' + _CJK_NUMERALS + r']+[\uff09)]\s*\S+', text):
        return 2
    if re.match(r'^\d+\.\d+\s*(?:MWh|MW|kWh|kg|h|\u5428|\u5143|%|[+\-\u2212=,\uff0c\uff09\u3002])', text, re.I):
        return 0
    if re.match(r'^\d+\.\d+\.\d+\s*', text):
        return 3
    if re.match(r'^\d+\.\d+\s*', text):
        return 2

    if len(text) > 200:
        heading_patterns = [
            r'^(\d+\.\s+\w)', r'^(\d+\.\d+\s+\w)',
        ]
        for pat in heading_patterns:
            if re.match(pat, text):
                return 1 if re.match(r'^(\d+\.\s+\w)', text) else 2
        return 0

    if text.startswith('\uff08') and (text.endswith('\uff09') or len(text) < 60):
        return 0

    style_level = _heading_level_from_style(para)
    if style_level and len(text) <= 80 and not text.endswith('\u3002'):
        if re.search(r'[=\u00d7\u00f7\u03a3\u03a0\u222b\u221a\u221e\u2248\u00b1]', text) and len(text) < 100:
            return 0
        return style_level

    max_size = 0
    is_bold = False
    for r in para.runs:
        rPr = r._element.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')
        if rPr is not None:
            for child in rPr:
                tag = child.tag.split('}')[-1]
                if tag == 'sz' or tag == 'szCs':
                    try:
                        max_size = max(max_size, int(child.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '0')) / 2.0)
                    except Exception:
                        pass
                if tag == 'b':
                    val = child.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                    if str(val).lower() not in ('0', 'false', 'off'):
                        is_bold = True

    if not is_bold or max_size < 12:
        return 0

    if re.search(r'\(\d{4}[-\u2013]\d{4}\)', text):
        return 0
    if re.search(r'\d{6}', text):
        return 0
    if 'University' in text and len(text) > 100:
        return 0
    if len(text.split()) <= 2 and not any(c.isdigit() for c in text):
        return 0
    if re.search(r'[=\u00d7\u00f7\u03a3\u03a0\u222b\u221a\u221e\u2248\u00b1]', text) and len(text) < 100:
        return 0
    if len(text) > 80:
        return 0
    if text.endswith('\u3002'):
        return 0

    numbered = bool(re.match(r'^[\d]+\.[\d]*\s', text))

    if max_size >= 15:
        return 1
    if max_size >= 14:
        return 2 if (numbered or len(text) < 60) else 1
    if max_size >= 12:
        return 2 if numbered else 3
    return 0


def normalize_role_heading(text):
    return re.sub(r'\s+', ' ', str(text or '').strip())


def ascii_alpha_ratio(text):
    text = str(text or '')
    if not text:
        return 0.0
    return sum(1 for c in text if c.isascii() and c.isalpha()) / max(len(text), 1)


def classify_section_role(heading, level=0):
    """Map a detected heading to a semantic role used by the renderer."""
    h = normalize_role_heading(heading)
    h_compact = re.sub(r'[\s\uff1a:]+', '', h).lower()
    if h_compact in ('\u6458\u8981', '\u4e2d\u6587\u6458\u8981'):
        return 'cn_abstract'
    if h_compact in ('\u5173\u952e\u8bcd', '\u5173\u952e\u5b57') or h.startswith('\u5173\u952e\u8bcd'):
        return 'cn_keywords'
    if h_compact in ('abstract', 'englishabstract'):
        return 'en_abstract'
    if h.upper().replace(' ', '').startswith('KEYWORDS') or re.match(r'(?i)^key\s*words?', h):
        return 'en_keywords'
    if re.match(r'(?i)^references?$', h) or h.startswith('\u53c2\u8003\u6587\u732e'):
        return 'references'
    if re.match(r'(?i)^acknowledg(?:e)?ments?\b|^acknowledgment\b', h):
        return 'acknowledgement'
    if re.search(r'\u81f4\s*\u8c22', h):
        return 'acknowledgement'
    if re.match(r'(?i)^append(?:ix|ices)\b', h):
        return 'appendix'
    if re.search(r'\u9644\s*\u5f55', h):
        return 'appendix'
    if level and level > 0:
        return 'heading'
    return 'body'


def is_backmatter_heading(text):
    h = normalize_role_heading(text)
    return bool(
        re.match(r'(?i)^acknowledg(?:e)?ments?\b|^acknowledgment\b', h)
        or re.match(r'(?i)^append(?:ix|ices)\b', h)
        or re.search(r'(\u81f4\s*\u8c22|\u9644\s*\u5f55)', h)
    )


def split_heading_number(text):
    """Return (number, title) for headings such as numbered Chinese chapters."""
    t = str(text or '').strip()
    m = re.match(r'^(\u7b2c[\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341\u767e\u5343\u4e07\d]+\u7ae0)\s*(.+)$', t)
    if m:
        return m.group(1), m.group(2).strip()
    m = re.match(r'^(\d+(?:\.\d+)*)\s*(.+)$', t)
    if m:
        return m.group(1), m.group(2).strip()
    return '', t


def normalize_heading_spacing(text):
    num, title = split_heading_number(text)
    return f'{num} {title}'.strip() if num and title else str(text or '').strip()
