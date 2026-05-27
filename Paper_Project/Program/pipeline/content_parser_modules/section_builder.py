"""Section construction and finalization helpers for content_parser.py."""
from typing import Any, Dict, Iterable, List, Optional

try:
    from content_parser_modules.caption_flow import pair_figure_blocks
    from content_parser_modules.formula_extractor import repair_split_formula_layouts
except ImportError:  # pragma: no cover - package-style imports
    from .caption_flow import pair_figure_blocks
    from .formula_extractor import repair_split_formula_layouts


BODY_HEADING = '\u6b63\u6587'
FRONT_ROLES = frozenset({'cn_abstract', 'cn_keywords', 'en_abstract', 'en_keywords'})


def make_section(
    heading: str,
    level: int = 1,
    role: str = 'body',
    *,
    paragraphs: Optional[List[Any]] = None,
    images: Optional[List[str]] = None,
    page_break_before: bool = False,
) -> Dict[str, Any]:
    section: Dict[str, Any] = {
        'heading': heading,
        'level': level,
        'role': role,
        'paragraphs': list(paragraphs or []),
        'images': list(images or []),
    }
    if page_break_before:
        section['page_break_before'] = True
    return section


def make_body_section() -> Dict[str, Any]:
    return make_section(BODY_HEADING, level=1, role='body')


def filter_content_sections(sections: Iterable[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Drop the initial body placeholder while keeping real structural headings."""
    items = list(sections or [])
    filtered: List[Dict[str, Any]] = []
    for section in items:
        heading = section.get('heading')
        if heading == BODY_HEADING and len(items) > 1:
            continue
        if section.get('paragraphs') or section.get('images') or (heading and heading != BODY_HEADING):
            filtered.append(section)
    return filtered


def mark_first_body_page_break(sections: Iterable[Dict[str, Any]]) -> List[Dict[str, Any]]:
    items = list(sections or [])
    for section in items:
        if section.get('role') not in FRONT_ROLES and not section.get('page_break_before'):
            section.setdefault('page_break_before', True)
            break
    return items


def postprocess_section_paragraphs(sections: Iterable[Dict[str, Any]]) -> List[Dict[str, Any]]:
    items = list(sections or [])
    for section in items:
        paragraphs = repair_split_formula_layouts(section.get('paragraphs') or [])
        section['paragraphs'] = pair_figure_blocks(paragraphs)
    return items
