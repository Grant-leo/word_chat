"""DOCX paragraph style helpers for content parsing."""
from __future__ import annotations

import re
from typing import Any


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def compact_text(text: str) -> str:
    return re.sub(r"[\s\u3000]+", "", str(text or "")).upper()


def paragraph_has_page_or_section_break(p_elem: Any) -> bool:
    if p_elem.find(f".//{{{W_NS}}}sectPr") is not None:
        return True
    for br in p_elem.iter(f"{{{W_NS}}}br"):
        if br.get(f"{{{W_NS}}}type") == "page":
            return True
    return False


def paragraph_style_id(para: Any) -> str:
    try:
        p_pr = para._element.find(f"{{{W_NS}}}pPr")
        if p_pr is not None:
            style = p_pr.find(f"{{{W_NS}}}pStyle")
            if style is not None:
                return style.get(f"{{{W_NS}}}val") or ""
    except Exception:
        pass
    return ""


def heading_level_from_style(para: Any) -> int:
    style_id = paragraph_style_id(para)
    try:
        style_name = para.style.name if para.style else ""
    except Exception:
        style_name = ""
    compact = compact_text(style_id + style_name)
    m = re.search(r"HEADING([1-6])", compact)
    if m:
        return int(m.group(1))
    m = re.search(r"\u6807\u9898([1-6])", compact)
    if m:
        return int(m.group(1))
    if re.search(r"HEADING|TITLE|CHAPTER", compact) or "\u6807\u9898" in compact or "\u7ae0" in compact:
        return 1
    return 0


def looks_like_heading_style(para: Any) -> bool:
    return bool(heading_level_from_style(para))
