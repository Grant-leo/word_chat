"""Source-document TOC detection and skip-window logic.

This module only decides whether a TOC already present in the content source
should be skipped.  It must not decide how to render the final generated TOC.
"""
from __future__ import annotations

import re
from typing import Any, Iterable, Optional, Tuple

from .placeholders import is_unfilled_placeholder_text
from .style import compact_text, looks_like_heading_style, paragraph_has_page_or_section_break


def is_source_toc_title(text: str) -> bool:
    compact = compact_text(text)
    return compact in {"\u76ee\u5f55", "\u76ee\u6b21", "CONTENTS", "TABLEOFCONTENTS"}


def is_source_toc_entry(text: str) -> bool:
    t = str(text or "").strip()
    if not t:
        return True
    if len(t) > 160:
        return False
    if is_source_toc_title(t):
        return True
    if re.search(r"(?:\.{2,}|\u2026+|\u00b7{2,}|_{2,})\s*(?:[ivxlcdm]+|\d+)\s*$", t, re.I):
        return True
    toc_prefix = (
        r"(?:\u7b2c[\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341\u767e\u5343\u4e07\d]+\u7ae0"
        r"|[\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341]+[\u3001.\uff0e]"
        r"|\d+(?:\.\d+)*|\u6458\u8981|ABSTRACT|\u5173\u952e\u8bcd|KEY\s*WORDS?"
        r"|\u53c2\u8003\u6587\u732e|\u81f4\u8c22|\u9644\u5f55|APPENDIX|ACKNOWLEDGEMENTS?)"
    )
    return bool(re.match(r"^" + toc_prefix + r"\s+.+\s+(?:[ivxlcdm]+|\d+)\s*$", t, re.I))


def is_unpaged_source_toc_entry(text: str, para: Any = None) -> bool:
    t = str(text or "").strip()
    if not t or len(t) > 100:
        return False
    if is_source_toc_entry(t):
        return True
    if re.search(r"[\u3002\uff01\uff1f!?；;]\s*$", t):
        return False
    if is_unfilled_placeholder_text(t):
        return False
    prefix = (
        r"(?:\u7b2c[\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341\u767e\u5343\u4e07\d]+\u7ae0\s*\S+"
        r"|[\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341]+[\u3001.\uff0e]\s*\S+"
        r"|\d+(?:\.\d+)*\s+\S+"
        r"|\u6458\u8981|ABSTRACT|\u5173\u952e\u8bcd|KEY\s*WORDS?"
        r"|\u53c2\u8003\u6587\u732e|REFERENCES?|\u81f4\u8c22|ACKNOWLEDGEMENTS?"
        r"|\u9644\u5f55|APPENDIX(?:\s+\S+)?)"
    )
    if re.match(r"^" + prefix + r"\s*$", t, re.I):
        return True
    if para is not None and looks_like_heading_style(para) and len(t) <= 80:
        return True
    return False


def simple_cn_number(value: str) -> Optional[int]:
    s = str(value or "").strip()
    if not s:
        return None
    if s.isdigit():
        return int(s)
    digits = {
        "\u96f6": 0,
        "\u4e00": 1,
        "\u4e8c": 2,
        "\u4e09": 3,
        "\u56db": 4,
        "\u4e94": 5,
        "\u516d": 6,
        "\u4e03": 7,
        "\u516b": 8,
        "\u4e5d": 9,
    }
    if s == "\u5341":
        return 10
    if "\u5341" in s:
        left, right = s.split("\u5341", 1)
        tens = digits.get(left, 1 if left == "" else 0)
        ones = digits.get(right, 0) if right else 0
        return tens * 10 + ones
    if len(s) == 1 and s in digits:
        return digits[s]
    return None


def toc_entry_key(text: str) -> str:
    t = str(text or "").strip()
    t = re.sub(r"(?:\.{2,}|\u2026+|\u00b7{2,}|_{2,})\s*(?:[ivxlcdm]+|\d+)\s*$", "", t, flags=re.I)
    t = re.sub(r"\s+(?:[ivxlcdm]+|\d+)\s*$", "", t, flags=re.I)
    return compact_text(t)


def toc_entry_order(text: str) -> Optional[Tuple[str, int]]:
    t = str(text or "").strip()
    m = re.match(r"^\u7b2c([\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341\u767e\u5343\u4e07\d]+)\u7ae0", t)
    if m:
        n = simple_cn_number(m.group(1))
        return ("chapter", n) if n is not None else None
    m = re.match(r"^([\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341]+)[\u3001.\uff0e]", t)
    if m:
        n = simple_cn_number(m.group(1))
        return ("cn", n) if n is not None else None
    m = re.match(r"^(\d+)(?:\.\d+)*\b", t)
    if m:
        return ("num", int(m.group(1)))
    return None


def source_toc_skip_count_after_title(paragraphs: Iterable[Any], title_idx: int, max_scan: int = 160) -> int:
    """Return how many paragraphs after a source TOC title should be skipped."""
    plist = list(paragraphs)
    n = len(plist)
    if title_idx < 0 or title_idx >= n:
        return 0
    first_visible = None
    scan_end = min(n, title_idx + 1 + max_scan)
    for idx in range(title_idx + 1, scan_end):
        text = str(getattr(plist[idx], "text", "") or "").strip()
        if text:
            first_visible = idx
            break
    if first_visible is None:
        return 0
    first_text = str(getattr(plist[first_visible], "text", "") or "").strip()
    if not is_unpaged_source_toc_entry(first_text, plist[first_visible]):
        return 0

    first_boundary_idx = None
    for idx in range(title_idx + 1, scan_end):
        try:
            if paragraph_has_page_or_section_break(plist[idx]._element):
                first_boundary_idx = idx
                break
        except Exception:
            continue

    visible_count = 0
    saw_paged_entry = False
    saw_boundary = False
    boundary_idx = None
    title_has_heading_style = looks_like_heading_style(plist[title_idx])
    non_toc_before_boundary = False
    skip_until = title_idx
    seen_keys = set()
    last_order = None
    for idx in range(title_idx + 1, scan_end):
        para = plist[idx]
        text = str(getattr(para, "text", "") or "").strip()
        try:
            has_boundary = paragraph_has_page_or_section_break(para._element)
        except Exception:
            has_boundary = False
        if not text:
            skip_until = idx
            if has_boundary:
                saw_boundary = True
                boundary_idx = idx
                break
            continue
        paged_entry = is_source_toc_entry(text) and not is_source_toc_title(text)
        unpaged_entry = is_unpaged_source_toc_entry(text, para)
        if not (paged_entry or unpaged_entry):
            if first_boundary_idx is not None and not title_has_heading_style:
                non_toc_before_boundary = True
                skip_until = idx
                if has_boundary:
                    saw_boundary = True
                    boundary_idx = idx
                    break
                continue
            break

        key = toc_entry_key(text)
        order = toc_entry_order(text)
        if visible_count > 0:
            if key and key in seen_keys:
                if first_boundary_idx is not None and not title_has_heading_style:
                    non_toc_before_boundary = True
                    skip_until = idx
                    if has_boundary:
                        saw_boundary = True
                        boundary_idx = idx
                        break
                    continue
                break
            if saw_paged_entry and not paged_entry:
                if first_boundary_idx is not None and not title_has_heading_style:
                    non_toc_before_boundary = True
                    skip_until = idx
                    if has_boundary:
                        saw_boundary = True
                        boundary_idx = idx
                        break
                    continue
                break
            if order and last_order and order[0] == last_order[0] and order[1] <= last_order[1]:
                if first_boundary_idx is not None and not title_has_heading_style:
                    non_toc_before_boundary = True
                    skip_until = idx
                    if has_boundary:
                        saw_boundary = True
                        boundary_idx = idx
                        break
                    continue
                break

        visible_count += 1
        saw_paged_entry = saw_paged_entry or paged_entry
        if key:
            seen_keys.add(key)
        if order:
            last_order = order
        skip_until = idx
        if has_boundary:
            saw_boundary = True
            boundary_idx = idx
            break

    if saw_boundary:
        if saw_paged_entry or visible_count >= 2 or (visible_count >= 1 and non_toc_before_boundary and not title_has_heading_style):
            return max(skip_until, boundary_idx or skip_until) - title_idx
        return 0
    if visible_count < 2 and not saw_paged_entry:
        return 0
    return max(0, skip_until - title_idx)
