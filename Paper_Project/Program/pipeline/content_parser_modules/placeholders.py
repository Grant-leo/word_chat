"""Placeholder and labeled-title helpers for content parsing."""
from __future__ import annotations

import re
from typing import Any, Dict, Iterable, List


PLACEHOLDER_RE = re.compile(
    r'(\[[^\]\n]*(?:\u62a5\u540d|\u5e8f\u53f7|\u59d3\u540d|\u5b66\u53f7|\u5b66\u9662|\u4e13\u4e1a|\u73ed\u7ea7|\u9898\u76ee|\u6307\u5bfc|\u6559\u5e08|\u65e5\u671f|\u7f16\u7801|\u5f85\u586b|\u8bf7\u8f93\u5165|XX|XXX)[^\]\n]*\])'
    r'|(\{\{[^}]+\}\}|TODO|FIXME|\u5f85\u586b\u5199|\u5f85\u8865\u5168|XXXX)',
    re.I,
)


def is_unfilled_placeholder_text(text: str) -> bool:
    return bool(PLACEHOLDER_RE.search(str(text or "")))


def placeholder_samples(paragraphs: Iterable[Any], limit: int = 8) -> List[Dict[str, Any]]:
    out: List[Dict[str, Any]] = []
    for idx, para in enumerate(paragraphs, 1):
        text = str(getattr(para, "text", "") or "").strip()
        if text and is_unfilled_placeholder_text(text):
            out.append({"paragraph": idx, "text": text[:120]})
            if len(out) >= limit:
                break
    return out


def extract_labeled_title(text: str) -> str:
    t = str(text or "").strip()
    m = re.match(r"^\s*(?:\u8bba\u6587\u9898\u76ee|\u9898\u76ee|\u6807\u9898)\s*[:\uff1a]\s*(.+?)\s*$", t)
    if not m:
        return ""
    value = m.group(1).strip()
    return "" if is_unfilled_placeholder_text(value) else value
