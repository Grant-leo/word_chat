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


def is_template_instruction_text(text: str) -> bool:
    """Return True for template guidance that should not become paper content."""
    raw = str(text or "").strip()
    if not raw:
        return False
    compact = re.sub(r"\s+", "", raw)
    if not compact:
        return False
    if re.fullmatch(r"[（(]?空[一二两三四五六七八九十\d]+行[）)]?", compact):
        return True
    if "完成时间按照答辩时间填写" in compact:
        return True
    if "摘要是论文内容的总结概括" in compact and ("约200词" in compact or "第三人称" in compact):
        return True
    if "不标注引用编号" in compact and ("摘要" in compact or "关键词" in compact):
        return True
    font_signal = bool(re.search(r"TimesNewRoman|宋体|黑体|楷体|仿宋|华文|字号|[一二三四五六七八九十小]+号", compact, re.I))
    layout_signal = bool(re.search(r"居中|加粗|行距|倍行距|段前|段后|缩进|对齐|表格行高|固定值|页边距", compact))
    subject_signal = bool(re.search(r"英文题目|中文题目|目录内容|一级标题|二级标题|三级标题|图表题注|参考文献|页眉|页脚", compact))
    if font_signal and layout_signal:
        return True
    if subject_signal and (font_signal or layout_signal):
        return True
    return False


def strip_template_instruction_fragments(text: str) -> str:
    """Remove inline parenthesized format notes while keeping real content."""
    value = str(text or "")

    def replace(match: re.Match[str]) -> str:
        inner = match.group(0)[1:-1]
        return "" if is_template_instruction_text(inner) else match.group(0)

    previous = None
    while previous != value:
        previous = value
        value = re.sub(r"[（(][^（）()]{1,120}[）)]", replace, value)
    return re.sub(r"\s{2,}", " ", value).strip()


def is_probable_toc_entry_text(text: str) -> bool:
    """Detect source/template TOC sample rows such as '2. Heading 12'."""
    raw = str(text or "").strip()
    if not raw or len(raw) > 180:
        return False
    compact = re.sub(r"\s+", " ", raw)
    if re.match(r"^\d+(?:\.\d+)*\.?\s+.+\s+\d+$", compact):
        return True
    if re.match(r"^\d+(?:\.\d+)*\.?\s+.+\D\d+$", compact):
        return True
    if re.match(r"^(?:第?[一二三四五六七八九十\d]+章|\d+(?:\.\d+)*)\s+.+\s+(?:[ivxlcdm]+|\d+)$", compact, re.I):
        return True
    return False


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
