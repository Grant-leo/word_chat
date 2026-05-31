"""Content sample detectors for structural QA."""
from __future__ import annotations

import os
import re
from typing import Any, Dict, Iterable, List

try:
    from formula_semantics import (
        CATEGORY_CONTAMINATED,
        classify_formula_text,
        formula_text_looks_contaminated as semantic_formula_text_looks_contaminated,
    )
except ImportError:  # pragma: no cover - package-style imports
    from ..formula_semantics import (
        CATEGORY_CONTAMINATED,
        classify_formula_text,
        formula_text_looks_contaminated as semantic_formula_text_looks_contaminated,
    )

try:
    from qa_checker_modules.content_metrics import _iter_content_image_refs, _iter_paragraph_items
except ImportError:  # pragma: no cover - package-style imports
    from .content_metrics import _iter_content_image_refs, _iter_paragraph_items

_PLACEHOLDER_RE = re.compile(
    r"(\[[^\]\n]*(?:报名|序号|姓名|学号|学院|专业|班级|题目|指导|教师|日期|编码|待填|请输入|XX|XXX)[^\]\n]*\])"
    r"|(\{\{[^}]+\}\}|TODO|FIXME|待填写|待补全|XXXX)",
    re.I,
)


def _placeholder_samples_from_texts(texts: Iterable[str], limit: int = 8) -> List[str]:
    out: List[str] = []
    for text in texts:
        t = str(text or "").strip()
        if t and _PLACEHOLDER_RE.search(t):
            out.append(t[:120])
            if len(out) >= limit:
                break
    return out


def _heading_looks_like_body_or_formula(text: str) -> bool:
    t = str(text or "").strip()
    if not t:
        return False
    if re.match(r"^\d{1,2}(?:\.\d{1,2}){1,5}\s+[\w\u4e00-\u9fff]", t, re.I):
        return False
    if len(t) > 90:
        return True
    if re.search(r"(?:MWh|MW|Etotal|Esell|Ebuy|ERE|Pbuy|Psell|PRE|rself|rgreen|rup|max\s*\(|∑)", t):
        return True
    if re.search(r"\d+\s*\.\s*\d+$", t):
        return True
    if re.search(r"[=+\-*/×·]\s*\d", t) and len(t) > 20:
        return True
    if re.search(r"[。；;]\s*$", t) and len(t) > 35:
        return True
    return False


def _content_toc_pollution_samples(content: Dict[str, Any], limit: int = 8) -> List[str]:
    sections = content.get("sections") or []
    samples: List[str] = []
    seen_empty: Dict[str, int] = {}
    for idx, sec in enumerate(sections):
        heading = str(sec.get("heading") or "").strip()
        role = str(sec.get("role") or "")
        paras = sec.get("paragraphs") or []
        norm = re.sub(r"\s+", "", heading)
        if _heading_looks_like_body_or_formula(heading):
            samples.append(f"heading#{idx + 1}: {heading[:120]}")
        if role == "heading" and heading and not paras:
            seen_empty.setdefault(norm, idx + 1)
        elif norm in seen_empty:
            samples.append(f"empty duplicate before body: {heading[:80]}")
        if len(samples) >= limit:
            break
    return samples


def _formula_number_conflict_samples(content: Dict[str, Any], limit: int = 8) -> List[str]:
    out: List[str] = []
    for item in _iter_paragraph_items(content):
        if not isinstance(item, dict):
            continue
        if item.get("role") == "rich_text":
            continue
        if not (item.get("role") == "formula" or item.get("latex") or item.get("xml") or item.get("math")):
            continue
        text = str(item.get("text") or "")
        labels = re.findall(r"[\(\uff08]\s*\d+(?:\s*[-.]\s*\d+)?\s*[\)\uff09]", text)
        if len(labels) >= 2:
            out.append(text[:160])
            if len(out) >= limit:
                break
    return out


def _formula_text_looks_contaminated(text: str) -> bool:
    return semantic_formula_text_looks_contaminated(text)


def _fragmented_formula_samples(content: Dict[str, Any], limit: int = 8) -> List[str]:
    out: List[str] = []
    token_re = re.compile(r"^(?:E|RE|total|rself|rgreen|rup|max|t=\d+|\d+\s+\d+|\d+(?:\.\d+)?|[A-Za-z]{1,8})$")
    for sec in content.get("sections") or []:
        streak: List[str] = []
        for item in sec.get("paragraphs") or []:
            if isinstance(item, dict) and item.get("role") == "rich_text":
                streak = []
                continue
            if isinstance(item, dict) and item.get("role") == "formula_problem":
                formula_text = str(item.get("text") or "")
                semantic = item.get("formula_semantics") or classify_formula_text(formula_text).to_dict()
                out.append(
                    f"{sec.get('heading')}: formula semantic problem "
                    f"`{semantic.get('category', CATEGORY_CONTAMINATED)}` `{formula_text[:100]}`"
                )
                if len(out) >= limit:
                    return out
                streak = []
                continue
            if isinstance(item, dict) and (item.get("role") == "formula" or item.get("math") or item.get("latex") or item.get("xml")):
                formula_text = str(item.get("text") or "")
                semantic = item.get("formula_semantics") or classify_formula_text(formula_text).to_dict()
                if semantic.get("category") == CATEGORY_CONTAMINATED or _formula_text_looks_contaminated(formula_text):
                    out.append(f"{sec.get('heading')}: contaminated formula text `{formula_text[:100]}`")
                    if len(out) >= limit:
                        return out
                    streak = []
                    continue
            text = item if isinstance(item, str) else (item.get("text") if isinstance(item, dict) else "")
            t = str(text or "").strip()
            if token_re.match(t):
                streak.append(t)
                if len(streak) >= 3:
                    out.append(f"{sec.get('heading')}: {' / '.join(streak[-6:])}")
                    if len(out) >= limit:
                        return out
            else:
                streak = []
    return out

def _looks_like_low_res_text_fragment(width: int, height: int, context: str = "") -> bool:
    if width <= 0 or height <= 0:
        return False
    context = str(context or "")
    context_hint = bool(re.search(
        r"(fragment|shard|broken|label|formula|equation|text|ocr|"
        r"碎片|残片|公式|方程|标签|标注|文字|截图|图表)",
        context,
        re.I,
    ))
    wide_ratio = width / max(height, 1)
    tall_ratio = height / max(width, 1)
    if width < 160 and height < 45 and wide_ratio >= 2.4:
        return True
    if height < 80 and wide_ratio >= 4.0:
        return True
    if width < 120 and tall_ratio >= 4.0:
        return True
    if context_hint and width < 240 and height < 120 and max(wide_ratio, tall_ratio) >= 1.8:
        return True
    return False


def _low_res_image_fragment_samples(content: Dict[str, Any], out_dir: str, limit: int = 8) -> List[str]:
    try:
        from PIL import Image
    except Exception:
        return []
    meta = content.get("_meta") or {}
    images_dir = str(meta.get("images_dir") or "")
    candidates = []
    if images_dir:
        candidates.append(images_dir)
    candidates.append(os.path.join(out_dir, "figures"))
    samples: List[str] = []
    for ref in _iter_content_image_refs(content):
        name = ref["name"]
        path = next((os.path.join(base, name) for base in candidates if base and os.path.exists(os.path.join(base, name))), "")
        if not path:
            continue
        try:
            with Image.open(path) as im:
                width, height = im.size
        except Exception:
            continue
        context = ref.get("caption") or ref.get("heading") or ""
        if _looks_like_low_res_text_fragment(width, height, context):
            detail = f"{name} {width}x{height}"
            if context:
                detail += f" ({context[:60]})"
            samples.append(detail)
            if len(samples) >= limit:
                break
    return samples

