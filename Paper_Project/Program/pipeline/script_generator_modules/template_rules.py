"""Template page and layout-rule inference for script_generator.py."""
from __future__ import annotations

import re
from typing import Any, Dict


def _text_blob(fmt: Dict[str, Any]) -> str:
    return "\n".join(str(p.get("text") or "") for p in fmt.get("paragraphs") or [])


def _extract_page_and_header(fmt: Dict[str, Any]) -> Dict[str, Any]:
    sections = fmt.get("sections") or []
    s0 = sections[0] if sections else {}
    page = {
        "page_w": s0.get("page_width_cm", 21.0),
        "page_h": s0.get("page_height_cm", 29.7),
        "mt": s0.get("margin_top_cm", 2.54),
        "mb": s0.get("margin_bottom_cm", 2.54),
        "ml": s0.get("margin_left_cm", 2.54),
        "mr": s0.get("margin_right_cm", 2.54),
        "header": None,
    }
    for sec in sections:
        for h in sec.get("header", []) or []:
            text = (h.get("text") or "").strip()
            if not text:
                continue
            run = next(
                (r for r in h.get("runs", []) if str(r.get("text", "")).strip()),
                next((r for r in h.get("runs", []) if r.get("size_pt")), {}),
            )
            page["header"] = {
                "text": text,
                "align": h.get("alignment") if h.get("alignment") != "DEFAULT" else "CENTER",
                "font": run.get("font") or "宋体",
                "size": run.get("size_pt") or 9,
                "bold": bool(run.get("bold", False)),
                "italic": bool(run.get("italic", False)),
            }
            return page
    return page


def _infer_template_rules(fmt: Dict[str, Any]) -> Dict[str, Any]:
    """Infer layout rules from template instruction text, not fixed names."""
    texts = _text_blob(fmt)
    caption_samples = [
        str(p.get("text") or "").strip()
        for p in fmt.get("paragraphs") or []
        if re.match(r"^(图|表)\s*\d+(?:[.-]\d+)?\s+", str(p.get("text") or "").strip())
    ]
    caption_number_space = None
    for sample in caption_samples:
        if re.match(r"^(图|表)\s+\d", sample):
            caption_number_space = True
            break
        if re.match(r"^(图|表)\d", sample):
            caption_number_space = False
            break
    ref_indent_chars = (
        2.0
        if re.search(
            r"参考文献[^\n。；;]{0,120}悬挂缩进[^\n。；;]{0,20}2\s*字符|悬挂缩进[^\n。；;]{0,20}2\s*字符",
            texts,
        )
        else None
    )
    return {
        "cn_abstract_single_paragraph": bool(re.search(r"中文摘要[^\n。；;]{0,80}不分自然段|不分自然段[^\n。；;]{0,80}中文摘要", texts)),
        "en_title_upper": bool(re.search(r"英文题目[^\n。；;]{0,120}(大写字母|大写)", texts)),
        "caption_number_space": caption_number_space,
        "formula_center": bool(re.search(r"公式[^\n。；;]{0,60}居中", texts)),
        "formula_numbered": bool(re.search(r"公式[^\n。；;]{0,80}(编号|括弧|括号)", texts)),
        "reference_hanging_chars": ref_indent_chars,
        "reference_english_left": bool(re.search(r"英文参考文献[^\n。；;]{0,80}左对齐", texts)),
        "toc_indents_cm": [0.0, 0.74, 1.48],
    }
