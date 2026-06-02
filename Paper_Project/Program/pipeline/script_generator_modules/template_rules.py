"""Template page and layout-rule inference for script_generator.py."""
from __future__ import annotations

import re
from typing import Any, Dict


def _text_blob(fmt: Dict[str, Any]) -> str:
    return "\n".join(str(p.get("text") or "") for p in fmt.get("paragraphs") or [])


def _is_template_instruction_fragment(text: str) -> bool:
    compact = re.sub(r"\s+", "", str(text or ""))
    if not compact:
        return False
    if "完成时间按照答辩时间填写" in compact:
        return True
    font_signal = bool(re.search(r"TimesNewRoman|宋体|黑体|楷体|仿宋|华文|字号|[一二三四五六七八九十小]+号", compact, re.I))
    layout_signal = bool(re.search(r"居中|加粗|行距|倍行距|段前|段后|缩进|对齐|表格行高|固定值|页边距", compact))
    subject_signal = bool(re.search(r"英文题目|中文题目|目录内容|一级标题|二级标题|三级标题|图表题注|参考文献|页眉|页脚|新罗马字体", compact))
    return bool((font_signal and layout_signal) or (subject_signal and (font_signal or layout_signal)))


def _strip_template_instruction_fragments(text: str) -> str:
    value = str(text or "")

    def replace(match: re.Match[str]) -> str:
        inner = match.group(0)[1:-1]
        return "" if _is_template_instruction_fragment(inner) else match.group(0)

    previous = None
    while previous != value:
        previous = value
        value = re.sub(r"[（(][^（）()]{1,120}[）)]", replace, value)
    return re.sub(r"\s{2,}", " ", value).strip()


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
            text = _strip_template_instruction_fragments(h.get("text") or "")
            if not text:
                continue
            clean_runs = []
            for r in h.get("runs", []) or []:
                rt = _strip_template_instruction_fragments(r.get("text") or "")
                if rt:
                    nr = dict(r)
                    nr["text"] = rt
                    clean_runs.append(nr)
            run = next(
                (r for r in clean_runs if str(r.get("text", "")).strip()),
                next((r for r in clean_runs if r.get("size_pt")), {}),
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
