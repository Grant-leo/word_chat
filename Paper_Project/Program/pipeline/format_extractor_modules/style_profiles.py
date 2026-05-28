"""Semantic style-profile inference from extracted template paragraphs."""
from __future__ import annotations

import re


def first_real_run(paragraph):
    for run in paragraph.get("runs", []):
        if run.get("text", "").strip() or run.get("size_pt"):
            return run
    return (paragraph.get("runs") or [{}])[0] if paragraph.get("runs") else {}


def profile_from_paragraph(paragraph):
    run = first_real_run(paragraph)
    return {
        "font": run.get("font") or "宋体",
        "size": run.get("size_pt") or 12,
        "bold": bool(run.get("bold", False)),
        "italic": bool(run.get("italic", False)),
        "align": paragraph.get("alignment") or paragraph.get("align") or "LEFT",
        "line_spacing_val": paragraph.get("line_spacing_val") or paragraph.get("ls"),
        "line_spacing_rule": paragraph.get("line_spacing_rule"),
        "line_spacing_fixed_pt": paragraph.get("line_spacing_fixed_pt"),
        "space_before_pt": paragraph.get("space_before_pt"),
        "space_after_pt": paragraph.get("space_after_pt"),
        "first_indent_cm": paragraph.get("first_indent_cm")
        if paragraph.get("first_indent_cm") is not None
        else paragraph.get("indent", 0),
        "left_indent_cm": paragraph.get("left_indent_cm"),
        "right_indent_cm": paragraph.get("right_indent_cm"),
        "hanging_indent_cm": paragraph.get("hanging_indent_cm"),
    }


def build_style_profiles(fmt):
    """Infer semantic style profiles from the template examples/instructions."""
    profiles = {}
    paras = fmt.get("paragraphs", [])

    def put(role, paragraph):
        if role not in profiles and paragraph:
            profiles[role] = profile_from_paragraph(paragraph)

    for paragraph in paras:
        text = (paragraph.get("text") or "").strip()
        if not text:
            continue
        no_space = text.replace(" ", "")
        if "论文" in text and "题目" in text and ("居中" in text or paragraph.get("style") == "论文题目"):
            put("cn_title", paragraph)
        if no_space.startswith("摘要") or no_space.startswith("摘要（") or no_space.startswith("摘要("):
            if len(text) < 30:
                put("cn_abstract_heading", paragraph)
        if text.startswith("摘要是") or ("中文摘要300" in text and len(text) > 50):
            put("cn_abstract_body", paragraph)
        if text.startswith("关键词"):
            put("cn_keywords", paragraph)
        if "英文题目" in text and ("Times" in text or "Roman" in text):
            put("en_title", paragraph)
        if text.upper().startswith("ABSTRACT") and len(text) < 40:
            put("en_abstract_heading", paragraph)
        if len(text) > 80 and sum(1 for c in text[:120] if c.isascii() and c.isalpha()) > 50:
            put("en_abstract_body", paragraph)
        if text.upper().startswith("KEY WORD") or text.upper().startswith("KEYWORDS"):
            put("en_keywords", paragraph)
        if no_space in ("目录", "目 录".replace(" ", "")) or no_space.startswith("目录"):
            put("toc_title", paragraph)
        if ("一级标题" in text or "第1章" in text) and len(text) < 80:
            put("h1", paragraph)
        if ("二级标题" in text or re.match(r"^1\.1\s+", text)) and len(text) < 80:
            put("h2", paragraph)
        if ("三级标题" in text or re.match(r"^1\.1\.1\s+", text)) and len(text) < 80:
            put("h3", paragraph)

    def heading_score(paragraph, level):
        text = (paragraph.get("text") or "").strip()
        run = first_real_run(paragraph)
        font = run.get("font") or ""
        size = run.get("size_pt") or 0
        score = 0
        if font and font not in ("Arial", "Times New Roman", "Calibri"):
            score += 10
        if level == 1 and re.match(r"^第[一二三四五六七八九十\d]+章\s+", text):
            score += 8
        if level == 2 and re.match(r"^\d+\.\d+\s+", text):
            score += 8
        if level == 3 and re.match(r"^\d+\.\d+\.\d+\s+", text):
            score += 8
        if size:
            score += min(float(size), 20) / 2
        if "标题" in text and "（" in text:
            score -= 4
        if "目录" in text:
            score -= 8
        return score

    for role, level, pattern in [
        ("h1", 1, r"^第[一二三四五六七八九十\d]+章\s+"),
        ("h2", 2, r"^\d+\.\d+\s+"),
        ("h3", 3, r"^\d+\.\d+\.\d+\s+"),
    ]:
        candidates = [
            paragraph
            for paragraph in paras
            if re.match(pattern, (paragraph.get("text") or "").strip())
            and len((paragraph.get("text") or "").strip()) < 80
        ]
        if candidates:
            best = max(candidates, key=lambda paragraph: heading_score(paragraph, level))
            profiles[role] = profile_from_paragraph(best)
            profiles[role]["bold"] = bool(profiles[role].get("bold")) or True
            profiles[role]["first_indent_cm"] = 0 if level == 1 else profiles[role].get("first_indent_cm", 0)

    candidates = []
    for paragraph in paras:
        text = (paragraph.get("text") or "").strip()
        if len(text) < 80:
            continue
        if any(key in text[:60] for key in ["本人郑重声明", "本人在导师", "格式", "要求", "行距", "字号", "页眉"]):
            continue
        if any("\u4e00" <= char <= "\u9fff" for char in text[:120]):
            candidates.append(paragraph)
    if candidates:
        put("body", candidates[0])

    body = profiles.get("body") or {
        "font": "宋体",
        "size": 12,
        "align": "JUSTIFY",
        "line_spacing_fixed_pt": 28,
        "first_indent_cm": 0.74,
    }
    profiles.setdefault("body", body)
    profiles.setdefault("h1", {**body, "font": "黑体", "size": 16, "bold": True, "align": "CENTER", "first_indent_cm": 0})
    profiles.setdefault("h2", {**body, "font": "黑体", "size": 14, "bold": True, "align": "LEFT", "first_indent_cm": 0})
    profiles.setdefault("h3", {**body, "font": "黑体", "size": 12, "bold": True, "align": "LEFT", "first_indent_cm": 0})
    profiles.setdefault("cn_title", profiles["h1"])
    profiles.setdefault("cn_abstract_heading", profiles["h1"])
    profiles.setdefault("cn_abstract_body", body)
    profiles.setdefault("cn_keywords", {**body, "first_indent_cm": 0})
    profiles.setdefault("en_title", {**profiles["h1"], "font": "Times New Roman"})
    profiles.setdefault(
        "en_abstract_heading",
        {**profiles["h1"], "font": "Times New Roman", "size": 16, "bold": True, "align": "CENTER", "first_indent_cm": 0},
    )
    profiles.setdefault(
        "en_abstract_body",
        {**body, "font": "Times New Roman", "line_spacing_fixed_pt": None, "line_spacing_val": 1.5, "first_indent_cm": 0.9},
    )
    profiles.setdefault(
        "figure_caption",
        {
            **body,
            "font": "宋体",
            "size": 10.5,
            "align": "CENTER",
            "first_indent_cm": 0,
            "space_before_pt": 6,
            "space_after_pt": 6,
            "line_spacing_fixed_pt": 28,
        },
    )
    profiles.setdefault("table_caption", {**profiles["figure_caption"]})
    profiles.setdefault(
        "code",
        {**body, "font": "Consolas", "size": 10.5, "align": "LEFT", "first_indent_cm": 0, "line_spacing_fixed_pt": None, "line_spacing_val": 1.0},
    )
    profiles.setdefault(
        "reference",
        {
            **body,
            "font": "宋体",
            "size": 12,
            "align": "JUSTIFY",
            "first_indent_cm": 0,
            "line_spacing_fixed_pt": 28,
            "space_before_pt": 6,
            "space_after_pt": 6,
        },
    )
    profiles.setdefault("en_keywords", {**profiles["en_abstract_body"], "bold": True, "first_indent_cm": 0})
    profiles.setdefault("toc_title", profiles["h1"])

    for role in ("h1", "h2", "h3", "cn_title", "cn_abstract_heading", "toc_title"):
        profile = profiles.get(role, {})
        try:
            line_spacing = float(profile.get("line_spacing_val") or 0)
        except Exception:
            line_spacing = 0
        if line_spacing > 10 and not profile.get("line_spacing_fixed_pt"):
            profile["line_spacing_val"] = body.get("line_spacing_val")
            profile["line_spacing_fixed_pt"] = body.get("line_spacing_fixed_pt")
            profile["line_spacing_rule"] = body.get("line_spacing_rule")
    return profiles

