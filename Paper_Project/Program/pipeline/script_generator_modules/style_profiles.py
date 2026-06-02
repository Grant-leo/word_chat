"""Style profile inference for script_generator.py."""
from __future__ import annotations

import re
from typing import Any, Dict, Optional


def _is_cjk(text: str) -> bool:
    return any("\u4e00" <= c <= "\u9fff" for c in str(text or ""))


def _ascii_ratio(text: str) -> float:
    text = str(text or "")
    if not text:
        return 0.0
    return sum(1 for c in text if c.isascii() and c.isalpha()) / max(len(text), 1)


_CN_SIZE_PATTERNS = [
    ("小二", 18.0),
    ("二号", 22.0),
    ("小三", 15.0),
    ("三号", 16.0),
    ("小四", 12.0),
    ("四号", 14.0),
    ("小五", 9.0),
    ("五号", 10.5),
]


def _text_blob(fmt: Dict[str, Any]) -> str:
    return "\n".join(str(p.get("text") or "") for p in fmt.get("paragraphs") or [])


def _find_instruction(texts: str, *needles: str) -> str:
    """Return a compact instruction line/paragraph that contains all needles."""
    chunks = re.split(r"[\r\n]+", texts)
    chunks += re.split(r"[。；;]\s*", texts)
    for chunk in chunks:
        if all(n in chunk for n in needles):
            return chunk.strip()
    return ""


def _find_regex_instruction(texts: str, pattern: str) -> str:
    m = re.search(pattern, texts, re.S)
    return m.group(1).strip() if m else ""


def _size_from_text(text: str, default: Optional[float] = None) -> Optional[float]:
    text = str(text or "")
    m = re.search(r"(\d+(?:\.\d+)?)\s*pt", text, re.I)
    if m:
        return float(m.group(1))
    for name, size in _CN_SIZE_PATTERNS:
        if name in text:
            return size
    return default


def _font_from_text(text: str, default: Optional[str] = None) -> Optional[str]:
    text = str(text or "")
    if "timesnewroman" in re.sub(r"\s+", "", text).lower():
        return "Times New Roman"
    for font in ("Times New Roman", "黑体", "宋体", "楷体_GB2312", "楷体", "仿宋", "微软雅黑", "华文中宋", "方正小标宋简体"):
        if font.lower() in text.lower():
            return font
    return default


def _align_from_text(text: str, default: Optional[str] = None) -> Optional[str]:
    if "居中" in text:
        return "CENTER"
    if "右对齐" in text or "靠右" in text:
        return "RIGHT"
    if "左对齐" in text or "靠左" in text or "左侧" in text:
        return "LEFT"
    if "两端对齐" in text or "右侧也要对齐" in text:
        return "JUSTIFY"
    return default


def _line_spacing_from_text(text: str) -> Dict[str, Any]:
    text = str(text or "")
    m = re.search(r"固定值\s*(\d+(?:\.\d+)?)\s*磅", text)
    if m:
        v = float(m.group(1))
        return {"line_spacing_val": v, "line_spacing_rule": "exact", "line_spacing_fixed_pt": v}
    if re.search(r"1\.5\s*倍|1\.5倍", text):
        return {"line_spacing_val": 1.5, "line_spacing_rule": "auto", "line_spacing_fixed_pt": None}
    if "单倍" in text:
        return {"line_spacing_val": 1.0, "line_spacing_rule": "auto", "line_spacing_fixed_pt": None}
    return {}


def _has_format_instruction(text: str) -> bool:
    text = str(text or "")
    if not text:
        return False
    compact = re.sub(r"\s+", "", text).lower()
    if re.search(r"\d+(?:\.\d+)?\s*(?:pt|磅)", text, re.I):
        return True
    terms = (
        "宋体",
        "黑体",
        "楷体",
        "楷体_GB2312",
        "仿宋",
        "微软雅黑",
        "华文宋体",
        "华文中宋",
        "方正小标宋简体",
        "Times New Roman",
        "小二",
        "二号",
        "小三",
        "三号",
        "小四",
        "四号",
        "小五",
        "五号",
        "加粗",
        "不加粗",
        "居中",
        "左对齐",
        "右对齐",
        "两端对齐",
        "固定值",
        "行距",
        "倍行距",
        "单倍",
        "缩进",
        "首行",
        "段前",
        "段后",
    )
    return any(re.sub(r"\s+", "", term).lower() in compact for term in terms)


def _spacing_before_after_from_text(text: str, line_pt: Optional[float] = None) -> Dict[str, Any]:
    text = str(text or "")
    out: Dict[str, Any] = {}
    m = re.search(r"段前段后各?\s*(\d+(?:\.\d+)?)\s*磅", text)
    if m:
        v = float(m.group(1))
        out["space_before_pt"] = v
        out["space_after_pt"] = v
        return out
    m = re.search(r"段前段后\s*(\d+(?:\.\d+)?)\s*磅", text)
    if m:
        v = float(m.group(1))
        out["space_before_pt"] = v
        out["space_after_pt"] = v
        return out
    if re.search(r"段前段后各?\s*1\s*行", text):
        v = float(line_pt or 28.0)
        out["space_before_pt"] = v
        out["space_after_pt"] = v
    elif re.search(r"段前段后\s*0(?:\.5|点5|半)\s*行", text):
        v = float(line_pt or 28.0) * 0.5
        out["space_before_pt"] = v
        out["space_after_pt"] = v
    elif re.search(r"段前段后\s*0\s*行|段前段后0", text):
        out["space_before_pt"] = 0.0
        out["space_after_pt"] = 0.0
    return out


def _indent_from_text(text: str, size_pt: Optional[float] = None, default: Optional[float] = None) -> Optional[float]:
    text = str(text or "")
    m = re.search(r"缩进\s*(\d+(?:\.\d+)?)\s*(?:个)?(?:汉)?字(?:符)?", text)
    if m:
        chars = float(m.group(1))
        return round(chars * float(size_pt or 12.0) * 0.0352778, 2)
    return default


def _profile_from_instruction(text: str, base: Dict[str, Any], allow_bold: bool = True, **defaults: Any) -> Dict[str, Any]:
    prof = dict(base)
    prof.update(defaults)
    font = _font_from_text(text)
    size = _size_from_text(text)
    align = _align_from_text(text)
    if font:
        prof["font"] = font
    if size:
        prof["size"] = size
    if align:
        prof["align"] = align
    if allow_bold:
        if "不加粗" in text:
            prof["bold"] = False
        elif "加粗" in text:
            prof["bold"] = True
    prof.update(_line_spacing_from_text(text))
    line_pt = prof.get("line_spacing_fixed_pt") or (float(prof.get("size") or 12) * float(prof.get("line_spacing_val") or 1.5))
    prof.update(_spacing_before_after_from_text(text, line_pt))
    ind = _indent_from_text(text, prof.get("size"))
    if ind is not None:
        prof["first_indent_cm"] = ind
    return _normalize_profile(prof, base)


def _find_body_rule(texts: str) -> str:
    candidates = [
        _find_regex_instruction(texts, r"摘要内容为([^。；;\n]*)"),
        _find_regex_instruction(texts, r"论文正文[^。；;\n]*?(宋体[^。；;\n]*?(?:固定值|行距)[^。；;\n]*)"),
        _find_instruction(texts, "论文正文", "行距"),
        _find_instruction(texts, "正文从", "行距"),
        _find_instruction(texts, "正文", "首行缩进", "两端对齐"),
    ]
    for rule in candidates:
        if not rule:
            continue
        if ("目录" in rule and "正文从" not in rule and "论文正文" not in rule) or "正文的一级标题" in rule or "正文其他层次标题" in rule:
            continue
        if _has_format_instruction(rule) or "行距" in rule or "首行缩进" in rule:
            return rule
    return ""


def _find_reference_rule(texts: str) -> str:
    candidates = [
        _find_regex_instruction(texts, r"参考文献中中文使用([^。；;\n]*)"),
        _find_instruction(texts, "参考文献格式要求"),
        _find_instruction(texts, "参考文献", "宋体", "小四"),
    ]
    for rule in candidates:
        if not rule:
            continue
        if any(marker in rule for marker in ("目录", "一级标题", "正文的", "正文其他层次标题")):
            continue
        if _has_format_instruction(rule):
            return rule
    return ""


def _reference_sample_profile(paras: list[Dict[str, Any]], body: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    for p in paras:
        txt = (p.get("text") or "").strip()
        if not re.match(r"^\[\d+\]\s+\S", txt):
            continue
        if _has_format_instruction(txt):
            continue
        prof = _profile_from_para_first_text(p, body)
        prof["font"] = "宋体"
        prof["bold"] = False
        prof["align"] = "LEFT" if prof.get("align") in ("CENTER", "RIGHT") else (prof.get("align") or "LEFT")
        prof["first_indent_cm"] = 0.0
        prof["space_before_pt"] = 0.0
        prof["space_after_pt"] = 0.0
        return _normalize_profile(prof, body)
    return None


def _first_run(p: Dict[str, Any]) -> Dict[str, Any]:
    runs = [r for r in (p.get("runs") or []) if r.get("text", "").strip() or r.get("size_pt")]
    if not runs:
        return (p.get("runs") or [{}])[0] if p.get("runs") else {}

    def score(r: Dict[str, Any]) -> float:
        txt = r.get("text", "") or ""
        font = r.get("font", "") or ""
        size = float(r.get("size_pt") or 0)
        return size + (5 if _is_cjk(txt) else 0) + (3 if font and font not in ("Arial", "Times New Roman", "Calibri") else 0)

    return max(runs, key=score)


def _first_text_run(p: Dict[str, Any]) -> Dict[str, Any]:
    for r in p.get("runs") or []:
        if str(r.get("text") or "").strip():
            return r
    return _first_run(p)


def _profile_from_para_first_text(p: Dict[str, Any], fallback: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    r = _first_text_run(p)
    prof = {
        "font": r.get("font") or (fallback or {}).get("font") or "宋体",
        "size": r.get("size_pt") or (fallback or {}).get("size") or 12,
        "bold": bool(r.get("bold", (fallback or {}).get("bold", False))),
        "italic": bool(r.get("italic", (fallback or {}).get("italic", False))),
        "align": p.get("alignment") or p.get("align") or (fallback or {}).get("align") or "LEFT",
        "line_spacing_val": p.get("line_spacing_val") if p.get("line_spacing_val") is not None else p.get("ls", (fallback or {}).get("line_spacing_val")),
        "line_spacing_rule": p.get("line_spacing_rule") or (fallback or {}).get("line_spacing_rule"),
        "line_spacing_fixed_pt": p.get("line_spacing_fixed_pt") or (fallback or {}).get("line_spacing_fixed_pt"),
        "space_before_pt": p.get("space_before_pt") if p.get("space_before_pt") is not None else (fallback or {}).get("space_before_pt", 0),
        "space_after_pt": p.get("space_after_pt") if p.get("space_after_pt") is not None else (fallback or {}).get("space_after_pt", 0),
        "first_indent_cm": p.get("first_indent_cm") if p.get("first_indent_cm") is not None else p.get("indent", (fallback or {}).get("first_indent_cm", 0)),
        "left_indent_cm": p.get("left_indent_cm", (fallback or {}).get("left_indent_cm")),
        "right_indent_cm": p.get("right_indent_cm", (fallback or {}).get("right_indent_cm")),
        "hanging_indent_cm": p.get("hanging_indent_cm", (fallback or {}).get("hanging_indent_cm")),
    }
    return _normalize_profile(prof, fallback)


def _normalize_profile(prof: Optional[Dict[str, Any]], fallback: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    fallback = fallback or {}
    p = dict(fallback)
    if prof:
        for k, v in prof.items():
            if v is not None or k in ("line_spacing_fixed_pt",):
                p[k] = v
    p.setdefault("font", "宋体")
    p.setdefault("size", 12)
    p.setdefault("bold", False)
    p.setdefault("italic", False)
    p.setdefault("align", "LEFT")
    if p.get("align") == "DEFAULT":
        p["align"] = fallback.get("align", "LEFT")
    p.setdefault("line_spacing_val", 1.5)
    p.setdefault("line_spacing_fixed_pt", None)
    p.setdefault("space_before_pt", 0)
    p.setdefault("space_after_pt", 0)
    p.setdefault("first_indent_cm", 0)
    try:
        p["size"] = float(p.get("size") or fallback.get("size") or 12)
    except Exception:
        p["size"] = 12.0
    for k in ("space_before_pt", "space_after_pt", "first_indent_cm"):
        try:
            p[k] = float(p.get(k) or 0)
        except Exception:
            p[k] = 0.0
    for k in ("left_indent_cm", "right_indent_cm", "hanging_indent_cm"):
        if k in p and p.get(k) is not None:
            try:
                p[k] = float(p.get(k) or 0)
            except Exception:
                p.pop(k, None)
    try:
        if p.get("line_spacing_fixed_pt") is not None:
            p["line_spacing_fixed_pt"] = float(p.get("line_spacing_fixed_pt"))
    except Exception:
        p["line_spacing_fixed_pt"] = None
    try:
        if p.get("line_spacing_val") is not None:
            p["line_spacing_val"] = float(p.get("line_spacing_val"))
    except Exception:
        p["line_spacing_val"] = fallback.get("line_spacing_val", 1.5)
    if p.get("line_spacing_val") and p.get("line_spacing_val") > 20 and not p.get("line_spacing_fixed_pt"):
        p["line_spacing_val"] = fallback.get("line_spacing_val", 1.5)
    return p


def _profile_from_para(p: Dict[str, Any], fallback: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    r = _first_run(p)
    prof = {
        "font": r.get("font") or (fallback or {}).get("font") or "宋体",
        "size": r.get("size_pt") or (fallback or {}).get("size") or 12,
        "bold": bool(r.get("bold", (fallback or {}).get("bold", False))),
        "italic": bool(r.get("italic", (fallback or {}).get("italic", False))),
        "align": p.get("alignment") or p.get("align") or (fallback or {}).get("align") or "LEFT",
        "line_spacing_val": p.get("line_spacing_val") if p.get("line_spacing_val") is not None else p.get("ls", (fallback or {}).get("line_spacing_val")),
        "line_spacing_rule": p.get("line_spacing_rule") or (fallback or {}).get("line_spacing_rule"),
        "line_spacing_fixed_pt": p.get("line_spacing_fixed_pt") or (fallback or {}).get("line_spacing_fixed_pt"),
        "space_before_pt": p.get("space_before_pt") if p.get("space_before_pt") is not None else (fallback or {}).get("space_before_pt", 0),
        "space_after_pt": p.get("space_after_pt") if p.get("space_after_pt") is not None else (fallback or {}).get("space_after_pt", 0),
        "first_indent_cm": p.get("first_indent_cm") if p.get("first_indent_cm") is not None else p.get("indent", (fallback or {}).get("first_indent_cm", 0)),
        "left_indent_cm": p.get("left_indent_cm", (fallback or {}).get("left_indent_cm")),
        "right_indent_cm": p.get("right_indent_cm", (fallback or {}).get("right_indent_cm")),
        "hanging_indent_cm": p.get("hanging_indent_cm", (fallback or {}).get("hanging_indent_cm")),
    }
    return _normalize_profile(prof, fallback)


def _infer_style_profiles(fmt: Dict[str, Any]) -> Dict[str, Dict[str, Any]]:
    existing = {k: _normalize_profile(v) for k, v in (fmt.get("style_profiles") or {}).items()}
    paras = fmt.get("paragraphs") or []
    profiles: Dict[str, Dict[str, Any]] = dict(existing)

    def put(role: str, p: Dict[str, Any]) -> None:
        if p and role not in profiles:
            profiles[role] = _profile_from_para(p)

    def style_sample_candidate(p: Dict[str, Any], txt: str) -> bool:
        style = str(p.get("style") or "").lower()
        if "toc" in style or "table of figures" in style:
            return False
        if len(txt) > 120:
            return False
        note_markers = ("本模板", "样式", "格式要求", "设置为", "目录采用", "正文中", "建议")
        return not any(marker in txt for marker in note_markers)

    for p in paras:
        txt = (p.get("text") or "").strip()
        if not txt:
            continue
        compact = re.sub(r"\s+", "", txt)
        up = txt.upper()
        if "论文" in txt and "题目" in txt and len(txt) < 100 and ("居中" in txt or p.get("style") == "论文题目"):
            put("cn_title", p)
        if re.match(r"^摘\s*要(?:[（(]|$)", txt) and len(txt) < 40:
            put("cn_abstract_heading", p)
        if txt.startswith("摘要是") or ("中文摘要" in txt and len(txt) > 60):
            put("cn_abstract_body", p)
        if txt.startswith("关键词"):
            put("cn_keywords", p)
        if "英文题目" in txt and ("Times" in txt or "Roman" in txt):
            profiles.setdefault("en_title", _profile_from_para_first_text(p))
        if up.startswith("ABSTRACT") and len(txt) < 40:
            profiles.setdefault("en_abstract_heading", _profile_from_para_first_text(p))
        if len(txt) > 80 and _ascii_ratio(txt[:160]) > 0.55:
            put("en_abstract_body", p)
        if up.startswith("KEY WORD") or up.startswith("KEYWORDS"):
            profiles.setdefault("en_keywords", _profile_from_para_first_text(p))
        if compact in ("目录", "目目录") or compact.startswith("目录"):
            put("toc_title", p)
        if ("一级标题" in txt or re.match(r"^第[一二三四五六七八九十\d]+章\s+", txt)) and len(txt) < 100 and style_sample_candidate(p, txt):
            put("h1", p)
        if ("二级标题" in txt or re.match(r"^\d+\.\d+\s+", txt)) and len(txt) < 100 and style_sample_candidate(p, txt):
            put("h2", p)
        if ("三级标题" in txt or re.match(r"^\d+\.\d+\.\d+\s+", txt)) and len(txt) < 100 and style_sample_candidate(p, txt):
            put("h3", p)
        if re.match(r"^(图|表)\s*\d+", txt) and len(txt) < 80 and style_sample_candidate(p, txt):
            put("figure_caption" if txt.startswith("图") else "table_caption", p)
        if "参考文献" in txt and len(txt) < 40:
            put("reference_heading", p)

    def heading_score(p: Dict[str, Any], level: int) -> float:
        txt = (p.get("text") or "").strip()
        r = _first_run(p)
        font = r.get("font") or ""
        size = float(r.get("size_pt") or 0)
        score = size
        if p.get("style", "").lower().startswith("heading"):
            score += 20
        if "\t" in txt or re.search(r"\s\d+$", txt):
            score -= 18
        if font and font not in ("Arial", "Times New Roman", "Calibri"):
            score += 8
        if level == 1 and re.match(r"^第[一二三四五六七八九十\d]+章\s+", txt):
            score += 10
        if level == 2 and re.match(r"^\d+\.\d+\s+", txt):
            score += 10
        if level == 3 and re.match(r"^\d+\.\d+\.\d+\s+", txt):
            score += 10
        return score

    for role, level, pat in [
        ("h1", 1, r"^第[一二三四五六七八九十\d]+章\s+"),
        ("h2", 2, r"^\d+\.\d+\s+"),
        ("h3", 3, r"^\d+\.\d+\.\d+\s+"),
    ]:
        cands = [
            p
            for p in paras
            if re.match(pat, (p.get("text") or "").strip())
            and len((p.get("text") or "").strip()) < 100
            and style_sample_candidate(p, (p.get("text") or "").strip())
        ]
        if cands and role not in profiles:
            profiles[role] = _profile_from_para(max(cands, key=lambda x: heading_score(x, level)))
            profiles[role]["bold"] = True
            profiles[role]["first_indent_cm"] = 0
            if level == 1:
                profiles[role]["align"] = "CENTER"

    def body_sample_score(p: Dict[str, Any]) -> Optional[float]:
        txt = (p.get("text") or "").strip()
        if len(txt) < 80:
            return None
        style = str(p.get("style") or "").lower()
        if "toc" in style or "table of figures" in style:
            return None
        note_markers = ("本模板", "格式", "格式要求", "设置为", "目录采用", "正文中", "建议", "字体要求", "页眉页脚")
        if any(marker in txt[:120] for marker in note_markers):
            return None
        if "\t" in txt or re.search(r"\s(?:[ivxlcdm]+|\d+)$", txt, re.I):
            return None
        if txt[:1] in ("(", "（") and _has_format_instruction(txt):
            return None
        r = _first_text_run(p)
        try:
            size = float(r.get("size_pt") or 0)
        except Exception:
            size = 0.0
        if size and size > 13.5:
            return None
        score = 0.0
        ar = _ascii_ratio(txt[:240])
        if ar > 0.55 and len(txt) >= 120:
            score += 80.0
        elif _is_cjk(txt):
            score += 50.0
        else:
            return None
        if not bool(r.get("bold", False)):
            score += 20.0
        if size:
            score += max(0.0, 12.0 - abs(size - 12.0))
        if (p.get("alignment") or p.get("align")) in ("JUSTIFY", "BOTH"):
            score += 4.0
        return score

    body_cands = []
    for p in paras:
        score = body_sample_score(p)
        if score is not None:
            body_cands.append((score, p))
    if body_cands:
        profiles["body"] = _profile_from_para(max(body_cands, key=lambda item: item[0])[1])

    body = _normalize_profile(profiles.get("body") or {"font": "宋体", "size": 12, "align": "JUSTIFY", "line_spacing_fixed_pt": 28, "first_indent_cm": 0.74})
    profiles["body"] = body
    profiles.setdefault("h1", _normalize_profile({"font": "黑体", "size": 16, "bold": True, "align": "CENTER", "line_spacing_fixed_pt": body.get("line_spacing_fixed_pt"), "first_indent_cm": 0}, body))
    profiles.setdefault("h2", _normalize_profile({"font": "黑体", "size": 14, "bold": True, "align": "LEFT", "line_spacing_fixed_pt": body.get("line_spacing_fixed_pt"), "first_indent_cm": 0}, body))
    profiles.setdefault("h3", _normalize_profile({"font": "黑体", "size": 12, "bold": True, "align": "LEFT", "line_spacing_fixed_pt": body.get("line_spacing_fixed_pt"), "first_indent_cm": 0}, body))
    profiles.setdefault("cn_title", profiles["h1"])
    profiles.setdefault("cn_abstract_heading", profiles["h1"])
    profiles.setdefault("cn_abstract_body", body)
    profiles.setdefault("cn_keywords", _normalize_profile({"first_indent_cm": 0, "bold": False}, body))
    profiles.setdefault("en_title", _normalize_profile({"font": "Times New Roman", "bold": True, "align": "CENTER"}, profiles["h1"]))
    profiles.setdefault("en_abstract_heading", _normalize_profile({"font": "Times New Roman", "size": 16, "bold": True, "align": "CENTER", "first_indent_cm": 0}, profiles["h1"]))
    profiles.setdefault("en_abstract_body", _normalize_profile({"font": "Times New Roman", "line_spacing_fixed_pt": None, "line_spacing_val": 1.5, "first_indent_cm": 0.9, "align": "JUSTIFY"}, body))
    profiles.setdefault("en_keywords", _normalize_profile({"font": "Times New Roman", "bold": False, "first_indent_cm": 0, "align": "LEFT"}, profiles["en_abstract_body"]))
    profiles.setdefault("toc_title", profiles["h1"])
    profiles.setdefault("figure_caption", _normalize_profile({"font": "宋体", "size": 10.5, "align": "CENTER", "first_indent_cm": 0, "space_before_pt": 6, "space_after_pt": 6, "line_spacing_fixed_pt": 28}, body))
    profiles.setdefault("table_caption", dict(profiles["figure_caption"]))
    profiles.setdefault("table_body", _normalize_profile({"font": "宋体", "size": 10.5, "align": "CENTER", "first_indent_cm": 0, "line_spacing_fixed_pt": None, "line_spacing_val": 1.0}, body))
    profiles.setdefault("table_header", _normalize_profile({"bold": True}, profiles["table_body"]))
    profiles.setdefault("formula", _normalize_profile({"align": "CENTER", "first_indent_cm": 0}, body))
    profiles.setdefault(
        "code",
        _normalize_profile(
            {
                "font": body.get("font") or "宋体",
                "size": body.get("size") or 12,
                "align": "LEFT",
                "first_indent_cm": 0,
                "line_spacing_fixed_pt": body.get("line_spacing_fixed_pt"),
                "line_spacing_val": body.get("line_spacing_val"),
                "line_spacing_rule": body.get("line_spacing_rule"),
                "space_before_pt": 0,
                "space_after_pt": 0,
            },
            body,
        ),
    )
    profiles.setdefault("reference", _normalize_profile({"font": "宋体", "size": 12, "align": "JUSTIFY", "first_indent_cm": 0, "space_before_pt": 6, "space_after_pt": 6, "line_spacing_fixed_pt": 28}, body))
    profiles.setdefault("reference_heading", profiles["h1"])
    return _apply_template_text_rules(fmt, profiles)


def _apply_template_text_rules(fmt: Dict[str, Any], profiles: Dict[str, Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:
    """Apply prose rules found in the template, without school-specific constants."""
    texts = _text_blob(fmt)
    paras = fmt.get("paragraphs") or []
    body = _normalize_profile(profiles.get("body") or {})
    body_rule = _find_body_rule(texts)
    if body_rule:
        body = _profile_from_instruction(body_rule, body, allow_bold=False, align="JUSTIFY")
    abstract_body_rule = _find_regex_instruction(texts, r"摘要内容为([^。；;\n]*)") or _find_instruction(texts, "摘要内容")
    if abstract_body_rule and ("宋体" in abstract_body_rule or "小四" in abstract_body_rule):
        body = _profile_from_instruction(abstract_body_rule, body, allow_bold=False, align="JUSTIFY")
    body.setdefault("font", "宋体")
    body.setdefault("size", 12.0)
    body.setdefault("align", "JUSTIFY")
    body.setdefault("line_spacing_fixed_pt", 28.0)
    body.setdefault("line_spacing_val", 28.0)
    body.setdefault("line_spacing_rule", "exact")
    profiles["body"] = body

    def role(name: str, **kw: Any) -> None:
        base = dict(body)
        base.update(kw)
        profiles[name] = _normalize_profile(base, body)

    cn_title_rule = (
        _find_regex_instruction(texts, r"(?:毕业)?论文(?:（设计）)?题目为([^。；;\n]*黑体[^。；;\n]*)")
        or _find_instruction(texts, "毕业论文", "题目", "黑体")
        or _find_instruction(texts, "论文题目", "黑体", "居中")
        or _find_instruction(texts, "宋体小三号", "加粗", "居中")
    )
    if cn_title_rule:
        profiles["cn_title"] = _profile_from_instruction(cn_title_rule, body, first_indent_cm=0.0)
    cn_abs_head_rule = (
        _find_regex_instruction(texts, r'[“"]?摘要[”"]?二字([^。；;\n]*宋体[^。；;\n]*)')
        or _find_regex_instruction(texts, r'[“"]?摘要[”"]?为([^。；;\n]*)')
        or _find_instruction(texts, "摘要", "宋体", "四号", "加粗")
        or _find_instruction(texts, "摘要", "居中")
    )
    if cn_abs_head_rule:
        profiles["cn_abstract_heading"] = _profile_from_instruction(cn_abs_head_rule, body, align="CENTER", first_indent_cm=0.0)
    if abstract_body_rule:
        cn_abs_base = profiles.get("cn_abstract_body") or body
        profiles["cn_abstract_body"] = _profile_from_instruction(
            abstract_body_rule,
            cn_abs_base,
            allow_bold=False,
            first_indent_cm=_indent_from_text(abstract_body_rule, 12.0, body.get("first_indent_cm")),
        )
    kw_rule = (
        _find_regex_instruction(texts, r"([“\"]?关键词[^。；;\n]*宋体[^。；;\n]*四号[^。；;\n]*)")
        or _find_instruction(texts, "关键词三字", "宋体", "四号")
        or _find_instruction(texts, "关键词", "宋体", "四号", "加粗")
    )
    kw_content_rule = _find_instruction(texts, "摘要和关键词内容", "小四")
    if kw_rule or kw_content_rule:
        kw_content_rule = kw_content_rule or kw_rule
        kw_base = profiles.get("cn_abstract_body") or body
        profiles["cn_keywords"] = _profile_from_instruction(
            kw_content_rule,
            kw_base,
            allow_bold=False,
            align="LEFT",
            first_indent_cm=0.0,
        )
        if kw_rule:
            profiles["cn_keywords_label"] = _profile_from_instruction(
                kw_rule,
                profiles["cn_keywords"],
                align="LEFT",
                first_indent_cm=0.0,
            )

    en_rule = _find_regex_instruction(texts, r"英文标题和摘要([^。；;\n]*)") or _find_instruction(texts, "英文标题", "摘要") or _find_instruction(texts, "英文题目", "摘要")
    en_title_rule = _find_regex_instruction(texts, r"论文题目为([^。；;\n]*Times\s*New\s*Roman[^。；;\n]*)") or _find_instruction(texts, "英文题目") or en_rule
    if en_title_rule:
        profiles["en_title"] = _profile_from_instruction(en_title_rule, body, font="Times New Roman", align="CENTER", first_indent_cm=0.0)
    if en_rule:
        en_spacing = _line_spacing_from_text(en_rule)
        en_spacing.update(_spacing_before_after_from_text(en_rule, float(body.get("size") or 12) * 1.5))
        if profiles.get("en_title"):
            profiles["en_title"] = _normalize_profile({**profiles["en_title"], **en_spacing}, profiles["en_title"])
        en_body = _normalize_profile({"font": "Times New Roman", "size": body.get("size", 12), "bold": False, "align": "JUSTIFY", "first_indent_cm": body.get("first_indent_cm"), **en_spacing}, body)
        profiles["en_abstract_body"] = en_body
        profiles["en_abstract_heading"] = _normalize_profile({"font": "Times New Roman", "size": profiles.get("en_title", en_body).get("size", 16), "bold": False, "align": "CENTER", "first_indent_cm": 0.0, **en_spacing}, en_body)
        profiles["en_keywords"] = _normalize_profile({"font": "Times New Roman", "size": body.get("size", 12), "bold": False, "align": "LEFT", "first_indent_cm": 0.0, **en_spacing}, en_body)
    en_kw_rule = _find_instruction(texts, "Key words") or _find_instruction(texts, "KEY WORD")
    if en_kw_rule:
        profiles["en_keywords"] = _profile_from_instruction(
            en_kw_rule,
            _normalize_profile({"font": "Times New Roman", "size": body.get("size", 12), "bold": False, "align": "LEFT", "first_indent_cm": 0.0}, body),
            allow_bold=False,
            font="Times New Roman",
            size=body.get("size", 12),
            bold=False,
            align="LEFT",
            first_indent_cm=0.0,
        )
        profiles["en_keywords_label"] = _profile_from_instruction(
            en_kw_rule,
            profiles["en_keywords"],
            font="Times New Roman",
            align="LEFT",
            first_indent_cm=0.0,
        )

    toc_title_rule = _find_instruction(texts, "目录", "黑体") or _find_instruction(texts, "【目录】")
    if toc_title_rule:
        profiles["toc_title"] = _profile_from_instruction(toc_title_rule, body, first_indent_cm=0.0)
    toc_rule = _find_regex_instruction(texts, r"中文：([^。；;\n]*一级标题[^。；;\n]*)") or _find_instruction(texts, "一级标题", "二级", "三级", "目录") or _find_instruction(texts, "一级标题", "宋体", "四号")
    if toc_rule:
        role("toc1", font=_font_from_text(toc_rule, "宋体"), size=14.0, bold=bool(re.search(r"一级标题[^。；;\n]*加粗", toc_rule)), align="LEFT", first_indent_cm=0.0)
        role("toc2", font=_font_from_text(toc_rule, "宋体"), size=12.0, bold=False, align="LEFT", first_indent_cm=0.0)
        role("toc3", font=_font_from_text(toc_rule, "宋体"), size=12.0, bold=False, align="LEFT", first_indent_cm=0.0)

    h_rules = [
        ("h1", _find_regex_instruction(texts, r"(第1章[^。；;\n]*一级标题[^。；;\n]*)") or _find_instruction(texts, "第1章", "标题") or _find_instruction(texts, "一级标题")),
        ("h2", _find_regex_instruction(texts, r"(1\.1[^。；;\n]*二级标题[^。；;\n]*)") or _find_instruction(texts, "二级标题")),
        ("h3", _find_regex_instruction(texts, r"(1\.1\.1[^。；;\n]*三级标题[^。；;\n]*)") or _find_instruction(texts, "三级标题")),
    ]
    for h_role, h_rule in h_rules:
        if h_rule and _has_format_instruction(h_rule):
            defaults = {
                "align": "CENTER" if h_role == "h1" else "LEFT",
                "first_indent_cm": 0.0 if h_role == "h1" else _indent_from_text(h_rule, _size_from_text(h_rule, body.get("size")), body.get("first_indent_cm")),
            }
            profiles[h_role] = _profile_from_instruction(h_rule, body, **defaults)

    fig_rule = _find_instruction(texts, "图标题") or _find_instruction(texts, "图题")
    tab_rule = _find_instruction(texts, "表标题") or _find_instruction(texts, "表题")
    table_detail_rule = _find_instruction(texts, "表内容") or _find_instruction(texts, "表格", "五号")
    if fig_rule and _has_format_instruction(fig_rule):
        profiles["figure_caption"] = _profile_from_instruction(fig_rule, body, font="宋体", size=10.5, bold=False, align="CENTER", first_indent_cm=0.0)
    if tab_rule and _has_format_instruction(tab_rule):
        profiles["table_caption"] = _profile_from_instruction(tab_rule, body, font="宋体", size=10.5, bold=False, align="CENTER", first_indent_cm=0.0)
    if table_detail_rule and _has_format_instruction(table_detail_rule):
        table_body = _profile_from_instruction(table_detail_rule, body, font="宋体", size=10.5, align="CENTER", first_indent_cm=0.0)
        table_body["font"] = "宋体"
        if "单倍" in table_detail_rule:
            table_body.update({"line_spacing_fixed_pt": None, "line_spacing_val": 1.0, "line_spacing_rule": "auto"})
        profiles["table_body"] = _normalize_profile(table_body, body)
        profiles["table_header"] = _normalize_profile({"bold": True}, profiles["table_body"])

    reference_sample = _reference_sample_profile(paras, body)
    if reference_sample:
        profiles["reference"] = reference_sample
    ref_rule = _find_reference_rule(texts)
    if ref_rule:
        reference = _profile_from_instruction(
            ref_rule,
            profiles.get("reference", body),
            allow_bold=False,
            font="宋体",
            size=12.0,
            bold=False,
            align="JUSTIFY",
            first_indent_cm=0.0,
            space_before_pt=0.0,
            space_after_pt=0.0,
        )
        reference["font"] = "宋体"
        profiles["reference"] = reference
    if _find_instruction(texts, "英文参考文献", "左对齐"):
        profiles["reference_english"] = _normalize_profile({"align": "LEFT"}, profiles.get("reference", body))
    profiles["reference_heading"] = profiles.get("reference_heading") or profiles["h1"]
    formula_rule = _find_instruction(texts, "公式应") or _find_instruction(texts, "公式", "居中")
    if formula_rule:
        profiles["formula"] = _profile_from_instruction(formula_rule, body, align="CENTER", first_indent_cm=0.0)
    return profiles
