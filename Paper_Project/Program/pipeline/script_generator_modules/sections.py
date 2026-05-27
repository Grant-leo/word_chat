"""Section planning helpers for script_generator.py."""
from __future__ import annotations

import re
from typing import Any, Dict, List, Optional


def _ascii_ratio(text: str) -> float:
    text = str(text or "")
    if not text:
        return 0.0
    return sum(1 for c in text if c.isascii() and c.isalpha()) / max(len(text), 1)


def _section_role(sec: Dict[str, Any]) -> str:
    role = (sec.get("role") or "").strip()
    if role:
        return role
    h = (sec.get("heading") or "").strip()
    compact = re.sub(r"[\s：:]+", "", h).lower()
    if compact in ("摘要", "中文摘要"):
        return "cn_abstract"
    if h.startswith("关键词") or compact in ("关键词", "关键字"):
        return "cn_keywords"
    if compact == "abstract":
        return "en_abstract"
    if h.upper().replace(" ", "").startswith("KEYWORDS") or re.match(r"(?i)^key\s*words?", h):
        return "en_keywords"
    if h.startswith("参考文献") or re.match(r"(?i)^references?$", h):
        return "references"
    if re.match(r"(?i)^acknowledg(?:e)?ments?\b|^acknowledgment\b", h):
        return "acknowledgement"
    if re.search(r"致\s*谢", h):
        return "acknowledgement"
    if re.match(r"(?i)^append(?:ix|ices)\b", h):
        return "appendix"
    if re.search(r"附\s*录", h):
        return "appendix"
    return "body"


def _front_matter_sections(cnt: Dict[str, Any]) -> Dict[str, Any]:
    result: Dict[str, Any] = {
        "cn_abs": None,
        "cn_kw": None,
        "en_title": str((cnt.get("title_info") or {}).get("title_en") or "").strip(),
        "en_abs": None,
        "en_kw": None,
        "front_indices": set(),
    }
    sections = cnt.get("sections") or []
    body_started = False
    for idx, sec in enumerate(sections):
        role = _section_role(sec)
        h = (sec.get("heading") or "").strip()
        if role == "cn_abstract":
            result["cn_abs"] = sec
            result["front_indices"].add(idx)
        elif role == "cn_keywords":
            result["cn_kw"] = sec
            result["front_indices"].add(idx)
        elif role == "en_abstract":
            result["en_abs"] = sec
            result["front_indices"].add(idx)
        elif role == "en_keywords":
            result["en_kw"] = sec
            result["front_indices"].add(idx)
        elif (
            not body_started
            and sec.get("level") == 1
            and _ascii_ratio(h) > 0.55
            and not re.match(r"(?i)^chapter\s+\d+", h)
            and not re.match(r"^\d+(?:\.\d+)*\s+", h)
            and role not in ("references", "acknowledgement", "appendix")
        ):
            if not result["en_title"]:
                result["en_title"] = h
                result["front_indices"].add(idx)
        elif role not in ("cn_abstract", "cn_keywords", "en_abstract", "en_keywords"):
            body_started = True
    return result


def _heading_num_tuple(text: str) -> Optional[tuple]:
    m = re.match(r"^(\d+(?:\.\d+)*)\b", str(text or "").strip())
    if not m:
        return None
    try:
        return tuple(int(x) for x in m.group(1).split("."))
    except Exception:
        return None


def _chapter_num(text: str) -> Optional[int]:
    s = str(text or "").strip()
    m = re.match(r"^第(\d+)章", s)
    if m:
        return int(m.group(1))
    cn = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9, "十": 10}
    m = re.match(r"^第([一二三四五六七八九十]+)章", s)
    if m:
        t = m.group(1)
        if t == "十":
            return 10
        if t.startswith("十"):
            return 10 + cn.get(t[1:], 0)
        if t.endswith("十"):
            return cn.get(t[0], 1) * 10
        if "十" in t:
            a, b = t.split("十", 1)
            return cn.get(a, 1) * 10 + cn.get(b, 0)
        return cn.get(t)
    return None


def _sort_h3_inside_h2(block: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    if len(block) <= 2:
        return block
    head = block[0]
    chunks: List[List[Dict[str, Any]]] = []
    cur: List[Dict[str, Any]] = []
    for sec in block[1:]:
        num = _heading_num_tuple(sec.get("heading", ""))
        if sec.get("level") == 3 and num and len(num) == 3:
            if cur:
                chunks.append(cur)
            cur = [sec]
        else:
            if cur:
                cur.append(sec)
            else:
                chunks.append([sec])
    if cur:
        chunks.append(cur)
    nums = [_heading_num_tuple(c[0].get("heading", "")) for c in chunks]
    if len(chunks) >= 2 and all(nums) and nums != sorted(nums):
        chunks = sorted(chunks, key=lambda c: _heading_num_tuple(c[0].get("heading", "")))
    return [head] + [sec for c in chunks for sec in c]


def _normalize_numbered_section_order(sections: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Safely reorder numeric subsection blocks within each chapter.

    It fixes order drift such as 2.3 appearing after 2.9 while preserving every
    paragraph/table/image. It does not rename duplicate numbers, because that
    is a content judgment rather than a formatting operation.
    """
    out: List[Dict[str, Any]] = []
    i = 0
    n = len(sections or [])
    while i < n:
        sec = sections[i]
        ch = _chapter_num(sec.get("heading", "")) if sec.get("level") == 1 else None
        if ch is None:
            out.append(sec)
            i += 1
            continue
        out.append(sec)
        i += 1
        chapter_items: List[Dict[str, Any]] = []
        while i < n:
            nxt = sections[i]
            if nxt.get("level") == 1 and _chapter_num(nxt.get("heading", "")) is not None:
                break
            chapter_items.append(nxt)
            i += 1
        blocks: List[List[Dict[str, Any]]] = []
        cur: List[Dict[str, Any]] = []
        for item in chapter_items:
            num = _heading_num_tuple(item.get("heading", ""))
            if item.get("level") == 2 and num and len(num) == 2 and num[0] == ch:
                if cur:
                    blocks.append(cur)
                cur = [item]
            else:
                if cur:
                    cur.append(item)
                else:
                    blocks.append([item])
        if cur:
            blocks.append(cur)
        nums = [_heading_num_tuple(b[0].get("heading", "")) for b in blocks]
        if blocks and all(num and len(num) == 2 and num[0] == ch for num in nums) and nums != sorted(nums):
            blocks = sorted(blocks, key=lambda b: _heading_num_tuple(b[0].get("heading", "")))
        for b in blocks:
            out.extend(_sort_h3_inside_h2(b))
    return out
