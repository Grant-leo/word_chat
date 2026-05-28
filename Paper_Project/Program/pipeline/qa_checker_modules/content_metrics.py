"""Content count and traversal helpers for structural QA."""
from __future__ import annotations

from typing import Any, Dict, Iterable, List

def _iter_paragraph_items(content: Dict[str, Any]) -> Iterable[Any]:
    for sec in content.get("sections") or []:
        for item in sec.get("paragraphs") or []:
            yield item


def _count_content_formulas(content: Dict[str, Any]) -> int:
    total = 0
    for item in _iter_paragraph_items(content):
        if not isinstance(item, dict):
            continue
        math_items = item.get("math") or []
        if math_items:
            total += len(math_items)
        elif item.get("role") == "formula" or item.get("latex"):
            total += 1
    return total


def _count_content_tables(content: Dict[str, Any]) -> int:
    total = 0
    saw_table_rows = False
    for item in _iter_paragraph_items(content):
        if isinstance(item, dict) and item.get("table_rows"):
            saw_table_rows = True
            if item.get("role") != "code":
                total += 1
    if saw_table_rows:
        return total
    return int((content.get("_meta") or {}).get("tables_count") or 0)


def _count_content_images(content: Dict[str, Any]) -> int:
    inline_total = 0
    inline_names: List[str] = []
    section_total = 0
    section_names: List[str] = []
    for sec in content.get("sections") or []:
        section_images = [str(x or "") for x in (sec.get("images") or [])]
        section_total += len(section_images)
        section_names.extend(section_images)
        for item in sec.get("paragraphs") or []:
            if not isinstance(item, dict):
                continue
            if item.get("role") in ("image", "figure") and (item.get("image") or item.get("filename") or item.get("asset")):
                inline_total += 1
                inline_names.append(str(item.get("image") or item.get("filename") or item.get("asset") or ""))
    if inline_total:
        extra_section_only = [name for name in section_names if name and name not in inline_names]
        return inline_total + len(extra_section_only)
    if section_total:
        return section_total
    return int((content.get("_meta") or {}).get("images_extracted") or 0)

def _iter_content_image_refs(content: Dict[str, Any]) -> Iterable[Dict[str, str]]:
    seen: set[str] = set()
    for sec in content.get("sections") or []:
        heading = str(sec.get("heading") or "")
        for item in sec.get("paragraphs") or []:
            if not isinstance(item, dict):
                continue
            if item.get("role") in ("image", "figure") and (item.get("image") or item.get("filename") or item.get("asset")):
                name = str(item.get("image") or item.get("filename") or item.get("asset") or "")
                if name and name not in seen:
                    seen.add(name)
                    yield {"name": name, "heading": heading, "caption": str(item.get("caption") or "")}
        for name in sec.get("images") or []:
            name = str(name or "")
            if name and name not in seen:
                seen.add(name)
                yield {"name": name, "heading": heading, "caption": ""}

def _content_text_chars(content: Dict[str, Any]) -> int:
    parts: List[str] = []
    title_info = content.get("title_info") or {}
    parts.extend(str(v or "") for v in title_info.values())
    for sec in content.get("sections") or []:
        parts.append(str(sec.get("heading") or ""))
        for item in sec.get("paragraphs") or []:
            if isinstance(item, dict):
                parts.append(str(item.get("text") or item.get("code") or ""))
                for row in item.get("table_rows") or []:
                    parts.extend(str(cell or "") for cell in row)
            else:
                parts.append(str(item or ""))
    for ref in content.get("references") or []:
        if isinstance(ref, dict):
            parts.append(str(ref.get("text") or ref.get("code") or ""))
        else:
            parts.append(str(ref or ""))
    return sum(len(p.strip()) for p in parts if p and p.strip())
