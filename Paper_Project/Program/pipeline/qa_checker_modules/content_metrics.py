"""Content count and traversal helpers for structural QA."""
from __future__ import annotations

from typing import Any, Dict, Iterable, List

def _iter_paragraph_items(content: Dict[str, Any]) -> Iterable[Any]:
    for sec in content.get("sections") or []:
        for item in sec.get("paragraphs") or []:
            yield item


def _iter_child_items(item: Dict[str, Any]) -> Iterable[Any]:
    for nested in item.get("items") or []:
        yield nested
    for cell in item.get("table_cell_items") or []:
        if not isinstance(cell, dict):
            continue
        for nested in cell.get("items") or []:
            yield nested
    for run in item.get("runs") or []:
        if not isinstance(run, dict):
            continue
        yield run


def _iter_item_tree(item: Any) -> Iterable[Dict[str, Any]]:
    if not isinstance(item, dict):
        return
    yield item
    for child in _iter_child_items(item):
        yield from _iter_item_tree(child)


def _direct_formula_count(item: Dict[str, Any]) -> int:
    math_items = item.get("math") or []
    if math_items:
        return len(math_items)
    if item.get("role") == "formula" or item.get("latex") or item.get("xml"):
        return 1
    return 0


def _count_content_formulas(content: Dict[str, Any]) -> int:
    def count_item(item: Any, skip_direct: bool = False) -> int:
        if not isinstance(item, dict):
            return 0
        total = 0 if skip_direct else _direct_formula_count(item)
        parent_has_math = bool(item.get("math"))
        for nested in item.get("items") or []:
            total += count_item(nested)
        for cell in item.get("table_cell_items") or []:
            if not isinstance(cell, dict):
                continue
            for nested in cell.get("items") or []:
                total += count_item(nested)
        for run in item.get("runs") or []:
            if not isinstance(run, dict):
                continue
            total += count_item(run, skip_direct=parent_has_math)
        return total

    total = 0
    for item in _iter_paragraph_items(content):
        total += count_item(item)
    return total


def _count_content_tables(content: Dict[str, Any]) -> int:
    def count_item(item: Any) -> int:
        if not isinstance(item, dict):
            return 0
        total = 1 if item.get("table_rows") and item.get("role") != "code" else 0
        for nested in _iter_child_items(item):
            total += count_item(nested)
        return total

    total = 0
    saw_table_rows = False
    for item in _iter_paragraph_items(content):
        total += count_item(item)
        for node in _iter_item_tree(item):
            if node.get("table_rows"):
                saw_table_rows = True
    if saw_table_rows:
        return total
    return int((content.get("_meta") or {}).get("tables_count") or 0)


def _count_content_note_refs(content: Dict[str, Any]) -> Dict[str, int]:
    counts = {"footnote": 0, "endnote": 0}
    for item in _iter_paragraph_items(content):
        for node in _iter_item_tree(item):
            note_type = str(node.get("note_type") or "").strip().lower()
            if (node.get("type") == "note_ref" or node.get("role") == "note_ref") and note_type in counts:
                counts[note_type] += 1
    return counts


def _image_name_from_item(item: Dict[str, Any]) -> str:
    name = str(item.get("image") or item.get("filename") or item.get("asset") or "")
    if name and (item.get("role") in ("image", "figure") or item.get("image") or item.get("filename") or item.get("asset")):
        return name
    return ""


def _image_names_from_item(item: Dict[str, Any]) -> List[str]:
    names: List[str] = []
    name = _image_name_from_item(item)
    if name:
        names.append(name)
    for nested in _iter_child_items(item):
        if isinstance(nested, dict):
            names.extend(_image_names_from_item(nested))
    return names


def _count_content_images(content: Dict[str, Any]) -> int:
    inline_total = 0
    inline_names: List[str] = []
    section_total = 0
    section_names: List[str] = []
    for sec in content.get("sections") or []:
        section_images = [str(x or "") for x in (sec.get("images") or []) if str(x or "")]
        section_total += len(section_images)
        section_names.extend(section_images)
        for item in sec.get("paragraphs") or []:
            if not isinstance(item, dict):
                continue
            item_names = _image_names_from_item(item)
            if item_names:
                inline_total += len(item_names)
                inline_names.extend(item_names)
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
            for name in _image_names_from_item(item):
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
                for node in _iter_item_tree(item):
                    parts.append(str(node.get("text") or node.get("code") or ""))
                    for row in node.get("table_rows") or []:
                        parts.extend(str(cell or "") for cell in row)
            else:
                parts.append(str(item or ""))
    for ref in content.get("references") or []:
        if isinstance(ref, dict):
            parts.append(str(ref.get("text") or ref.get("code") or ""))
        else:
            parts.append(str(ref or ""))
    return sum(len(p.strip()) for p in parts if p and p.strip())
