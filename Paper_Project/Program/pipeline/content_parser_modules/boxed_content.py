"""Fallback extraction for visible text inside DOCX boxed structures."""
from __future__ import annotations

import re
import zipfile
from typing import Any, Callable, Dict, Iterable, List

from lxml import etree

try:
    from content_parser_modules.paragraph_stream import visible_text_from_ooxml
    from content_parser_modules.placeholders import is_template_instruction_text, is_unfilled_placeholder_text
    from content_parser_modules.text_cleaner import clean_text_artifacts
except ImportError:  # pragma: no cover - package-style imports
    from .paragraph_stream import visible_text_from_ooxml
    from .placeholders import is_template_instruction_text, is_unfilled_placeholder_text
    from .text_cleaner import clean_text_artifacts


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


def _local_name(element: Any) -> str:
    return element.tag.split("}")[-1] if "}" in element.tag else element.tag


def _has_ancestor(element: Any, local_name: str) -> bool:
    parent = element.getparent()
    while parent is not None:
        if _local_name(parent) == local_name:
            return True
        parent = parent.getparent()
    return False


def _paragraph_text(paragraph: Any) -> str:
    return visible_text_from_ooxml(paragraph)


def _clean_visible_text(text: str) -> str:
    cleaned = clean_text_artifacts(text, preserve_newlines=True).strip()
    if not cleaned:
        return ""
    if is_unfilled_placeholder_text(cleaned) or is_template_instruction_text(cleaned):
        return ""
    return cleaned


def _normalize_for_dedupe(text: str) -> str:
    return re.sub(r"\s+", "", str(text or "")).casefold()


def recover_boxed_text_records(docx_path: str) -> List[Dict[str, Any]]:
    """Return visible text from textboxes and content controls in document order.

    These structures are often skipped by python-docx's paragraph iterator.  The
    fallback intentionally returns plain text only; original floating position
    remains a visual-review concern.
    """
    try:
        with zipfile.ZipFile(docx_path) as zf:
            xml = zf.read("word/document.xml")
    except Exception:
        return []
    try:
        root = etree.fromstring(xml)
    except Exception:
        return []

    records: List[Dict[str, Any]] = []
    seen = set()

    def add(source: str, text: str, *, in_table_cell: bool = False) -> None:
        cleaned = _clean_visible_text(text)
        key = (source, _normalize_for_dedupe(cleaned))
        if not cleaned or not key[1] or key in seen:
            return
        seen.add(key)
        record: Dict[str, Any] = {"source": source, "text": cleaned}
        if in_table_cell:
            record["context"] = "table_cell"
        records.append(record)

    for txbx in root.xpath(".//w:txbxContent", namespaces=NS):
        for para in txbx.xpath(".//w:p", namespaces=NS):
            add("textbox", _paragraph_text(para))

    for sdt_content in root.xpath(".//w:sdtContent", namespaces=NS):
        if _has_ancestor(sdt_content, "txbxContent"):
            continue
        for para in sdt_content.xpath(".//w:p", namespaces=NS):
            if _has_ancestor(para, "txbxContent"):
                continue
            add(
                "content_control",
                _paragraph_text(para),
                in_table_cell=_has_ancestor(sdt_content, "tc") or _has_ancestor(para, "tc"),
            )

    return records


def _iter_table_text(item: Any) -> Iterable[str]:
    if not isinstance(item, dict):
        return
    for row in item.get("table_rows") or []:
        if not isinstance(row, list):
            continue
        for cell in row:
            yield str(cell or "")
    for cell_entry in item.get("table_cell_items") or []:
        if not isinstance(cell_entry, dict):
            continue
        for cell_item in cell_entry.get("items") or []:
            yield from _iter_table_text(cell_item)


def _visible_item_text(item: Any) -> str:
    if isinstance(item, str):
        return item
    if not isinstance(item, dict):
        return ""
    runs = item.get("runs") or []
    if runs:
        return "".join(
            str(run.get("text") or "")
            for run in runs
            if run.get("type") in {"text", "math"}
        )
    return str(item.get("text") or item.get("code") or "")


def _iter_existing_text(content: Dict[str, Any]) -> Iterable[str]:
    def iter_item_text(item: Any) -> Iterable[str]:
        text = _visible_item_text(item)
        if text:
            yield text

    for value in (content.get("title_info") or {}).values():
        yield str(value or "")
    for value in (content.get("cover_info") or {}).values():
        yield str(value or "")
    for section in content.get("sections") or []:
        yield str(section.get("heading") or "")
        for item in section.get("paragraphs") or []:
            yield from iter_item_text(item)
    for item in content.get("references") or []:
        yield from iter_item_text(item)


def _iter_existing_body_text_containers(content: Dict[str, Any]) -> Iterable[str]:
    for section in content.get("sections") or []:
        pieces: List[str] = []
        for item in section.get("paragraphs") or []:
            if isinstance(item, dict) and item.get("role") == "table":
                continue
            text = _visible_item_text(item)
            if text:
                pieces.append(text)
        joined = "".join(pieces)
        if joined:
            yield joined


def _iter_existing_table_text(content: Dict[str, Any]) -> Iterable[str]:
    for section in content.get("sections") or []:
        for item in section.get("paragraphs") or []:
            yield from _iter_table_text(item)


def _target_section(content: Dict[str, Any]) -> Dict[str, Any]:
    sections = content.setdefault("sections", [])
    for section in reversed(sections):
        if section.get("role") not in {"cn_abstract", "cn_keywords", "en_abstract", "en_keywords", "references"}:
            return section
    section = {"heading": "正文", "level": 1, "role": "body", "paragraphs": [], "images": []}
    sections.append(section)
    return section


def append_recovered_boxed_text(
    content: Dict[str, Any],
    records: List[Dict[str, Any]],
    append_text_or_code_func: Callable[[Dict[str, Any], str], None],
) -> Dict[str, int]:
    """Append non-duplicate recovered text to the body content stream."""
    existing = {_normalize_for_dedupe(text) for text in _iter_existing_text(content) if _normalize_for_dedupe(text)}
    existing_table_containers = {
        _normalize_for_dedupe(text)
        for text in _iter_existing_table_text(content)
        if _normalize_for_dedupe(text)
    }
    existing_body_containers = {
        _normalize_for_dedupe(text)
        for text in _iter_existing_body_text_containers(content)
        if _normalize_for_dedupe(text)
    }
    counts = {"textbox": 0, "content_control": 0, "duplicates": 0}
    section: Dict[str, Any] | None = None
    for record in records or []:
        text = str(record.get("text") or "")
        key = _normalize_for_dedupe(text)
        source = str(record.get("source") or "")
        is_table_cell_content_control = source == "content_control" and record.get("context") == "table_cell"
        is_body_content_control_duplicate = (
            source == "content_control"
            and len(key) >= 16
            and any(key in value for value in existing_body_containers)
        )
        if not key or key in existing or (
            is_table_cell_content_control and any(key in value for value in existing_table_containers)
        ) or is_body_content_control_duplicate:
            counts["duplicates"] += 1
            continue
        if section is None:
            section = _target_section(content)
        append_text_or_code_func(section, text)
        existing.add(key)
        if source in counts:
            counts[source] += 1
    return counts
