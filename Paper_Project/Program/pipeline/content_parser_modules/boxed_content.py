"""Fallback extraction for visible text inside DOCX boxed structures."""
from __future__ import annotations

import re
import zipfile
from typing import Any, Callable, Dict, Iterable, List

from lxml import etree

try:
    from content_parser_modules.placeholders import is_template_instruction_text, is_unfilled_placeholder_text
    from content_parser_modules.text_cleaner import clean_text_artifacts
except ImportError:  # pragma: no cover - package-style imports
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
    pieces: List[str] = []
    for node in paragraph.iter():
        name = _local_name(node)
        if name == "t":
            pieces.append(node.text or "")
        elif name == "tab":
            pieces.append("\t")
        elif name in {"br", "cr"}:
            pieces.append("\n")
    return "".join(pieces)


def _clean_visible_text(text: str) -> str:
    cleaned = clean_text_artifacts(text, preserve_newlines=True).strip()
    if not cleaned:
        return ""
    if is_unfilled_placeholder_text(cleaned) or is_template_instruction_text(cleaned):
        return ""
    return cleaned


def _normalize_for_dedupe(text: str) -> str:
    return re.sub(r"\s+", "", str(text or "")).casefold()


def recover_boxed_text_records(docx_path: str) -> List[Dict[str, str]]:
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

    records: List[Dict[str, str]] = []
    seen = set()

    def add(source: str, text: str) -> None:
        cleaned = _clean_visible_text(text)
        key = (source, _normalize_for_dedupe(cleaned))
        if not cleaned or not key[1] or key in seen:
            return
        seen.add(key)
        records.append({"source": source, "text": cleaned})

    for txbx in root.xpath(".//w:txbxContent", namespaces=NS):
        for para in txbx.xpath(".//w:p", namespaces=NS):
            add("textbox", _paragraph_text(para))

    for sdt_content in root.xpath(".//w:sdtContent", namespaces=NS):
        if _has_ancestor(sdt_content, "txbxContent"):
            continue
        for para in sdt_content.xpath(".//w:p", namespaces=NS):
            if _has_ancestor(para, "txbxContent"):
                continue
            add("content_control", _paragraph_text(para))

    return records


def _iter_existing_text(content: Dict[str, Any]) -> Iterable[str]:
    for value in (content.get("title_info") or {}).values():
        yield str(value or "")
    for value in (content.get("cover_info") or {}).values():
        yield str(value or "")
    for section in content.get("sections") or []:
        yield str(section.get("heading") or "")
        for item in section.get("paragraphs") or []:
            if isinstance(item, str):
                yield item
            elif isinstance(item, dict):
                yield str(item.get("text") or item.get("code") or "")
    for item in content.get("references") or []:
        if isinstance(item, str):
            yield item
        elif isinstance(item, dict):
            yield str(item.get("text") or item.get("code") or "")


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
    records: List[Dict[str, str]],
    append_text_or_code_func: Callable[[Dict[str, Any], str], None],
) -> Dict[str, int]:
    """Append non-duplicate recovered text to the body content stream."""
    existing = {_normalize_for_dedupe(text) for text in _iter_existing_text(content) if _normalize_for_dedupe(text)}
    counts = {"textbox": 0, "content_control": 0, "duplicates": 0}
    section: Dict[str, Any] | None = None
    for record in records or []:
        text = str(record.get("text") or "")
        key = _normalize_for_dedupe(text)
        source = str(record.get("source") or "")
        if not key or key in existing:
            counts["duplicates"] += 1
            continue
        if section is None:
            section = _target_section(content)
        append_text_or_code_func(section, text)
        existing.add(key)
        if source in counts:
            counts[source] += 1
    return counts
