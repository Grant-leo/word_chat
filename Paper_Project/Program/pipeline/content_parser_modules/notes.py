"""DOCX footnote/endnote extraction helpers."""
from __future__ import annotations

import zipfile
from typing import Any, Dict

from lxml import etree

try:
    from content_parser_modules.text_cleaner import clean_text_artifacts
except ImportError:  # pragma: no cover - package-style imports
    from .text_cleaner import clean_text_artifacts


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


def _local_name(element: Any) -> str:
    return element.tag.split("}")[-1] if "}" in element.tag else element.tag


def _node_text(node: Any) -> str:
    pieces = []
    for child in node.iter():
        name = _local_name(child)
        if name == "t":
            pieces.append(child.text or "")
        elif name == "tab":
            pieces.append("\t")
        elif name in {"br", "cr"}:
            pieces.append("\n")
    return "".join(pieces)


def _read_part(zf: zipfile.ZipFile, name: str) -> bytes | None:
    try:
        return zf.read(name)
    except KeyError:
        return None


def _extract_note_part(xml_bytes: bytes | None, note_tag: str) -> Dict[str, str]:
    if not xml_bytes:
        return {}
    try:
        root = etree.fromstring(xml_bytes)
    except Exception:
        return {}
    notes: Dict[str, str] = {}
    for note in root.xpath(f".//w:{note_tag}", namespaces=NS):
        note_type = str(note.get(f"{{{W_NS}}}type") or "")
        if note_type in {"separator", "continuationSeparator", "continuationNotice"}:
            continue
        note_id = str(note.get(f"{{{W_NS}}}id") or "").strip()
        if not note_id:
            continue
        paras = []
        for para in note.xpath(".//w:p", namespaces=NS):
            text = clean_text_artifacts(_node_text(para), preserve_newlines=True).strip()
            if text:
                paras.append(text)
        text = "\n".join(paras).strip()
        if text:
            notes[note_id] = text
    return notes


def extract_note_maps(docx_path: str) -> Dict[str, Dict[str, str]]:
    """Extract visible footnote/endnote bodies keyed by source note id."""
    try:
        with zipfile.ZipFile(docx_path) as zf:
            footnotes = _extract_note_part(_read_part(zf, "word/footnotes.xml"), "footnote")
            endnotes = _extract_note_part(_read_part(zf, "word/endnotes.xml"), "endnote")
    except Exception:
        return {"footnote": {}, "endnote": {}}
    return {"footnote": footnotes, "endnote": endnotes}


def count_note_runs(content: Dict[str, Any]) -> Dict[str, int]:
    counts = {"footnote": 0, "endnote": 0}

    def visit_item(item: Any) -> None:
        if not isinstance(item, dict):
            return
        for run in item.get("runs") or []:
            if not isinstance(run, dict) or run.get("type") != "note_ref":
                continue
            note_type = str(run.get("note_type") or "footnote")
            if note_type in counts:
                counts[note_type] += 1
        for cell in item.get("table_cell_items") or []:
            if not isinstance(cell, dict):
                continue
            for nested in cell.get("items") or []:
                visit_item(nested)

    for section in content.get("sections") or []:
        for item in section.get("paragraphs") or []:
            visit_item(item)
    return counts
