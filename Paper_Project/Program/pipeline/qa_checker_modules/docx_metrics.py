"""DOCX XML text and heading sample helpers for structural QA."""
from __future__ import annotations

import re
import zipfile
from typing import Any, Dict, List

def _read_docx_xml(docx_path: str) -> str:
    with zipfile.ZipFile(docx_path) as zf:
        return zf.read("word/document.xml").decode("utf-8", errors="replace")


def _xml_plain_text(xml: str) -> str:
    texts = re.findall(r"<w:t[^>]*>(.*?)</w:t>", xml, flags=re.S)
    return "".join(re.sub(r"<[^>]+>", "", t) for t in texts)


def _xml_paragraph_texts(xml: str) -> List[str]:
    paragraphs = re.findall(r"<w:p\b[^>]*>.*?</w:p>", xml, flags=re.S)
    return [_xml_plain_text(p).strip() for p in paragraphs]


def _front_matter_heading_role(text: str) -> str:
    compact = re.sub(r"[\s\u3000:：]+", "", str(text or "")).upper()
    if compact in {"摘要", "中文摘要"}:
        return "cn_abstract"
    if compact == "ABSTRACT":
        return "en_abstract"
    if compact in {"关键词", "关键字"}:
        return "cn_keywords"
    if compact in {"KEYWORDS", "KEYWORD"}:
        return "en_keywords"
    return ""


def _content_front_roles(content: Dict[str, Any]) -> set[str]:
    roles: set[str] = set()
    for sec in content.get("sections") or []:
        role = str(sec.get("role") or "")
        if role in {"cn_abstract", "cn_keywords", "en_abstract", "en_keywords"}:
            roles.add(role)
            continue
        heading_role = _front_matter_heading_role(str(sec.get("heading") or ""))
        if heading_role:
            roles.add(heading_role)
    return roles


def _duplicate_front_matter_headings(content: Dict[str, Any], xml: str) -> List[str]:
    expected_roles = _content_front_roles(content)
    if not expected_roles:
        return []
    role_labels = {
        "cn_abstract": "摘要",
        "en_abstract": "ABSTRACT",
        "cn_keywords": "关键词",
        "en_keywords": "KEY WORDS",
    }
    counts: Dict[str, int] = {}
    samples: Dict[str, List[int]] = {}
    for idx, text in enumerate(_xml_paragraph_texts(xml), 1):
        role = _front_matter_heading_role(text)
        if not role or role not in expected_roles:
            continue
        counts[role] = counts.get(role, 0) + 1
        samples.setdefault(role, []).append(idx)
    return [
        f"{role_labels.get(role, role)} x{count} at paragraphs {','.join(map(str, samples.get(role, [])[:6]))}"
        for role, count in sorted(counts.items())
        if count > 1
    ]

def _compact_text(text: str) -> str:
    return re.sub(r"\s+", "", str(text or "")).lower()


def _missing_heading_samples(content: Dict[str, Any], plain_text: str, limit: int = 6) -> List[str]:
    compact = _compact_text(plain_text)
    missing: List[str] = []
    front_roles = {"cn_abstract", "cn_keywords", "en_abstract", "en_keywords"}
    for sec in content.get("sections") or []:
        if sec.get("role") in front_roles:
            continue
        heading = str(sec.get("heading") or "").strip()
        if not heading or heading == "正文" or len(heading) > 80:
            continue
        sample = _compact_text(heading)
        if sample and sample not in compact:
            missing.append(heading)
        if len(missing) >= limit:
            break
    return missing
