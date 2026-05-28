"""Content paragraph expectations and style checks for strict conformance QA."""
from __future__ import annotations

import re
from typing import Any, Dict, List, Optional
from xml.etree import ElementTree as ET

try:
    from qa_conformance_modules.ooxml import (
        _cm_indent_to_twips,
        _compact,
        _expected_align,
        _expected_line_twips,
        _first_cjk_run_props,
        _first_run_props,
        _is_cjk_font,
        _is_cjk_text,
        _para_props,
        _pt_to_twips,
        _text_of_para,
    )
    from qa_conformance_modules.registry import BACKMATTER_ROLES, FRONTMATTER_ROLES
except ImportError:  # pragma: no cover - package-style imports
    from .ooxml import (
        _cm_indent_to_twips,
        _compact,
        _expected_align,
        _expected_line_twips,
        _first_cjk_run_props,
        _first_run_props,
        _is_cjk_font,
        _is_cjk_text,
        _para_props,
        _pt_to_twips,
        _text_of_para,
    )
    from .registry import BACKMATTER_ROLES, FRONTMATTER_ROLES


def _style_issues(role: str, text: str, p: ET.Element, profile: Dict[str, Any]) -> List[str]:
    issues: List[str] = []
    snippet = re.sub(r"\s+", " ", str(text or "")).strip()
    label = role if not snippet else f"{role} `{snippet[:50]}`"
    run = _first_run_props(p)
    para = _para_props(p)
    if profile.get("size") is not None and run.get("size") is not None:
        if abs(float(run["size"]) - float(profile["size"])) > 0.05:
            issues.append(f"{label}: size {run['size']} != {profile['size']}")
    for key in ("bold", "italic"):
        if key in profile and bool(run.get(key)) != bool(profile.get(key)):
            issues.append(f"{label}: {key} {run.get(key)} != {bool(profile.get(key))}")
    expected = _expected_align(profile.get("align"))
    if expected and para.get("align") != expected:
        issues.append(f"{label}: align {para.get('align')} != {expected}")
    font = profile.get("font")
    if font:
        if _is_cjk_font(str(font)):
            font_run = _first_cjk_run_props(p)
            if _is_cjk_text(text) and font_run.get("east_asia_font") != font:
                issues.append(f"{label}: eastAsia font {font_run.get('east_asia_font')} != {font}")
        elif run.get("ascii_font") != font:
            issues.append(f"{label}: ascii font {run.get('ascii_font')} != {font}")
    expected_line = _expected_line_twips(profile)
    if expected_line is not None and para.get("line") is not None and abs(int(para["line"]) - expected_line) > 1:
        issues.append(f"{label}: line {para.get('line')} != {expected_line}")
    if "space_before_pt" in profile and abs(int(para["space_before"]) - _pt_to_twips(profile.get("space_before_pt"))) > 1:
        issues.append(f"{label}: space_before {para.get('space_before')} != {_pt_to_twips(profile.get('space_before_pt'))}")
    if "space_after_pt" in profile and abs(int(para["space_after"]) - _pt_to_twips(profile.get("space_after_pt"))) > 1:
        issues.append(f"{label}: space_after {para.get('space_after')} != {_pt_to_twips(profile.get('space_after_pt'))}")
    if "first_indent_cm" in profile and abs(int(para["first_line"]) - _cm_indent_to_twips(profile.get("first_indent_cm"))) > 3:
        issues.append(f"{label}: first_line {para.get('first_line')} != {_cm_indent_to_twips(profile.get('first_indent_cm'))}")
    if "left_indent_cm" in profile and profile.get("left_indent_cm") is not None and abs(int(para["left"]) - _cm_indent_to_twips(profile.get("left_indent_cm"))) > 3:
        issues.append(f"{label}: left {para.get('left')} != {_cm_indent_to_twips(profile.get('left_indent_cm'))}")
    if "hanging_indent_cm" in profile and profile.get("hanging_indent_cm") is not None and abs(int(para["hanging"]) - _cm_indent_to_twips(profile.get("hanging_indent_cm"))) > 3:
        issues.append(f"{label}: hanging {para.get('hanging')} != {_cm_indent_to_twips(profile.get('hanging_indent_cm'))}")
    return issues


def _find_para_by_text(paragraphs: List[ET.Element], text: str) -> Optional[ET.Element]:
    target = _compact(text)
    if not target:
        return None
    for p in paragraphs:
        if _compact(_text_of_para(p)) == target:
            return p
    sample = target[:80]
    for p in paragraphs:
        actual = _compact(_text_of_para(p))
        if sample and sample in actual:
            return p
    return None


def _find_body_start_index(paragraphs: List[ET.Element], expected: List[Dict[str, str]]) -> int:
    """Skip cover/front matter/TOC paragraphs before style checks.

    Heading text appears both in the generated static TOC and in the actual
    body.  The earlier implementation used a hard-coded "1 Introduction"
    marker, which made non-English or differently titled papers compare TOC
    lines against body heading profiles.  Prefer the last occurrence of the
    first expected heading, because the body occurrence follows the TOC.
    """
    heading_texts = [
        str(item.get("text") or "").strip()
        for item in expected
        if str(item.get("role") or "").startswith("h") and str(item.get("text") or "").strip()
    ]
    for heading in heading_texts[:5]:
        target = _compact(heading)
        matches = [idx for idx, para in enumerate(paragraphs) if _compact(_text_of_para(para)) == target]
        if matches:
            return matches[-1]
    for idx, para in enumerate(paragraphs):
        if _compact(_text_of_para(para)) in {"目录", "contents"}:
            return min(idx + 1, len(paragraphs))
    return 0


def _is_reference_heading(text: str) -> bool:
    return bool(re.match(r"(?i)^(references?|参考文献)$", str(text or "").strip()))


def _is_backmatter(role: str, heading: str) -> bool:
    if role in BACKMATTER_ROLES:
        return True
    return bool(re.search(r"致\s*谢|附\s*录|^appendix\b|^acknowledgements?$", str(heading or ""), re.I))


def _is_caption_text(text: str) -> bool:
    return bool(re.match(r"^(Fig\.|Figure|Table|图|表)\s*\d+", str(text or "").strip(), re.I))


def _is_table_item(item: Any) -> bool:
    return isinstance(item, dict) and bool(item.get("table_rows")) and item.get("role") != "code"


def _is_formula_item(item: Any) -> bool:
    return isinstance(item, dict) and (item.get("role") == "formula" or item.get("latex") or item.get("xml") or item.get("math"))


def _expected_paragraphs(content: Dict[str, Any]) -> List[Dict[str, str]]:
    expected: List[Dict[str, str]] = []
    for sec in content.get("sections") or []:
        heading = str(sec.get("heading") or "").strip()
        role = str(sec.get("role") or "")
        if role in FRONTMATTER_ROLES:
            continue
        if _is_backmatter(role, heading):
            continue
        level = int(sec.get("level") or 1)
        if heading and heading != "正文" and not _is_backmatter(role, heading) and not _is_caption_text(heading):
            expected.append({"role": f"h{max(1, min(level, 3))}", "text": heading})
        for item in sec.get("paragraphs") or []:
            if _is_table_item(item) or _is_formula_item(item):
                continue
            if isinstance(item, dict):
                item_role = str(item.get("role") or "")
                text = str(item.get("text") or item.get("caption") or "").strip()
                if not text or item_role in {"image", "figure", "code"}:
                    continue
                if item_role in {"figure_caption", "table_caption"}:
                    expected.append({"role": item_role, "text": text})
                elif _is_caption_text(text):
                    expected.append({"role": "figure_caption" if re.match(r"^(Fig\.|Figure|图)", text, re.I) else "table_caption", "text": text})
                else:
                    expected.append({"role": "body", "text": text})
            else:
                text = str(item or "").strip()
                if text:
                    expected.append({"role": "body", "text": text})
    for ref in content.get("references") or []:
        text = str(ref.get("text") if isinstance(ref, dict) else ref or "").strip()
        if text and not _is_reference_heading(text):
            expected.append({"role": "reference", "text": text})
    return expected

