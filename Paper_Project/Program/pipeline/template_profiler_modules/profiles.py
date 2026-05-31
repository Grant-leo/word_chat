"""Template capability profile construction."""
from __future__ import annotations

import os
import re
from collections import Counter
from datetime import datetime
from typing import Any, Dict, Iterable, List

STYLE_ROLES = [
    "cn_title", "en_title", "cn_abstract_heading", "cn_abstract_body",
    "cn_keywords", "en_abstract_heading", "en_abstract_body", "en_keywords",
    "toc_title", "toc1", "toc2", "toc3", "body", "h1", "h2", "h3",
    "figure_caption", "table_caption", "table_body", "table_header",
    "formula", "reference",
]


def _texts(fmt: Dict[str, Any]) -> List[str]:
    return [str(p.get("text") or "").strip() for p in fmt.get("paragraphs") or []]


def _text_blob(fmt: Dict[str, Any]) -> str:
    return "\n".join(t for t in _texts(fmt) if t)


def _walk(value: Any) -> Iterable[Any]:
    if isinstance(value, dict):
        yield value
        for child in value.values():
            yield from _walk(child)
    elif isinstance(value, list):
        for child in value:
            yield from _walk(child)


def _profile_style(style: Dict[str, Any]) -> Dict[str, Any]:
    keys = [
        "font", "size", "bold", "italic", "align",
        "line_spacing_val", "line_spacing_fixed_pt",
        "space_before_pt", "space_after_pt", "first_indent_cm",
    ]
    return {k: style.get(k) for k in keys if k in style and style.get(k) is not None}


def _first_section_page(fmt: Dict[str, Any]) -> Dict[str, Any]:
    section = (fmt.get("sections") or [{}])[0] or {}
    keys = [
        "page_width_cm", "page_height_cm",
        "margin_top_cm", "margin_bottom_cm",
        "margin_left_cm", "margin_right_cm",
        "diff_first_page",
    ]
    return {k: section.get(k) for k in keys if k in section}


def _is_landscape_page(page: Dict[str, Any]) -> bool:
    try:
        width = float(page.get("page_width_cm") or 0)
        height = float(page.get("page_height_cm") or 0)
    except (TypeError, ValueError):
        return False
    return width > 0 and height > 0 and width > height


def _header_footer_profile(fmt: Dict[str, Any]) -> Dict[str, Any]:
    sections = fmt.get("sections") or []
    header_count = 0
    footer_count = 0
    diff_first_page = False
    for sec in sections:
        header_count += len(sec.get("header") or [])
        footer_count += len(sec.get("footer") or [])
        diff_first_page = diff_first_page or bool(sec.get("diff_first_page"))
    return {
        "has_header": header_count > 0,
        "has_footer": footer_count > 0,
        "header_count": header_count,
        "footer_count": footer_count,
        "diff_first_page": diff_first_page,
    }


def _cover_profile(fmt: Dict[str, Any]) -> Dict[str, Any]:
    cover = fmt.get("cover") or []
    type_counts = Counter(str(el.get("type") or "unknown") for el in cover if isinstance(el, dict))
    role_counts = Counter(str(el.get("role") or "unknown") for el in cover if isinstance(el, dict))
    image_count = sum(1 for el in cover if isinstance(el, dict) and el.get("type") == "image")
    table_count = sum(1 for el in cover if isinstance(el, dict) and el.get("type") == "table")
    para_count = sum(1 for el in cover if isinstance(el, dict) and el.get("type") in {"para", "empty"})
    if table_count and (image_count or para_count):
        layout_type = "mixed"
    elif table_count:
        layout_type = "table"
    elif image_count or para_count:
        layout_type = "paragraphs"
    else:
        layout_type = "none"
    return {
        "has_cover": bool(cover),
        "layout_type": layout_type,
        "element_count": len(cover),
        "image_count": image_count,
        "table_count": table_count,
        "type_counts": dict(type_counts),
        "role_counts": dict(role_counts),
    }


def _toc_profile(fmt: Dict[str, Any], text_blob: str) -> Dict[str, Any]:
    styles = fmt.get("style_profiles") or {}
    has_toc_styles = any(k in styles for k in ("toc_title", "toc1", "toc2", "toc3"))
    has_toc_text = bool(re.search(r"(鐩甛s*褰晐contents)", text_blob, re.I))
    return {
        "has_toc_style": has_toc_styles,
        "has_toc_text": has_toc_text,
        "toc_levels": [k for k in ("toc1", "toc2", "toc3") if k in styles],
    }


def _risk_flags(fmt: Dict[str, Any], text_blob: str) -> Dict[str, Any]:
    cover = fmt.get("cover") or []
    all_nodes = list(_walk(fmt))
    floating = any("anchor" in str(node).lower() for node in all_nodes)
    textbox = any("textbox" in str(node).lower() or "txbx" in str(node).lower() for node in all_nodes)
    pdf_meta = (fmt.get("_meta") or {}).get("pdf_template") or {}
    pdf_warnings = [str(item) for item in (pdf_meta.get("warnings") or [])]
    pdf_dependency_missing = any(item.startswith(("PDFINFO_MISSING", "PDFTOTEXT_MISSING")) for item in pdf_warnings)
    pdf_protected = any(item.startswith("PDF_TEMPLATE_PROTECTED") for item in pdf_warnings)
    pdf_read_failed = any(item.startswith(("PDFINFO_FAILED", "PDFTOTEXT_FAILED")) for item in pdf_warnings)
    pdf_instruction_incomplete = any(item.startswith("PDF_TEMPLATE_INSTRUCTION_INCOMPLETE") for item in pdf_warnings)
    pdf_visual_approximation = any(item.startswith("PDF_TEMPLATE_VISUAL_APPROXIMATION") for item in pdf_warnings)
    pdf_unsupported = bool(pdf_meta.get("errors")) or pdf_meta.get("type") == "scanned_or_unsupported_pdf"
    pdf_unsupported = bool(pdf_unsupported and not pdf_dependency_missing and not pdf_read_failed and not pdf_protected)
    page = _first_section_page(fmt)
    return {
        "complex_cover": len(cover) > 20,
        "uses_textbox": textbox,
        "uses_floating_images": floating,
        "many_sections": len(fmt.get("sections") or []) > 3,
        "many_template_tables": len(fmt.get("tables") or []) > 8,
        "pdf_template": bool(pdf_meta),
        "pdf_template_limited_confidence": bool(pdf_meta.get("warnings")) or pdf_meta.get("type") == "visual_sample_pdf",
        "pdf_template_instruction_incomplete": bool(pdf_instruction_incomplete),
        "pdf_template_visual_approximation": bool(pdf_visual_approximation),
        "pdf_template_landscape_page": bool(pdf_meta and _is_landscape_page(page)),
        "pdf_template_dependency_missing": bool(pdf_dependency_missing),
        "pdf_template_protected": bool(pdf_protected),
        "pdf_template_read_failed": bool(
            pdf_read_failed
            and not pdf_protected
            and (pdf_meta.get("errors") or int(pdf_meta.get("text_chars") or 0) == 0)
        ),
        "pdf_template_unsupported": pdf_unsupported,
        "mentions_formula_rules": bool(re.search(r"(公式|formula)", text_blob, re.I)),
        "mentions_reference_rules": bool(re.search(r"(参考文献|references?)", text_blob, re.I)),
    }


def _safe_source_meta(meta: Dict[str, Any]) -> Dict[str, Any]:
    source = str(meta.get("source") or "")
    safe = {
        "sha256": meta.get("sha256") or meta.get("source_hash"),
        "source_ext": os.path.splitext(source)[1].lower() if source else "",
        "paragraphs": meta.get("paragraphs"),
        "tables": meta.get("tables"),
        "sections": meta.get("sections"),
        "has_assets_dir": bool(meta.get("assets_dir")),
    }
    return {k: v for k, v in safe.items() if v not in (None, "")}


def profile_format(fmt: Dict[str, Any], project_root: str | None = None) -> Dict[str, Any]:
    text_blob = _text_blob(fmt)
    styles = fmt.get("style_profiles") or {}
    meta = fmt.get("_meta") or {}
    pdf_meta = meta.get("pdf_template") or {}
    style_roles = {
        role: _profile_style(styles[role])
        for role in STYLE_ROLES
        if isinstance(styles.get(role), dict)
    }

    return {
        "schema_version": 1,
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "source": _safe_source_meta(meta),
        "pdf_template": {
            "type": pdf_meta.get("type"),
            "confidence": pdf_meta.get("confidence"),
            "page_count": pdf_meta.get("page_count"),
            "text_chars": pdf_meta.get("text_chars"),
            "warnings": list(pdf_meta.get("warnings") or []),
            "errors": list(pdf_meta.get("errors") or []),
            "backend": dict(pdf_meta.get("backend") or {}),
        } if pdf_meta else {},
        "counts": {
            "paragraphs": len(fmt.get("paragraphs") or []),
            "tables": len(fmt.get("tables") or []),
            "sections": len(fmt.get("sections") or []),
            "cover_elements": len(fmt.get("cover") or []),
            "style_roles": len(style_roles),
        },
        "page": _first_section_page(fmt),
        "header_footer": _header_footer_profile(fmt),
        "cover": _cover_profile(fmt),
        "toc": _toc_profile(fmt, text_blob),
        "style_roles": style_roles,
        "capabilities": {
            "has_cover": bool(fmt.get("cover")),
            "has_body_style": "body" in style_roles,
            "has_heading_styles": all(k in style_roles for k in ("h1", "h2", "h3")),
            "has_caption_styles": any(k in style_roles for k in ("figure_caption", "table_caption")),
            "has_reference_style": "reference" in style_roles,
            "has_formula_style": "formula" in style_roles,
            "has_page_geometry": bool(_first_section_page(fmt)),
        },
        "risk_flags": _risk_flags(fmt, text_blob),
    }
