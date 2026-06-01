"""Template/content requirement generation for strict conformance QA."""
from __future__ import annotations

import json
import os
from datetime import datetime
from typing import Any, Dict, Optional

try:
    from privacy import sanitize_value
except Exception:  # pragma: no cover
    def sanitize_value(value: Any, project_root: str | None = None) -> Any:
        return value

try:
    from script_generator import _infer_style_profiles
except Exception:  # pragma: no cover
    _infer_style_profiles = None

try:
    from qa_conformance_modules.ooxml import _profile_subset
    from qa_conformance_modules.registry import STYLE_ROLES
except ImportError:  # pragma: no cover - package-style imports
    from .ooxml import _profile_subset
    from .registry import STYLE_ROLES


def _write_json(path: str, value: Dict[str, Any]) -> None:
    with open(path, "w", encoding="utf-8") as f:
        json.dump(value, f, ensure_ascii=False, indent=2)


def _profile_from_format(fmt: Dict[str, Any]) -> Dict[str, Dict[str, Any]]:
    if _infer_style_profiles is not None:
        try:
            return _infer_style_profiles(fmt)
        except Exception:
            pass
    return fmt.get("style_profiles") or {}


def _content_counts(content: Dict[str, Any]) -> Dict[str, int]:
    inline_images = 0
    inline_names = []
    section_images = 0
    section_names = []
    tables = 0
    formulas = 0
    for sec in content.get("sections") or []:
        current_section_images = [str(image or "") for image in (sec.get("images") or []) if str(image or "")]
        section_images += len(current_section_images)
        section_names.extend(current_section_images)
        for item in sec.get("paragraphs") or []:
            if not isinstance(item, dict):
                continue
            item_image_names = []
            if item.get("role") in {"image", "figure"} or item.get("image") or item.get("filename"):
                item_image_names.append(str(item.get("image") or item.get("filename") or item.get("asset") or ""))
            for cell in item.get("table_cell_items") or []:
                if not isinstance(cell, dict):
                    continue
                for nested in cell.get("items") or []:
                    if isinstance(nested, dict) and (nested.get("role") in {"image", "figure"} or nested.get("image") or nested.get("filename")):
                        item_image_names.append(str(nested.get("image") or nested.get("filename") or nested.get("asset") or ""))
            if item_image_names:
                inline_images += len(item_image_names)
                inline_names.extend(name for name in item_image_names if name)
            if item.get("table_rows") and item.get("role") != "code":
                tables += 1
            math_items = item.get("math") or []
            if math_items:
                formulas += len(math_items)
            elif item.get("role") == "formula" or item.get("latex") or item.get("xml"):
                formulas += 1
    if inline_images:
        extra_section_only = [name for name in section_names if name and name not in inline_names]
        images = inline_images + len(extra_section_only)
    elif section_images:
        images = section_images
    else:
        images = int((content.get("_meta") or {}).get("images_extracted") or 0)
    if not images:
        images = 0
    if not tables:
        tables = int((content.get("_meta") or {}).get("tables_count") or 0)
    return {"images": images, "tables": tables, "formulas": formulas}


def build_requirements(fmt: Dict[str, Any], content: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """Build machine-checkable requirements from extracted template/content."""
    profiles = _profile_from_format(fmt)
    sections = fmt.get("sections") or []
    first_section = sections[0] if sections else {}
    page = {
        "page_width_cm": first_section.get("page_width_cm", 21.0),
        "page_height_cm": first_section.get("page_height_cm", 29.7),
        "margin_top_cm": first_section.get("margin_top_cm", 2.54),
        "margin_bottom_cm": first_section.get("margin_bottom_cm", 2.54),
        "margin_left_cm": first_section.get("margin_left_cm", 2.54),
        "margin_right_cm": first_section.get("margin_right_cm", 2.54),
    }
    header_footer = {
        "header_count": sum(len(s.get("header") or []) for s in sections),
        "footer_count": sum(len(s.get("footer") or []) for s in sections),
        "diff_first_page": any(bool(s.get("diff_first_page")) for s in sections),
    }
    counts = _content_counts(content or {})
    return {
        "schema_version": 1,
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "source": sanitize_value((fmt.get("_meta") or {}).get("source") or ""),
        "page": page,
        "header_footer": header_footer,
        "style_roles": {role: _profile_subset(profiles[role]) for role in STYLE_ROLES if isinstance(profiles.get(role), dict)},
        "expected_counts": counts,
        "content_contract": {
            "check_all_headings": True,
            "check_all_body_paragraphs": True,
            "check_all_captions": True,
            "check_all_references": True,
        },
        "specialty": {
            "tables_require_three_line_borders": True,
            "images_must_have_positive_extent": True,
            "images_must_fit_text_width": True,
            "formulas_must_be_native_omml": True,
            "omml_runs_require_rpr": True,
        },
    }


def requirements_to_markdown(req: Dict[str, Any]) -> str:
    lines = [
        "# Template Requirements",
        "",
        f"- Source: `{req.get('source') or ''}`",
        "",
        "## Page",
        "",
    ]
    for key, value in (req.get("page") or {}).items():
        lines.append(f"- `{key}`: {value}")
    lines.extend(["", "## Style Roles", ""])
    for role, profile in sorted((req.get("style_roles") or {}).items()):
        bits = ", ".join(f"{k}={v}" for k, v in profile.items())
        lines.append(f"- `{role}`: {bits}")
    lines.extend(["", "## Expected Counts", ""])
    for key, value in sorted((req.get("expected_counts") or {}).items()):
        lines.append(f"- `{key}`: {value}")
    lines.extend(["", "## Specialty Checks", ""])
    for key, value in sorted((req.get("specialty") or {}).items()):
        lines.append(f"- `{key}`: {value}")
    lines.append("")
    return "\n".join(lines)


def write_requirements(fmt: Dict[str, Any], content: Dict[str, Any], out_dir: str) -> Dict[str, Any]:
    req = build_requirements(fmt, content)
    _write_json(os.path.join(out_dir, "template_requirements.json"), req)
    with open(os.path.join(out_dir, "template_requirements.md"), "w", encoding="utf-8") as f:
        f.write(requirements_to_markdown(req))
    return req
