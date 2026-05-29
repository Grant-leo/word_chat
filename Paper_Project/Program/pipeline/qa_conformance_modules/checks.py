"""Strict DOCX conformance check orchestration."""
from __future__ import annotations

import json
import os
import re
import zipfile
from typing import Any, Dict, List, Tuple
from xml.etree import ElementTree as ET

try:
    from qa_conformance_modules.content_checks import (
        _expected_paragraphs,
        _find_body_start_index,
        _find_para_by_text,
        _style_issues,
    )
    from qa_conformance_modules.docx_checks import (
        _content_tables,
        _find_table,
        _formula_issues,
        _image_issues,
        _page_geometry_issues,
        _table_border_issues,
    )
    from qa_conformance_modules.ooxml import W, _text_of_para
    from qa_conformance_modules.registry import TEXT_ONLY_ROLES, VALID_MODES
    from qa_conformance_modules.reports import build_report as _report
    from qa_conformance_modules.requirements import build_requirements
except ImportError:  # pragma: no cover - package-style imports
    from .content_checks import (
        _expected_paragraphs,
        _find_body_start_index,
        _find_para_by_text,
        _style_issues,
    )
    from .docx_checks import (
        _content_tables,
        _find_table,
        _formula_issues,
        _image_issues,
        _page_geometry_issues,
        _table_border_issues,
    )
    from .ooxml import W, _text_of_para
    from .registry import TEXT_ONLY_ROLES, VALID_MODES
    from .reports import build_report as _report
    from .requirements import build_requirements

def _load_json(path: str) -> Dict[str, Any]:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

def _read_docx(out_dir: str, output_docx_name: str) -> Tuple[ET.Element, Dict[str, bytes]]:
    path = os.path.join(out_dir, output_docx_name)
    with zipfile.ZipFile(path) as zf:
        data = {name: zf.read(name) for name in zf.namelist() if name.startswith("word/")}
    return ET.fromstring(data["word/document.xml"]), data

def _issue(code: str, severity: str, message: str, detail: str = "") -> Dict[str, Any]:
    return {"code": code, "severity": severity, "message": message, "detail": detail}

def check_conformance(out_dir: str, mode: str = "user", output_docx_name: str = "最终论文.docx", project_root: str | None = None) -> Dict[str, Any]:
    mode = mode if mode in VALID_MODES else "user"
    out_dir = os.path.abspath(out_dir)
    issues: List[Dict[str, Any]] = []
    counts: Dict[str, Any] = {}

    def add(code: str, severity: str, message: str, detail: str = "") -> None:
        issues.append(_issue(code, severity, message, detail))

    paths = {
        "format": os.path.join(out_dir, "format.json"),
        "content": os.path.join(out_dir, "content.json"),
        "requirements": os.path.join(out_dir, "template_requirements.json"),
        "manifest": os.path.join(out_dir, "build_manifest.json"),
        "docx": os.path.join(out_dir, output_docx_name),
    }
    missing = [key for key, path in paths.items() if key != "requirements" and not os.path.exists(path)]
    if missing:
        for key in missing:
            add("CONFORMANCE_INPUT_MISSING", "error", f"Missing required conformance input: {key}", paths[key])
        return _report(out_dir, mode, counts, issues, project_root)

    fmt = _load_json(paths["format"])
    content = _load_json(paths["content"])
    req = _load_json(paths["requirements"]) if os.path.exists(paths["requirements"]) else build_requirements(fmt, content)
    manifest = _load_json(paths["manifest"])
    manifest_counts = manifest.get("counts") or {}

    try:
        root, _parts = _read_docx(out_dir, output_docx_name)
    except Exception as exc:
        add("DOCX_XML_UNREADABLE", "error", "Could not read final DOCX XML.", str(exc))
        return _report(out_dir, mode, counts, issues, project_root)

    paragraphs = [p for p in root.iter(W + "p") if _text_of_para(p)]
    tables_xml = list(root.iter(W + "tbl"))
    all_text = "\n".join(_text_of_para(p) for p in paragraphs)
    counts["paragraphs_with_text"] = len(paragraphs)
    counts["docx_tables"] = len(tables_xml)

    section_count, page_issues = _page_geometry_issues(root, req)
    counts["docx_sections"] = section_count
    for detail in page_issues:
        add("PAGE_GEOMETRY_MISMATCH", "error", "Final DOCX page geometry does not match template requirements.", detail)

    style_roles = req.get("style_roles") or {}
    expected = _expected_paragraphs(content)
    body_start = _find_body_start_index(paragraphs, expected)
    body_paragraphs = paragraphs[body_start:]
    used_body_paragraphs: set[int] = set()
    counts["expected_content_paragraphs"] = len(expected)
    missing_samples = []
    style_mismatches = []
    for item in expected:
        role = item["role"]
        text = item["text"]
        profile = style_roles.get(role if role in style_roles else "body")
        para = _find_para_by_text(body_paragraphs, text, used_body_paragraphs)
        if para is None:
            missing_samples.append(text[:80])
            continue
        if profile and role in TEXT_ONLY_ROLES:
            style_mismatches.extend(_style_issues(role, _text_of_para(para), para, profile))
    counts["missing_content_paragraphs"] = len(missing_samples)
    counts["style_mismatches"] = len(style_mismatches)
    if missing_samples:
        add("CONTENT_PARAGRAPH_MISSING", "error", "Some extracted content paragraphs are missing from final DOCX.", " / ".join(missing_samples[:8]))
    if style_mismatches:
        add("STYLE_MISMATCH", "error", "Some final DOCX paragraphs do not match template role styles.", " / ".join(style_mismatches[:12]))

    content_counts = req.get("expected_counts") or {}
    for key, manifest_key in [
        ("images", "content_images_rendered"),
        ("tables", "content_tables_rendered"),
        ("formulas", "content_formulas_rendered"),
    ]:
        expected_count = int(content_counts.get(key) or 0)
        rendered_count = int(manifest_counts.get(manifest_key) or 0)
        counts[f"expected_{key}"] = expected_count
        counts[f"rendered_{key}"] = rendered_count
        if expected_count and rendered_count < expected_count:
            add("RENDER_COUNT_MISMATCH", "error", f"Rendered {key} count is lower than template/content requirement.", f"expected={expected_count} rendered={rendered_count}")

    for idx, rows in enumerate(_content_tables(content), 1):
        table_xml = _find_table(tables_xml, rows)
        if table_xml is None:
            add("TABLE_NOT_FOUND", "error", "A content table could not be found in final DOCX.", f"table={idx}")
            continue
        for detail in _table_border_issues(table_xml, idx)[:12]:
            add("TABLE_BORDER_MISMATCH", "error", "A content table does not satisfy three-line border requirements.", detail)

    drawing_count, drawing_issues = _image_issues(root, req)
    counts["docx_drawing_extents"] = drawing_count
    if content_counts.get("images") and drawing_count < int(content_counts.get("images") or 0):
        add("IMAGE_COUNT_MISMATCH", "error", "Final DOCX has fewer image drawings than expected.", f"expected={content_counts.get('images')} drawing_extents={drawing_count}")
    for detail in drawing_issues[:12]:
        add("IMAGE_LAYOUT_MISMATCH", "error", "An image layout constraint failed.", detail)

    formula_counts, formula_issue_list = _formula_issues(root)
    counts.update({f"docx_{k}": v for k, v in formula_counts.items()})
    if int(content_counts.get("formulas") or 0) and formula_counts["oMath"] < int(content_counts.get("formulas") or 0):
        add("FORMULA_COUNT_MISMATCH", "error", "Final DOCX has fewer native OOXML formulas than expected.", f"expected={content_counts.get('formulas')} oMath={formula_counts['oMath']}")
    for detail in formula_issue_list:
        add("OMML_WPS_COMPAT", "error", "A native math run is missing m:rPr, which can break WPS rendering.", detail)

    if re.search(r"LATEX_ERROR|FORMULA_ERROR|\[LaTeX error", all_text, re.I):
        add("FORMULA_ERROR_TEXT", "error", "Formula conversion error text remains in final DOCX.")
    if re.search(r"(\{\{[^}]+\}\}|TODO|FIXME|\(Insert\b|\(E\.g\.\s*X{2,})", all_text, re.I):
        add("PLACEHOLDER_TEXT_LEFT", "error", "Template or generated placeholder text remains in final DOCX.")
    if re.search(r"Error!\s*(Reference source not found|Bookmark not defined)|错误！未找到", all_text, re.I):
        add("WORD_FIELD_ERROR", "error", "Word field error text remains in final DOCX.")

    return _report(out_dir, mode, counts, issues, project_root)

