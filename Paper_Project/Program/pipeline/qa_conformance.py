"""
qa_conformance.py - strict DOCX conformance QA for generated thesis outputs.

This module turns extracted template data into a machine-checkable requirement
file, then validates the final DOCX package directly.  It is intentionally more
specific than qa_checker.py:

- template_requirements.json/md records page, style, cover/header/footer, and
  expected rendered content counts
- conformance_report.json/md validates page geometry, all expected content
  paragraphs/headings/captions/references, tables, images, and native formulas
- checks read the DOCX XML rather than trusting generated reports alone
"""
from __future__ import annotations

import json
import os
import re
import zipfile
from datetime import datetime
from typing import Any, Dict, Iterable, List, Optional, Tuple
from xml.etree import ElementTree as ET

try:
    from privacy import sanitize_value
except Exception:  # pragma: no cover
    def sanitize_value(value: Any, project_root: str | None = None) -> Any:
        return value

try:
    from script_generator import _infer_style_profiles
except Exception:  # pragma: no cover
    _infer_style_profiles = None


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
W = f"{{{W_NS}}}"
M = f"{{{M_NS}}}"
WP = f"{{{WP_NS}}}"

VALID_MODES = {"user", "developer"}

STYLE_ROLES = [
    "body", "h1", "h2", "h3",
    "cn_title", "en_title", "cn_abstract_heading", "cn_abstract_body",
    "en_abstract_heading", "en_abstract_body", "figure_caption",
    "table_caption", "table_body", "table_header", "formula",
    "reference", "reference_english",
]

TEXT_ONLY_ROLES = {"body", "h1", "h2", "h3", "figure_caption", "table_caption", "reference"}
BACKMATTER_ROLES = {"references", "acknowledgement", "appendix"}
FRONTMATTER_ROLES = {"cn_abstract", "cn_keywords", "en_abstract", "en_keywords"}


def _load_json(path: str) -> Dict[str, Any]:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def _write_json(path: str, value: Dict[str, Any]) -> None:
    with open(path, "w", encoding="utf-8") as f:
        json.dump(value, f, ensure_ascii=False, indent=2)


def _qn(name: str) -> str:
    return W + name


def _attr(el: Optional[ET.Element], name: str) -> Optional[str]:
    return el.attrib.get(_qn(name)) if el is not None else None


def _bool_w(el: Optional[ET.Element]) -> bool:
    if el is None:
        return False
    val = _attr(el, "val")
    return str(val).lower() not in {"0", "false", "off"}


def _cm_to_twips(value: Any) -> int:
    return int(round(float(value or 0) / 2.54 * 1440))


def _pt_to_twips(value: Any) -> int:
    return int(round(float(value or 0) * 20))


def _cm_indent_to_twips(value: Any) -> int:
    return int(round(float(value or 0) * 567))


def _is_cjk_text(text: str) -> bool:
    return any("\u4e00" <= c <= "\u9fff" for c in str(text or ""))


def _is_cjk_font(font: str) -> bool:
    return _is_cjk_text(font)


def _compact(text: str) -> str:
    return re.sub(r"\s+", "", str(text or "")).lower()


def _text_of_para(p: ET.Element) -> str:
    return "".join(t.text or "" for t in p.iter(W + "t")).strip()


def _text_of_table(tbl: ET.Element) -> str:
    return "\n".join(t.text or "" for t in tbl.iter(W + "t")).strip()


def _first_run_props(p: ET.Element) -> Dict[str, Any]:
    run = p.find(W + "r")
    rpr = run.find(W + "rPr") if run is not None else None
    fonts = rpr.find(W + "rFonts") if rpr is not None else None
    size = rpr.find(W + "sz") if rpr is not None else None
    size_val = _attr(size, "val")
    return {
        "size": int(size_val) / 2.0 if size_val else None,
        "bold": _bool_w(rpr.find(W + "b")) if rpr is not None else False,
        "italic": _bool_w(rpr.find(W + "i")) if rpr is not None else False,
        "ascii_font": _attr(fonts, "ascii"),
        "hansi_font": _attr(fonts, "hAnsi"),
        "east_asia_font": _attr(fonts, "eastAsia"),
    }


def _para_props(p: ET.Element) -> Dict[str, Any]:
    ppr = p.find(W + "pPr")
    jc = ppr.find(W + "jc") if ppr is not None else None
    spacing = ppr.find(W + "spacing") if ppr is not None else None
    ind = ppr.find(W + "ind") if ppr is not None else None
    return {
        "align": _attr(jc, "val"),
        "line": int(_attr(spacing, "line")) if _attr(spacing, "line") else None,
        "line_rule": _attr(spacing, "lineRule"),
        "space_before": int(_attr(spacing, "before")) if _attr(spacing, "before") else 0,
        "space_after": int(_attr(spacing, "after")) if _attr(spacing, "after") else 0,
        "first_line": int(_attr(ind, "firstLine")) if _attr(ind, "firstLine") else 0,
        "left": int(_attr(ind, "left")) if _attr(ind, "left") else 0,
        "hanging": int(_attr(ind, "hanging")) if _attr(ind, "hanging") else 0,
    }


def _expected_align(value: Any) -> Optional[str]:
    return {
        "CENTER": "center",
        "LEFT": "left",
        "RIGHT": "right",
        "JUSTIFY": "both",
    }.get(str(value or "").upper())


def _expected_line_twips(profile: Dict[str, Any]) -> Optional[int]:
    if profile.get("line_spacing_fixed_pt") is not None:
        return _pt_to_twips(profile.get("line_spacing_fixed_pt"))
    if profile.get("line_spacing_val") is not None:
        return int(round(float(profile.get("line_spacing_val")) * 240))
    return None


def _profile_subset(profile: Dict[str, Any]) -> Dict[str, Any]:
    keys = [
        "font", "size", "bold", "italic", "align",
        "line_spacing_val", "line_spacing_rule", "line_spacing_fixed_pt",
        "space_before_pt", "space_after_pt", "first_indent_cm",
    ]
    return {k: profile.get(k) for k in keys if k in profile and profile.get(k) is not None}


def _profile_from_format(fmt: Dict[str, Any]) -> Dict[str, Dict[str, Any]]:
    if _infer_style_profiles is not None:
        try:
            return _infer_style_profiles(fmt)
        except Exception:
            pass
    return fmt.get("style_profiles") or {}


def _content_counts(content: Dict[str, Any]) -> Dict[str, int]:
    images = 0
    tables = 0
    formulas = 0
    for sec in content.get("sections") or []:
        for item in sec.get("paragraphs") or []:
            if not isinstance(item, dict):
                continue
            if item.get("role") in {"image", "figure"} or item.get("image") or item.get("filename"):
                images += 1
            if item.get("table_rows") and item.get("role") != "code":
                tables += 1
            math_items = item.get("math") or []
            if math_items:
                formulas += len(math_items)
            elif item.get("role") == "formula" or item.get("latex") or item.get("xml"):
                formulas += 1
        if not images:
            images += len(sec.get("images") or [])
    if not images:
        images = int((content.get("_meta") or {}).get("images_extracted") or 0)
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


def _read_docx(out_dir: str, output_docx_name: str) -> Tuple[ET.Element, Dict[str, bytes]]:
    path = os.path.join(out_dir, output_docx_name)
    with zipfile.ZipFile(path) as zf:
        data = {name: zf.read(name) for name in zf.namelist() if name.startswith("word/")}
    return ET.fromstring(data["word/document.xml"]), data


def _issue(code: str, severity: str, message: str, detail: str = "") -> Dict[str, Any]:
    return {"code": code, "severity": severity, "message": message, "detail": detail}


def _style_issues(role: str, text: str, p: ET.Element, profile: Dict[str, Any]) -> List[str]:
    issues: List[str] = []
    run = _first_run_props(p)
    para = _para_props(p)
    if profile.get("size") is not None and run.get("size") is not None:
        if abs(float(run["size"]) - float(profile["size"])) > 0.05:
            issues.append(f"{role}: size {run['size']} != {profile['size']}")
    for key in ("bold", "italic"):
        if key in profile and bool(run.get(key)) != bool(profile.get(key)):
            issues.append(f"{role}: {key} {run.get(key)} != {bool(profile.get(key))}")
    expected = _expected_align(profile.get("align"))
    if expected and para.get("align") != expected:
        issues.append(f"{role}: align {para.get('align')} != {expected}")
    font = profile.get("font")
    if font:
        if _is_cjk_font(str(font)):
            if _is_cjk_text(text) and run.get("east_asia_font") != font:
                issues.append(f"{role}: eastAsia font {run.get('east_asia_font')} != {font}")
        elif run.get("ascii_font") != font:
            issues.append(f"{role}: ascii font {run.get('ascii_font')} != {font}")
    expected_line = _expected_line_twips(profile)
    if expected_line is not None and para.get("line") is not None and abs(int(para["line"]) - expected_line) > 1:
        issues.append(f"{role}: line {para.get('line')} != {expected_line}")
    if "space_before_pt" in profile and abs(int(para["space_before"]) - _pt_to_twips(profile.get("space_before_pt"))) > 1:
        issues.append(f"{role}: space_before {para.get('space_before')} != {_pt_to_twips(profile.get('space_before_pt'))}")
    if "space_after_pt" in profile and abs(int(para["space_after"]) - _pt_to_twips(profile.get("space_after_pt"))) > 1:
        issues.append(f"{role}: space_after {para.get('space_after')} != {_pt_to_twips(profile.get('space_after_pt'))}")
    if "first_indent_cm" in profile and abs(int(para["first_line"]) - _cm_indent_to_twips(profile.get("first_indent_cm"))) > 3:
        issues.append(f"{role}: first_line {para.get('first_line')} != {_cm_indent_to_twips(profile.get('first_indent_cm'))}")
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


def _page_geometry_issues(root: ET.Element, req: Dict[str, Any]) -> Tuple[int, List[str]]:
    issues: List[str] = []
    page = req.get("page") or {}
    exp_w = _cm_to_twips(page.get("page_width_cm", 21.0))
    exp_h = _cm_to_twips(page.get("page_height_cm", 29.7))
    margins = {
        "top": _cm_to_twips(page.get("margin_top_cm", 2.54)),
        "bottom": _cm_to_twips(page.get("margin_bottom_cm", 2.54)),
        "left": _cm_to_twips(page.get("margin_left_cm", 2.54)),
        "right": _cm_to_twips(page.get("margin_right_cm", 2.54)),
    }
    sections = list(root.iter(W + "sectPr"))
    for idx, sect in enumerate(sections):
        pg_size = sect.find(W + "pgSz")
        pg_margin = sect.find(W + "pgMar")
        got_w = int(_attr(pg_size, "w") or 0)
        got_h = int(_attr(pg_size, "h") or 0)
        if abs(got_w - exp_w) > 4 or abs(got_h - exp_h) > 4:
            issues.append(f"section {idx}: page size {got_w}x{got_h} != {exp_w}x{exp_h}")
        for key, exp in margins.items():
            got = int(_attr(pg_margin, key) or 0)
            if abs(got - exp) > 4:
                issues.append(f"section {idx}: margin {key} {got} != {exp}")
    return len(sections), issues


def _tc_border_val(cell: ET.Element, side: str) -> Optional[str]:
    tc_pr = cell.find(W + "tcPr")
    borders = tc_pr.find(W + "tcBorders") if tc_pr is not None else None
    el = borders.find(W + side) if borders is not None else None
    return _attr(el, "val")


def _table_border_issues(table: ET.Element, index: int) -> List[str]:
    rows = table.findall(W + "tr")
    if not rows:
        return [f"table {index}: no rows"]
    issues: List[str] = []
    first_cells = rows[0].findall(W + "tc")
    last_cells = rows[-1].findall(W + "tc")
    for ci, cell in enumerate(first_cells):
        if _tc_border_val(cell, "top") in {None, "nil", "none"}:
            issues.append(f"table {index}: first-row cell {ci} missing top border")
        if _tc_border_val(cell, "bottom") in {None, "nil", "none"}:
            issues.append(f"table {index}: first-row cell {ci} missing header bottom border")
    for ci, cell in enumerate(last_cells):
        if _tc_border_val(cell, "bottom") in {None, "nil", "none"}:
            issues.append(f"table {index}: last-row cell {ci} missing bottom border")
    return issues


def _content_tables(content: Dict[str, Any]) -> List[List[List[str]]]:
    tables: List[List[List[str]]] = []
    for sec in content.get("sections") or []:
        for item in sec.get("paragraphs") or []:
            if _is_table_item(item):
                tables.append([[str(cell or "") for cell in row] for row in item.get("table_rows") or []])
    return tables


def _find_table(tables_xml: List[ET.Element], rows: List[List[str]]) -> Optional[ET.Element]:
    samples = [_compact(cell) for row in rows[:2] for cell in row[:4] if str(cell or "").strip()]
    samples = [s for s in samples if s][:6]
    if not samples:
        return None
    for tbl in tables_xml:
        text = _compact(_text_of_table(tbl))
        if all(sample in text for sample in samples[: min(3, len(samples))]):
            return tbl
    return None


def _image_issues(root: ET.Element, req: Dict[str, Any]) -> Tuple[int, List[str]]:
    issues: List[str] = []
    extents = list(root.iter(WP + "extent"))
    page = req.get("page") or {}
    text_width_cm = float(page.get("page_width_cm", 21.0) or 21.0) - float(page.get("margin_left_cm", 2.54) or 0) - float(page.get("margin_right_cm", 2.54) or 0)
    max_cx = int(max(text_width_cm, 1.0) * 360000 * 1.05)
    for idx, extent in enumerate(extents):
        cx = int(extent.attrib.get("cx") or 0)
        cy = int(extent.attrib.get("cy") or 0)
        if cx <= 0 or cy <= 0:
            issues.append(f"image {idx}: non-positive extent {cx}x{cy}")
        if cx > max_cx:
            issues.append(f"image {idx}: width {cx} exceeds text width {max_cx}")
    return len(extents), issues


def _formula_issues(root: ET.Element) -> Tuple[Dict[str, int], List[str]]:
    issues: List[str] = []
    math_runs = list(root.iter(M + "r"))
    for idx, run in enumerate(math_runs):
        if run.find(M + "rPr") is None:
            issues.append(f"math run {idx}: missing m:rPr")
            if len(issues) >= 12:
                break
    return {
        "oMathPara": len(list(root.iter(M + "oMathPara"))),
        "oMath": len(list(root.iter(M + "oMath"))),
        "math_runs": len(math_runs),
    }, issues


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
    body_start = 0
    for idx, para in enumerate(paragraphs):
        if _compact(_text_of_para(para)) == _compact("1 Introduction"):
            body_start = idx
            break
    body_paragraphs = paragraphs[body_start:]
    tables_xml = list(root.iter(W + "tbl"))
    all_text = "\n".join(_text_of_para(p) for p in paragraphs)
    counts["paragraphs_with_text"] = len(paragraphs)
    counts["docx_tables"] = len(tables_xml)

    section_count, page_issues = _page_geometry_issues(root, req)
    counts["sections"] = section_count
    for detail in page_issues:
        add("PAGE_GEOMETRY_MISMATCH", "error", "Final DOCX page geometry does not match template requirements.", detail)

    style_roles = req.get("style_roles") or {}
    expected = _expected_paragraphs(content)
    counts["expected_content_paragraphs"] = len(expected)
    missing_samples = []
    style_mismatches = []
    for item in expected:
        role = item["role"]
        text = item["text"]
        profile = style_roles.get(role if role in style_roles else "body")
        para = _find_para_by_text(body_paragraphs, text)
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


def _report(out_dir: str, mode: str, counts: Dict[str, Any], issues: List[Dict[str, Any]], project_root: str | None) -> Dict[str, Any]:
    return {
        "schema_version": 1,
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "mode": mode,
        "output_dir_name": os.path.basename(os.path.abspath(out_dir)),
        "passed": not any(i.get("severity") == "error" for i in issues),
        "counts": counts,
        "issues": sanitize_value(issues, project_root),
        "next_action": (
            "Strict conformance passed for machine-checkable template requirements."
            if not any(i.get("severity") == "error" for i in issues)
            else ("Fix Outputs/<run>/build_generated.py and rerun it." if mode == "user" else "Fix core pipeline scripts and rerun the full pipeline.")
        ),
    }


def report_to_markdown(report: Dict[str, Any]) -> str:
    lines = [
        "# Conformance QA Report",
        "",
        f"- Result: {'passed' if report.get('passed') else 'failed'}",
        f"- Mode: `{report.get('mode')}`",
        f"- Output: `{report.get('output_dir_name')}`",
        f"- Next action: {report.get('next_action')}",
        "",
        "## Counts",
        "",
    ]
    for key, value in sorted((report.get("counts") or {}).items()):
        lines.append(f"- `{key}`: {value}")
    lines.extend(["", "## Issues", ""])
    if not report.get("issues"):
        lines.append("- No conformance issues detected.")
    else:
        for item in report.get("issues") or []:
            lines.append(f"- **{item.get('severity')}** `{item.get('code')}`: {item.get('message')}")
            if item.get("detail"):
                lines.append(f"  Detail: `{item.get('detail')}`")
    lines.append("")
    return "\n".join(lines)


def write_reports(report: Dict[str, Any], out_dir: str) -> None:
    _write_json(os.path.join(out_dir, "conformance_report.json"), report)
    with open(os.path.join(out_dir, "conformance_report.md"), "w", encoding="utf-8") as f:
        f.write(report_to_markdown(report))


def check_and_write(out_dir: str, mode: str = "user", output_docx_name: str = "最终论文.docx", project_root: str | None = None) -> Dict[str, Any]:
    report = check_conformance(out_dir, mode=mode, output_docx_name=output_docx_name, project_root=project_root)
    write_reports(report, out_dir)
    return report


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Run strict DOCX conformance QA on a generated output directory.")
    parser.add_argument("out_dir")
    parser.add_argument("--mode", choices=sorted(VALID_MODES), default="user")
    parser.add_argument("--docx", default="最终论文.docx")
    args = parser.parse_args()

    result = check_and_write(args.out_dir, mode=args.mode, output_docx_name=args.docx)
    print(report_to_markdown(result))
    raise SystemExit(0 if result.get("passed") else 1)
