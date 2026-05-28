"""DOCX XML element checks for strict conformance QA."""
from __future__ import annotations

from typing import Any, Dict, List, Optional, Tuple
from xml.etree import ElementTree as ET

try:
    from qa_conformance_modules.content_checks import _is_table_item
    from qa_conformance_modules.ooxml import M, W, WP, _attr, _cm_to_twips, _compact, _text_of_table
except ImportError:  # pragma: no cover - package-style imports
    from .content_checks import _is_table_item
    from .ooxml import M, W, WP, _attr, _cm_to_twips, _compact, _text_of_table


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

