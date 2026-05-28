"""OOXML scalar and paragraph-format helpers for template extraction."""
from __future__ import annotations

from docx.oxml.ns import qn

ALIGN_MAP = {0: "LEFT", 1: "CENTER", 2: "RIGHT", 3: "JUSTIFY", None: "DEFAULT"}


def tag(el):
    return el.tag.split("}")[-1] if "}" in el.tag else el.tag


def val(el, attr="w:val", default=None):
    return el.get(qn(attr), default)


def pt(half_pts_str):
    """Convert half-points string to float. '28' -> 14.0."""
    try:
        return int(half_pts_str) / 2.0
    except Exception:
        return None


def emu_to_pt(emu):
    try:
        return round(int(emu) / 12700, 1)
    except Exception:
        return None


def twips_to_pt(value):
    try:
        return round(int(value) / 20.0, 2)
    except Exception:
        return None


def twips_to_cm(value):
    try:
        return round(int(value) / 567.0, 2)
    except Exception:
        return None


def paragraph_metrics(p_elem):
    """Extract paragraph metrics directly from OOXML."""
    info = {
        "alignment": "DEFAULT",
        "line_spacing_val": None,
        "line_spacing_rule": None,
        "line_spacing_fixed_pt": None,
        "space_before_pt": None,
        "space_after_pt": None,
        "first_indent_cm": None,
        "left_indent_cm": None,
        "right_indent_cm": None,
        "hanging_indent_cm": None,
    }
    pPr = p_elem.find(qn("w:pPr"))
    if pPr is None:
        return info
    jc = pPr.find(qn("w:jc"))
    if jc is not None:
        value = val(jc)
        info["alignment"] = {
            "left": "LEFT",
            "center": "CENTER",
            "right": "RIGHT",
            "both": "JUSTIFY",
            "distribute": "DISTRIBUTE",
        }.get(value, "DEFAULT")
    spacing = pPr.find(qn("w:spacing"))
    if spacing is not None:
        line = spacing.get(qn("w:line"))
        rule = spacing.get(qn("w:lineRule"))
        info["line_spacing_rule"] = rule
        if line:
            try:
                n = int(line)
                if rule in ("exact", "atLeast"):
                    info["line_spacing_fixed_pt"] = round(n / 20.0, 2)
                    info["line_spacing_val"] = info["line_spacing_fixed_pt"]
                else:
                    info["line_spacing_val"] = round(n / 240.0, 4)
            except Exception:
                pass
        info["space_before_pt"] = twips_to_pt(spacing.get(qn("w:before")))
        info["space_after_pt"] = twips_to_pt(spacing.get(qn("w:after")))
    ind = pPr.find(qn("w:ind"))
    if ind is not None:
        left = ind.get(qn("w:left"))
        if left is not None:
            info["left_indent_cm"] = twips_to_cm(left)
        right = ind.get(qn("w:right"))
        if right is not None:
            info["right_indent_cm"] = twips_to_cm(right)
        hanging = ind.get(qn("w:hanging"))
        if hanging is not None:
            info["hanging_indent_cm"] = twips_to_cm(hanging)
        first = ind.get(qn("w:firstLine"))
        if first is not None:
            info["first_indent_cm"] = twips_to_cm(first) or 0
    return info

