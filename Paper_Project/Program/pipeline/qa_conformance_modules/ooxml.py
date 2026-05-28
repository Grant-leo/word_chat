"""OOXML/XML helper functions for strict conformance checks."""
from __future__ import annotations

import re
from typing import Any, Dict, Optional
from xml.etree import ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
W = f"{{{W_NS}}}"
M = f"{{{M_NS}}}"
WP = f"{{{WP_NS}}}"


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
        "left_indent_cm", "hanging_indent_cm",
    ]
    return {k: profile.get(k) for k in keys if k in profile and profile.get(k) is not None}

