"""PDF template extraction support.

PDF templates are not editable Word templates, so this module extracts a
best-effort format profile from either:

* an instruction-style PDF that describes page/style requirements in text; or
* a visual sample PDF whose typography and margins can be estimated from text
  bounding boxes.

Scanned or textless PDFs are surfaced as unsupported with explicit metadata so
QA can block the pipeline with a repair prompt.
"""

from __future__ import annotations

import hashlib
import os
import re
import shutil
import subprocess
import tempfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path
from statistics import median
from typing import Any

try:
    from format_extractor_modules.style_profiles import build_style_profiles
    from md_parser_modules.format_extractor import _build_format_dict
except ImportError:  # pragma: no cover - package-style imports
    from .style_profiles import build_style_profiles
    from ..md_parser_modules.format_extractor import _build_format_dict


PT_TO_CM = 2.54 / 72.0
DEFAULT_PAGE = {
    "page_width_cm": 21.0,
    "page_height_cm": 29.7,
    "width_cm": 21.0,
    "height_cm": 29.7,
    "margin_top_cm": 2.5,
    "margin_bottom_cm": 2.5,
    "margin_left_cm": 3.0,
    "margin_right_cm": 2.5,
}


@dataclass
class PdfWord:
    page: int
    x_min: float
    y_min: float
    x_max: float
    y_max: float
    text: str

    @property
    def height(self) -> float:
        return max(0.1, self.y_max - self.y_min)

    @property
    def center_y(self) -> float:
        return (self.y_min + self.y_max) / 2.0


def extract_pdf_template(pdf_path: str | os.PathLike[str]) -> tuple[dict[str, Any], str]:
    """Extract a format dictionary and markdown report from a PDF template."""

    path = Path(pdf_path)
    meta = _base_pdf_meta(path)
    info = _read_pdf_info(path)
    meta["backend"]["pdfinfo"] = info["ok"]
    if info.get("error"):
        meta["warnings"].append(info["error"])
    page = _page_from_pdfinfo(info) or dict(DEFAULT_PAGE)
    page_count = int(info.get("page_count") or 0)
    if page_count:
        meta["page_count"] = page_count

    text_result = _pdftotext_layout(path)
    meta["backend"]["pdftotext"] = text_result["ok"]
    text = text_result["text"]
    if text_result["error"]:
        meta["warnings"].append(text_result["error"])

    words_result = _pdftotext_bbox(path)
    meta["backend"]["bbox"] = words_result["ok"]
    words = words_result["words"]
    if words_result["page"] is not None:
        page = {**page, **words_result["page"]}
    if words_result["error"]:
        meta["warnings"].append(words_result["error"])
    if not text.strip() and words:
        text = "\n".join(line["text"] for line in _group_words_into_lines(words))

    meta["text_chars"] = len(text.strip())
    meta["word_count"] = _word_count(text)

    template_type = _classify_pdf_template(text, words, bool(text_result["ok"] or words))
    meta["type"] = template_type

    if template_type == "instruction_pdf":
        fmt = _format_from_instruction_pdf(path, text, page, meta)
    elif template_type == "visual_sample_pdf":
        fmt = _format_from_visual_pdf(path, text, words, page, meta)
    else:
        meta["errors"].append("PDF_TEMPLATE_NO_TEXT")
        fmt = _minimal_unsupported_format(path, page, meta)

    fmt.setdefault("_meta", {})["pdf_template"] = meta
    fmt["_meta"]["source_ext"] = ".pdf"
    fmt["_meta"]["source_type"] = "pdf_template"
    fmt["style_profiles"] = build_style_profiles(fmt)

    return fmt, _pdf_report(path, fmt, text, words)


def _base_pdf_meta(path: Path) -> dict[str, Any]:
    return {
        "source": str(path),
        "source_name": path.name,
        "type": "unknown",
        "confidence": 0.0,
        "page_count": 0,
        "text_chars": 0,
        "word_count": 0,
        "backend": {"pdfinfo": False, "pdftotext": False, "bbox": False},
        "warnings": [],
        "errors": [],
    }


def _read_pdf_info(path: Path) -> dict[str, Any]:
    exe = shutil.which("pdfinfo")
    if not exe:
        return {"ok": False, "error": "PDFINFO_MISSING", "raw": ""}
    try:
        proc = subprocess.run(
            [exe, str(path)],
            capture_output=True,
            check=False,
            timeout=20,
        )
    except Exception as exc:  # pragma: no cover - platform dependent
        return {"ok": False, "error": f"PDFINFO_FAILED: {exc}", "raw": ""}
    raw = _decode_bytes(proc.stdout) + _decode_bytes(proc.stderr)
    if proc.returncode != 0:
        return {"ok": False, "error": "PDFINFO_FAILED", "raw": raw}
    page_count = 0
    match = re.search(r"^Pages:\s*(\d+)", raw, re.MULTILINE)
    if match:
        page_count = int(match.group(1))
    return {"ok": True, "error": "", "raw": raw, "page_count": page_count}


def _page_from_pdfinfo(info: dict[str, Any]) -> dict[str, float] | None:
    raw = str(info.get("raw") or "")
    match = re.search(
        r"Page\s+size:\s*([0-9.]+)\s+x\s+([0-9.]+)\s+pts",
        raw,
        re.IGNORECASE,
    )
    if not match:
        return None
    width_pt = float(match.group(1))
    height_pt = float(match.group(2))
    return {
        "page_width_cm": round(width_pt * PT_TO_CM, 2),
        "page_height_cm": round(height_pt * PT_TO_CM, 2),
        "width_cm": round(width_pt * PT_TO_CM, 2),
        "height_cm": round(height_pt * PT_TO_CM, 2),
        "margin_top_cm": DEFAULT_PAGE["margin_top_cm"],
        "margin_bottom_cm": DEFAULT_PAGE["margin_bottom_cm"],
        "margin_left_cm": DEFAULT_PAGE["margin_left_cm"],
        "margin_right_cm": DEFAULT_PAGE["margin_right_cm"],
    }


def _pdftotext_layout(path: Path) -> dict[str, Any]:
    exe = shutil.which("pdftotext")
    if not exe:
        return {"ok": False, "error": "PDFTOTEXT_MISSING", "text": ""}
    try:
        proc = subprocess.run(
            [exe, "-layout", "-enc", "UTF-8", str(path), "-"],
            capture_output=True,
            check=False,
            timeout=30,
        )
    except Exception as exc:  # pragma: no cover - platform dependent
        return {"ok": False, "error": f"PDFTOTEXT_FAILED: {exc}", "text": ""}
    text = _decode_bytes(proc.stdout)
    if proc.returncode != 0:
        return {"ok": False, "error": "PDFTOTEXT_FAILED", "text": text}
    return {"ok": True, "error": "", "text": text}


def _pdftotext_bbox(path: Path) -> dict[str, Any]:
    exe = shutil.which("pdftotext")
    if not exe:
        return {"ok": False, "error": "PDFTOTEXT_BBOX_MISSING", "words": [], "page": None}
    with tempfile.TemporaryDirectory(prefix="pdf_tpl_") as tmp_dir:
        out_path = Path(tmp_dir) / "bbox.html"
        try:
            proc = subprocess.run(
                [exe, "-bbox-layout", "-enc", "UTF-8", str(path), str(out_path)],
                capture_output=True,
                check=False,
                timeout=30,
            )
        except Exception as exc:  # pragma: no cover - platform dependent
            return {"ok": False, "error": f"PDFTOTEXT_BBOX_FAILED: {exc}", "words": [], "page": None}
        if proc.returncode != 0 or not out_path.exists():
            return {
                "ok": False,
                "error": "PDFTOTEXT_BBOX_FAILED",
                "words": [],
                "page": None,
            }
        try:
            root = ET.fromstring(out_path.read_text(encoding="utf-8", errors="ignore"))
        except ET.ParseError as exc:
            return {"ok": False, "error": f"PDFTOTEXT_BBOX_PARSE_FAILED: {exc}", "words": [], "page": None}
    words: list[PdfWord] = []
    page_geom: dict[str, float] | None = None
    for page_idx, page_el in enumerate(_iter_tags(root, "page"), start=1):
        width = _float_attr(page_el, "width")
        height = _float_attr(page_el, "height")
        if width and height and page_geom is None:
            page_geom = {
                "page_width_cm": round(width * PT_TO_CM, 2),
                "page_height_cm": round(height * PT_TO_CM, 2),
                "width_cm": round(width * PT_TO_CM, 2),
                "height_cm": round(height * PT_TO_CM, 2),
            }
        page_no = int(page_el.attrib.get("number") or page_idx)
        for word_el in _iter_tags(page_el, "word"):
            text = "".join(word_el.itertext()).strip()
            if not text:
                continue
            x_min = _float_attr(word_el, "xMin")
            y_min = _float_attr(word_el, "yMin")
            x_max = _float_attr(word_el, "xMax")
            y_max = _float_attr(word_el, "yMax")
            if None in (x_min, y_min, x_max, y_max):
                continue
            words.append(PdfWord(page_no, x_min, y_min, x_max, y_max, text))
    if page_geom:
        margins = _estimate_margins(words, page_geom)
        page_geom.update(margins)
    return {"ok": True, "error": "", "words": words, "page": page_geom}


def _iter_tags(root: ET.Element, local_name: str):
    for elem in root.iter():
        if elem.tag.rsplit("}", 1)[-1] == local_name:
            yield elem


def _float_attr(elem: ET.Element, name: str) -> float | None:
    raw = elem.attrib.get(name)
    if raw is None:
        return None
    try:
        return float(raw)
    except ValueError:
        return None


def _estimate_margins(words: list[PdfWord], page: dict[str, float]) -> dict[str, float]:
    page_words = [w for w in words if w.page <= 3]
    if not page_words:
        return {
            "margin_top_cm": DEFAULT_PAGE["margin_top_cm"],
            "margin_bottom_cm": DEFAULT_PAGE["margin_bottom_cm"],
            "margin_left_cm": DEFAULT_PAGE["margin_left_cm"],
            "margin_right_cm": DEFAULT_PAGE["margin_right_cm"],
        }
    page_w_pt = _page_width_cm(page) / PT_TO_CM
    page_h_pt = _page_height_cm(page) / PT_TO_CM
    min_x = min(w.x_min for w in page_words)
    max_x = max(w.x_max for w in page_words)
    min_y = min(w.y_min for w in page_words)
    max_y = max(w.y_max for w in page_words)
    return {
        "margin_top_cm": _clamped_cm(min_y * PT_TO_CM, 0.5, 6.0),
        "margin_bottom_cm": _clamped_cm((page_h_pt - max_y) * PT_TO_CM, 0.5, 6.0),
        "margin_left_cm": _clamped_cm(min_x * PT_TO_CM, 0.5, 6.0),
        "margin_right_cm": _clamped_cm((page_w_pt - max_x) * PT_TO_CM, 0.5, 6.0),
    }


def _clamped_cm(value: float, lo: float, hi: float) -> float:
    return round(max(lo, min(hi, value)), 2)


def _page_width_cm(page: dict[str, Any]) -> float:
    return float(page.get("page_width_cm") or page.get("width_cm") or DEFAULT_PAGE["page_width_cm"])


def _page_height_cm(page: dict[str, Any]) -> float:
    return float(page.get("page_height_cm") or page.get("height_cm") or DEFAULT_PAGE["page_height_cm"])


def _section_page(page: dict[str, Any]) -> dict[str, Any]:
    return {
        "index": 0,
        "page_width_cm": round(_page_width_cm(page), 2),
        "page_height_cm": round(_page_height_cm(page), 2),
        "margin_top_cm": page.get("margin_top_cm", DEFAULT_PAGE["margin_top_cm"]),
        "margin_bottom_cm": page.get("margin_bottom_cm", DEFAULT_PAGE["margin_bottom_cm"]),
        "margin_left_cm": page.get("margin_left_cm", DEFAULT_PAGE["margin_left_cm"]),
        "margin_right_cm": page.get("margin_right_cm", DEFAULT_PAGE["margin_right_cm"]),
        "diff_first_page": False,
        "header": [],
        "footer": [],
    }


def _classify_pdf_template(text: str, words: list[PdfWord], text_ok: bool) -> str:
    normalized = re.sub(r"\s+", " ", text).strip().lower()
    if not text_ok or len(normalized) < 20 or _word_count(normalized) < 4:
        return "scanned_or_unsupported_pdf"

    instruction_terms = [
        "格式",
        "要求",
        "模板说明",
        "正文",
        "一级标题",
        "二级标题",
        "三级标题",
        "字体",
        "字号",
        "行距",
        "页边距",
        "参考文献",
        "format",
        "requirement",
        "body",
        "heading",
        "font",
        "font size",
        "line spacing",
        "margin",
        "reference",
    ]
    instruction_score = sum(1 for term in instruction_terms if term in normalized)
    visual_terms = ["abstract", "introduction", "conclusion", "figure", "table", "摘要", "引言", "结论", "图", "表"]
    visual_score = sum(1 for term in visual_terms if term in normalized)

    if instruction_score >= 4 or ("format" in normalized and "requirement" in normalized):
        return "instruction_pdf"
    if words and (_has_varied_font_sizes(words) or visual_score >= 3):
        return "visual_sample_pdf"
    if instruction_score >= 2:
        return "instruction_pdf"
    return "visual_sample_pdf"


def _has_varied_font_sizes(words: list[PdfWord]) -> bool:
    heights = [w.height for w in words if w.height > 1]
    if len(heights) < 8:
        return False
    return max(heights) - min(heights) >= 3.0


def _format_from_instruction_pdf(
    path: Path,
    text: str,
    page: dict[str, float],
    meta: dict[str, Any],
) -> dict[str, Any]:
    fmt = _build_format_dict(str(path), text, page_override=_section_page(page))
    fmt["_meta"] = {
        **fmt.get("_meta", {}),
        "source": str(path),
        "source_name": path.name,
        "source_hash": _sha16(path),
        "source_ext": ".pdf",
        "source_type": "pdf_template",
    }
    sections = fmt.setdefault("sections", [])
    if sections:
        sections[0].update(_section_page(page))
    else:
        sections.append(_section_page(page))
    fmt.setdefault("paragraphs", [])
    fmt.setdefault("tables", [])
    fmt.setdefault("cover", [])
    fmt.setdefault("normal_style", _default_normal_style())
    meta["confidence"] = 0.72 if meta["backend"].get("pdftotext") else 0.25
    missing_roles = _instruction_missing_roles(text)
    if "heading" in missing_roles or len(missing_roles) >= 2:
        meta["warnings"].append("PDF_TEMPLATE_INSTRUCTION_INCOMPLETE:" + ",".join(missing_roles))
        meta["confidence"] = min(float(meta["confidence"]), 0.52)
    if not meta["backend"].get("bbox"):
        meta["warnings"].append("PDF_TEMPLATE_BBOX_UNAVAILABLE")
    return fmt


def _format_from_visual_pdf(
    path: Path,
    text: str,
    words: list[PdfWord],
    page: dict[str, float],
    meta: dict[str, Any],
) -> dict[str, Any]:
    lines = _group_words_into_lines(words)
    paragraphs = [_paragraph_from_line(i, line, page) for i, line in enumerate(lines)]
    if not paragraphs and text.strip():
        paragraphs = _paragraphs_from_layout_text(text)
    fmt = {
        "_meta": {
            "source": str(path),
            "source_name": path.name,
            "source_hash": _sha16(path),
            "source_ext": ".pdf",
            "source_type": "pdf_template",
        },
        "sections": [_section_page(page)],
        "paragraphs": paragraphs,
        "tables": [],
        "cover": [],
        "normal_style": _default_normal_style(),
    }
    meta["confidence"] = 0.64 if paragraphs and meta["backend"].get("bbox") else 0.38
    meta["warnings"].append("PDF_TEMPLATE_VISUAL_APPROXIMATION")
    if not meta["backend"].get("bbox"):
        meta["warnings"].append("PDF_TEMPLATE_BBOX_UNAVAILABLE")
    return fmt


def _minimal_unsupported_format(path: Path, page: dict[str, float], meta: dict[str, Any]) -> dict[str, Any]:
    meta["confidence"] = 0.0
    return {
        "_meta": {
            "source": str(path),
            "source_name": path.name,
            "source_hash": _sha16(path),
            "source_ext": ".pdf",
            "source_type": "pdf_template",
        },
        "sections": [_section_page(page)],
        "paragraphs": [],
        "tables": [],
        "cover": [],
        "normal_style": _default_normal_style(),
    }


def _group_words_into_lines(words: list[PdfWord]) -> list[dict[str, Any]]:
    if not words:
        return []
    lines: list[list[PdfWord]] = []
    for word in sorted(words, key=lambda w: (w.page, w.center_y, w.x_min)):
        if not lines:
            lines.append([word])
            continue
        prev = lines[-1]
        prev_y = median([w.center_y for w in prev])
        same_page = prev[0].page == word.page
        if same_page and abs(word.center_y - prev_y) <= max(3.0, median([w.height for w in prev]) * 0.45):
            prev.append(word)
        else:
            lines.append([word])

    result: list[dict[str, Any]] = []
    for row in lines[:320]:
        row = sorted(row, key=lambda w: w.x_min)
        text = _normalize_pdf_line(" ".join(w.text for w in row))
        if not text:
            continue
        result.append(
            {
                "page": row[0].page,
                "text": text,
                "x_min": min(w.x_min for w in row),
                "x_max": max(w.x_max for w in row),
                "y_min": min(w.y_min for w in row),
                "y_max": max(w.y_max for w in row),
                "height": median([w.height for w in row]),
            }
        )
    return result


def _paragraph_from_line(index: int, line: dict[str, Any], page: dict[str, float]) -> dict[str, Any]:
    text = line["text"]
    size_pt = _round_word_half_point(max(8.0, min(24.0, float(line.get("height") or 12.0))))
    page_w_pt = _page_width_cm(page) / PT_TO_CM
    line_width = float(line.get("x_max", 0.0)) - float(line.get("x_min", 0.0))
    center = (float(line.get("x_min", 0.0)) + float(line.get("x_max", 0.0))) / 2.0
    centered = abs(center - page_w_pt / 2.0) < page_w_pt * 0.08 and line_width < page_w_pt * 0.72
    bold = size_pt >= 14.0 or _looks_like_heading(text)
    return {
        "index": index,
        "style": "PDFTemplateLine",
        "text": text,
        "runs": [
            {
                "text": text,
                "font": "宋体" if _contains_cjk(text) else "Times New Roman",
                "size_pt": size_pt,
                "bold": bold,
                "italic": False,
            }
        ],
        "alignment": "CENTER" if centered else "JUSTIFY",
        "align": "CENTER" if centered else "JUSTIFY",
        "line_spacing_val": 1.25,
        "line_spacing_rule": "auto",
        "line_spacing_fixed_pt": None,
        "space_before_pt": 0,
        "space_after_pt": 0 if len(text) > 24 else 3,
        "first_indent_cm": 0.0 if centered or bold else 0.74,
        "left_indent_cm": 0.0,
        "right_indent_cm": 0.0,
        "has_page_break": False,
        "numbering": None,
    }


def _paragraphs_from_layout_text(text: str) -> list[dict[str, Any]]:
    paragraphs: list[dict[str, Any]] = []
    for idx, line in enumerate([ln.strip() for ln in text.splitlines() if ln.strip()][:240]):
        paragraphs.append(
            {
                "index": idx,
                "style": "PDFTemplateText",
                "text": line,
                "runs": [
                    {
                        "text": line,
                        "font": "宋体" if _contains_cjk(line) else "Times New Roman",
                        "size_pt": 12,
                        "bold": _looks_like_heading(line),
                        "italic": False,
                    }
                ],
                "alignment": "CENTER" if idx == 0 else "JUSTIFY",
                "align": "CENTER" if idx == 0 else "JUSTIFY",
                "line_spacing_val": 1.25,
                "line_spacing_rule": "auto",
                "space_before_pt": 0,
                "space_after_pt": 0,
                "first_indent_cm": 0.0 if idx == 0 else 0.74,
                "has_page_break": False,
            }
        )
    return paragraphs


def _default_normal_style() -> dict[str, Any]:
    return {
        "font": "宋体",
        "size_pt": 12,
        "bold": False,
        "italic": False,
        "alignment": "JUSTIFY",
        "line_spacing_val": 1.25,
        "space_before_pt": 0,
        "space_after_pt": 0,
        "first_indent_cm": 0.74,
    }


def _round_word_half_point(size_pt: float) -> float:
    """Normalize PDF bbox estimates to Word's half-point font-size grid."""
    return round(size_pt * 2.0) / 2.0


def _instruction_missing_roles(text: str) -> list[str]:
    normalized = re.sub(r"\s+", " ", text).strip().lower()
    role_terms = {
        "page": ["page", "a4", "margin", "页边距", "页面", "纸张"],
        "body": ["body", "正文", "font", "line spacing", "行距"],
        "heading": ["heading", "一级标题", "二级标题", "三级标题", "h1", "标题"],
        "caption": ["caption", "figure caption", "table caption", "图题", "表题", "题注"],
        "reference": ["reference", "references", "参考文献"],
    }
    missing = []
    for role, terms in role_terms.items():
        if not any(term in normalized for term in terms):
            missing.append(role)
    return missing


def _pdf_report(path: Path, fmt: dict[str, Any], text: str, words: list[PdfWord]) -> str:
    pdf_meta = fmt.get("_meta", {}).get("pdf_template", {})
    sections = fmt.get("sections") or []
    page = sections[0] if isinstance(sections, list) and sections else {}
    lines = [
        f"# PDF 模板格式提取 - {path.name}",
        "",
        f"- 类型: {pdf_meta.get('type', 'unknown')}",
        f"- 置信度: {pdf_meta.get('confidence', 0):.2f}",
        f"- 页数: {pdf_meta.get('page_count', 0)}",
        f"- 可提取文本字符数: {pdf_meta.get('text_chars', 0)}",
        f"- bbox词元数: {len(words)}",
        f"- 页面: {page.get('page_width_cm')}cm x {page.get('page_height_cm')}cm",
        f"- 页边距估计: 上{page.get('margin_top_cm')} / 下{page.get('margin_bottom_cm')} / 左{page.get('margin_left_cm')} / 右{page.get('margin_right_cm')} cm",
        "",
    ]
    warnings = pdf_meta.get("warnings") or []
    errors = pdf_meta.get("errors") or []
    if errors:
        lines.extend(["## 错误", *[f"- {item}" for item in errors], ""])
    if warnings:
        lines.extend(["## 警告", *[f"- {item}" for item in warnings], ""])
    sample_lines = [ln.strip() for ln in text.splitlines() if ln.strip()][:20]
    if sample_lines:
        lines.extend(["## 文本样例", *[f"- {ln}" for ln in sample_lines], ""])
    lines.append("## 说明")
    lines.append("- PDF 模板无法像 DOCX 一样读取完整样式树；精排 PDF 使用文本位置和字号估计格式。")
    lines.append("- 扫描件或不可复制文字的 PDF 会进入 QA 错误，需要用户提供 DOCX 模板、文字说明 PDF，或先做 OCR。")
    return "\n".join(lines)


def _looks_like_heading(text: str) -> bool:
    stripped = text.strip()
    return bool(
        re.match(r"^(\d+(\.\d+){0,3}\s+|第[一二三四五六七八九十]+[章节])", stripped)
        or stripped.lower() in {"abstract", "references", "introduction", "conclusion"}
        or stripped in {"摘要", "目录", "参考文献", "结论", "致谢"}
    )


def _contains_cjk(text: str) -> bool:
    return bool(re.search(r"[\u3400-\u9fff]", text))


def _normalize_pdf_line(text: str) -> str:
    text = re.sub(r"\s+", " ", text).strip()
    text = re.sub(r"\s+([,.;:!?，。；：！？）\]\}])", r"\1", text)
    text = re.sub(r"([（\[\{])\s+", r"\1", text)
    return text


def _word_count(text: str) -> int:
    if not text:
        return 0
    cjk = len(re.findall(r"[\u3400-\u9fff]", text))
    latin = len(re.findall(r"[A-Za-z0-9_]+", text))
    return cjk + latin


def _decode_bytes(data: bytes) -> str:
    for enc in ("utf-8", "gb18030", "latin-1"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="ignore")


def _sha16(path: Path) -> str:
    try:
        return hashlib.sha256(path.read_bytes()).hexdigest()[:16]
    except OSError:
        return ""
