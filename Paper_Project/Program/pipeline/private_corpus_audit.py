"""Private real-data corpus inventory and classifier.

This module is intentionally local/report-oriented. It records file locations,
structural features, and issue codes, but not document body text.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import subprocess
import zipfile
from datetime import datetime
from typing import Any, Dict, Iterable, List, Tuple

try:
    from content_parser_modules.source_audit import audit_docx_source
except ImportError:  # pragma: no cover - package-style import fallback
    from .content_parser_modules.source_audit import audit_docx_source


CLASSIFICATIONS = {
    "template_candidate",
    "content_candidate",
    "reference_candidate",
    "attachment_or_nonpaper",
    "unsupported_or_conversion_needed",
}
DOCX_EXT = ".docx"
PDF_EXT = ".pdf"
LEGACY_EXTS = {".doc", ".wps"}
ATTACHMENT_EXTS = {
    ".7z", ".bmp", ".csv", ".gif", ".htm", ".html", ".jpeg", ".jpg", ".lnk",
    ".png", ".rar", ".svg", ".tif", ".tiff", ".txt", ".webp", ".xls", ".xlsx",
    ".zip",
}


def _default_output_dir() -> str:
    return os.path.abspath(os.path.join(os.getcwd(), "Outputs", "_private_realdata_audit"))


def _sha256_prefix(path: str, limit_bytes: int = 64 * 1024 * 1024) -> str:
    h = hashlib.sha256()
    read_total = 0
    with open(path, "rb") as f:
        while True:
            chunk = f.read(1024 * 1024)
            if not chunk:
                break
            h.update(chunk)
            read_total += len(chunk)
            if read_total >= limit_bytes:
                break
    return h.hexdigest()[:16]


def _safe_rel(path: str, root: str) -> str:
    try:
        return os.path.relpath(path, root).replace("\\", "/")
    except Exception:
        return os.path.basename(path)


def _docx_features(path: str) -> Tuple[Dict[str, Any], List[str], bool]:
    features: Dict[str, Any] = {}
    reasons: List[str] = []
    try:
        with zipfile.ZipFile(path) as zf:
            names = zf.namelist()
            document_xml = zf.read("word/document.xml").decode("utf-8", errors="ignore")
            features.update(
                {
                    "paragraph_count": len(re.findall(r"<w:p\b", document_xml)),
                    "table_count": len(re.findall(r"<w:tbl\b", document_xml)),
                    "image_count": len([n for n in names if n.startswith("word/media/")]),
                    "formula_count": len(re.findall(r"<m:oMath\b|<m:oMathPara\b", document_xml)),
                    "section_count": len(re.findall(r"<w:sectPr\b", document_xml)),
                    "header_count": len([n for n in names if n.startswith("word/header") and n.endswith(".xml")]),
                    "footer_count": len([n for n in names if n.startswith("word/footer") and n.endswith(".xml")]),
                    "field_count": len(re.findall(r"<w:fldChar\b|<w:instrText\b", document_xml)),
                    "toc_field_count": len(re.findall(r">\s*TOC\b", document_xml, flags=re.IGNORECASE)),
                    "footnote_part": any(n == "word/footnotes.xml" for n in names),
                    "endnote_part": any(n == "word/endnotes.xml" for n in names),
                    "comments_part": any(n == "word/comments.xml" for n in names),
                    "tracked_change_count": len(re.findall(r"<w:(?:ins|del|moveFrom|moveTo)\b", document_xml)),
                    "textbox_count": len(re.findall(r"<w:txbxContent\b|<wps:txbx\b|<v:textbox\b", document_xml)),
                    "content_control_count": len(re.findall(r"<w:sdt\b", document_xml)),
                    "landscape_section_count": len(re.findall(r'w:orient=["\']landscape["\']', document_xml)),
                    "embedded_object_count": len([n for n in names if n.startswith("word/embeddings/")]),
                    "merged_cell_count": len(re.findall(r"<w:gridSpan\b|<w:hMerge\b|<w:vMerge\b", document_xml)),
                    "text_char_count": len(re.sub(r"<[^>]+>", "", document_xml).strip()),
                }
            )
            audit = audit_docx_source(path)
            features["source_audit_issue_codes"] = [
                str(issue.get("code")) for issue in audit.get("issues") or []
                if isinstance(issue, dict) and issue.get("code")
            ]
            return features, reasons, True
    except Exception as exc:
        reasons.append(f"bad_docx:{type(exc).__name__}")
        return features, reasons, False


def _run_text_command(args: List[str], timeout: int = 15) -> Tuple[int, str]:
    try:
        completed = subprocess.run(
            args,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="ignore",
            timeout=timeout,
            check=False,
        )
        return completed.returncode, (completed.stdout or "") + "\n" + (completed.stderr or "")
    except Exception as exc:
        return 999, type(exc).__name__


def _pdf_features(path: str) -> Tuple[Dict[str, Any], List[str], bool]:
    features: Dict[str, Any] = {"pdf_tools_available": bool(shutil.which("pdfinfo") and shutil.which("pdftotext"))}
    reasons: List[str] = []
    if not features["pdf_tools_available"]:
        reasons.append("pdf_tools_missing")
        return features, reasons, True
    rc, info = _run_text_command(["pdfinfo", path])
    if rc != 0:
        lowered = info.lower()
        if "password" in lowered or "encrypted" in lowered or "permission" in lowered:
            reasons.append("protected_pdf")
        else:
            reasons.append("pdf_read_failed")
        return features, reasons, False
    pages_match = re.search(r"^Pages:\s+(\d+)", info, flags=re.MULTILINE)
    size_match = re.search(r"^Page size:\s+([0-9.]+)\s+x\s+([0-9.]+)", info, flags=re.MULTILINE)
    page_count = int(pages_match.group(1)) if pages_match else 0
    width = float(size_match.group(1)) if size_match else 0.0
    height = float(size_match.group(2)) if size_match else 0.0
    features.update(
        {
            "page_count": page_count,
            "page_width_pt": width,
            "page_height_pt": height,
            "landscape_page": bool(width and height and width > height),
        }
    )
    rc, text = _run_text_command(["pdftotext", "-layout", path, "-"])
    if rc != 0:
        reasons.append("pdf_text_extract_failed")
        return features, reasons, True
    text_chars = len(re.sub(r"\s+", "", text or ""))
    features["text_char_count"] = text_chars
    features["text_chars_per_page"] = round(text_chars / max(1, page_count), 2)
    if page_count and text_chars / max(1, page_count) < 20:
        reasons.append("scanned_or_textless_pdf")
        return features, reasons, False
    return features, reasons, True


def _classify_docx(features: Dict[str, Any]) -> Tuple[str, float, List[str]]:
    reasons: List[str] = []
    paragraphs = int(features.get("paragraph_count") or 0)
    tables = int(features.get("table_count") or 0)
    fields = int(features.get("field_count") or 0)
    toc_fields = int(features.get("toc_field_count") or 0)
    headers = int(features.get("header_count") or 0)
    footers = int(features.get("footer_count") or 0)
    text_chars = int(features.get("text_char_count") or 0)
    risk_codes = set(features.get("source_audit_issue_codes") or [])

    if toc_fields or (headers + footers and fields):
        reasons.append("docx_has_toc_or_layout_fields")
        return "reference_candidate", 0.78, reasons
    if text_chars < 3000 and (tables >= 2 or headers + footers or fields):
        reasons.append("short_docx_with_format_structures")
        return "template_candidate", 0.72, reasons
    if paragraphs >= 120 or text_chars >= 12000:
        reasons.append("long_body_like_docx")
        return "content_candidate", 0.76, reasons
    if risk_codes & {"SOURCE_LANDSCAPE_SECTION_UNSUPPORTED", "COMPLEX_TABLE_UNSUPPORTED", "TABLE_MERGE_UNSUPPORTED"}:
        reasons.append("docx_has_layout_risk_reference_value")
        return "reference_candidate", 0.62, reasons
    reasons.append("ordinary_docx_body_candidate")
    return "content_candidate", 0.58, reasons


def _classify_pdf(features: Dict[str, Any], reasons: List[str]) -> Tuple[str, float, List[str]]:
    if not features.get("pdf_tools_available"):
        reasons.append("pdf_not_fully_classified_without_tools")
        return "template_candidate", 0.42, reasons
    if "scanned_or_textless_pdf" in reasons or "protected_pdf" in reasons or "pdf_read_failed" in reasons:
        return "unsupported_or_conversion_needed", 0.9, reasons
    pages = int(features.get("page_count") or 0)
    density = float(features.get("text_chars_per_page") or 0)
    if pages <= 8 and density > 40:
        reasons.append("short_text_pdf_likely_template_or_rules")
        return "template_candidate", 0.65, reasons
    if pages >= 9:
        reasons.append("multi_page_pdf_likely_reference_sample")
        return "reference_candidate", 0.68, reasons
    reasons.append("pdf_requires_human_review")
    return "template_candidate", 0.48, reasons


def _classify_file(path: str, root: str) -> Dict[str, Any]:
    ext = os.path.splitext(path)[1].lower()
    item: Dict[str, Any] = {
        "relative_path": _safe_rel(path, root),
        "extension": ext or "<no_ext>",
        "size_bytes": os.path.getsize(path),
        "classification": "attachment_or_nonpaper",
        "confidence": 0.4,
        "reasons": [],
        "features": {},
        "sha256_prefix": _sha256_prefix(path),
    }
    if ext in LEGACY_EXTS:
        item.update(
            classification="unsupported_or_conversion_needed",
            confidence=0.95,
            reasons=[f"legacy_format:{ext}", "manual_save_as_docx_required"],
        )
        return item
    if ext in ATTACHMENT_EXTS:
        item.update(
            classification="attachment_or_nonpaper",
            confidence=0.9,
            reasons=[f"attachment_or_nonpaper_ext:{ext}"],
        )
        return item
    if ext == DOCX_EXT:
        features, reasons, ok = _docx_features(path)
        item["features"] = features
        if not ok:
            item.update(classification="unsupported_or_conversion_needed", confidence=0.92, reasons=reasons)
            return item
        classification, confidence, extra_reasons = _classify_docx(features)
        item.update(classification=classification, confidence=confidence, reasons=reasons + extra_reasons)
        return item
    if ext == PDF_EXT:
        features, reasons, ok = _pdf_features(path)
        item["features"] = features
        if not ok:
            item.update(classification="unsupported_or_conversion_needed", confidence=0.9, reasons=reasons)
            return item
        classification, confidence, extra_reasons = _classify_pdf(features, reasons)
        item.update(classification=classification, confidence=confidence, reasons=extra_reasons)
        return item
    item.update(
        classification="attachment_or_nonpaper",
        confidence=0.65,
        reasons=[f"unrecognized_ext:{ext or '<no_ext>'}"],
    )
    return item


def _mark_duplicates(items: List[Dict[str, Any]]) -> None:
    first_seen: Dict[str, str] = {}
    for item in items:
        digest = str(item.get("sha256_prefix") or "")
        if not digest:
            continue
        if digest in first_seen:
            item["duplicate_of"] = first_seen[digest]
        else:
            first_seen[digest] = str(item.get("relative_path") or "")


def _summary(items: Iterable[Dict[str, Any]]) -> Dict[str, Any]:
    by_class: Dict[str, int] = {}
    by_ext: Dict[str, int] = {}
    risk_codes: Dict[str, int] = {}
    for item in items:
        by_class[str(item.get("classification"))] = by_class.get(str(item.get("classification")), 0) + 1
        by_ext[str(item.get("extension"))] = by_ext.get(str(item.get("extension")), 0) + 1
        for code in item.get("features", {}).get("source_audit_issue_codes") or []:
            risk_codes[str(code)] = risk_codes.get(str(code), 0) + 1
    return {
        "total_files": sum(by_class.values()),
        "by_classification": dict(sorted(by_class.items())),
        "by_extension": dict(sorted(by_ext.items())),
        "source_audit_issue_distribution": dict(sorted(risk_codes.items())),
    }


def _review_queue(items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    queue = []
    for item in items:
        reasons = list(item.get("reasons") or [])
        risk_codes = list(item.get("features", {}).get("source_audit_issue_codes") or [])
        if float(item.get("confidence") or 0) < 0.7 or item.get("classification") == "unsupported_or_conversion_needed" or risk_codes:
            queue.append(
                {
                    "relative_path": item.get("relative_path"),
                    "classification": item.get("classification"),
                    "confidence": item.get("confidence"),
                    "reasons": reasons[:8],
                    "source_audit_issue_codes": risk_codes[:12],
                    "next_action": _queue_next_action(str(item.get("classification")), reasons, risk_codes),
                }
            )
    queue.sort(key=lambda x: (str(x.get("classification")), float(x.get("confidence") or 0), str(x.get("relative_path"))))
    return queue


def _queue_next_action(classification: str, reasons: List[str], risk_codes: List[str]) -> str:
    if classification == "unsupported_or_conversion_needed":
        if any("legacy_format" in reason for reason in reasons):
            return "用 Word/WPS 另存为 DOCX 后再纳入测试矩阵。"
        if "scanned_or_textless_pdf" in reasons:
            return "改用 DOCX、可复制文字 PDF，或先 OCR 后再测试。"
        return "先确认文件能正常打开/读取，再决定转换或排除。"
    if risk_codes:
        return "进入高风险 E2E 样本池，运行 strict/visual QA 并人工复核相关页面。"
    return "人工复核分类；确认后再加入模板、正文或参照样本池。"


def _write_reports(output_dir: str, root: str, items: List[Dict[str, Any]], queue: List[Dict[str, Any]]) -> None:
    os.makedirs(output_dir, exist_ok=True)
    summary = _summary(items)
    payload = {
        "schema_version": 1,
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "corpus_root": os.path.abspath(root),
        "privacy": "local_only_do_not_commit_private_paths_or_reports",
        "summary": summary,
        "items": items,
    }
    with open(os.path.join(output_dir, "inventory.json"), "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    with open(os.path.join(output_dir, "review_queue.json"), "w", encoding="utf-8") as f:
        json.dump(
            {
                "schema_version": 1,
                "generated_at": payload["generated_at"],
                "privacy": payload["privacy"],
                "review_queue": queue,
            },
            f,
            ensure_ascii=False,
            indent=2,
        )
    lines = [
        "# Private Real-Data Inventory",
        "",
        "- Privacy: local only; do not commit this report or source files.",
        f"- Corpus root: `{os.path.abspath(root)}`",
        f"- Total files: {summary['total_files']}",
        "",
        "## By Classification",
    ]
    for key, value in summary["by_classification"].items():
        lines.append(f"- `{key}`: {value}")
    lines.extend(["", "## By Extension"])
    for key, value in summary["by_extension"].items():
        lines.append(f"- `{key}`: {value}")
    lines.extend(["", "## Source Risk Codes"])
    if summary["source_audit_issue_distribution"]:
        for key, value in summary["source_audit_issue_distribution"].items():
            lines.append(f"- `{key}`: {value}")
    else:
        lines.append("- None detected in readable DOCX files.")
    lines.extend(["", "## Review Queue", ""])
    for item in queue[:200]:
        codes = ", ".join(item.get("source_audit_issue_codes") or []) or "-"
        lines.append(
            f"- `{item.get('relative_path')}` -> `{item.get('classification')}` "
            f"(confidence={item.get('confidence')}, codes={codes})"
        )
        lines.append(f"  Next: {item.get('next_action')}")
    with open(os.path.join(output_dir, "inventory.md"), "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")


def audit_corpus(root: str, output_dir: str | None = None) -> Dict[str, Any]:
    """Scan a private corpus and write inventory/review reports."""
    root = os.path.abspath(root)
    output_dir = os.path.abspath(output_dir or _default_output_dir())
    items: List[Dict[str, Any]] = []
    for current, _dirs, files in os.walk(root):
        for name in sorted(files):
            path = os.path.join(current, name)
            try:
                items.append(_classify_file(path, root))
            except Exception as exc:
                items.append(
                    {
                        "relative_path": _safe_rel(path, root),
                        "extension": os.path.splitext(path)[1].lower() or "<no_ext>",
                        "size_bytes": os.path.getsize(path) if os.path.exists(path) else 0,
                        "classification": "unsupported_or_conversion_needed",
                        "confidence": 0.9,
                        "reasons": [f"audit_failed:{type(exc).__name__}"],
                        "features": {},
                    }
                )
    _mark_duplicates(items)
    queue = _review_queue(items)
    _write_reports(output_dir, root, items, queue)
    return {
        "output_dir": output_dir,
        "inventory_json": os.path.join(output_dir, "inventory.json"),
        "inventory_md": os.path.join(output_dir, "inventory.md"),
        "review_queue_json": os.path.join(output_dir, "review_queue.json"),
        "summary": _summary(items),
        "items": items,
        "review_queue_count": len(queue),
    }


def main(argv: List[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Audit a private real-data corpus without storing body text.")
    parser.add_argument("root", nargs="?", default=os.path.join("Templates", "20261"), help="Private corpus root to scan.")
    parser.add_argument("--output-dir", default=None, help="Local output directory for inventory reports.")
    args = parser.parse_args(argv)
    result = audit_corpus(args.root, args.output_dir)
    console_summary = {
        "output_dir": result.get("output_dir"),
        "inventory_json": result.get("inventory_json"),
        "inventory_md": result.get("inventory_md"),
        "review_queue_json": result.get("review_queue_json"),
        "summary": result.get("summary"),
        "review_queue_count": result.get("review_queue_count"),
    }
    print(json.dumps(console_summary, ensure_ascii=True, indent=2))
    return 0


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
