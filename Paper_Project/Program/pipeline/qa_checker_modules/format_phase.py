"""Format handoff checks for structural QA."""
from __future__ import annotations

import os
from typing import Any, Callable, Dict

try:
    from qa_checker_modules.metrics import _load_json
except ImportError:  # pragma: no cover - package-style imports
    from .metrics import _load_json

AddIssue = Callable[..., None]

PDF_INSTRUCTION_ROLE_LABELS = {
    "page": "页面/页边距",
    "body": "正文",
    "heading": "标题",
    "caption": "图表题注",
    "reference": "参考文献",
}


def _instruction_incomplete_detail(warnings: list[Any]) -> str:
    missing_roles: list[str] = []
    raw_warnings: list[str] = []
    for item in warnings:
        text = str(item)
        if not text.startswith("PDF_TEMPLATE_INSTRUCTION_INCOMPLETE"):
            continue
        raw_warnings.append(text)
        _, _, role_text = text.partition(":")
        for role in role_text.split(","):
            role = role.strip()
            if role and role not in missing_roles:
                missing_roles.append(role)
    labels = [PDF_INSTRUCTION_ROLE_LABELS.get(role, role) for role in missing_roles]
    if labels:
        return "缺少：" + "、".join(labels) + "。原始警告：" + "; ".join(raw_warnings)
    return "; ".join(raw_warnings)


def run_format_checks(paths: Dict[str, str], counts: Dict[str, Any], add: AddIssue) -> Dict[str, Any]:
    fmt: Dict[str, Any] = {}
    if os.path.exists(paths["format"]):
        try:
            fmt = _load_json(paths["format"])
            counts["format_paragraphs"] = len(fmt.get("paragraphs") or [])
            counts["format_tables"] = len(fmt.get("tables") or [])
            counts["format_sections"] = len(fmt.get("sections") or [])
            counts["cover_elements"] = len(fmt.get("cover") or [])
            counts["style_profiles"] = len(fmt.get("style_profiles") or {})
            meta = fmt.get("_meta") or {}
            source = str(meta.get("source") or "")
            is_md_format = source.lower().endswith(".md")
            pdf_meta = meta.get("pdf_template") or {}
            is_pdf_format = bool(pdf_meta) or source.lower().endswith(".pdf")
            if pdf_meta:
                counts["pdf_template_type"] = pdf_meta.get("type")
                counts["pdf_template_confidence"] = pdf_meta.get("confidence", 0)
                counts["pdf_template_text_chars"] = pdf_meta.get("text_chars", 0)
                pdf_errors = pdf_meta.get("errors") or []
                pdf_warnings = pdf_meta.get("warnings") or []
                missing_pdf_tools = [
                    str(item)
                    for item in pdf_warnings
                    if str(item).startswith(("PDFINFO_MISSING", "PDFTOTEXT_MISSING"))
                ]
                failed_pdf_reads = [
                    str(item)
                    for item in pdf_warnings
                    if str(item).startswith(("PDFINFO_FAILED", "PDFTOTEXT_FAILED"))
                ]
                instruction_incomplete = [
                    str(item)
                    for item in pdf_warnings
                    if str(item).startswith("PDF_TEMPLATE_INSTRUCTION_INCOMPLETE")
                ]
                if missing_pdf_tools:
                    add(
                        "PDF_TEMPLATE_DEPENDENCY_MISSING",
                        "error",
                        "PDF 模板解析缺少 Poppler 命令行工具。",
                        "; ".join(missing_pdf_tools),
                    )
                elif failed_pdf_reads and (
                    pdf_meta.get("type") == "scanned_or_unsupported_pdf"
                    or pdf_errors
                    or int(pdf_meta.get("text_chars") or 0) == 0
                ):
                    detail_parts = list(failed_pdf_reads)
                    detail_parts.extend(str(item) for item in pdf_errors)
                    add(
                        "PDF_TEMPLATE_READ_FAILED",
                        "error",
                        "PDF 模板文件无法被 Poppler 正常读取。",
                        "; ".join(detail_parts),
                    )
                elif pdf_meta.get("type") == "scanned_or_unsupported_pdf" or pdf_errors:
                    detail = "; ".join(str(item) for item in pdf_errors) or "PDF 模板没有可提取文字。"
                    add(
                        "PDF_TEMPLATE_UNSUPPORTED",
                        "error",
                        "PDF 模板无法可靠提取格式。",
                        detail,
                    )
                if instruction_incomplete:
                    add(
                        "PDF_TEMPLATE_INSTRUCTION_INCOMPLETE",
                        "warning",
                        "PDF 文字说明模板缺少关键格式规则。",
                        _instruction_incomplete_detail(instruction_incomplete),
                    )
                if pdf_meta.get("type") == "visual_sample_pdf" or pdf_warnings:
                    add(
                        "PDF_TEMPLATE_LIMITED_CONFIDENCE",
                        "warning",
                        "PDF 模板只能估计格式，不能像 DOCX 一样读取完整样式树。",
                        "; ".join(str(item) for item in pdf_warnings[:5]),
                    )
            if not fmt.get("sections"):
                add("FORMAT_EMPTY", "error", "格式提取结果没有 section。")
            if not fmt.get("paragraphs"):
                add("FORMAT_EMPTY", "warning", "格式提取结果没有 paragraph，可能大量使用默认格式。")
            expected = {"body", "h1", "h2", "h3"}
            missing = sorted(expected - set((fmt.get("style_profiles") or {}).keys()))
            if missing and not is_md_format:
                add("STYLE_PROFILE_MISSING", "warning", "关键样式 profile 不完整。", ", ".join(missing))
            if not fmt.get("cover") and not is_md_format and not is_pdf_format:
                add("COVER_NOT_EXTRACTED", "warning", "没有提取到封面结构；纯 MD 或无封面模板时可以忽略。")
        except Exception as exc:
            add("MISSING_FORMAT_JSON", "error", "format.json 无法读取。", str(exc))
    return fmt

