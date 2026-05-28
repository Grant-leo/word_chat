"""Final DOCX XML checks for structural QA."""
from __future__ import annotations

import os
import re
from typing import Any, Callable, Dict

try:
    from qa_checker_modules.metrics import (
        _duplicate_front_matter_headings,
        _missing_heading_samples,
        _placeholder_samples_from_texts,
        _read_docx_xml,
        _xml_paragraph_texts,
        _xml_plain_text,
    )
except ImportError:  # pragma: no cover - package-style imports
    from .metrics import (
        _duplicate_front_matter_headings,
        _missing_heading_samples,
        _placeholder_samples_from_texts,
        _read_docx_xml,
        _xml_paragraph_texts,
        _xml_plain_text,
    )

AddIssue = Callable[..., None]
def run_docx_checks(paths: Dict[str, str], counts: Dict[str, Any], content: Dict[str, Any], manifest_counts: Dict[str, Any], add: AddIssue) -> None:
    if os.path.exists(paths["docx"]):
        try:
            xml = _read_docx_xml(paths["docx"])
            plain = _xml_plain_text(xml)
            counts["docx_oMathPara"] = len(re.findall(r"<(?:[A-Za-z_][\w.-]*:)?oMathPara\b", xml))
            counts["docx_oMath"] = len(re.findall(r"<(?:[A-Za-z_][\w.-]*:)?oMath\b", xml))
            counts["docx_drawings"] = len(re.findall(r"<wp:(?:inline|anchor)\b", xml))
            counts["docx_tables"] = len(re.findall(r"<w:tbl\b", xml))
            counts["docx_text_chars"] = len(plain)

            if "[LaTeX error" in plain or "[LaTeX error" in xml:
                add("LATEX_ERROR_TEXT", "error", "最终文档中仍包含 LaTeX 转换错误占位。")
            if re.search(r"\$\$[^$]{2,}\$\$|\$[^\n$]{2,}\$", plain):
                add("LATEX_DELIMITER_TEXT", "error", "最终文档中仍残留 LaTeX 公式分隔符，可能有公式未转换。")
            if "M|b|p|s" in plain or "M|b|p|s" in xml:
                add("FORMULA_PIPE_ARTIFACT", "error", "公式出现 run 分隔伪影，例如 M|b|p|s。")
            rendered_formulas = int(manifest_counts["content_formulas_rendered"]) if "content_formulas_rendered" in manifest_counts else int(counts.get("docx_oMath", 0) or 0)
            rendered_images = int(manifest_counts["content_images_rendered"]) if "content_images_rendered" in manifest_counts else int(counts.get("docx_drawings", 0) or 0)
            rendered_tables = int(manifest_counts["content_tables_rendered"]) if "content_tables_rendered" in manifest_counts else int(counts.get("docx_tables", 0) or 0)
            if counts.get("content_formulas", 0) and counts.get("docx_oMath", 0) == 0:
                add("FORMULA_NOT_NATIVE", "error", "内容中有公式，但最终 docx 未检测到原生 OOXML Math。")
            if counts.get("content_formulas", 0) and rendered_formulas < counts.get("content_formulas", 0):
                add(
                    "FORMULA_COUNT_MISMATCH",
                    "warning",
                    "最终 docx 中的原生公式数量少于内容提取数量，可能有公式被丢失或转成普通文本。",
                    f"content={counts.get('content_formulas')} rendered={rendered_formulas} docx={counts.get('docx_oMath')}",
                )
            if counts.get("content_images", 0) and counts.get("docx_drawings", 0) == 0:
                add("IMAGE_NOT_RENDERED", "error", "内容中有图片，但最终 docx 未检测到 drawing。")
            if counts.get("content_images", 0) and rendered_images < counts.get("content_images", 0):
                add(
                    "IMAGE_COUNT_MISMATCH",
                    "error",
                    "最终 docx 中的图片数量少于内容提取数量，可能有图片未插入。",
                    f"content={counts.get('content_images')} rendered={rendered_images} docx={counts.get('docx_drawings')}",
                )
            if counts.get("content_tables", 0) and rendered_tables < counts.get("content_tables", 0):
                add(
                    "TABLE_COUNT_MISMATCH",
                    "warning",
                    "最终 docx 中的表格数量少于内容提取数量，可能有表格未渲染。",
                    f"content={counts.get('content_tables')} rendered={rendered_tables} docx={counts.get('docx_tables')}",
                )
            if counts.get("content_text_chars", 0) > 200 and counts.get("docx_text_chars", 0) < counts.get("content_text_chars", 0) * 0.6:
                add(
                    "DOCX_TEXT_TOO_SHORT",
                    "error",
                    "最终 docx 文本量明显少于提取内容，可能发生正文丢失。",
                    f"content={counts.get('content_text_chars')} docx={counts.get('docx_text_chars')}",
                )
            missing_headings = _missing_heading_samples(content, plain)
            if missing_headings:
                add("CONTENT_HEADING_MISSING", "warning", "部分内容标题没有出现在最终 docx 中。", " / ".join(missing_headings))
            duplicate_front = _duplicate_front_matter_headings(content, xml)
            if duplicate_front:
                add(
                    "DUPLICATE_FRONT_MATTER_HEADING",
                    "error",
                    "最终 docx 中检测到重复的摘要/关键词等前置章节标题。",
                    " / ".join(duplicate_front),
                )
            final_placeholders = _placeholder_samples_from_texts(_xml_paragraph_texts(xml))
            if final_placeholders:
                add("PLACEHOLDER_TEXT_LEFT", "error", "最终 docx 中残留模板占位符或待补全文本。", " / ".join(final_placeholders[:8]))
            if re.search(r"Error!\s*(Reference source not found|Bookmark not defined)|错误！未找到", plain, re.I):
                add("WORD_FIELD_ERROR", "warning", "最终 docx 中可能存在 Word 域错误文本。")
            plain_compact = re.sub(r"\s+", "", plain)
            has_toc_text = "目录" in plain_compact or "Contents" in plain
            has_toc_field = r"TOC \o" in xml or r"TOC\\o" in xml
            if len(content.get("sections") or []) >= 3 and not (has_toc_text or has_toc_field):
                add("TOC_MISSING", "warning", "最终文档中未检测到目录文本。")
        except Exception as exc:
            add("DOCX_XML_UNREADABLE", "error", "最终 docx 无法读取 document.xml。", str(exc))

