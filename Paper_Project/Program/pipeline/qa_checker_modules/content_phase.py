"""Content handoff checks for structural QA."""
from __future__ import annotations

import json
import os
from typing import Any, Callable, Dict

try:
    from privacy import sanitize_value
except ImportError:  # pragma: no cover - standalone fallback
    def sanitize_value(value: Any, project_root: str | None = None) -> Any:
        return value

try:
    from qa_checker_modules.metrics import (
        _content_text_chars,
        _content_toc_pollution_samples,
        _count_content_formulas,
        _count_content_images,
        _count_content_tables,
        _formula_number_conflict_samples,
        _fragmented_formula_samples,
        _load_json,
        _low_res_image_fragment_samples,
        _placeholder_samples_from_texts,
    )
except ImportError:  # pragma: no cover - package-style imports
    from .metrics import (
        _content_text_chars,
        _content_toc_pollution_samples,
        _count_content_formulas,
        _count_content_images,
        _count_content_tables,
        _formula_number_conflict_samples,
        _fragmented_formula_samples,
        _load_json,
        _low_res_image_fragment_samples,
        _placeholder_samples_from_texts,
    )

AddIssue = Callable[..., None]
def run_content_checks(out_dir: str, paths: Dict[str, str], counts: Dict[str, Any], manifest_counts: Dict[str, Any], add: AddIssue) -> Dict[str, Any]:
    content: Dict[str, Any] = {}
    if os.path.exists(paths["content"]):
        try:
            content = _load_json(paths["content"])
            counts["content_sections"] = len(content.get("sections") or [])
            counts["references"] = len(content.get("references") or [])
            counts["content_images"] = _count_content_images(content)
            counts["content_formulas"] = _count_content_formulas(content)
            counts["content_tables"] = _count_content_tables(content)
            counts["content_text_chars"] = _content_text_chars(content)
            meta = content.get("_meta") or {}
            missing_images = meta.get("missing_images") or []
            image_failures = meta.get("image_extract_failures") or []
            non_body_images = meta.get("non_body_images") or []
            counts["content_missing_images"] = len(missing_images)
            counts["content_image_extract_failures"] = len(image_failures)
            counts["content_non_body_images"] = len(non_body_images)
            remote_missing_images = [
                item for item in missing_images
                if isinstance(item, dict) and str(item.get("reason") or "").lower() == "remote"
            ]
            unreadable_images = [
                item for item in missing_images
                if isinstance(item, dict) and str(item.get("reason") or "").lower() == "unreadable"
            ]
            local_missing_images = [
                item for item in missing_images
                if not (
                    isinstance(item, dict)
                    and str(item.get("reason") or "").lower() in {"remote", "unreadable"}
                )
            ]
            if remote_missing_images:
                add(
                    "CONTENT_IMAGE_REMOTE_UNSUPPORTED",
                    "error",
                    "Markdown 引用了远程图片 URL，流水线不会自动联网下载图片。",
                    json.dumps(sanitize_value(remote_missing_images[:5], os.getcwd()), ensure_ascii=False),
                )
            if unreadable_images:
                add(
                    "CONTENT_IMAGE_UNREADABLE",
                    "error",
                    "Markdown 图片引用存在，但无法作为可渲染图片读取。",
                    json.dumps(sanitize_value(unreadable_images[:5], os.getcwd()), ensure_ascii=False),
                )
            if local_missing_images:
                add(
                    "CONTENT_IMAGE_MISSING",
                    "error",
                    "内容中存在未能解析或复制的图片引用，最终文档会丢图。",
                    json.dumps(sanitize_value(local_missing_images[:5], os.getcwd()), ensure_ascii=False),
                )
            if image_failures:
                add(
                    "IMAGE_EXTRACT_FAILED",
                    "error",
                    "DOCX 图片关系读取失败，最终文档可能缺图。",
                    json.dumps(sanitize_value(image_failures[:5], os.getcwd()), ensure_ascii=False),
                )
            if non_body_images:
                add(
                    "NON_BODY_IMAGE_UNSUPPORTED",
                    "error",
                    "源 DOCX 页眉/页脚中存在图片，当前不会作为正文图片渲染。",
                    json.dumps(sanitize_value(non_body_images[:5], os.getcwd()), ensure_ascii=False),
                )
            low_res_fragments = _low_res_image_fragment_samples(content, out_dir)
            if low_res_fragments:
                contained_fragments = int(manifest_counts.get("content_image_fragments_contained") or 0)
                contained = contained_fragments >= len(low_res_fragments)
                add(
                    "LOW_RES_IMAGE_FRAGMENT",
                    "warning" if contained else "error",
                    "内容中存在疑似图表/公式碎片的低分辨率图片；已按原始尺寸保留，避免放大。" if contained else "内容中存在疑似图表/公式碎片的低分辨率图片，可能形成模糊文字图。",
                    " / ".join(low_res_fragments),
                )
            if not content.get("sections"):
                add("CONTENT_EMPTY", "error", "内容提取结果没有正文 section。")
            toc_pollution = _content_toc_pollution_samples(content)
            if toc_pollution:
                add(
                    "CONTENT_TOC_POLLUTION",
                    "error",
                    "内容提取结果疑似混入源文档目录页，正文长句/公式碎片被当成标题。",
                    " / ".join(toc_pollution),
                )
            source_placeholders = []
            meta = content.get("_meta") or {}
            for item in meta.get("source_placeholders") or []:
                if isinstance(item, dict):
                    source_placeholders.append(str(item.get("text") or ""))
                else:
                    source_placeholders.append(str(item))
            content_texts = []
            title_info = content.get("title_info") or {}
            content_texts.extend(str(v or "") for v in title_info.values())
            for sec in content.get("sections") or []:
                content_texts.append(str(sec.get("heading") or ""))
                for item in sec.get("paragraphs") or []:
                    if isinstance(item, str):
                        content_texts.append(item)
                    elif isinstance(item, dict):
                        content_texts.append(str(item.get("text") or item.get("code") or ""))
            placeholder_samples = source_placeholders[:8] or _placeholder_samples_from_texts(content_texts)
            if placeholder_samples:
                placeholders_auto_removed = int(meta.get("source_placeholders_auto_removed") or 0)
                add(
                    "UNFILLED_PLACEHOLDER_TEXT",
                    "warning" if placeholders_auto_removed else "error",
                    "输入内容中存在未填写占位符，流水线已尽量从内容流中过滤。" if placeholders_auto_removed else "输入内容或提取结果中存在未填写的模板占位符。",
                    " / ".join(x[:120] for x in placeholder_samples[:8] if x),
                )
            formula_conflicts = _formula_number_conflict_samples(content)
            if formula_conflicts:
                add(
                    "FORMULA_NUMBER_CONFLICT",
                    "error",
                    "内容公式中存在重复或冲突编号，可能导致最终公式编号叠加。",
                    " / ".join(formula_conflicts),
                )
            fragmented_formulas = _fragmented_formula_samples(content)
            if fragmented_formulas:
                add(
                    "FORMULA_TEXT_FRAGMENTED",
                    "warning",
                    "内容中存在疑似碎裂公式文本；不安全公式已降级为普通正文，建议人工核对或改用原生公式/LaTeX。",
                    " / ".join(fragmented_formulas),
                )
            title_info = content.get("title_info") or {}
            if not any(str(v or "").strip() for v in title_info.values()):
                add("TITLE_MISSING", "warning", "未识别到论文标题信息，封面和标题页可能需要人工核对。")
            if not content.get("references"):
                add("REFERENCES_MISSING", "warning", "未识别到参考文献。")
        except Exception as exc:
            add("MISSING_CONTENT_JSON", "error", "content.json 无法读取。", str(exc))
    return content

