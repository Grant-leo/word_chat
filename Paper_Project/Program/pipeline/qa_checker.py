"""
qa_checker.py - lightweight QA checks for generated Word pipeline outputs.

The checker does not fix files. It writes a structured report that tells the AI
which artifact should be edited in user mode and which core engine owns the same
class of issue in developer mode.
"""
from __future__ import annotations

import json
import os
import re
import zipfile
from datetime import datetime
from typing import Any, Dict, Iterable, List

try:
    from privacy import sanitize_value
except ImportError:  # pragma: no cover - standalone fallback
    def sanitize_value(value: Any, project_root: str | None = None) -> Any:
        return value


VALID_MODES = {"user", "developer"}

OWNER_BY_CODE = {
    "MISSING_DOCX": "script_generator.py",
    "MISSING_BUILD_SCRIPT": "script_generator.py",
    "MISSING_FORMAT_JSON": "format_extractor.py / md_parser.py",
    "MISSING_CONTENT_JSON": "content_parser.py / md_parser.py",
    "FORMAT_EMPTY": "format_extractor.py / md_parser.py",
    "CONTENT_EMPTY": "content_parser.py / md_parser.py",
    "STYLE_PROFILE_MISSING": "format_extractor.py / script_generator.py",
    "COVER_NOT_EXTRACTED": "format_extractor.py / script_generator.py",
    "TITLE_MISSING": "content_parser.py / md_parser.py",
    "REFERENCES_MISSING": "content_parser.py / md_parser.py",
    "DOCX_XML_UNREADABLE": "script_generator.py",
    "LATEX_ERROR_TEXT": "latex_omath.py / script_generator.py",
    "FORMULA_PIPE_ARTIFACT": "latex_omath.py",
    "FORMULA_COUNT_MISMATCH": "content_parser.py / md_parser.py / script_generator.py / latex_omath.py",
    "FORMULA_NOT_NATIVE": "content_parser.py / script_generator.py / latex_omath.py",
    "IMAGE_NOT_RENDERED": "content_parser.py / script_generator.py",
    "IMAGE_COUNT_MISMATCH": "content_parser.py / script_generator.py",
    "CONTENT_IMAGE_MISSING": "md_parser.py / content_parser.py",
    "IMAGE_EXTRACT_FAILED": "content_parser.py",
    "TABLE_COUNT_MISMATCH": "content_parser.py / md_parser.py / script_generator.py",
    "BUILD_MANIFEST_MISSING": "script_generator.py",
    "DOCX_TEXT_TOO_SHORT": "content_parser.py / md_parser.py / script_generator.py",
    "CONTENT_HEADING_MISSING": "content_parser.py / md_parser.py / script_generator.py",
    "PLACEHOLDER_TEXT_LEFT": "script_generator.py",
    "WORD_FIELD_ERROR": "script_generator.py",
    "TOC_MISSING": "script_generator.py",
    "WORKFLOW_MODE_INVALID": "run_pipeline.py",
}


def _load_json(path: str) -> Dict[str, Any]:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def _load_manifest_counts(out_dir: str) -> Dict[str, Any]:
    path = os.path.join(out_dir, "build_manifest.json")
    if not os.path.exists(path):
        return {}
    data = _load_json(path)
    return data.get("counts") or {}


def _read_docx_xml(docx_path: str) -> str:
    with zipfile.ZipFile(docx_path) as zf:
        return zf.read("word/document.xml").decode("utf-8", errors="replace")


def _xml_plain_text(xml: str) -> str:
    texts = re.findall(r"<w:t[^>]*>(.*?)</w:t>", xml, flags=re.S)
    return "".join(re.sub(r"<[^>]+>", "", t) for t in texts)


def _iter_paragraph_items(content: Dict[str, Any]) -> Iterable[Any]:
    for sec in content.get("sections") or []:
        for item in sec.get("paragraphs") or []:
            yield item


def _count_content_formulas(content: Dict[str, Any]) -> int:
    total = 0
    for item in _iter_paragraph_items(content):
        if not isinstance(item, dict):
            continue
        math_items = item.get("math") or []
        if math_items:
            total += len(math_items)
        elif item.get("role") == "formula" or item.get("latex"):
            total += 1
    return total


def _count_content_tables(content: Dict[str, Any]) -> int:
    total = 0
    saw_table_rows = False
    for item in _iter_paragraph_items(content):
        if isinstance(item, dict) and item.get("table_rows"):
            saw_table_rows = True
            if item.get("role") != "code":
                total += 1
    if saw_table_rows:
        return total
    return int((content.get("_meta") or {}).get("tables_count") or 0)


def _count_content_images(content: Dict[str, Any]) -> int:
    inline_total = 0
    inline_names: List[str] = []
    section_total = 0
    section_names: List[str] = []
    for sec in content.get("sections") or []:
        section_images = [str(x or "") for x in (sec.get("images") or [])]
        section_total += len(section_images)
        section_names.extend(section_images)
        for item in sec.get("paragraphs") or []:
            if not isinstance(item, dict):
                continue
            if item.get("role") in ("image", "figure") and (item.get("image") or item.get("filename") or item.get("asset")):
                inline_total += 1
                inline_names.append(str(item.get("image") or item.get("filename") or item.get("asset") or ""))
    if inline_total:
        extra_section_only = [name for name in section_names if name and name not in inline_names]
        return inline_total + len(extra_section_only)
    if section_total:
        return section_total
    return int((content.get("_meta") or {}).get("images_extracted") or 0)


def _content_text_chars(content: Dict[str, Any]) -> int:
    parts: List[str] = []
    title_info = content.get("title_info") or {}
    parts.extend(str(v or "") for v in title_info.values())
    for sec in content.get("sections") or []:
        parts.append(str(sec.get("heading") or ""))
        for item in sec.get("paragraphs") or []:
            if isinstance(item, dict):
                parts.append(str(item.get("text") or item.get("code") or ""))
                for row in item.get("table_rows") or []:
                    parts.extend(str(cell or "") for cell in row)
            else:
                parts.append(str(item or ""))
    for ref in content.get("references") or []:
        if isinstance(ref, dict):
            parts.append(str(ref.get("text") or ref.get("code") or ""))
        else:
            parts.append(str(ref or ""))
    return sum(len(p.strip()) for p in parts if p and p.strip())


def _compact_text(text: str) -> str:
    return re.sub(r"\s+", "", str(text or "")).lower()


def _missing_heading_samples(content: Dict[str, Any], plain_text: str, limit: int = 6) -> List[str]:
    compact = _compact_text(plain_text)
    missing: List[str] = []
    for sec in content.get("sections") or []:
        heading = str(sec.get("heading") or "").strip()
        if not heading or heading == "正文" or len(heading) > 80:
            continue
        sample = _compact_text(heading)
        if sample and sample not in compact:
            missing.append(heading)
        if len(missing) >= limit:
            break
    return missing


def _issue(code: str, severity: str, message: str, mode: str, detail: str = "") -> Dict[str, Any]:
    owner_dev = OWNER_BY_CODE.get(code, "script_generator.py")
    return {
        "code": code,
        "severity": severity,
        "message": message,
        "detail": detail,
        "owner_user": "Outputs/<run>/build_generated.py",
        "owner_developer": owner_dev,
        "active_owner": "Outputs/<run>/build_generated.py" if mode == "user" else owner_dev,
    }


def check_output(out_dir: str, mode: str = "user", output_docx_name: str = "最终论文.docx") -> Dict[str, Any]:
    mode = mode if mode in VALID_MODES else "user"
    issues: List[Dict[str, Any]] = []
    counts: Dict[str, Any] = {}

    def add(code: str, severity: str, message: str, detail: str = "") -> None:
        issues.append(_issue(code, severity, message, mode, detail))

    paths = {
        "docx": os.path.join(out_dir, output_docx_name),
        "build": os.path.join(out_dir, "build_generated.py"),
        "format": os.path.join(out_dir, "format.json"),
        "content": os.path.join(out_dir, "content.json"),
        "workflow": os.path.join(out_dir, "workflow_mode.json"),
        "manifest": os.path.join(out_dir, "build_manifest.json"),
    }

    if not os.path.exists(paths["workflow"]):
        add("WORKFLOW_MODE_INVALID", "warning", "未找到 workflow_mode.json，无法确认本轮应按用户模式还是开发者模式修复。")
    else:
        try:
            workflow = _load_json(paths["workflow"])
            if workflow.get("mode") not in VALID_MODES:
                add("WORKFLOW_MODE_INVALID", "warning", "workflow_mode.json 中的 mode 无效。", str(workflow.get("mode")))
        except Exception as exc:
            add("WORKFLOW_MODE_INVALID", "warning", "workflow_mode.json 无法读取。", str(exc))

    for key, code, message in [
        ("docx", "MISSING_DOCX", "缺少最终 docx。"),
        ("build", "MISSING_BUILD_SCRIPT", "缺少 build_generated.py。"),
        ("format", "MISSING_FORMAT_JSON", "缺少 format.json。"),
        ("content", "MISSING_CONTENT_JSON", "缺少 content.json。"),
    ]:
        if not os.path.exists(paths[key]):
            add(code, "error", message, paths[key])

    manifest_counts: Dict[str, Any] = {}
    if os.path.exists(paths["manifest"]):
        try:
            manifest_counts = _load_manifest_counts(out_dir)
            for key, value in manifest_counts.items():
                counts[f"manifest_{key}"] = value
        except Exception as exc:
            add("BUILD_MANIFEST_MISSING", "warning", "build_manifest.json 无法读取，正文元素数量只能退回 XML 总量检测。", str(exc))
    elif os.path.exists(paths["docx"]):
        add("BUILD_MANIFEST_MISSING", "warning", "未找到 build_manifest.json，图片/表格数量检测可能被封面或页眉元素干扰。")

    fmt: Dict[str, Any] = {}
    content: Dict[str, Any] = {}
    if os.path.exists(paths["format"]):
        try:
            fmt = _load_json(paths["format"])
            counts["format_paragraphs"] = len(fmt.get("paragraphs") or [])
            counts["format_tables"] = len(fmt.get("tables") or [])
            counts["format_sections"] = len(fmt.get("sections") or [])
            counts["cover_elements"] = len(fmt.get("cover") or [])
            counts["style_profiles"] = len(fmt.get("style_profiles") or {})
            source = str((fmt.get("_meta") or {}).get("source") or "")
            is_md_format = source.lower().endswith(".md")
            if not fmt.get("sections"):
                add("FORMAT_EMPTY", "error", "格式提取结果没有 section。")
            if not fmt.get("paragraphs"):
                add("FORMAT_EMPTY", "warning", "格式提取结果没有 paragraph，可能大量使用默认格式。")
            expected = {"body", "h1", "h2", "h3"}
            missing = sorted(expected - set((fmt.get("style_profiles") or {}).keys()))
            if missing and not is_md_format:
                add("STYLE_PROFILE_MISSING", "warning", "关键样式 profile 不完整。", ", ".join(missing))
            if not fmt.get("cover") and not is_md_format:
                add("COVER_NOT_EXTRACTED", "warning", "没有提取到封面结构；纯 MD 或无封面模板时可以忽略。")
        except Exception as exc:
            add("MISSING_FORMAT_JSON", "error", "format.json 无法读取。", str(exc))

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
            counts["content_missing_images"] = len(missing_images)
            counts["content_image_extract_failures"] = len(image_failures)
            if missing_images:
                add(
                    "CONTENT_IMAGE_MISSING",
                    "error",
                    "内容中存在未能解析或复制的图片引用，最终文档会丢图。",
                    json.dumps(sanitize_value(missing_images[:5], os.getcwd()), ensure_ascii=False),
                )
            if image_failures:
                add(
                    "IMAGE_EXTRACT_FAILED",
                    "error",
                    "DOCX 图片关系读取失败，最终文档可能缺图。",
                    json.dumps(sanitize_value(image_failures[:5], os.getcwd()), ensure_ascii=False),
                )
            if not content.get("sections"):
                add("CONTENT_EMPTY", "error", "内容提取结果没有正文 section。")
            title_info = content.get("title_info") or {}
            if not any(str(v or "").strip() for v in title_info.values()):
                add("TITLE_MISSING", "warning", "未识别到论文标题信息，封面和标题页可能需要人工核对。")
            if not content.get("references"):
                add("REFERENCES_MISSING", "warning", "未识别到参考文献。")
        except Exception as exc:
            add("MISSING_CONTENT_JSON", "error", "content.json 无法读取。", str(exc))

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
            if re.search(r"(\{\{[^}]+\}\}|TODO|FIXME|待填写|待补充|请输入|XXXX|XX学院|某某)", plain, re.I):
                add("PLACEHOLDER_TEXT_LEFT", "warning", "最终 docx 中可能残留模板占位符或待补全文本。")
            if re.search(r"Error!\s*(Reference source not found|Bookmark not defined)|错误！未找到", plain, re.I):
                add("WORD_FIELD_ERROR", "warning", "最终 docx 中可能存在 Word 域错误文本。")
            plain_compact = re.sub(r"\s+", "", plain)
            has_toc_text = "目录" in plain_compact or "Contents" in plain
            has_toc_field = r"TOC \o" in xml or r"TOC\\o" in xml
            if len(content.get("sections") or []) >= 3 and not (has_toc_text or has_toc_field):
                add("TOC_MISSING", "warning", "最终文档中未检测到目录文本。")
        except Exception as exc:
            add("DOCX_XML_UNREADABLE", "error", "最终 docx 无法读取 document.xml。", str(exc))

    passed = not any(i["severity"] == "error" for i in issues)
    next_action = (
        "通过 QA。仍建议用 WPS/Word 做最终视觉核对。"
        if passed else
        ("用户模式：根据 active_owner 修改当前输出目录的 build_generated.py 后重跑该脚本。"
         if mode == "user" else
         "开发者模式：根据 active_owner 修改核心引擎脚本后重跑完整流水线。")
    )

    return {
        "schema_version": 1,
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "mode": mode,
        "output_dir_name": os.path.basename(os.path.abspath(out_dir)),
        "passed": passed,
        "counts": counts,
        "issues": issues,
        "next_action": next_action,
    }


def report_to_markdown(report: Dict[str, Any]) -> str:
    lines = [
        "# QA 检测报告",
        "",
        f"- 模式：`{report.get('mode')}`",
        f"- 结果：{'通过' if report.get('passed') else '未通过'}",
        f"- 输出目录：`{report.get('output_dir_name')}`",
        f"- 下一步：{report.get('next_action')}",
        "",
        "## 统计",
        "",
    ]
    counts = report.get("counts") or {}
    if counts:
        for key in sorted(counts):
            lines.append(f"- `{key}`: {counts[key]}")
    else:
        lines.append("- 无统计信息")

    lines.extend(["", "## 问题", ""])
    issues = report.get("issues") or []
    if not issues:
        lines.append("- 未发现结构性问题。")
    else:
        for item in issues:
            lines.append(
                f"- **{item.get('severity')}** `{item.get('code')}`：{item.get('message')} "
                f"修复目标：`{item.get('active_owner')}`"
            )
            if item.get("detail"):
                lines.append(f"  细节：`{item.get('detail')}`")
    lines.append("")
    return "\n".join(lines)


def write_reports(report: Dict[str, Any], out_dir: str) -> None:
    json_path = os.path.join(out_dir, "qa_report.json")
    md_path = os.path.join(out_dir, "qa_report.md")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)
    with open(md_path, "w", encoding="utf-8") as f:
        f.write(report_to_markdown(report))


def check_and_write(out_dir: str, mode: str = "user", output_docx_name: str = "最终论文.docx") -> Dict[str, Any]:
    report = check_output(out_dir, mode=mode, output_docx_name=output_docx_name)
    write_reports(report, out_dir)
    return report


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Check generated pipeline output.")
    parser.add_argument("out_dir")
    parser.add_argument("--mode", choices=sorted(VALID_MODES), default="user")
    parser.add_argument("--docx", default="最终论文.docx")
    args = parser.parse_args()

    result = check_and_write(args.out_dir, mode=args.mode, output_docx_name=args.docx)
    print(report_to_markdown(result))
    raise SystemExit(0 if result.get("passed") else 1)
