"""Artifact-writing helpers for pipeline JSON and markdown handoffs."""
from __future__ import annotations

import json
import os


def _sanitize_report_value(value, project_root=None):
    try:
        from privacy import sanitize_value
    except Exception:  # pragma: no cover - report hardening fallback
        return value
    try:
        return sanitize_value(value, project_root=project_root)
    except Exception:
        return value


def _truncate_detail(value, limit=4000):
    text = str(value or "").strip()
    if len(text) <= limit:
        return text
    return text[:limit] + "\n...[truncated]"


def write_extraction_failure_report(out_dir, *, mode, label, error, target):
    """Write a QA-shaped report when extraction verification cannot converge."""
    try:
        from qa_checker_modules.repair import build_repair_plan
        from qa_checker_modules.reports import write_reports
    except ImportError:  # pragma: no cover - package-style imports
        from ..qa_checker_modules.repair import build_repair_plan
        from ..qa_checker_modules.reports import write_reports

    label_text = str(label or "Extraction").strip() or "Extraction"
    issue = {
        "code": "EXTRACTION_VERIFICATION_FAILED",
        "severity": "error",
        "message": f"{label_text} extraction could not be verified across repeated runs.",
        "detail": str(error or "")[:2000],
        "active_owner": target,
        "owner_user": "User input/template file",
        "owner_developer": target,
    }
    report = {
        "schema_version": 1,
        "mode": mode,
        "passed": False,
        "counts": {},
        "issues": [issue],
        "output_dir_name": os.path.basename(os.path.abspath(out_dir)),
        "next_action": "查看 qa_repair_plan.md；优先核对输入文件是否稳定可提取，开发者再检查对应提取器。",
    }
    report["repair_plan"] = build_repair_plan(report, out_dir)
    report["repair_plan"]["open_first"] = [
        "qa_repair_plan.md",
        "qa_report.md",
        "workflow_mode.json",
    ]
    report["repair_plan"].setdefault("commands", {})["rebuild_current_docx"] = ""
    write_reports(report, out_dir)
    return report


def write_build_failure_report(
    out_dir,
    *,
    mode,
    stderr,
    stdout="",
    output_docx_name="最终论文.docx",
    project_root=None,
):
    """Write a QA-shaped report when build_generated.py fails before QA."""
    try:
        from qa_checker_modules.repair import build_repair_plan
        from qa_checker_modules.reports import write_reports
    except ImportError:  # pragma: no cover - package-style imports
        from ..qa_checker_modules.repair import build_repair_plan
        from ..qa_checker_modules.reports import write_reports

    detail_parts = []
    if stderr:
        detail_parts.append("stderr:\n" + str(stderr))
    if stdout:
        detail_parts.append("stdout:\n" + str(stdout))
    detail = _sanitize_report_value(_truncate_detail("\n\n".join(detail_parts)), project_root)
    active_owner = "Outputs/<run>/build_generated.py" if mode == "user" else "script_generator.py / script_generator_modules/runtime_build.py"
    issue = {
        "code": "MISSING_DOCX",
        "severity": "error",
        "message": f"生成脚本执行失败，`{output_docx_name}` 没有生成。",
        "detail": detail,
        "active_owner": active_owner,
        "owner_user": "Outputs/<run>/build_generated.py",
        "owner_developer": "script_generator.py / script_generator_modules/runtime_build.py",
    }
    report = {
        "schema_version": 1,
        "mode": mode,
        "passed": False,
        "counts": {},
        "issues": [issue],
        "output_dir_name": os.path.basename(os.path.abspath(out_dir)),
    }
    report["repair_plan"] = build_repair_plan(report, out_dir)
    report["repair_plan"]["open_first"] = [
        "qa_repair_plan.md",
        "qa_report.md",
        "build_generated.py",
        "workflow_mode.json",
        "格式提取.md",
        "内容提取.md",
    ]
    report["next_action"] = report["repair_plan"].get("next_action") or "先打开 build_generated.py 查看构建错误，修复后重建当前 DOCX。"
    write_reports(report, out_dir)
    return report


def write_format_blocker_report_if_needed(fmt_json_path, out_dir, *, mode):
    """Fail closed on template-format blockers before content extraction/build."""
    try:
        from qa_checker_modules.format_phase import run_format_checks
        from qa_checker_modules.registry import OWNER_BY_CODE, REPAIR_GUIDES
        from qa_checker_modules.report_phase import build_report
        from qa_checker_modules.reports import write_reports
    except ImportError:  # pragma: no cover - package-style imports
        from ..qa_checker_modules.format_phase import run_format_checks
        from ..qa_checker_modules.registry import OWNER_BY_CODE, REPAIR_GUIDES
        from ..qa_checker_modules.report_phase import build_report
        from ..qa_checker_modules.reports import write_reports

    issues = []
    counts = {}

    def add(code, severity, message, detail=""):
        guide = REPAIR_GUIDES.get(code) or {}
        auto_level = str(guide.get("auto_level") or "")
        if auto_level == "needs_environment":
            owner_user = "Local environment / Poppler"
        elif auto_level.startswith("needs_user"):
            owner_user = "User input/template file"
        else:
            owner_user = "Outputs/<run>/build_generated.py"
        owner_developer = OWNER_BY_CODE.get(code, "format_extractor.py")
        issues.append(
            {
                "code": code,
                "severity": severity,
                "message": message,
                "detail": detail,
                "owner_user": owner_user,
                "owner_developer": owner_developer,
                "active_owner": owner_user if mode == "user" else owner_developer,
            }
        )

    run_format_checks({"format": fmt_json_path}, counts, add)
    blocking_codes = {"PDF_TEMPLATE_UNSUPPORTED", "PDF_TEMPLATE_DEPENDENCY_MISSING"}
    if not any(item.get("severity") == "error" and item.get("code") in blocking_codes for item in issues):
        return None

    report = build_report(out_dir, mode, counts, issues)
    report["repair_plan"]["open_first"] = [
        "qa_repair_plan.md",
        "qa_report.md",
        "格式提取.md",
        "template_profile.md",
        "workflow_mode.json",
    ]
    report["repair_plan"].setdefault("commands", {})["rebuild_current_docx"] = ""
    write_reports(report, out_dir)
    return report


def _math_count(item):
    count = len(item.get("math") or [])
    if count:
        return count
    count = 0
    for run in item.get("runs") or []:
        if isinstance(run, dict):
            count += len(run.get("math") or [])
    return count


def _table_shape(rows):
    rows = rows or []
    cols = max((len(row or []) for row in rows), default=0)
    return len(rows), cols


def _paragraph_summary(paragraph):
    if not isinstance(paragraph, dict):
        return str(paragraph)

    role = str(paragraph.get("role") or paragraph.get("type") or "")
    if role in {"figure", "image"} or paragraph.get("image") or paragraph.get("filename") or paragraph.get("asset"):
        image = paragraph.get("image") or paragraph.get("path") or ""
        if not image:
            image = paragraph.get("filename") or paragraph.get("asset") or ""
        caption = paragraph.get("caption") or ""
        text = f"[图片] {image}".strip()
        return f"{text} — {caption}" if caption else text
    if paragraph.get("table_rows") or role == "table":
        rows, cols = _table_shape(paragraph.get("table_rows") or [])
        return f"[表格] {rows}行 x {cols}列"
    if role in {"figure_caption", "table_caption"}:
        return paragraph.get("text") or "[题注]"
    if role == "formula_problem":
        return f"[公式问题] {paragraph.get('text') or paragraph.get('latex') or ''}".strip()
    if role == "formula" or paragraph.get("latex") or paragraph.get("xml"):
        return paragraph.get("text") or paragraph.get("latex") or "[公式]"
    if role == "code":
        return f"[代码] {paragraph.get('code') or paragraph.get('text') or ''}".strip()

    text = paragraph.get("text") or paragraph.get("code") or "[结构化内容]"
    count = _math_count(paragraph)
    if count:
        text += f" (+{count}公式)"
    return text


def write_format_artifacts(fmt, md_text, out_dir):
    fmt_json_path = os.path.join(out_dir, "format.json")
    fmt_md_path = os.path.join(out_dir, "格式提取.md")
    with open(fmt_json_path, "w", encoding="utf-8") as f:
        json.dump(fmt, f, ensure_ascii=False, indent=2)
    with open(fmt_md_path, "w", encoding="utf-8") as f:
        f.write(md_text)
    return fmt_json_path, fmt_md_path


def build_content_markdown(content, content_path):
    lines = [f"# 内容提取 — {os.path.basename(content_path)}\n"]
    for sec in content.get("sections", []):
        lines.append(f'## {sec.get("heading", "")}\n')
        inline_images = {
            str(paragraph.get("image") or paragraph.get("filename") or paragraph.get("asset") or "")
            for paragraph in sec.get("paragraphs", [])
            if isinstance(paragraph, dict) and (paragraph.get("role") in {"figure", "image"} or paragraph.get("image"))
        }
        for img in sec.get("images", []):
            if str(img) not in inline_images:
                lines.append(f"- [图片] {img}")
        for paragraph in sec.get("paragraphs", []):
            text = _paragraph_summary(paragraph)
            text = text[:120] + "..." if len(text) > 120 else text
            lines.append(f"- {text}")
        lines.append("")
    if content.get("references"):
        lines.append("## 参考文献\n")
        for ref in content["references"]:
            if isinstance(ref, dict):
                text = ref.get("text") or ref.get("code") or "[结构化内容]"
            else:
                text = str(ref)
            lines.append(f"- {text[:120]}")
    return "\n".join(lines)


def write_content_artifacts(content, out_dir, content_path):
    cnt_json_path = os.path.join(out_dir, "content.json")
    cnt_md_path = os.path.join(out_dir, "内容提取.md")
    with open(cnt_json_path, "w", encoding="utf-8") as f:
        json.dump(content, f, ensure_ascii=False, indent=2)
    with open(cnt_md_path, "w", encoding="utf-8") as f:
        f.write(build_content_markdown(content, content_path))
    return cnt_json_path, cnt_md_path
