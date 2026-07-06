"""Report rendering for visual QA."""
from __future__ import annotations

import json
import os
from typing import Any, Dict


def _has_warnings(issues: list[Dict[str, Any]]) -> bool:
    return any(item.get("severity") == "warning" for item in issues or [])


def _result_label(passed: bool, issues: list[Dict[str, Any]]) -> str:
    if not passed:
        return "未通过"
    if _has_warnings(issues):
        return "通过但有警告"
    return "通过"


def _append_path_list(lines: list[str], label: str, paths: Any, limit: int = 6) -> None:
    if isinstance(paths, str) and paths.strip():
        lines.append(f"- {label}: `{paths}`")
        return
    if not isinstance(paths, list) or not paths:
        return
    shown = [str(item) for item in paths[:limit] if str(item).strip()]
    if not shown:
        return
    lines.append(f"- {label}: `{len(paths)}` 个")
    for item in shown:
        lines.append(f"  - `{item}`")
    if len(paths) > len(shown):
        lines.append(f"  - 另有 `{len(paths) - len(shown)}` 个，见对应目录。")


def _append_artifacts(lines: list[str], artifacts: Dict[str, Any]) -> None:
    artifact_lines: list[str] = []
    _append_path_list(artifact_lines, "Word PDF", artifacts.get("pdf"))
    _append_path_list(artifact_lines, "Word 文本诊断", artifacts.get("word_text") or artifacts.get("rendered_text"))
    _append_path_list(artifact_lines, "WPS PDF", artifacts.get("wps_pdf"))
    _append_path_list(artifact_lines, "WPS 文本诊断", artifacts.get("wps_text"))
    _append_path_list(artifact_lines, "Word 样张 PNG", artifacts.get("samples"), limit=8)
    _append_path_list(artifact_lines, "WPS 样张 PNG", artifacts.get("wps_samples"), limit=8)
    _append_path_list(artifact_lines, "全页 PNG", artifacts.get("all_pages"), limit=3)
    if artifact_lines:
        lines.extend(["", "## 诊断产物", "", *artifact_lines])


def report_to_markdown(report: Dict[str, Any]) -> str:
    issues = report.get("issues") or []
    lines = [
        "# 视觉 QA 报告",
        "",
        f"- 结果：{_result_label(bool(report.get('passed')), issues)}",
        f"- 输出目录：`{report.get('output_dir_name')}`",
        f"- 下一步：{report.get('next_action') or '打开最终 DOCX 和 PDF/PNG 渲染样张，检查视觉问题。'}",
        "",
        "## 统计",
        "",
    ]
    for key, value in sorted((report.get("counts") or {}).items()):
        lines.append(f"- `{key}`: {value}")
    artifacts = report.get("artifacts") or {}
    _append_artifacts(lines, artifacts)
    golden = artifacts.get("golden_baseline") or {}
    if golden:
        lines.extend(["", "## 黄金基线", ""])
        lines.append(f"- `status`: {golden.get('status')}")
        if golden.get("key"):
            lines.append(f"- `key`: {golden.get('key')}")
        if golden.get("path"):
            lines.append(f"- `path`: {golden.get('path')}")
        for issue in golden.get("issues") or []:
            lines.append(f"- `issue`: {issue}")
    lines.extend(["", "## 问题", ""])
    if not issues:
        lines.append("- 自动视觉检查未发现问题。")
    else:
        for item in issues:
            lines.append(f"- **{item.get('severity')}** `{item.get('code')}`: {item.get('message')}")
            if item.get("detail"):
                lines.append(f"  细节：`{item.get('detail')}`")
    lines.append("")
    return "\n".join(lines)


def write_reports(report: Dict[str, Any], out_dir: str) -> None:
    with open(os.path.join(out_dir, "visual_report.json"), "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)
    with open(os.path.join(out_dir, "visual_report.md"), "w", encoding="utf-8") as f:
        f.write(report_to_markdown(report))
