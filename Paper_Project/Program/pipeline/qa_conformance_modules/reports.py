"""Report builders and writers for strict conformance QA."""
from __future__ import annotations

import json
import os
from datetime import datetime
from typing import Any, Dict, List

try:
    from privacy import sanitize_value
except Exception:  # pragma: no cover
    def sanitize_value(value: Any, project_root: str | None = None) -> Any:
        return value


def _write_json(path: str, value: Dict[str, Any]) -> None:
    with open(path, "w", encoding="utf-8") as f:
        json.dump(value, f, ensure_ascii=False, indent=2)


def _next_action(mode: str, issues: List[Dict[str, Any]]) -> str:
    error_codes = {str(item.get("code") or "") for item in issues if item.get("severity") == "error"}
    if not error_codes:
        return "Strict conformance passed for machine-checkable template requirements."
    if error_codes & {"CONFORMANCE_INPUT_MISSING"}:
        return "重新运行完整流水线，确保 format.json、content.json、build_manifest.json 和最终 DOCX 都已生成。"
    if error_codes & {"DOCX_XML_UNREADABLE"}:
        return "先确认最终 DOCX 能被 Word/WPS 正常打开；若文件损坏，重新生成最终论文。"
    if error_codes & {"CONTENT_PARAGRAPH_MISSING"}:
        return "对照 content.json/内容提取.md 与最终 DOCX，修复遗漏段落；普通用户优先检查 build_generated.py 的正文遍历。"
    if error_codes & {"RENDER_COUNT_MISMATCH", "TABLE_NOT_FOUND", "FORMULA_COUNT_MISMATCH", "IMAGE_COUNT_MISMATCH"}:
        return "查看 build_manifest.json 的渲染数量，定位图片、表格或公式在哪个生成分支被跳过。"
    if error_codes & {"STYLE_MISMATCH", "PAGE_GEOMETRY_MISMATCH", "TABLE_BORDER_MISMATCH", "IMAGE_LAYOUT_MISMATCH"}:
        return "按 conformance_report.md 的 detail 修复样式/页边距/表格线/图片尺寸，修复后重跑 strict QA。"
    if error_codes & {"OMML_WPS_COMPAT", "FORMULA_ERROR_TEXT"}:
        return "检查公式渲染链路和 latex_omath.py，确保公式输出为 WPS 兼容的原生 OOXML Math。"
    return "Fix Outputs/<run>/build_generated.py and rerun it." if mode == "user" else "Fix core pipeline scripts and rerun the full pipeline."


def build_report(
    out_dir: str,
    mode: str,
    counts: Dict[str, Any],
    issues: List[Dict[str, Any]],
    project_root: str | None,
) -> Dict[str, Any]:
    return {
        "schema_version": 1,
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "mode": mode,
        "output_dir_name": os.path.basename(os.path.abspath(out_dir)),
        "passed": not any(i.get("severity") == "error" for i in issues),
        "counts": counts,
        "issues": sanitize_value(issues, project_root),
        "next_action": _next_action(mode, issues),
    }


def report_to_markdown(report: Dict[str, Any]) -> str:
    lines = [
        "# Conformance QA Report",
        "",
        f"- Result: {'passed' if report.get('passed') else 'failed'}",
        f"- Mode: `{report.get('mode')}`",
        f"- Output: `{report.get('output_dir_name')}`",
        f"- Next action: {report.get('next_action')}",
        "",
        "## Counts",
        "",
    ]
    for key, value in sorted((report.get("counts") or {}).items()):
        lines.append(f"- `{key}`: {value}")
    lines.extend(["", "## Issues", ""])
    if not report.get("issues"):
        lines.append("- No conformance issues detected.")
    else:
        for item in report.get("issues") or []:
            lines.append(f"- **{item.get('severity')}** `{item.get('code')}`: {item.get('message')}")
            if item.get("detail"):
                lines.append(f"  Detail: `{item.get('detail')}`")
    lines.append("")
    return "\n".join(lines)


def write_reports(report: Dict[str, Any], out_dir: str) -> None:
    _write_json(os.path.join(out_dir, "conformance_report.json"), report)
    with open(os.path.join(out_dir, "conformance_report.md"), "w", encoding="utf-8") as f:
        f.write(report_to_markdown(report))
