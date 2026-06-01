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


def _has_warnings(issues: List[Dict[str, Any]]) -> bool:
    return any(item.get("severity") == "warning" for item in issues or [])


def _result_label(passed: bool, issues: List[Dict[str, Any]]) -> str:
    if not passed:
        return "未通过"
    if _has_warnings(issues):
        return "通过但有警告"
    return "通过"


def _ordered_codes(issues: List[Dict[str, Any]], severity: str) -> List[str]:
    codes: List[str] = []
    seen = set()
    for item in issues:
        if item.get("severity") != severity:
            continue
        code = str(item.get("code") or "").strip()
        if code and code not in seen:
            seen.add(code)
            codes.append(code)
    return codes


def _action_for_codes(codes: List[str], mode: str) -> str:
    code_set = set(codes)
    if code_set & {"CONFORMANCE_INPUT_MISSING"}:
        return "重新运行完整流水线，确保 format.json、content.json、build_manifest.json 和最终 DOCX 都已生成。"
    if code_set & {"DOCX_XML_UNREADABLE"}:
        return "先确认最终 DOCX 能被 Word/WPS 正常打开；若文件损坏，重新生成最终论文。"
    if code_set & {"CONTENT_PARAGRAPH_MISSING"}:
        return "对照 content.json/内容提取.md 与最终 DOCX，修复遗漏段落；普通用户优先检查 build_generated.py 的正文遍历。"
    if code_set & {"RENDER_COUNT_MISMATCH", "TABLE_NOT_FOUND", "FORMULA_COUNT_MISMATCH", "IMAGE_COUNT_MISMATCH"}:
        return "查看 build_manifest.json 的渲染数量，定位图片、表格或公式在哪个生成分支被跳过。"
    if code_set & {"STYLE_MISMATCH", "PAGE_GEOMETRY_MISMATCH", "TABLE_BORDER_MISMATCH", "IMAGE_LAYOUT_MISMATCH"}:
        return "按 conformance_report.md 的 detail 修复样式/页边距/表格线/图片尺寸，修复后重跑 strict QA。"
    if code_set & {"OMML_WPS_COMPAT", "FORMULA_ERROR_TEXT"}:
        return "检查公式渲染链路和 latex_omath.py，确保公式输出为 WPS 兼容的原生 OOXML Math。"
    if code_set & {"PLACEHOLDER_TEXT_LEFT"}:
        return "最终 DOCX 里还残留模板占位符；补齐输入信息或过滤占位符后重跑 strict QA。"
    if code_set & {"WORD_FIELD_ERROR"}:
        return "最终 DOCX 里还残留 Word 域错误；更新或修复目录、交叉引用、页码字段后重跑 strict QA。"
    return "普通用户模式：让 Agent 修复 Outputs/<本轮>/build_generated.py 后重跑。" if mode == "user" else "开发者模式：修复核心流水线脚本后重跑完整流水线。"


def _next_action(mode: str, issues: List[Dict[str, Any]]) -> str:
    error_codes = _ordered_codes(issues, "error")
    warning_codes = _ordered_codes(issues, "warning")
    if not error_codes and warning_codes:
        leading = "、".join(f"`{code}`" for code in warning_codes[:3])
        suffix = f" 另有 {len(warning_codes) - 3} 类警告。" if len(warning_codes) > 3 else ""
        action = _action_for_codes(warning_codes, mode)
        return f"strict 合规 QA 没有阻断错误，但有警告 {leading} 需要人工确认；{action} 若确认不影响交付，可继续用 Word/WPS 做最终核对。{suffix}"
    if not error_codes:
        return "strict 合规 QA 的机器检查已通过；仍建议用 Word/WPS 打开最终 DOCX 做人工核对。"
    return _action_for_codes(error_codes, mode)


def _append_review_artifacts(lines: List[str], output_dir_name: Any) -> None:
    folder = str(output_dir_name or "").strip()
    if not folder:
        return
    artifacts = [
        ("最终 DOCX", "最终论文.docx"),
        ("内容摘要", "内容提取.md"),
        ("结构化内容", "content.json"),
        ("渲染清单", "build_manifest.json"),
        ("模板要求", "template_requirements.json"),
        ("格式数据", "format.json"),
    ]
    lines.extend([
        "",
        "## 核对入口",
        "",
        "- 以下是本轮常用核对产物；若某项还未生成，先按顶部“下一步”恢复流水线。",
    ])
    for label, filename in artifacts:
        lines.append(f"- {label}: `Outputs/{folder}/{filename}`")


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
    issues = report.get("issues") or []
    lines = [
        "# DOCX/XML 合规 QA 报告",
        "",
        f"- 结果：{_result_label(bool(report.get('passed')), issues)}",
        f"- 模式：`{report.get('mode')}`",
        f"- 输出目录：`{report.get('output_dir_name')}`",
        f"- 下一步：{report.get('next_action')}",
        "",
        "## 统计",
        "",
    ]
    for key, value in sorted((report.get("counts") or {}).items()):
        lines.append(f"- `{key}`: {value}")
    _append_review_artifacts(lines, report.get("output_dir_name"))
    lines.extend(["", "## 问题", ""])
    if not issues:
        lines.append("- 自动合规检查未发现问题。")
    else:
        for item in issues:
            lines.append(f"- **{item.get('severity')}** `{item.get('code')}`: {item.get('message')}")
            if item.get("detail"):
                lines.append(f"  细节：`{item.get('detail')}`")
    lines.append("")
    return "\n".join(lines)


def write_reports(report: Dict[str, Any], out_dir: str) -> None:
    _write_json(os.path.join(out_dir, "conformance_report.json"), report)
    with open(os.path.join(out_dir, "conformance_report.md"), "w", encoding="utf-8") as f:
        f.write(report_to_markdown(report))
