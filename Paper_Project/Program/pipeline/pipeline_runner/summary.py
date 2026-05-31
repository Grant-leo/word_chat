"""Completion-summary helpers for the one-click pipeline runner."""
from __future__ import annotations

from datetime import datetime
import json
import os


REPORT_SPECS = (
    ("structural", "结构 QA", "qa_report.json", "qa_report.md"),
    ("conformance", "DOCX/XML 合规 QA", "conformance_report.json", "conformance_report.md"),
    ("visual", "视觉 QA", "visual_report.json", "visual_report.md"),
)

REPORT_LABELS = {key: label for key, label, _json_name, _md_name in REPORT_SPECS}


def _read_json(path, default=None):
    try:
        with open(path, "r", encoding="utf-8") as handle:
            return json.load(handle)
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        return {} if default is None else default


def _issue_counts(report):
    issues = report.get("issues") or []
    errors = sum(1 for item in issues if item.get("severity") == "error")
    warnings = sum(1 for item in issues if item.get("severity") == "warning")
    return errors, warnings, len(issues)


def _repair_step_actions(label, report, limit=5):
    repair_plan = report.get("repair_plan") or {}
    steps = repair_plan.get("steps") or []
    if not steps:
        return []
    errors = [item for item in steps if item.get("severity") == "error"]
    warnings = [item for item in steps if item.get("severity") == "warning"]
    ordered = errors + warnings + [item for item in steps if item not in errors and item not in warnings]
    actions = []
    for step in ordered[:limit]:
        code = str(step.get("code") or "").strip()
        user_action = str(step.get("user_action") or "").strip()
        title = str(step.get("title") or "").strip()
        if user_action:
            prefix = f"{label} `{code}`" if code else label
            actions.append(f"{prefix}：{user_action}")
        elif title:
            prefix = f"{label} `{code}`" if code else label
            actions.append(f"{prefix}：{title}")
    if len(ordered) > limit:
        actions.append(f"{label} 还有 {len(ordered) - limit} 项问题；请继续按对应报告逐项处理。")
    return actions


def _rel_output(folder_name, *parts):
    normalized_parts = [str(part).replace(os.sep, "/") for part in parts if part]
    return "/".join(["Outputs", folder_name, *normalized_parts])


def _report_summary(out_dir, folder_name):
    reports = {}
    total_errors = 0
    total_warnings = 0
    next_actions = []
    for key, label, json_name, md_name in REPORT_SPECS:
        report = _read_json(os.path.join(out_dir, json_name), default={})
        exists = bool(report)
        errors, warnings, total = _issue_counts(report) if exists else (0, 0, 0)
        if exists:
            total_errors += errors
            total_warnings += warnings
            next_action = str(report.get("next_action") or "").strip()
            if not report.get("passed"):
                step_actions = _repair_step_actions(label, report)
                if step_actions:
                    next_actions.extend(step_actions)
                elif next_action:
                    next_actions.append(f"{label}: {next_action}")
        reports[key] = {
            "label": label,
            "exists": exists,
            "passed": bool(report.get("passed")) if exists else None,
            "errors": errors,
            "warnings": warnings,
            "issues": total,
            "report": _rel_output(folder_name, md_name),
        }
    return reports, total_errors, total_warnings, next_actions


def _expected_report_keys(workflow):
    if not workflow or not workflow.get("qa_enabled"):
        return []
    level = str(workflow.get("qa_level") or "basic").strip().lower()
    keys = ["structural"]
    if level in ("strict", "visual"):
        keys.append("conformance")
    if level == "visual":
        keys.append("visual")
    return keys


def _missing_report_actions(reports, missing_keys):
    actions = []
    for key in missing_keys:
        label = REPORT_LABELS.get(key, key)
        report_path = (reports.get(key) or {}).get("report") or ""
        if key == "conformance":
            actions.append(
                f"{label} 未生成；请重新运行完整流水线。若仍缺失，请先修复 strict QA 依赖并查看 {report_path}。"
            )
        elif key == "visual":
            actions.append(
                f"{label} 未生成；请确认 Word COM 和 Poppler 工具可用后重新运行 visual QA，并查看 {report_path}。"
            )
        else:
            actions.append(
                f"{label} 未生成；请重新运行完整流水线，若仍缺失请查看终端错误和 qa_report.md。"
            )
    return actions


def _friendly_manual_check(item):
    text = str(item or "").strip()
    lowered = text.lower()
    if "open the final docx" in lowered and "word/wps" in lowered:
        return "用 Word/WPS 打开最终 DOCX，核对分页、图片、公式、表格和目录。"
    if "review remaining warnings" in lowered or "剩余 warning" in text:
        return "查看 QA 报告中的剩余 warning，并确认它们不会影响交付。"
    if "automatic qa convergence is not a 100%" in lowered:
        return "自动 QA 通过不等于 100% 保证，交付前仍需人工视觉核对。"
    return text


def _manual_checks(repair_report, next_actions):
    checks = []
    if repair_report:
        final_warnings = int(repair_report.get("final_warnings") or 0)
        for item in repair_report.get("manual_check_required") or []:
            text = str(item or "").strip()
            if not text:
                continue
            if final_warnings == 0 and ("remaining warnings" in text.lower() or "剩余 warning" in text):
                continue
            checks.append(_friendly_manual_check(text))
        remaining = str(repair_report.get("remaining_manual_note") or "").strip()
        if remaining and "no remaining warnings" not in remaining.lower() and "未报告剩余 warning" not in remaining:
            checks.append(_friendly_manual_check(remaining))
    checks.extend(_friendly_manual_check(item) for item in next_actions)
    if not checks:
        checks = [
            "用 Word/WPS 打开最终 DOCX，核对分页、图片、公式、表格和目录。",
            "若学校或期刊有特殊细则，人工确认封面、摘要、参考文献和页眉页脚。",
        ]
    seen = set()
    unique = []
    for item in checks:
        if item not in seen:
            seen.add(item)
            unique.append(item)
    return unique


def build_agent_summary(
    out_dir,
    folder_name=None,
    output_docx="最终论文.docx",
    mode="user",
    pipeline_status="completed",
    note=None,
):
    folder_name = folder_name or os.path.basename(os.path.abspath(out_dir))
    workflow = _read_json(os.path.join(out_dir, "workflow_mode.json"), default={})
    repair_report = _read_json(os.path.join(out_dir, "repair_loop_report.json"), default={})
    reports, report_errors, report_warnings, next_actions = _report_summary(out_dir, folder_name)

    final_errors = int(repair_report.get("final_errors", report_errors) or 0) if repair_report else report_errors
    final_warnings = int(repair_report.get("final_warnings", report_warnings) or 0) if repair_report else report_warnings
    output_docx_path = os.path.join(out_dir, output_docx)
    final_docx_exists = os.path.exists(output_docx_path)
    reports_present = any(item["exists"] for item in reports.values())
    should_require_qa_reports = pipeline_status == "completed" or final_docx_exists
    expected_report_keys = _expected_report_keys(workflow) if should_require_qa_reports else []
    missing_required_reports = [
        key for key in expected_report_keys if not reports.get(key, {}).get("exists")
    ]
    all_existing_reports_passed = all(
        item["passed"] is not False for item in reports.values() if item["exists"]
    )
    required_reports_present = bool(expected_report_keys) and not missing_required_reports
    automatic_qa_passed = required_reports_present and final_errors == 0 and all_existing_reports_passed
    next_actions.extend(_missing_report_actions(reports, missing_required_reports))
    if pipeline_status != "completed":
        status_label = "需要继续处理"
    elif automatic_qa_passed and final_docx_exists:
        status_label = "自动 QA 已通过"
    elif final_docx_exists and missing_required_reports:
        status_label = "已生成 DOCX，但自动 QA 报告不完整"
    elif final_docx_exists and not reports_present:
        status_label = "已生成 DOCX，未运行自动 QA"
    elif final_docx_exists:
        status_label = "已生成 DOCX，但自动 QA 未完全通过"
    else:
        status_label = "未生成最终 DOCX"
    manual_checks = _manual_checks(repair_report, next_actions)
    if pipeline_status != "completed" and not final_docx_exists and not next_actions:
        manual_checks = [
            note or "流水线已中断，请先查看终端错误或对应报告。",
            "修复输入、依赖或构建错误后，让 Agent 重新运行完整流水线。",
        ]

    return {
        "schema_version": 1,
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "status": pipeline_status,
        "status_label": status_label,
        "note": note or "",
        "output_dir": _rel_output(folder_name),
        "output_docx": _rel_output(folder_name, output_docx),
        "final_docx_exists": final_docx_exists,
        "mode": workflow.get("mode") or mode,
        "qa_level": workflow.get("qa_level") or "",
        "auto_repair": bool(workflow.get("auto_repair")),
        "agent_auto": bool(workflow.get("agent_auto")),
        "missing_required_reports": missing_required_reports,
        "repair_loop": {
            "exists": bool(repair_report),
            "status": repair_report.get("status") if repair_report else "",
            "rounds_run": int(repair_report.get("rounds_run") or 0) if repair_report else 0,
            "final_errors": final_errors,
            "final_warnings": final_warnings,
            "report": _rel_output(folder_name, "repair_loop_report.md"),
        },
        "reports": reports,
        "next_actions": next_actions,
        "manual_check_required": manual_checks,
    }


def _agent_summary_markdown(summary):
    lines = [
        "# Agent 排版摘要",
        "",
        f"- 状态：{summary['status_label']}",
        f"- 最终论文：`{summary['output_docx']}`",
        f"- 输出目录：`{summary['output_dir']}/`",
        f"- 工作模式：`{summary['mode']}`",
        f"- QA 等级：`{summary.get('qa_level') or '未记录'}`",
        f"- Agent 自动入口：`{'是' if summary.get('agent_auto') else '否'}`",
        f"- 自动修复：`{'是' if summary.get('auto_repair') else '否'}`",
    ]
    if summary.get("note"):
        lines.append(f"- 说明：{summary['note']}")

    repair = summary["repair_loop"]
    if repair["exists"]:
        lines.extend(
            [
                "",
                "## 自动修复",
                "",
                f"- 状态：`{repair.get('status')}`",
                f"- 轮次：`{repair.get('rounds_run')}`",
                f"- 最终错误：`{repair.get('final_errors')}`",
                f"- 最终警告：`{repair.get('final_warnings')}`",
                f"- 报告：`{repair.get('report')}`",
            ]
        )

    lines.extend(["", "## QA 报告", ""])
    for report in summary["reports"].values():
        if not report["exists"]:
            lines.append(f"- {report['label']}：未生成")
            continue
        result = "通过" if report["passed"] else "未通过"
        lines.append(
            f"- {report['label']}：{result}，错误 `{report['errors']}`，警告 `{report['warnings']}`，报告 `{report['report']}`"
        )

    lines.extend(["", "## 仍需人工查看", ""])
    for item in summary["manual_check_required"]:
        lines.append(f"- {item}")
    lines.extend(
        [
            "",
            "## 下一句可以这样对 Agent 说",
            "",
            "请打开最新 `agent_summary.md`、`qa_report.md` 和最终 DOCX，帮我继续检查需要人工确认的地方。",
        ]
    )
    return "\n".join(lines) + "\n"


def write_agent_summary(
    out_dir,
    folder_name=None,
    output_docx="最终论文.docx",
    mode="user",
    pipeline_status="completed",
    note=None,
):
    summary = build_agent_summary(
        out_dir,
        folder_name=folder_name,
        output_docx=output_docx,
        mode=mode,
        pipeline_status=pipeline_status,
        note=note,
    )
    json_path = os.path.join(out_dir, "agent_summary.json")
    md_path = os.path.join(out_dir, "agent_summary.md")
    with open(json_path, "w", encoding="utf-8") as handle:
        json.dump(summary, handle, ensure_ascii=False, indent=2)
    with open(md_path, "w", encoding="utf-8") as handle:
        handle.write(_agent_summary_markdown(summary))
    return json_path, md_path


def _agent_preflight_markdown(report):
    lines = [
        "# Agent 预检中断报告",
        "",
        f"- 状态：{report.get('status_label')}",
        f"- 原因：{report.get('message')}",
        f"- 说明：这不是正式排版输出目录，流水线尚未开始生成 DOCX。",
        "",
        "## 下一步",
        "",
    ]
    for item in report.get("next_steps") or []:
        lines.append(f"- {item}")
    candidates = report.get("candidates") or {}
    if candidates:
        lines.extend(["", "## 候选文件", ""])
        for label, files in candidates.items():
            lines.append(f"### {label}")
            for filename in files or []:
                lines.append(f"- `{filename}`")
            lines.append("")
    return "\n".join(lines).rstrip() + "\n"


def write_agent_preflight_report(outputs_dir, *, status, message, next_steps, candidates=None):
    out_dir = os.path.join(outputs_dir, "_agent_preflight_latest")
    os.makedirs(out_dir, exist_ok=True)
    report = {
        "schema_version": 1,
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "status": status,
        "status_label": "需要用户补充信息后重跑",
        "message": str(message or "").strip(),
        "next_steps": [str(item or "").strip() for item in (next_steps or []) if str(item or "").strip()],
        "candidates": candidates or {},
        "output_dir": "Outputs/_agent_preflight_latest",
    }
    json_path = os.path.join(out_dir, "agent_preflight_report.json")
    md_path = os.path.join(out_dir, "agent_preflight_report.md")
    with open(json_path, "w", encoding="utf-8") as handle:
        json.dump(report, handle, ensure_ascii=False, indent=2)
    with open(md_path, "w", encoding="utf-8") as handle:
        handle.write(_agent_preflight_markdown(report))
    return json_path, md_path


def build_completion_summary(folder_name, output_docx, mode):
    active_mode = "普通用户" if mode == "user" else "开发者"
    return f"""
  输出目录: Outputs/{folder_name}/
    ├── 格式提取.md          <- 核对模版格式
    ├── 内容提取.md          <- 核对文本内容
    ├── format.json
    ├── content.json
    ├── template_profile.json <- 模板能力画像
    ├── template_requirements.json <- 机器可核查模板要求
    ├── workflow_mode.json <- 用户/开发者模式
    ├── build_manifest.json <- 正文元素渲染数量
    ├── qa_report.md       <- 自动检测报告
    ├── agent_summary.md   <- 给用户和 Agent 的最终交接摘要
    ├── conformance_report.md <- strict DOCX/XML 合规报告
    ├── repair_loop_report.md <- 自动修复闭环报告（--auto-repair 时生成）
    ├── visual_report.md   <- PDF 渲染 QA（--qa-level visual 时生成）
    ├── build_generated.py   <- 生成脚本
    └── {output_docx}        <- 最终文件

  修复工作流:
    当前模式: {active_mode}
    普通用户模式: 让 Agent 修改本次输出目录中的 build_generated.py，然后重跑该脚本
    自动修复模式: 使用 --auto-repair 后读取 repair_loop_report.md 查看每轮修改和停止原因
    开发者模式: 修改 Paper_Project/Program/pipeline/ 下的核心脚本后重跑完整流水线
    目录: 生成脚本会优先用 Word COM 解析正文标题页码；不可用时仍保留静态目录行
"""


def print_completion_summary(folder_name, output_docx, mode):
    print(build_completion_summary(folder_name, output_docx, mode))
