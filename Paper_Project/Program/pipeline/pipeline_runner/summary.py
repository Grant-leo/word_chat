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

REPORT_ISSUE_ACTIONS = {
    "CONFORMANCE_INPUT_MISSING": "重新运行完整流水线，确认 format.json、content.json、build_manifest.json 和最终 DOCX 都已生成；若仍失败，打开 {report_path} 查看缺失项。",
    "CONFORMANCE_QA_UNAVAILABLE": "修复 strict conformance QA 依赖后重跑；先查看 {report_path} 里的缺失模块或导入错误。",
    "DOCX_XML_UNREADABLE": "先确认最终 DOCX 能用 Word/WPS 正常打开；如果文件损坏，让 Agent 重新生成最终论文后再重跑 strict QA。",
    "PAGE_GEOMETRY_MISMATCH": "打开 {report_path} 查看页边距/纸张 detail，确认模板页面设置后重跑 strict QA。",
    "CONTENT_PARAGRAPH_MISSING": "对照 内容提取.md 和最终 DOCX 找缺失段落；普通用户先让 Agent 修本次 build_generated.py，开发者再检查正文遍历引擎并重跑 strict QA。",
    "STYLE_MISMATCH": "打开 {report_path} 查看样式 detail；普通用户先让 Agent 修本次 build_generated.py，开发者再检查样式生成规则并重跑 strict QA。",
    "RENDER_COUNT_MISMATCH": "查看 build_manifest.json 的图片/表格/公式渲染数量，定位被跳过的生成分支后重跑 strict QA。",
    "TABLE_NOT_FOUND": "对照 内容提取.md、build_manifest.json 和最终 DOCX 找缺失表格，修复表格渲染分支后重跑 strict QA。",
    "TABLE_BORDER_MISMATCH": "打开 {report_path} 查看三线表 detail，修复表格边框规则后重跑 strict QA。",
    "IMAGE_COUNT_MISMATCH": "对照 内容提取.md、figures/ 和 build_manifest.json 找缺失图片，修复图片输入或渲染分支后重跑 strict QA。",
    "IMAGE_LAYOUT_MISMATCH": "打开 {report_path} 查看图片尺寸/版心 detail，修复图片缩放或居中规则后重跑 strict QA。",
    "FORMULA_COUNT_MISMATCH": "对照 内容提取.md 和 build_manifest.json 找缺失公式，修复公式渲染分支后重跑 strict QA。",
    "OMML_WPS_COMPAT": "检查公式 OOXML，确保每个数学 run 有 WPS 兼容的 m:rPr；修复后重跑 strict QA。",
    "FORMULA_ERROR_TEXT": "最终 DOCX 里还残留公式转换错误文本；检查对应公式源文本或 latex_omath.py，修复后重跑 strict QA。",
    "PLACEHOLDER_TEXT_LEFT": "最终 DOCX 里还残留模板占位符；补齐输入信息或过滤占位符后重跑 strict QA。",
    "WORD_FIELD_ERROR": "最终 DOCX 里还残留 Word 域错误；更新/修复目录、交叉引用或页码字段后重跑 strict QA。",
    "MISSING_DOCX": "先修复构建阶段，确保最终论文 DOCX 生成后再运行 visual QA。",
    "VISUAL_QA_UNAVAILABLE": "修复 visual QA 依赖后重跑；先查看 {report_path}，确认 Word COM 和 Poppler 工具是否可用。",
    "PDF_EXPORT_FAILED": "先确认最终 DOCX 能用 Word 打开，再修复 Word COM/PDF 导出环境并重跑 visual QA。",
    "PDFINFO_UNAVAILABLE": "安装或修复 Poppler 命令行工具（pdfinfo、pdftotext、pdftoppm）后重跑 visual QA。",
    "PDFINFO_FAILED": "打开 visual_report.md 查看 pdfinfo 错误；修复 PDF 导出文件或 Poppler 环境后重跑 visual QA。",
    "PDF_PAGE_COUNT_INVALID": "PDF 导出后没有有效页面；先用 Word 打开 DOCX 检查文件，再重新导出并重跑 visual QA。",
    "PDFTOTEXT_UNAVAILABLE": "安装或修复 Poppler 命令行工具（pdfinfo、pdftotext、pdftoppm）后重跑 visual QA。",
    "PDFTOTEXT_FAILED": "打开 visual_report.md 查看 pdftotext 错误；修复 PDF 导出或 Poppler 环境后重跑 visual QA。",
    "SAMPLE_RENDER_FAILED": "安装或修复 Poppler 的 pdftoppm 后重跑 visual QA，并检查 visual_qa/samples/ 是否生成。",
    "ALL_PAGE_RENDER_FAILED": "安装或修复 Poppler 的 pdftoppm 后重跑 visual QA；若仍失败，先打开导出的 PDF 检查是否损坏。",
    "PAGE_IMAGE_UNREADABLE": "打开 visual_report.md 查看不可读页面，修复 PDF 渲染或页面图片生成后重跑 visual QA。",
    "MANY_BLANK_PAGES": "打开导出的 PDF 核对空白页；如果是异常空白，修复分页/分节逻辑后重跑 visual QA。",
    "TOC_TEXT_NOT_FOUND": "打开导出的 PDF 核对目录页；如果目录缺失，检查 TOC 生成或 Word 字段更新后重跑 visual QA。",
    "MANY_BLANK_PAGE_IMAGES": "打开 visual_qa/samples/ 核对空白样张；如果是异常空白，修复分页或 PDF 渲染后重跑 visual QA。",
    "GOLDEN_BASELINE_MISMATCH": "打开 visual_report.md 和 visual_qa/samples/ 对比页面；确认变化正确则用 --update-golden 更新基线，否则继续修复排版。",
    "GOLDEN_BASELINE_MISSING": "首次建立视觉基线时可用 --update-golden 生成；如果不需要基线，取消 golden 参数后重跑 visual QA。",
    "WPS_PAGE_COUNT_MISMATCH": "分别打开 Word 与 WPS 导出的 PDF 比对分页差异，确认是兼容性差异还是排版脚本问题后再修复。",
    "WPS_EXPORT_UNAVAILABLE": "若启用了 --require-wps，安装/配置 WPS COM；否则可取消 --require-wps 后重跑 visual QA。",
}


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


def _issue_actions(label, report, report_path, limit=5):
    issues = report.get("issues") or []
    if not issues:
        return []
    errors = [item for item in issues if item.get("severity") == "error"]
    warnings = [item for item in issues if item.get("severity") == "warning"]
    ordered = errors + warnings + [item for item in issues if item not in errors and item not in warnings]
    actions = []
    seen = set()
    next_action = str(report.get("next_action") or "").strip()
    for issue in ordered:
        code = str(issue.get("code") or "").strip()
        key = code or str(issue.get("message") or "").strip()
        if key in seen:
            continue
        seen.add(key)
        action_template = REPORT_ISSUE_ACTIONS.get(code)
        if action_template:
            action = action_template.format(report_path=report_path)
        elif next_action:
            action = f"{next_action} 请打开 {report_path} 查看 detail，并在处理后重跑对应 QA。"
        else:
            action = f"打开 {report_path} 查看 detail，按问题码逐项处理后重跑对应 QA。"
        prefix = f"{label} `{code}`" if code else label
        actions.append(f"{prefix}：{action}")
        if len(actions) >= limit:
            break
    remaining = len({str(item.get("code") or item.get("message") or "") for item in ordered}) - len(actions)
    if remaining > 0:
        actions.append(f"{label} 还有 {remaining} 类问题；请继续按对应报告逐项处理。")
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
        report_path = _rel_output(folder_name, md_name)
        exists = bool(report)
        errors, warnings, total = _issue_counts(report) if exists else (0, 0, 0)
        if exists:
            total_errors += errors
            total_warnings += warnings
            next_action = str(report.get("next_action") or "").strip()
            if not report.get("passed"):
                step_actions = _repair_step_actions(label, report)
                issue_actions = [] if step_actions else _issue_actions(label, report, report_path)
                report_actions = step_actions or issue_actions
                if report_actions:
                    next_actions.extend(report_actions)
                elif next_action:
                    next_actions.append(f"{label}: {next_action}")
        reports[key] = {
            "label": label,
            "exists": exists,
            "passed": bool(report.get("passed")) if exists else None,
            "errors": errors,
            "warnings": warnings,
            "issues": total,
            "report": report_path,
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


def _repair_loop_actions(repair_report):
    if not repair_report or repair_report.get("ok") is True:
        return []
    action = str(repair_report.get("next_action") or "").strip()
    if not action:
        return []
    status = str(repair_report.get("status") or "stopped").strip()
    scope = str(repair_report.get("resume_scope") or "").strip()
    command = str(repair_report.get("resume_command") or "").strip()
    parts = [f"自动修复 `{status}`：{action}"]
    if scope:
        parts.append(f"修复范围：`{scope}`。")
    if command and command not in action:
        parts.append(f"恢复命令：`{command}`。")
    return [" ".join(parts)]


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
    next_actions.extend(_repair_loop_actions(repair_report))
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
