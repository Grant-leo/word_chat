"""QA phase orchestration for the one-click pipeline runner."""
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
import json
import os

from .contracts import validate_qa_report
from .reports import print_contract_issues, print_repair_hint, qa_status_fields

try:
    from privacy import sanitize_value
except Exception:  # pragma: no cover - best-effort report hardening
    def sanitize_value(value, project_root=None):
        return value


@dataclass(frozen=True)
class QADependencies:
    qa_check_and_write: object
    conformance_check_and_write: object
    visual_check_and_write: object
    optional_import_detail: object


def _safe_report_value(value, project_root=None):
    try:
        return sanitize_value(value, project_root)
    except Exception:
        return value


def _optional_detail(deps: QADependencies, name: str, project_root=None) -> str:
    if deps.optional_import_detail is None:
        return ""
    return str(_safe_report_value(deps.optional_import_detail(name), project_root) or "")


def _exception_detail(exc: Exception, project_root=None) -> str:
    return str(_safe_report_value(f"{exc.__class__.__name__}: {exc}", project_root) or "")


def _issue_counts(report):
    issues = report.get("issues") or []
    error_count = sum(1 for item in issues if item.get("severity") == "error")
    warning_count = sum(1 for item in issues if item.get("severity") == "warning")
    return issues, error_count, warning_count


def _print_report_summary(label, report, *, show_owner=False, report_file="qa_report.md"):
    issues, error_count, warning_count = _issue_counts(report)
    if not report.get("passed"):
        status = "未通过"
    elif warning_count:
        status = "通过但有警告"
    else:
        status = "通过"
    print(f"  [{label}] {status}: {error_count} error(s), {warning_count} warning(s), {len(issues)} issue(s)")
    for item in issues[:8]:
        print(f'   - {item.get("severity")} {item.get("code")}: {item.get("message")}')
        if show_owner:
            print(f'     修复目标: {item.get("active_owner")}')
    if len(issues) > 8:
        print(f"   ... 还有 {len(issues) - 8} 项，请看 {report_file}")
    if report.get("passed") and warning_count:
        next_action = str(report.get("next_action") or "").strip()
        if next_action:
            print(f"     下一步: {next_action[:220]}")


def _print_report_contract(filename, report):
    print_contract_issues(filename, validate_qa_report(report))


def _print_failed_report_hint(qa_report, failed_reports):
    if qa_report and not qa_report.get("passed"):
        print_repair_hint(qa_report, None)
    if not failed_reports:
        print("  [ERROR] QA 未通过。已保留输出目录和最终论文初稿；请按失败报告修复后重跑。")
        return
    print("  [NEXT] 失败报告:")
    for label, filename, report in failed_reports:
        print(f"   - {label}: {filename}")
        next_action = str((report or {}).get("next_action") or "").strip()
        if next_action:
            print(f"     下一步: {next_action[:180]}")
    if qa_report and not qa_report.get("passed"):
        print("  [NEXT] 结构 QA 的逐步修复向导: qa_repair_plan.md / qa_repair_plan.json")
    else:
        print("  [NEXT] 结构 QA 已通过；本轮请优先查看上面的 conformance/visual 报告。")
    print("  [ERROR] QA 未通过。已保留输出目录和最终论文初稿；请按失败报告修复后重跑。")


def _write_dependency_report(out_dir, *, report_name, mode, code, message, detail, next_action, project_root=None):
    message = _safe_report_value(message, project_root)
    detail = _safe_report_value(detail, project_root)
    next_action = _safe_report_value(next_action, project_root)
    report = {
        "schema_version": 1,
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "mode": mode,
        "output_dir_name": os.path.basename(os.path.abspath(out_dir)),
        "passed": False,
        "counts": {},
        "issues": [_safe_report_value({"code": code, "severity": "error", "message": message, "detail": detail}, project_root)],
        "next_action": next_action,
    }
    report.update(qa_status_fields(report["passed"], report["issues"]))
    with open(os.path.join(out_dir, f"{report_name}.json"), "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)
    lines = [
        f"# {report_name.replace('_', ' ')}",
        "",
        "- 结果：未通过",
        f"- 问题码：`{code}`",
        f"- 信息：{message}",
        f"- 下一步：{next_action}",
    ]
    if detail:
        lines.append(f"- 细节：`{detail}`")
    with open(os.path.join(out_dir, f"{report_name}.md"), "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    return report


def _quote_command_arg(value):
    text = str(value)
    if not text:
        return '""'
    if any(ch.isspace() for ch in text) or any(ch in text for ch in '"&|<>^'):
        return '"' + text.replace('"', r'\"') + '"'
    return text


def _workflow_resume_command(out_dir, fallback_mode):
    workflow_path = os.path.join(out_dir, "workflow_mode.json")
    try:
        with open(workflow_path, "r", encoding="utf-8") as f:
            workflow = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        return ""

    mode = str(workflow.get("mode") or fallback_mode or "user")
    if workflow.get("input_location_warnings"):
        return ""
    template = workflow.get("template")
    content = workflow.get("content")
    md_file = workflow.get("md")
    if not md_file and template and content and str(template).lower().endswith(".md") and template == content:
        md_file = template

    args = ["python", "run_pipeline.py", "--mode", mode]
    if md_file:
        args.extend(["--md", str(md_file)])
    elif template and content:
        args.extend(["--template", str(template), "--content", str(content)])
    else:
        return ""

    qa_level = str(workflow.get("qa_level") or "").strip().lower()
    if qa_level in {"basic", "strict", "visual"}:
        args.extend(["--qa-level", qa_level])
    if workflow.get("auto_repair"):
        args.append("--auto-repair")
        if workflow.get("repair_max_rounds"):
            args.extend(["--repair-max-rounds", str(workflow.get("repair_max_rounds"))])
        if workflow.get("repair_stop_no_improve"):
            args.extend(["--repair-stop-no-improve", str(workflow.get("repair_stop_no_improve"))])
    if workflow.get("require_wps"):
        args.append("--require-wps")
    if workflow.get("update_golden"):
        args.append("--update-golden")
    golden_dir = workflow.get("golden_dir")
    if golden_dir:
        args.extend(["--golden-dir", str(golden_dir)])
    return " ".join(_quote_command_arg(arg) for arg in args)


def _workflow_input_location_hint(out_dir):
    workflow_path = os.path.join(out_dir, "workflow_mode.json")
    try:
        with open(workflow_path, "r", encoding="utf-8") as f:
            workflow = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        return ""
    warnings = [str(item).strip() for item in (workflow.get("input_location_warnings") or []) if str(item).strip()]
    return " ".join(warnings)


def _write_structural_dependency_handoff(
    out_dir,
    *,
    mode,
    detail,
    project_root=None,
    code="STRUCTURAL_QA_UNAVAILABLE",
    message=None,
    title=None,
    why=None,
    user_action=None,
    developer_action=None,
    next_action=None,
):
    detail = _safe_report_value(detail, project_root)
    output_dir_name = os.path.basename(os.path.abspath(out_dir))
    resume_command = _workflow_resume_command(out_dir, mode)
    location_hint = _workflow_input_location_hint(out_dir)
    if next_action is None:
        next_action = (
            "结构 QA 无法运行。先修复 qa_checker.py 或 qa_checker_modules 的导入/依赖，"
            "然后重新运行完整流水线；先查看 qa_report.md 和 qa_repair_plan.md。"
        )
    if location_hint:
        next_action = f"{next_action} {location_hint}"
    if resume_command:
        next_action = f"{next_action} 修复后运行：`{resume_command}`"
    issue = {
        "code": code,
        "severity": "error",
        "message": message or "Required structural QA is unavailable, so the pipeline cannot prove this DOCX is safe to deliver.",
        "detail": detail,
        "active_owner": "Paper_Project/Program/pipeline/qa_checker.py / qa_checker_modules",
    }
    step = {
        "code": issue["code"],
        "severity": "error",
        "title": title or "结构 QA 不可用",
        "why": why or "qa_checker.py 是必备结构 QA；缺失时不能把本轮输出标记为已通过。",
        "detail": detail,
        "counts": {},
        "auto_level": "dependency_repair",
        "target": issue["active_owner"],
        "user_action": user_action or "让 Agent 修复结构 QA 依赖后重跑完整流水线；不要把这次输出当作已通过。",
        "developer_action": developer_action or "检查 qa_checker.py、qa_checker_modules 的导入错误和 Python 依赖，修复后重跑完整 regression 与真实流水线。",
    }
    plan = {
        "schema_version": 1,
        "passed": False,
        "summary": "结构 QA 没有运行成功。流水线已停止，避免误报通过。",
        "mode": mode,
        "blocking_errors": 1,
        "warnings": 0,
        "output_dir": output_dir_name,
        "next_action": next_action,
        "resume_scope": "full_pipeline",
        "resume_command": resume_command,
        "open_first": ["qa_report.md", "qa_repair_plan.md", "qa_fix_prompt.txt"],
        "commands": {"rerun_current_pipeline": resume_command, "rebuild_current_docx": ""},
        "steps": [step],
        "copy_to_ai_prompt": "\n".join(
            [
                "请修复本项目的结构 QA 依赖缺失问题。",
                f"输出目录：{output_dir_name}",
                "先阅读 `qa_report.md` 和 `qa_repair_plan.md`。",
                f"下一步：{next_action}",
                f"1. {step['code']}: {step['user_action']}",
            ]
        ),
    }
    report = {
        "schema_version": 1,
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "mode": mode,
        "output_dir_name": output_dir_name,
        "passed": False,
        "counts": {},
        "issues": [issue],
        "next_action": next_action,
        "repair_plan": plan,
    }
    report.update(qa_status_fields(report["passed"], report["issues"]))
    report = _safe_report_value(report, project_root)
    plan = report["repair_plan"]
    issue = report["issues"][0]
    step = plan["steps"][0]
    report_lines = [
        "# QA 检测报告",
        "",
        "- 结果：未通过",
        f"- 问题码：`{issue['code']}`",
        f"- 信息：{issue['message']}",
        f"- 下一步：{report['next_action']}",
        f"- 修复目标：`{issue['active_owner']}`",
    ]
    if issue.get("detail"):
        report_lines.append(f"- 细节：`{issue['detail']}`")
    report_lines.extend(
        [
            "",
            "## 修复计划",
            "",
            "- 先打开 `qa_repair_plan.md`。",
            "- 修复结构 QA 依赖后重新运行完整流水线。",
            "",
        ]
    )
    report_md = "\n".join(report_lines)
    repair_lines = [
        "# QA 修复向导",
        "",
        "- 结果：需要修复",
        f"- 摘要：{plan['summary']}",
        f"- 修复范围：`{plan['resume_scope']}`",
        f"- 优先动作：{plan['next_action']}",
        "",
        "## 先打开这些文件",
        "",
        "- `qa_report.md`",
        "- `qa_fix_prompt.txt`",
        "",
        "## 修复步骤",
        "",
        f"1. **{step['code']}**：{step['title']}",
        f"   - 小白用户下一步：{step['user_action']}",
        f"   - 开发者检查：{step['developer_action']}",
        f"   - 修复目标：`{step['target']}`",
    ]
    if step.get("detail"):
        repair_lines.append(f"   - 细节：`{step['detail']}`")
    repair_lines.append("")
    repair_md = "\n".join(repair_lines)
    files = {
        "qa_report.json": json.dumps(report, ensure_ascii=False, indent=2),
        "qa_report.md": report_md,
        "qa_repair_plan.json": json.dumps(plan, ensure_ascii=False, indent=2),
        "qa_repair_plan.md": repair_md,
        "qa_fix_prompt.txt": str(plan.get("copy_to_ai_prompt") or ""),
    }
    for filename, content in files.items():
        with open(os.path.join(out_dir, filename), "w", encoding="utf-8") as f:
            f.write(content if content.endswith("\n") else content + "\n")
    return report


def run_qa_phases(
    out_dir,
    *,
    mode,
    output_docx_name,
    qa_level,
    project_root,
    golden_dir=None,
    update_golden=False,
    require_wps=False,
    deps: QADependencies,
):
    qa_failed = False
    qa_report = None
    failed_reports = []

    if deps.qa_check_and_write is None:
        detail = _optional_detail(deps, "qa_checker", project_root)
        print(f'  [ERROR] qa_checker.py 不可用，无法执行必备 QA。{detail}')
        qa_report = _write_structural_dependency_handoff(
            out_dir,
            mode=mode,
            detail=detail,
            project_root=project_root,
        )
        _print_report_contract("qa_report.json", qa_report)
        print("  [ERROR] 结构 QA 依赖缺失，已写出 qa_report.md / qa_repair_plan.md / qa_fix_prompt.txt")
        failed_reports.append(("Structural QA", "qa_report.md / qa_repair_plan.md", qa_report))
        _print_failed_report_hint(qa_report, failed_reports)
        return False

    try:
        qa_report = deps.qa_check_and_write(out_dir, mode=mode, output_docx_name=output_docx_name)
    except Exception as exc:
        detail = _exception_detail(exc, project_root)
        print(f"  [ERROR] 结构 QA 执行中断，已写出 qa_report.md / qa_repair_plan.md。{detail}")
        qa_report = _write_structural_dependency_handoff(
            out_dir,
            mode=mode,
            detail=detail,
            project_root=project_root,
            code="STRUCTURAL_QA_FAILED",
            message="Structural QA crashed before it could finish, so the pipeline cannot prove this DOCX is safe to deliver.",
            title="结构 QA 执行失败",
            why="qa_checker.py 运行中抛出异常；不能把本轮输出标记为已通过。",
            user_action="让 Agent 先查看 qa_report.md 和 qa_repair_plan.md，修复 qa_checker.py / qa_checker_modules 后重跑完整流水线；不要把这次输出当作已通过。",
            developer_action="检查 qa_checker.py / qa_checker_modules 的异常堆栈、输入报告和依赖状态，修复后重跑 targeted regression、完整 regression 与真实流水线。",
            next_action="结构 QA 运行中断。先修复 qa_checker.py / qa_checker_modules 的异常，再重新运行完整流水线；先查看 qa_report.md 和 qa_repair_plan.md。",
        )
        _print_report_contract("qa_report.json", qa_report)
        failed_reports.append(("Structural QA", "qa_report.md / qa_repair_plan.md", qa_report))
        _print_failed_report_hint(qa_report, failed_reports)
        return False
    _print_report_contract("qa_report.json", qa_report)
    _print_report_summary("QA", qa_report, show_owner=True)
    print("  [OK] QA 报告 -> qa_report.json / qa_report.md")
    if not qa_report.get("passed"):
        qa_failed = True
        failed_reports.append(("Structural QA", "qa_report.md / qa_repair_plan.md", qa_report))

    if qa_level in ("strict", "visual"):
        if deps.conformance_check_and_write is None:
            detail = _optional_detail(deps, "qa_conformance", project_root)
            print(f'  [ERROR] qa_conformance.py 不可用，无法执行 strict conformance QA。{detail}')
            conformance = _write_dependency_report(
                out_dir,
                report_name="conformance_report",
                mode=mode,
                code="CONFORMANCE_QA_UNAVAILABLE",
                message="strict conformance QA is required but qa_conformance.py is unavailable.",
                detail=detail,
                next_action="修复 strict conformance QA 依赖后重跑；先查看 conformance_report.md。",
                project_root=project_root,
            )
            _print_report_contract("conformance_report.json", conformance)
            failed_reports.append(("Conformance QA", "conformance_report.md", conformance))
            _print_failed_report_hint(qa_report, failed_reports)
            return False
        try:
            conformance = deps.conformance_check_and_write(
                out_dir,
                mode=mode,
                output_docx_name=output_docx_name,
                project_root=project_root,
            )
        except Exception as exc:
            detail = _exception_detail(exc, project_root)
            print(f"  [ERROR] strict conformance QA 执行中断，已写出 conformance_report.md。{detail}")
            conformance = _write_dependency_report(
                out_dir,
                report_name="conformance_report",
                mode=mode,
                code="CONFORMANCE_QA_FAILED",
                message="strict conformance QA crashed before it could finish.",
                detail=detail,
                next_action="strict conformance QA 运行中断。先修复 qa_conformance.py / qa_conformance_modules 的异常后重跑；先查看 conformance_report.md。",
                project_root=project_root,
            )
            _print_report_contract("conformance_report.json", conformance)
            failed_reports.append(("Conformance QA", "conformance_report.md", conformance))
            _print_failed_report_hint(qa_report, failed_reports)
            return False
        _print_report_contract("conformance_report.json", conformance)
        _print_report_summary("Conformance QA", conformance, report_file="conformance_report.md")
        print("  [OK] strict conformance QA -> conformance_report.json / conformance_report.md")
        if not conformance.get("passed"):
            qa_failed = True
            failed_reports.append(("Conformance QA", "conformance_report.md", conformance))

    if qa_level == "visual":
        if deps.visual_check_and_write is None:
            detail = _optional_detail(deps, "qa_visual", project_root)
            print(f'  [ERROR] qa_visual.py 不可用，无法执行 visual QA。{detail}')
            visual = _write_dependency_report(
                out_dir,
                report_name="visual_report",
                mode=mode,
                code="VISUAL_QA_UNAVAILABLE",
                message="visual QA is required but qa_visual.py is unavailable.",
                detail=detail,
                next_action="修复 visual QA 依赖后重跑；先查看 visual_report.md。",
                project_root=project_root,
            )
            _print_report_contract("visual_report.json", visual)
            failed_reports.append(("Visual QA", "visual_report.md", visual))
            _print_failed_report_hint(qa_report, failed_reports)
            return False
        resolved_golden_dir = os.path.abspath(golden_dir) if golden_dir else None
        try:
            visual = deps.visual_check_and_write(
                out_dir,
                output_docx_name=output_docx_name,
                project_root=project_root,
                render_all_pages=True,
                require_wps=bool(require_wps),
                golden_dir=resolved_golden_dir,
                update_golden=bool(update_golden),
            )
        except Exception as exc:
            detail = _exception_detail(exc, project_root)
            print(f"  [ERROR] visual QA 执行中断，已写出 visual_report.md。{detail}")
            visual = _write_dependency_report(
                out_dir,
                report_name="visual_report",
                mode=mode,
                code="VISUAL_QA_FAILED",
                message="visual QA crashed before it could finish.",
                detail=detail,
                next_action="visual QA 运行中断。先修复 qa_visual.py / qa_visual_modules、Word COM 或 Poppler 渲染异常后重跑；先查看 visual_report.md。",
                project_root=project_root,
            )
            _print_report_contract("visual_report.json", visual)
            failed_reports.append(("Visual QA", "visual_report.md", visual))
            _print_failed_report_hint(qa_report, failed_reports)
            return False
        _print_report_contract("visual_report.json", visual)
        _print_report_summary("Visual QA", visual, report_file="visual_report.md")
        print("  [OK] PDF 渲染 QA -> visual_report.json / visual_report.md")
        if not visual.get("passed"):
            qa_failed = True
            failed_reports.append(("Visual QA", "visual_report.md / visual_qa/samples/", visual))

    if qa_failed:
        _print_failed_report_hint(qa_report, failed_reports)
        return False

    return True
