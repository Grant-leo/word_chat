"""Build user-facing repair plans from structural QA reports."""
from __future__ import annotations

import json
import os
import re
from typing import Any, Dict

try:
    from qa_checker_modules.registry import OWNER_BY_CODE, REPAIR_GUIDES
except ImportError:  # pragma: no cover - package-style imports
    from .registry import OWNER_BY_CODE, REPAIR_GUIDES

NEEDS_USER_AUTO_LEVELS = {
    "needs_user_file",
    "needs_user_input",
    "needs_user_confirmation",
    "optional_user_input",
}


def _load_json(path: str) -> Dict[str, Any]:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def _safe_rel(path: str, root: str | None = None) -> str:
    path_text = os.fspath(path)
    try:
        base = os.path.abspath(root or os.getcwd())
        return os.path.relpath(os.path.abspath(path_text), base).replace("\\", "/")
    except Exception:
        return path_text.replace("\\", "/")


def _workflow_commands(out_dir: str, mode: str) -> Dict[str, str]:
    workflow_path = os.path.join(out_dir, "workflow_mode.json")
    build_path = _safe_rel(os.path.join(out_dir, "build_generated.py"))
    commands = {
        "rebuild_current_docx": f"python {build_path}",
        "rerun_current_pipeline": "",
    }
    try:
        workflow = _load_json(workflow_path)
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
            args = []
        if args:
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
            commands["rerun_current_pipeline"] = " ".join(_quote_arg(arg) for arg in args)
    except Exception:
        pass
    return commands


def _quote_arg(value: str) -> str:
    text = str(value)
    if not text:
        return '""'
    if re.search(r'[\s"&|<>^]', text):
        return '"' + text.replace('"', r'\"') + '"'
    return text


def _parse_count_detail(detail: str) -> Dict[str, int]:
    found: Dict[str, int] = {}
    for key, value in re.findall(r"(content|rendered|docx)=([0-9]+)", str(detail or "")):
        try:
            found[key] = int(value)
        except ValueError:
            pass
    return found


def _default_owner(code: str, auto_level: str, mode: str) -> str:
    if auto_level in NEEDS_USER_AUTO_LEVELS:
        return "User input/template file"
    if mode == "user":
        return "Outputs/<run>/build_generated.py"
    return OWNER_BY_CODE.get(code, "script_generator.py")


def _repair_step(issue: Dict[str, Any], counts: Dict[str, Any], mode: str) -> Dict[str, Any]:
    code = str(issue.get("code") or "")
    guide = REPAIR_GUIDES.get(code, {})
    count_detail = _parse_count_detail(str(issue.get("detail") or ""))
    auto_level = str(guide.get("auto_level") or "manual_review")
    owner = issue.get("active_owner") or _default_owner(code, auto_level, mode)
    return {
        "code": code,
        "severity": issue.get("severity"),
        "title": guide.get("title") or code,
        "why": guide.get("why") or str(issue.get("message") or ""),
        "detail": issue.get("detail") or "",
        "counts": count_detail,
        "auto_level": auto_level,
        "target": owner,
        "user_action": guide.get("user_action") or "打开 `qa_report.md` 和最终 DOCX，对照问题详情核查。",
        "developer_action": guide.get("developer_action") or f"检查 `{OWNER_BY_CODE.get(code, 'script_generator.py')}`。",
    }


def _leading_step(steps: list[Dict[str, Any]]) -> Dict[str, Any]:
    errors = [step for step in steps if step.get("severity") == "error"]
    if errors:
        return errors[0]
    if steps:
        return steps[0]
    return {}


def _leading_action(step: Dict[str, Any]) -> str:
    code = str(step.get("code") or "").strip()
    action = str(step.get("user_action") or "").strip()
    if code and action:
        return f"优先处理 `{code}`：{action}"
    if code:
        return f"优先处理 `{code}`。"
    return action


def _resume_route(
    passed: bool,
    mode: str,
    steps: list[Dict[str, Any]],
    all_errors_need_user: bool,
    commands: Dict[str, str],
) -> Dict[str, str]:
    if passed or not steps:
        return {
            "resume_scope": "final_review",
            "resume_command": "",
            "route": "QA 已通过。仍建议用 WPS/Word 打开最终 DOCX 做视觉核对。",
        }
    if all_errors_need_user:
        return {
            "resume_scope": "input_files",
            "resume_command": commands.get("rerun_current_pipeline") or "",
            "route": "先补充或更换输入文件/模板，不要只改 `build_generated.py`；修复后重新运行完整流水线。",
        }
    if mode == "user":
        return {
            "resume_scope": "current_docx",
            "resume_command": commands.get("rebuild_current_docx") or commands.get("rerun_current_pipeline") or "",
            "route": "先按修复步骤处理当前输出目录；需要改输出脚本时只改本次 `build_generated.py`。",
        }
    return {
        "resume_scope": "full_pipeline",
        "resume_command": commands.get("rerun_current_pipeline") or "",
        "route": "先修复 `Paper_Project/Program/pipeline/` 下的核心引擎脚本，再重跑完整流水线。",
    }


def _next_action(
    passed: bool,
    mode: str,
    steps: list[Dict[str, Any]],
    all_errors_need_user: bool,
    commands: Dict[str, str],
) -> Dict[str, str]:
    route = _resume_route(passed, mode, steps, all_errors_need_user, commands)
    action = _leading_action(_leading_step(steps)) if steps else ""
    parts = []
    if action:
        parts.append(action)
    parts.append(route["route"])
    if route["resume_command"]:
        parts.append(f"修复后运行：`{route['resume_command']}`")
    return {
        "next_action": " ".join(part for part in parts if part).strip(),
        "resume_command": route["resume_command"],
        "resume_scope": route["resume_scope"],
    }


def build_repair_plan(report: Dict[str, Any], out_dir: str) -> Dict[str, Any]:
    issues = report.get("issues") or []
    counts = report.get("counts") or {}
    mode = str(report.get("mode") or "user")
    errors = [i for i in issues if i.get("severity") == "error"]
    warnings = [i for i in issues if i.get("severity") == "warning"]
    ordered_issues = errors + warnings + [i for i in issues if i not in errors and i not in warnings]
    steps = [_repair_step(item, counts, mode) for item in ordered_issues]
    commands = _workflow_commands(out_dir, mode)
    error_steps = [step for step in steps if step.get("severity") == "error"]
    all_errors_need_user = bool(error_steps) and all(str(step.get("auto_level") or "") in NEEDS_USER_AUTO_LEVELS for step in error_steps)
    if all_errors_need_user:
        commands["rebuild_current_docx"] = ""
    next_action = _next_action(bool(report.get("passed")), mode, steps, all_errors_need_user, commands)
    summary = (
        "QA 已通过，仍建议用 WPS/Word 做最终视觉核对。"
        if not errors else
        f"QA 发现 {len(errors)} 个阻断错误和 {len(warnings)} 个警告。最终 DOCX 已保留，但交付前需要按修复计划处理。"
    )
    user_prompt_lines = [
        "请继续修复本次 Word 论文流水线输出。",
        f"输出目录：{_safe_rel(out_dir)}",
        "先阅读 `qa_repair_plan.md`、`qa_report.md`、`内容提取.md`、`build_manifest.json`。",
        "目标：优先处理 error，再处理 warning；修复后重新生成最终 DOCX 并重新运行 QA。",
    ]
    if next_action.get("next_action"):
        user_prompt_lines.append(f"下一步：{next_action['next_action']}")
    if mode == "user":
        if all_errors_need_user:
            user_prompt_lines.append("当前是 user 模式，本轮需要先指导用户补充或更换输入文件；不要只修改 `build_generated.py`。")
        else:
            user_prompt_lines.append("当前是 user 模式，优先只修改本次输出目录里的 `build_generated.py` 或指导用户修正输入文件。")
    else:
        user_prompt_lines.append("当前是 developer 模式，可修改 `Paper_Project/Program/pipeline/` 下的核心引擎脚本并重跑完整流水线。")
    for idx, step in enumerate(steps[:5], 1):
        user_prompt_lines.append(f"{idx}. {step['code']}: {step['user_action']}")
    return {
        "schema_version": 1,
        "passed": bool(report.get("passed")),
        "summary": summary,
        "mode": mode,
        "blocking_errors": len(errors),
        "warnings": len(warnings),
        "output_dir": _safe_rel(out_dir),
        "next_action": next_action["next_action"],
        "resume_scope": next_action["resume_scope"],
        "resume_command": next_action["resume_command"],
        "open_first": [
            "qa_repair_plan.md",
            "qa_report.md",
            "内容提取.md",
            "build_manifest.json",
            "最终论文.docx",
        ],
        "commands": commands,
        "steps": steps,
        "copy_to_ai_prompt": "\n".join(user_prompt_lines),
    }
