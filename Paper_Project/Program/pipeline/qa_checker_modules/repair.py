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
        if template and content:
            commands["rerun_current_pipeline"] = (
                f"python run_pipeline.py --mode {mode} --template {template} --content {content}"
            )
    except Exception:
        pass
    return commands


def _parse_count_detail(detail: str) -> Dict[str, int]:
    found: Dict[str, int] = {}
    for key, value in re.findall(r"(content|rendered|docx)=([0-9]+)", str(detail or "")):
        try:
            found[key] = int(value)
        except ValueError:
            pass
    return found


def _repair_step(issue: Dict[str, Any], counts: Dict[str, Any], mode: str) -> Dict[str, Any]:
    code = str(issue.get("code") or "")
    guide = REPAIR_GUIDES.get(code, {})
    count_detail = _parse_count_detail(str(issue.get("detail") or ""))
    owner = issue.get("active_owner") or (
        "Outputs/<run>/build_generated.py" if mode == "user" else OWNER_BY_CODE.get(code, "script_generator.py")
    )
    return {
        "code": code,
        "severity": issue.get("severity"),
        "title": guide.get("title") or code,
        "why": guide.get("why") or str(issue.get("message") or ""),
        "detail": issue.get("detail") or "",
        "counts": count_detail,
        "auto_level": guide.get("auto_level") or "manual_review",
        "target": owner,
        "user_action": guide.get("user_action") or "打开 `qa_report.md` 和最终 DOCX，对照问题详情核查。",
        "developer_action": guide.get("developer_action") or f"检查 `{OWNER_BY_CODE.get(code, 'script_generator.py')}`。",
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
