"""Final structural QA report assembly."""
from __future__ import annotations

import os
from datetime import datetime
from typing import Any, Dict, List

try:
    from qa_checker_modules.repair import build_repair_plan
except ImportError:  # pragma: no cover - package-style imports
    from .repair import build_repair_plan


NEEDS_USER_AUTO_LEVELS = {"needs_user_file", "needs_user_input", "needs_user_confirmation", "optional_user_input"}
ENVIRONMENT_AUTO_LEVELS = {"needs_environment"}


def _leading_step_action(steps: List[Dict[str, Any]]) -> str:
    if not steps:
        return ""
    errors = [step for step in steps if step.get("severity") == "error"]
    step = (errors or steps)[0]
    code = str(step.get("code") or "").strip()
    user_action = str(step.get("user_action") or "").strip()
    if code and user_action:
        return f"优先处理 `{code}`：{user_action}"
    if code:
        return f"优先处理 `{code}`。"
    return user_action


def _has_warning_items(issues: List[Dict[str, Any]]) -> bool:
    return any(item.get("severity") == "warning" for item in issues or [])


def _status_label(passed: bool, issues: List[Dict[str, Any]]) -> str:
    if not passed:
        return "failed"
    if _has_warning_items(issues):
        return "passed_with_warnings"
    return "passed"


def _result_label(passed: bool, issues: List[Dict[str, Any]]) -> str:
    if not passed:
        return "未通过"
    if _has_warning_items(issues):
        return "通过但有警告"
    return "通过"


def _next_action(passed: bool, mode: str, issues: List[Dict[str, Any]], repair_plan: Dict[str, Any]) -> str:
    steps = repair_plan.get("steps") or []
    if passed:
        leading_action = _leading_step_action(steps)
        if leading_action:
            return (
                f"QA 已通过但有警告需要人工确认。{leading_action} "
                "如果确认该 warning 不影响交付，可继续用 WPS/Word 做最终视觉核对；如果处理了 warning，请重新运行 QA。"
            )
        return "通过 QA。仍建议用 WPS/Word 做最终视觉核对。"
    error_steps = [step for step in steps if step.get("severity") == "error"]
    leading_action = _leading_step_action(steps)
    if error_steps and all(str(step.get("auto_level") or "") in ENVIRONMENT_AUTO_LEVELS for step in error_steps):
        suffix = f" {leading_action}" if leading_action else " 请安装或修复缺失的本机依赖后重跑完整流水线。"
        return f"需要先修复本机依赖：{suffix.strip()}"
    if error_steps and all(str(step.get("auto_level") or "") in NEEDS_USER_AUTO_LEVELS for step in error_steps):
        suffix = f" {leading_action}" if leading_action else " 请按 qa_repair_plan.md 提供缺失图片、可提取模板、OCR 后 PDF，或修正源内容后重跑。"
        return f"需要用户确认或补充输入文件：{suffix.strip()}"
    if leading_action:
        return (
            f"用户模式：{leading_action} 需要改输出脚本时，根据 active_owner 修改当前输出目录的 build_generated.py 后重跑该脚本。"
            if mode == "user" else
            f"开发者模式：{leading_action} 根据 active_owner 修改核心引擎脚本后重跑完整流水线。"
        )
    return (
        "用户模式：根据 active_owner 修改当前输出目录的 build_generated.py 后重跑该脚本。"
        if mode == "user" else
        "开发者模式：根据 active_owner 修改核心引擎脚本后重跑完整流水线。"
    )


def _with_input_location_hint(action: str, repair_plan: Dict[str, Any]) -> str:
    commands = repair_plan.get("commands") or {}
    hint = str(commands.get("input_location_hint") or "").strip()
    if hint and hint not in action:
        return f"{action} {hint}"
    return action


def build_report(out_dir: str, mode: str, counts: Dict[str, Any], issues: List[Dict[str, Any]]) -> Dict[str, Any]:
    passed = not any(i["severity"] == "error" for i in issues)
    report = {
        "schema_version": 1,
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "mode": mode,
        "output_dir_name": os.path.basename(os.path.abspath(out_dir)),
        "passed": passed,
        "status": _status_label(passed, issues),
        "result_label": _result_label(passed, issues),
        "counts": counts,
        "issues": issues,
    }
    report["repair_plan"] = build_repair_plan(report, out_dir)
    report["next_action"] = _with_input_location_hint(
        _next_action(passed, mode, issues, report["repair_plan"]),
        report["repair_plan"],
    )
    return report

