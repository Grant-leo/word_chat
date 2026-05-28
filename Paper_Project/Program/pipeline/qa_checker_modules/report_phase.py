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


def _next_action(passed: bool, mode: str, issues: List[Dict[str, Any]], repair_plan: Dict[str, Any]) -> str:
    if passed:
        return "通过 QA。仍建议用 WPS/Word 做最终视觉核对。"
    steps = repair_plan.get("steps") or []
    error_steps = [step for step in steps if step.get("severity") == "error"]
    if error_steps and all(str(step.get("auto_level") or "") in NEEDS_USER_AUTO_LEVELS for step in error_steps):
        return "需要用户确认或补充输入文件：请按 qa_repair_plan.md 提供缺失图片、可提取模板、OCR 后 PDF，或修正源内容后重跑。"
    return (
        "用户模式：根据 active_owner 修改当前输出目录的 build_generated.py 后重跑该脚本。"
        if mode == "user" else
        "开发者模式：根据 active_owner 修改核心引擎脚本后重跑完整流水线。"
    )


def build_report(out_dir: str, mode: str, counts: Dict[str, Any], issues: List[Dict[str, Any]]) -> Dict[str, Any]:
    passed = not any(i["severity"] == "error" for i in issues)
    report = {
        "schema_version": 1,
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "mode": mode,
        "output_dir_name": os.path.basename(os.path.abspath(out_dir)),
        "passed": passed,
        "counts": counts,
        "issues": issues,
    }
    report["repair_plan"] = build_repair_plan(report, out_dir)
    report["next_action"] = _next_action(passed, mode, issues, report["repair_plan"])
    return report

