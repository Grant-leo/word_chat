"""Final structural QA report assembly."""
from __future__ import annotations

import os
from datetime import datetime
from typing import Any, Dict, List

try:
    from qa_checker_modules.repair import build_repair_plan
except ImportError:  # pragma: no cover - package-style imports
    from .repair import build_repair_plan
def build_report(out_dir: str, mode: str, counts: Dict[str, Any], issues: List[Dict[str, Any]]) -> Dict[str, Any]:
    passed = not any(i["severity"] == "error" for i in issues)
    next_action = (
        "通过 QA。仍建议用 WPS/Word 做最终视觉核对。"
        if passed else
        ("用户模式：根据 active_owner 修改当前输出目录的 build_generated.py 后重跑该脚本。"
         if mode == "user" else
         "开发者模式：根据 active_owner 修改核心引擎脚本后重跑完整流水线。")
    )

    report = {
        "schema_version": 1,
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "mode": mode,
        "output_dir_name": os.path.basename(os.path.abspath(out_dir)),
        "passed": passed,
        "counts": counts,
        "issues": issues,
        "next_action": next_action,
    }
    report["repair_plan"] = build_repair_plan(report, out_dir)
    return report

