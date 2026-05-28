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
        "next_action": (
            "Strict conformance passed for machine-checkable template requirements."
            if not any(i.get("severity") == "error" for i in issues)
            else ("Fix Outputs/<run>/build_generated.py and rerun it." if mode == "user" else "Fix core pipeline scripts and rerun the full pipeline.")
        ),
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

