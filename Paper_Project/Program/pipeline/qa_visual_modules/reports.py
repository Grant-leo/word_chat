"""Report rendering for visual QA."""
from __future__ import annotations

import json
import os
from typing import Any, Dict


def report_to_markdown(report: Dict[str, Any]) -> str:
    lines = [
        "# Visual QA Report",
        "",
        f"- Result: {'passed' if report.get('passed') else 'failed'}",
        f"- Output: `{report.get('output_dir_name')}`",
        "",
        "## Counts",
        "",
    ]
    for key, value in sorted((report.get("counts") or {}).items()):
        lines.append(f"- `{key}`: {value}")
    artifacts = report.get("artifacts") or {}
    golden = artifacts.get("golden_baseline") or {}
    if golden:
        lines.extend(["", "## Golden Baseline", ""])
        lines.append(f"- `status`: {golden.get('status')}")
        if golden.get("key"):
            lines.append(f"- `key`: {golden.get('key')}")
        if golden.get("path"):
            lines.append(f"- `path`: {golden.get('path')}")
        for issue in golden.get("issues") or []:
            lines.append(f"- `issue`: {issue}")
    lines.extend(["", "## Issues", ""])
    issues = report.get("issues") or []
    if not issues:
        lines.append("- No visual QA issues detected by automated checks.")
    else:
        for item in issues:
            lines.append(f"- **{item.get('severity')}** `{item.get('code')}`: {item.get('message')}")
            if item.get("detail"):
                lines.append(f"  Detail: `{item.get('detail')}`")
    lines.append("")
    return "\n".join(lines)


def write_reports(report: Dict[str, Any], out_dir: str) -> None:
    with open(os.path.join(out_dir, "visual_report.json"), "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)
    with open(os.path.join(out_dir, "visual_report.md"), "w", encoding="utf-8") as f:
        f.write(report_to_markdown(report))

