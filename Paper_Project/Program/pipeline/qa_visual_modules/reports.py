"""Report rendering for visual QA."""
from __future__ import annotations

import json
import os
from typing import Any, Dict


def report_to_markdown(report: Dict[str, Any]) -> str:
    lines = [
        "# 视觉 QA 报告",
        "",
        f"- 结果：{'通过' if report.get('passed') else '未通过'}",
        f"- 输出目录：`{report.get('output_dir_name')}`",
        f"- 下一步：{report.get('next_action') or '打开最终 DOCX 和 PDF/PNG 渲染样张，检查视觉问题。'}",
        "",
        "## 统计",
        "",
    ]
    for key, value in sorted((report.get("counts") or {}).items()):
        lines.append(f"- `{key}`: {value}")
    artifacts = report.get("artifacts") or {}
    golden = artifacts.get("golden_baseline") or {}
    if golden:
        lines.extend(["", "## 黄金基线", ""])
        lines.append(f"- `status`: {golden.get('status')}")
        if golden.get("key"):
            lines.append(f"- `key`: {golden.get('key')}")
        if golden.get("path"):
            lines.append(f"- `path`: {golden.get('path')}")
        for issue in golden.get("issues") or []:
            lines.append(f"- `issue`: {issue}")
    lines.extend(["", "## 问题", ""])
    issues = report.get("issues") or []
    if not issues:
        lines.append("- 自动视觉检查未发现问题。")
    else:
        for item in issues:
            lines.append(f"- **{item.get('severity')}** `{item.get('code')}`: {item.get('message')}")
            if item.get("detail"):
                lines.append(f"  细节：`{item.get('detail')}`")
    lines.append("")
    return "\n".join(lines)


def write_reports(report: Dict[str, Any], out_dir: str) -> None:
    with open(os.path.join(out_dir, "visual_report.json"), "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)
    with open(os.path.join(out_dir, "visual_report.md"), "w", encoding="utf-8") as f:
        f.write(report_to_markdown(report))
