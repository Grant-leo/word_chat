"""Report writing helpers for template profiles."""
from __future__ import annotations

import json
import os
from typing import Any, Dict

from .profiles import profile_format


def report_to_markdown(profile: Dict[str, Any]) -> str:
    cap = profile.get("capabilities") or {}
    risks = profile.get("risk_flags") or {}
    pdf = profile.get("pdf_template") or {}
    lines = [
        "# 模板能力画像",
        "",
        f"- 段落: `{profile.get('counts', {}).get('paragraphs', 0)}`",
        f"- 表格: `{profile.get('counts', {}).get('tables', 0)}`",
        f"- 节: `{profile.get('counts', {}).get('sections', 0)}`",
        f"- 封面元素: `{profile.get('counts', {}).get('cover_elements', 0)}`",
        "",
        "## 能力",
        "",
    ]
    for key in sorted(cap):
        lines.append(f"- `{key}`: {cap[key]}")
    if pdf:
        lines.extend([
            "",
            "## PDF 模板",
            "",
            f"- 类型: `{pdf.get('type')}`",
            f"- 置信度: `{pdf.get('confidence')}`",
            f"- 页数: `{pdf.get('page_count')}`",
            f"- 可提取文本字符数: `{pdf.get('text_chars')}`",
        ])
        if pdf.get("warnings"):
            lines.append(f"- 警告: `{'; '.join(map(str, pdf.get('warnings') or []))}`")
        if pdf.get("errors"):
            lines.append(f"- 错误: `{'; '.join(map(str, pdf.get('errors') or []))}`")
    lines.extend(["", "## 风险标记", ""])
    for key in sorted(risks):
        lines.append(f"- `{key}`: {risks[key]}")
    lines.append("")
    return "\n".join(lines)


def write_profile(fmt: Dict[str, Any], output_dir: str, project_root: str | None = None) -> Dict[str, Any]:
    profile = profile_format(fmt, project_root=project_root)
    json_path = os.path.join(output_dir, "template_profile.json")
    md_path = os.path.join(output_dir, "template_profile.md")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(profile, f, ensure_ascii=False, indent=2)
    with open(md_path, "w", encoding="utf-8") as f:
        f.write(report_to_markdown(profile))
    return profile
