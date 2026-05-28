"""Report writing helpers for template profiles."""
from __future__ import annotations

import json
import os
from typing import Any, Dict

from .profiles import profile_format


def report_to_markdown(profile: Dict[str, Any]) -> str:
    cap = profile.get("capabilities") or {}
    risks = profile.get("risk_flags") or {}
    lines = [
        "# 妯℃澘鐢诲儚",
        "",
        f"- 娈佃惤: `{profile.get('counts', {}).get('paragraphs', 0)}`",
        f"- 琛ㄦ牸: `{profile.get('counts', {}).get('tables', 0)}`",
        f"- 鑺? `{profile.get('counts', {}).get('sections', 0)}`",
        f"- 灏侀潰鍏冪礌: `{profile.get('counts', {}).get('cover_elements', 0)}`",
        "",
        "## 鑳藉姏",
        "",
    ]
    for key in sorted(cap):
        lines.append(f"- `{key}`: {cap[key]}")
    lines.extend(["", "## 椋庨櫓鏍囪", ""])
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
