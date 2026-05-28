"""Markdown/JSON writers for structural QA reports."""
from __future__ import annotations

import json
import os
from typing import Any, Dict


def report_to_markdown(report: Dict[str, Any]) -> str:
    lines = [
        "# QA 检测报告",
        "",
        f"- 模式：`{report.get('mode')}`",
        f"- 结果：{'通过' if report.get('passed') else '未通过'}",
        f"- 输出目录：`{report.get('output_dir_name')}`",
        f"- 下一步：{report.get('next_action')}",
        "",
        "## 统计",
        "",
    ]
    counts = report.get("counts") or {}
    if counts:
        for key in sorted(counts):
            lines.append(f"- `{key}`: {counts[key]}")
    else:
        lines.append("- 无统计信息")

    lines.extend(["", "## 问题", ""])
    issues = report.get("issues") or []
    if not issues:
        lines.append("- 未发现结构性问题。")
    else:
        for item in issues:
            lines.append(
                f"- **{item.get('severity')}** `{item.get('code')}`：{item.get('message')} "
                f"修复目标：`{item.get('active_owner')}`"
            )
            if item.get("detail"):
                lines.append(f"  细节：`{item.get('detail')}`")
    repair_plan = report.get("repair_plan") or {}
    if repair_plan:
        lines.extend(["", "## 修复计划", ""])
        lines.append(f"- 摘要：{repair_plan.get('summary')}")
        commands = repair_plan.get("commands") or {}
        if commands.get("rerun_current_pipeline"):
            lines.append(f"- 重新跑完整流水线：`{commands.get('rerun_current_pipeline')}`")
        if commands.get("rebuild_current_docx"):
            lines.append(f"- 只重建当前 DOCX：`{commands.get('rebuild_current_docx')}`")
        steps = repair_plan.get("steps") or []
        if steps:
            lines.extend(["", "### 建议步骤", ""])
            for idx, step in enumerate(steps, 1):
                lines.append(f"{idx}. **{step.get('code')}**：{step.get('title')}")
                lines.append(f"   - 原因：{step.get('why')}")
                lines.append(f"   - 小白用户下一步：{step.get('user_action')}")
                lines.append(f"   - 开发者检查：{step.get('developer_action')}")
                if step.get("detail"):
                    lines.append(f"   - 细节：`{step.get('detail')}`")
    lines.append("")
    return "\n".join(lines)


def repair_plan_to_markdown(plan: Dict[str, Any]) -> str:
    lines = [
        "# QA 修复向导",
        "",
        f"- 结果：{'已通过' if plan.get('passed') else '需要修复'}",
        f"- 摘要：{plan.get('summary') or ''}",
        f"- 输出目录：`{plan.get('output_dir') or ''}`",
        "",
        "## 先打开这些文件",
        "",
    ]
    for item in plan.get("open_first") or []:
        lines.append(f"- `{item}`")
    commands = plan.get("commands") or {}
    lines.extend(["", "## 可执行命令", ""])
    if commands.get("rerun_current_pipeline"):
        lines.append(f"- 重新跑完整流水线：`{commands.get('rerun_current_pipeline')}`")
    if commands.get("rebuild_current_docx"):
        lines.append(f"- 修改 `build_generated.py` 后只重建当前 DOCX：`{commands.get('rebuild_current_docx')}`")
    if not commands.get("rerun_current_pipeline") and not commands.get("rebuild_current_docx"):
        lines.append("- 暂无可自动推断的命令。")
    steps = plan.get("steps") or []
    lines.extend(["", "## 修复步骤", ""])
    if not steps:
        lines.append("- 当前没有 QA 问题。")
    else:
        for idx, step in enumerate(steps, 1):
            lines.append(f"### {idx}. {step.get('title') or step.get('code')}")
            lines.append("")
            lines.append(f"- 问题码：`{step.get('code')}`")
            lines.append(f"- 级别：`{step.get('severity')}`")
            lines.append(f"- 可能原因：{step.get('why')}")
            lines.append(f"- 小白用户下一步：{step.get('user_action')}")
            lines.append(f"- 开发者检查：{step.get('developer_action')}")
            lines.append(f"- 修复目标：`{step.get('target')}`")
            if step.get("detail"):
                lines.append(f"- 细节：`{step.get('detail')}`")
            counts = step.get("counts") or {}
            if counts:
                lines.append("- 数量： " + ", ".join(f"`{k}={v}`" for k, v in sorted(counts.items())))
            lines.append("")
    prompt = plan.get("copy_to_ai_prompt")
    if prompt:
        lines.extend(["## 可直接发给 AI 的修复请求", "", "```text", str(prompt), "```", ""])
    return "\n".join(lines)


def write_reports(report: Dict[str, Any], out_dir: str) -> None:
    json_path = os.path.join(out_dir, "qa_report.json")
    md_path = os.path.join(out_dir, "qa_report.md")
    repair_json_path = os.path.join(out_dir, "qa_repair_plan.json")
    repair_md_path = os.path.join(out_dir, "qa_repair_plan.md")
    prompt_path = os.path.join(out_dir, "qa_fix_prompt.txt")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)
    with open(md_path, "w", encoding="utf-8") as f:
        f.write(report_to_markdown(report))
    repair_plan = report.get("repair_plan") or {}
    with open(repair_json_path, "w", encoding="utf-8") as f:
        json.dump(repair_plan, f, ensure_ascii=False, indent=2)
    with open(repair_md_path, "w", encoding="utf-8") as f:
        f.write(repair_plan_to_markdown(repair_plan))
    with open(prompt_path, "w", encoding="utf-8") as f:
        f.write(str(repair_plan.get("copy_to_ai_prompt") or ""))

