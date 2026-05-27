"""Terminal reporting helpers for the one-click pipeline runner."""
from __future__ import annotations

from .contracts import format_contract_issues


def step(msg):
    print(f'\n{"=" * 50}')
    print(f"  {msg}")
    print(f'{"=" * 50}')


def print_repair_hint(report, out_dir):
    repair = (report or {}).get("repair_plan") or {}
    if not repair:
        return
    print("  [NEXT] 已生成 QA 修复向导 -> qa_repair_plan.md / qa_repair_plan.json")
    print("  [NEXT] 可直接发给 AI 的修复请求 -> qa_fix_prompt.txt")
    steps = repair.get("steps") or []
    for idx, item in enumerate(steps[:3], 1):
        print(f'   {idx}. {item.get("code")}: {item.get("title")}')
        action = str(item.get("user_action") or "")
        if action:
            print(f"      下一步: {action[:160]}")
    commands = repair.get("commands") or {}
    if commands.get("rerun_current_pipeline"):
        print(f'  [NEXT] 修正输入文件后重跑: {commands.get("rerun_current_pipeline")}')
    elif commands.get("rebuild_current_docx"):
        print(f'  [NEXT] 修改 build_generated.py 后重建: {commands.get("rebuild_current_docx")}')


def print_contract_issues(label, issues):
    issues = list(issues or [])
    if not issues:
        return
    print(f"  [CONTRACT] {label}: {len(issues)} issue(s); continuing so QA can report full context")
    for line in format_contract_issues(issues):
        print(f"   - {line}")
