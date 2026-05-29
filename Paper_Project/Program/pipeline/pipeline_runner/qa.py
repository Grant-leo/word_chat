"""QA phase orchestration for the one-click pipeline runner."""
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
import json
import os

from .contracts import validate_qa_report
from .reports import print_contract_issues, print_repair_hint


@dataclass(frozen=True)
class QADependencies:
    qa_check_and_write: object
    conformance_check_and_write: object
    visual_check_and_write: object
    optional_import_detail: object


def _optional_detail(deps: QADependencies, name: str) -> str:
    if deps.optional_import_detail is None:
        return ""
    return deps.optional_import_detail(name)


def _issue_counts(report):
    issues = report.get("issues") or []
    error_count = sum(1 for item in issues if item.get("severity") == "error")
    return issues, error_count


def _print_report_summary(label, report, *, show_owner=False, report_file="qa_report.md"):
    issues, error_count = _issue_counts(report)
    status = "通过" if report.get("passed") else "未通过"
    print(f"  [{label}] {status}: {error_count} error(s), {len(issues)} issue(s)")
    for item in issues[:8]:
        print(f'   - {item.get("severity")} {item.get("code")}: {item.get("message")}')
        if show_owner:
            print(f'     修复目标: {item.get("active_owner")}')
    if len(issues) > 8:
        print(f"   ... 还有 {len(issues) - 8} 项，请看 {report_file}")


def _print_failed_report_hint(qa_report, failed_reports):
    if qa_report and not qa_report.get("passed"):
        print_repair_hint(qa_report, None)
    if not failed_reports:
        print("  [ERROR] QA 未通过。已保留输出目录和最终论文初稿；请按失败报告修复后重跑。")
        return
    print("  [NEXT] 失败报告:")
    for label, filename, report in failed_reports:
        print(f"   - {label}: {filename}")
        next_action = str((report or {}).get("next_action") or "").strip()
        if next_action:
            print(f"     下一步: {next_action[:180]}")
    if qa_report and not qa_report.get("passed"):
        print("  [NEXT] 结构 QA 的逐步修复向导: qa_repair_plan.md / qa_repair_plan.json")
    else:
        print("  [NEXT] 结构 QA 已通过；本轮请优先查看上面的 conformance/visual 报告。")
    print("  [ERROR] QA 未通过。已保留输出目录和最终论文初稿；请按失败报告修复后重跑。")


def _write_dependency_report(out_dir, *, report_name, mode, code, message, detail, next_action):
    report = {
        "schema_version": 1,
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "mode": mode,
        "output_dir_name": os.path.basename(os.path.abspath(out_dir)),
        "passed": False,
        "counts": {},
        "issues": [{"code": code, "severity": "error", "message": message, "detail": detail}],
        "next_action": next_action,
    }
    with open(os.path.join(out_dir, f"{report_name}.json"), "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)
    lines = [
        f"# {report_name.replace('_', ' ').title()}",
        "",
        "- Result: failed",
        f"- Issue: `{code}`",
        f"- Message: {message}",
        f"- Next action: {next_action}",
    ]
    if detail:
        lines.append(f"- Detail: `{detail}`")
    with open(os.path.join(out_dir, f"{report_name}.md"), "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    return report


def run_qa_phases(
    out_dir,
    *,
    mode,
    output_docx_name,
    qa_level,
    project_root,
    golden_dir=None,
    update_golden=False,
    require_wps=False,
    deps: QADependencies,
):
    qa_failed = False
    qa_report = None
    failed_reports = []

    if deps.qa_check_and_write is None:
        print(f'  [ERROR] qa_checker.py 不可用，无法执行必备 QA。{_optional_detail(deps, "qa_checker")}')
        return False

    qa_report = deps.qa_check_and_write(out_dir, mode=mode, output_docx_name=output_docx_name)
    print_contract_issues("qa_report.json", validate_qa_report(qa_report))
    _print_report_summary("QA", qa_report, show_owner=True)
    print("  [OK] QA 报告 -> qa_report.json / qa_report.md")
    if not qa_report.get("passed"):
        qa_failed = True
        failed_reports.append(("Structural QA", "qa_report.md / qa_repair_plan.md", qa_report))

    if qa_level in ("strict", "visual"):
        if deps.conformance_check_and_write is None:
            detail = _optional_detail(deps, "qa_conformance")
            print(f'  [ERROR] qa_conformance.py 不可用，无法执行 strict conformance QA。{detail}')
            conformance = _write_dependency_report(
                out_dir,
                report_name="conformance_report",
                mode=mode,
                code="CONFORMANCE_QA_UNAVAILABLE",
                message="strict conformance QA is required but qa_conformance.py is unavailable.",
                detail=detail,
                next_action="修复 strict conformance QA 依赖后重跑；先查看 conformance_report.md。",
            )
            failed_reports.append(("Conformance QA", "conformance_report.md", conformance))
            _print_failed_report_hint(qa_report, failed_reports)
            return False
        conformance = deps.conformance_check_and_write(
            out_dir,
            mode=mode,
            output_docx_name=output_docx_name,
            project_root=project_root,
        )
        _print_report_summary("Conformance QA", conformance, report_file="conformance_report.md")
        print("  [OK] strict conformance QA -> conformance_report.json / conformance_report.md")
        if not conformance.get("passed"):
            qa_failed = True
            failed_reports.append(("Conformance QA", "conformance_report.md", conformance))

    if qa_level == "visual":
        if deps.visual_check_and_write is None:
            detail = _optional_detail(deps, "qa_visual")
            print(f'  [ERROR] qa_visual.py 不可用，无法执行 visual QA。{detail}')
            visual = _write_dependency_report(
                out_dir,
                report_name="visual_report",
                mode=mode,
                code="VISUAL_QA_UNAVAILABLE",
                message="visual QA is required but qa_visual.py is unavailable.",
                detail=detail,
                next_action="修复 visual QA 依赖后重跑；先查看 visual_report.md。",
            )
            failed_reports.append(("Visual QA", "visual_report.md", visual))
            _print_failed_report_hint(qa_report, failed_reports)
            return False
        resolved_golden_dir = os.path.abspath(golden_dir) if golden_dir else None
        visual = deps.visual_check_and_write(
            out_dir,
            output_docx_name=output_docx_name,
            project_root=project_root,
            render_all_pages=True,
            require_wps=bool(require_wps),
            golden_dir=resolved_golden_dir,
            update_golden=bool(update_golden),
        )
        _print_report_summary("Visual QA", visual, report_file="visual_report.md")
        print("  [OK] PDF 渲染 QA -> visual_report.json / visual_report.md")
        if not visual.get("passed"):
            qa_failed = True
            failed_reports.append(("Visual QA", "visual_report.md / visual_qa/samples/", visual))

    if qa_failed:
        _print_failed_report_hint(qa_report, failed_reports)
        return False

    return True
