"""QA phase orchestration for the one-click pipeline runner."""
from __future__ import annotations

from dataclasses import dataclass
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


def _print_report_summary(label, report, *, show_owner=False):
    issues, error_count = _issue_counts(report)
    status = "通过" if report.get("passed") else "未通过"
    print(f"  [{label}] {status}: {error_count} error(s), {len(issues)} issue(s)")
    for item in issues[:8]:
        print(f'   - {item.get("severity")} {item.get("code")}: {item.get("message")}')
        if show_owner:
            print(f'     修复目标: {item.get("active_owner")}')
    if len(issues) > 8:
        print(f"   ... 还有 {len(issues) - 8} 项，请看 qa_report.md")


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

    if deps.qa_check_and_write is None:
        print(f'  [ERROR] qa_checker.py 不可用，无法执行必备 QA。{_optional_detail(deps, "qa_checker")}')
        return False

    qa_report = deps.qa_check_and_write(out_dir, mode=mode, output_docx_name=output_docx_name)
    print_contract_issues("qa_report.json", validate_qa_report(qa_report))
    _print_report_summary("QA", qa_report, show_owner=True)
    print("  [OK] QA 报告 -> qa_report.json / qa_report.md")
    if not qa_report.get("passed"):
        qa_failed = True

    if qa_level in ("strict", "visual"):
        if deps.conformance_check_and_write is None:
            print(f'  [ERROR] qa_conformance.py 不可用，无法执行 strict conformance QA。{_optional_detail(deps, "qa_conformance")}')
            return False
        conformance = deps.conformance_check_and_write(
            out_dir,
            mode=mode,
            output_docx_name=output_docx_name,
            project_root=project_root,
        )
        _print_report_summary("Conformance QA", conformance)
        print("  [OK] strict conformance QA -> conformance_report.json / conformance_report.md")
        if not conformance.get("passed"):
            qa_failed = True

    if qa_level == "visual":
        if deps.visual_check_and_write is None:
            print(f'  [ERROR] qa_visual.py 不可用，无法执行 visual QA。{_optional_detail(deps, "qa_visual")}')
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
        _print_report_summary("Visual QA", visual)
        print("  [OK] PDF 渲染 QA -> visual_report.json / visual_report.md")
        if not visual.get("passed"):
            qa_failed = True

    if qa_failed:
        if qa_report:
            print_repair_hint(qa_report, out_dir)
        print("  [ERROR] QA 未通过。已保留输出目录和最终论文初稿；请按 qa_repair_plan.md 修复后重跑。")
        return False

    return True
