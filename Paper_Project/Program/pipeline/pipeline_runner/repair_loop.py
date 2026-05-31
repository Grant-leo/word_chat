"""Controlled user-level repair loop for generated build scripts.

The loop is intentionally conservative: it may rebuild the current output and
may patch only ``Outputs/<run>/build_generated.py`` with known-safe repairs.
Reusable engine changes remain developer work outside this module.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
import json
import os
from pathlib import Path
import re
from typing import Any, Callable, Dict, List


NEEDS_USER_AUTO_LEVELS = {
    "needs_user_file",
    "needs_user_input",
    "needs_user_confirmation",
    "optional_user_input",
}
NEEDS_USER_CODES = {
    "CONTENT_EMPTY",
    "CONTENT_IMAGE_MISSING",
    "IMAGE_EXTRACT_FAILED",
    "NON_BODY_IMAGE_UNSUPPORTED",
    "PDF_TEMPLATE_UNSUPPORTED",
    "MISSING_CONTENT_JSON",
    "MISSING_FORMAT_JSON",
}
REBUILD_ONLY_CODES = {"MISSING_DOCX", "DOCX_XML_UNREADABLE"}
PLACEHOLDER_MARKER = "# AUTO_REPAIR_PLACEHOLDER_CLEANUP_V1"
REFERENCE_EAST_ASIA_MARKER = "# AUTO_REPAIR_REFERENCE_EAST_ASIA_FONT_V1"
REPORT_BLOCKER_GUIDES = {
    "CONFORMANCE_QA_UNAVAILABLE": {
        "auto_level": "needs_user_input",
        "user_action": "修复 strict conformance QA 依赖后重跑；先查看 conformance_report.md。",
    },
    "VISUAL_QA_UNAVAILABLE": {
        "auto_level": "needs_user_input",
        "user_action": "修复 visual QA 依赖后重跑；先查看 visual_report.md。",
    },
    "PDF_EXPORT_FAILED": {
        "auto_level": "needs_user_input",
        "user_action": "安装或修复 Microsoft Word COM/PDF 导出环境后重跑 visual QA。",
    },
    "PDFINFO_UNAVAILABLE": {
        "auto_level": "needs_user_input",
        "user_action": "安装 Poppler 的 pdfinfo 后重跑 visual QA。",
    },
    "PDFINFO_FAILED": {
        "auto_level": "needs_user_input",
        "user_action": "检查 Poppler/pdfinfo 是否可运行，或打开 visual_report.md 查看 PDF 元信息读取错误。",
    },
    "PDF_PAGE_COUNT_INVALID": {
        "auto_level": "needs_user_confirmation",
        "user_action": "打开导出的 PDF/DOCX 确认是否为空白；若导出异常，先修复 Word/PDF 导出环境后重跑。",
    },
    "PDFTOTEXT_UNAVAILABLE": {
        "auto_level": "needs_user_input",
        "user_action": "安装 Poppler 的 pdftotext 后重跑 visual QA。",
    },
    "PDFTOTEXT_FAILED": {
        "auto_level": "needs_user_input",
        "user_action": "检查 Poppler/pdftotext 是否可运行，或打开 visual_report.md 查看文本提取错误。",
    },
    "SAMPLE_RENDER_FAILED": {
        "auto_level": "needs_user_input",
        "user_action": "安装 Poppler 的 pdftoppm 后重跑 visual QA。",
    },
    "ALL_PAGE_RENDER_FAILED": {
        "auto_level": "needs_user_input",
        "user_action": "安装或修复 Poppler 的 pdftoppm 后重跑 visual QA。",
    },
    "PAGE_IMAGE_UNREADABLE": {
        "auto_level": "needs_user_input",
        "user_action": "检查 visual_qa/all_pages 下的 PNG 是否损坏，并修复 PDF 渲染/图像读取环境后重跑。",
    },
    "WPS_EXPORT_UNAVAILABLE": {
        "auto_level": "needs_user_input",
        "user_action": "安装/配置 WPS COM，或取消 --require-wps 后重跑 visual QA。",
    },
    "WPS_PAGE_COUNT_MISMATCH": {
        "auto_level": "needs_user_confirmation",
        "user_action": "分别打开 Word 与 WPS 导出的 PDF 比对分页差异；确认是兼容性差异还是排版脚本问题后再修复。",
    },
    "GOLDEN_BASELINE_MISMATCH": {
        "auto_level": "needs_user_confirmation",
        "user_action": "打开 visual_report.md 和 visual_qa/samples/ 对比页面；确认变化正确后用 --update-golden 更新基线，或继续修复排版。",
    },
}


@dataclass(frozen=True)
class RepairLoopResult:
    ok: bool
    status: str
    report_path: str
    rounds: int
    final_errors: int


def run_repair_loop(
    out_dir: str,
    *,
    mode: str,
    output_docx_name: str,
    qa_level: str,
    project_root: str,
    max_rounds: int,
    stop_no_improve: int,
    deps: Any,
    run_generated_script: Callable[..., Any],
    python_executable: str,
    golden_dir: str | None = None,
    update_golden: bool = False,
    require_wps: bool = False,
) -> RepairLoopResult:
    """Run a bounded build-script repair loop and write audit reports."""

    max_rounds = max(1, int(max_rounds or 1))
    stop_no_improve = max(1, int(stop_no_improve or 1))
    out_path = Path(out_dir).resolve()
    history: List[Dict[str, Any]] = []

    state = _ensure_qa_state(
        out_path,
        mode=mode,
        output_docx_name=output_docx_name,
        qa_level=qa_level,
        project_root=project_root,
        deps=deps,
        golden_dir=golden_dir,
        update_golden=update_golden,
        require_wps=require_wps,
    )
    history.append(_round_snapshot(0, "initial", state, actions=[], previous_error_codes=[]))

    if state["total_errors"] == 0:
        return _finish(
            out_path,
            ok=True,
            status="converged",
            mode=mode,
            output_docx_name=output_docx_name,
            qa_level=qa_level,
            max_rounds=max_rounds,
            stop_no_improve=stop_no_improve,
            history=history,
        )

    best_errors = state["total_errors"]
    stagnant_rounds = 0

    for round_no in range(1, max_rounds + 1):
        blockers = _blocking_steps(state)
        if blockers:
            history.append(_round_snapshot(round_no, "blocked", state, actions=[], previous_error_codes=history[-1].get("error_codes") or []))
            status, stop_detail = _blocker_stop_reason(blockers)
            return _finish(
                out_path,
                ok=False,
                status=status,
                mode=mode,
                output_docx_name=output_docx_name,
                qa_level=qa_level,
                max_rounds=max_rounds,
                stop_no_improve=stop_no_improve,
                history=history,
                stop_detail=stop_detail,
                blockers=blockers,
            )

        actions = _apply_safe_repairs(out_path, state)
        if not actions:
            history.append(_round_snapshot(round_no, "no_supported_auto_repair", state, actions=[], previous_error_codes=history[-1].get("error_codes") or []))
            return _finish(
                out_path,
                ok=False,
                status="stopped_no_supported_auto_repair",
                mode=mode,
                output_docx_name=output_docx_name,
                qa_level=qa_level,
                max_rounds=max_rounds,
                stop_no_improve=stop_no_improve,
                history=history,
                stop_detail="No known-safe build_generated.py repair matched the current QA errors.",
            )

        build_result = _rebuild_current_docx(
            out_path,
            output_docx_name=output_docx_name,
            run_generated_script=run_generated_script,
            python_executable=python_executable,
        )
        if build_result["returncode"] != 0:
            history.append(_round_snapshot(round_no, "build_failed", state, actions=actions, build=build_result, previous_error_codes=history[-1].get("error_codes") or []))
            return _finish(
                out_path,
                ok=False,
                status="stopped_build_failed",
                mode=mode,
                output_docx_name=output_docx_name,
                qa_level=qa_level,
                max_rounds=max_rounds,
                stop_no_improve=stop_no_improve,
                history=history,
                stop_detail=build_result.get("stderr") or build_result.get("stdout") or "build_generated.py failed",
            )

        state = _run_qa_state(
            out_path,
            mode=mode,
            output_docx_name=output_docx_name,
            qa_level=qa_level,
            project_root=project_root,
            deps=deps,
            golden_dir=golden_dir,
            update_golden=update_golden,
            require_wps=require_wps,
        )
        history.append(_round_snapshot(round_no, "after_repair", state, actions=actions, build=build_result, previous_error_codes=history[-1].get("error_codes") or []))

        if state["total_errors"] == 0:
            return _finish(
                out_path,
                ok=True,
                status="converged",
                mode=mode,
                output_docx_name=output_docx_name,
                qa_level=qa_level,
                max_rounds=max_rounds,
                stop_no_improve=stop_no_improve,
                history=history,
            )

        if state["total_errors"] < best_errors:
            best_errors = state["total_errors"]
            stagnant_rounds = 0
        else:
            stagnant_rounds += 1
        if stagnant_rounds >= stop_no_improve:
            return _finish(
                out_path,
                ok=False,
                status="stopped_no_improvement",
                mode=mode,
                output_docx_name=output_docx_name,
                qa_level=qa_level,
                max_rounds=max_rounds,
                stop_no_improve=stop_no_improve,
                history=history,
                stop_detail=f"QA error count did not improve for {stagnant_rounds} consecutive repair rounds.",
            )

    return _finish(
        out_path,
        ok=False,
        status="stopped_max_rounds",
        mode=mode,
        output_docx_name=output_docx_name,
        qa_level=qa_level,
        max_rounds=max_rounds,
        stop_no_improve=stop_no_improve,
        history=history,
        stop_detail=f"Reached max repair rounds: {max_rounds}.",
    )


def _ensure_qa_state(out_path: Path, **kwargs: Any) -> Dict[str, Any]:
    if not (out_path / "qa_report.json").exists() or _required_report_missing(out_path, kwargs["qa_level"]):
        return _run_qa_state(out_path, **kwargs)
    return _read_state(out_path, kwargs["qa_level"])


def _required_report_missing(out_path: Path, qa_level: str) -> bool:
    if qa_level in ("strict", "visual") and not (out_path / "conformance_report.json").exists():
        return True
    if qa_level == "visual" and not (out_path / "visual_report.json").exists():
        return True
    return False


def _run_qa_state(
    out_path: Path,
    *,
    mode: str,
    output_docx_name: str,
    qa_level: str,
    project_root: str,
    deps: Any,
    golden_dir: str | None = None,
    update_golden: bool = False,
    require_wps: bool = False,
) -> Dict[str, Any]:
    if deps.qa_check_and_write is None:
        raise RuntimeError("qa_checker.py is required for auto repair")
    deps.qa_check_and_write(str(out_path), mode=mode, output_docx_name=output_docx_name)
    if qa_level in ("strict", "visual") and deps.conformance_check_and_write is None:
        _write_dependency_report(
            out_path,
            report_name="conformance_report",
            mode=mode,
            code="CONFORMANCE_QA_UNAVAILABLE",
            message="strict conformance QA is required but qa_conformance.py is unavailable.",
            detail=_optional_detail(deps, "qa_conformance"),
        )
        return _read_state(out_path, qa_level)
    if qa_level in ("strict", "visual"):
        deps.conformance_check_and_write(
            str(out_path),
            mode=mode,
            output_docx_name=output_docx_name,
            project_root=project_root,
        )
    if qa_level == "visual" and deps.visual_check_and_write is None:
        _write_dependency_report(
            out_path,
            report_name="visual_report",
            mode=mode,
            code="VISUAL_QA_UNAVAILABLE",
            message="visual QA is required but qa_visual.py is unavailable.",
            detail=_optional_detail(deps, "qa_visual"),
        )
        return _read_state(out_path, qa_level)
    if qa_level == "visual":
        deps.visual_check_and_write(
            str(out_path),
            output_docx_name=output_docx_name,
            project_root=project_root,
            render_all_pages=True,
            require_wps=bool(require_wps),
            golden_dir=os.path.abspath(golden_dir) if golden_dir else None,
            update_golden=bool(update_golden),
        )
    return _read_state(out_path, qa_level)


def _optional_detail(deps: Any, name: str) -> str:
    detail = getattr(deps, "optional_import_detail", None)
    if detail is None:
        return ""
    try:
        return str(detail(name) or "")
    except Exception:
        return ""


def _write_dependency_report(
    out_path: Path,
    *,
    report_name: str,
    mode: str,
    code: str,
    message: str,
    detail: str,
) -> None:
    next_action = (REPORT_BLOCKER_GUIDES.get(code) or {}).get("user_action") or "安装或修复所需 QA 依赖后，重新运行完整流水线。"
    issue = {
        "code": code,
        "severity": "error",
        "message": message,
        "detail": detail,
    }
    report = {
        "schema_version": 1,
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "mode": mode,
        "output_dir_name": out_path.name,
        "passed": False,
        "counts": {},
        "issues": [issue],
        "next_action": next_action,
    }
    (out_path / f"{report_name}.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    lines = [
        f"# {report_name.replace('_', ' ')}",
        "",
        "- 结果：未通过",
        f"- 问题码：`{code}`",
        f"- 信息：{message}",
        f"- 下一步：{next_action}",
    ]
    if detail:
        lines.append(f"- 细节：`{detail}`")
    (out_path / f"{report_name}.md").write_text("\n".join(lines) + "\n", encoding="utf-8")


def _read_state(out_path: Path, qa_level: str) -> Dict[str, Any]:
    qa_report = _load_json(out_path / "qa_report.json", {})
    repair_plan = _load_json(out_path / "qa_repair_plan.json", qa_report.get("repair_plan") or {})
    reports: Dict[str, Dict[str, Any]] = {"qa": qa_report}
    if qa_level in ("strict", "visual"):
        reports["conformance"] = _load_json(out_path / "conformance_report.json", {})
    if qa_level == "visual":
        reports["visual"] = _load_json(out_path / "visual_report.json", {})
    issues = []
    for label, report in reports.items():
        for item in report.get("issues") or []:
            issue = dict(item)
            issue["report"] = label
            issues.append(issue)
    errors = [item for item in issues if item.get("severity") == "error"]
    warnings = [item for item in issues if item.get("severity") == "warning"]
    passed = bool(qa_report.get("passed")) and all(
        bool(report.get("passed", True)) for label, report in reports.items() if label != "qa" and report
    )
    return {
        "passed": passed,
        "reports": reports,
        "qa_report": qa_report,
        "repair_plan": repair_plan,
        "issues": issues,
        "errors": errors,
        "warnings": warnings,
        "error_codes": [str(item.get("code") or "") for item in errors],
        "warning_codes": [str(item.get("code") or "") for item in warnings],
        "qa_errors": sum(1 for item in qa_report.get("issues") or [] if item.get("severity") == "error"),
        "qa_warnings": sum(1 for item in qa_report.get("issues") or [] if item.get("severity") == "warning"),
        "total_errors": len(errors),
        "total_warnings": len(warnings),
    }


def _blocking_steps(state: Dict[str, Any]) -> List[Dict[str, Any]]:
    steps = state.get("repair_plan", {}).get("steps") or []
    error_codes = set(state.get("error_codes") or [])
    blockers = []
    for step in steps:
        code = str(step.get("code") or "")
        auto_level = str(step.get("auto_level") or "")
        severity = str(step.get("severity") or "")
        if code in error_codes and severity == "error" and (auto_level in NEEDS_USER_AUTO_LEVELS or code in NEEDS_USER_CODES):
            blockers.append(
                {
                    "code": code,
                    "auto_level": auto_level,
                    "user_action": step.get("user_action") or "",
                    "detail": step.get("detail") or "",
                }
            )
    blocked_codes = {item.get("code") for item in blockers}
    for issue in state.get("errors") or []:
        code = str(issue.get("code") or "")
        if not code or code in blocked_codes:
            continue
        guide = REPORT_BLOCKER_GUIDES.get(code)
        if not guide:
            continue
        blockers.append(
            {
                "code": code,
                "auto_level": guide.get("auto_level") or "needs_user_input",
                "user_action": guide.get("user_action") or "查看对应 QA 报告后重跑。",
                "detail": issue.get("detail") or issue.get("message") or "",
                "report": issue.get("report") or "",
            }
        )
    return blockers


def _blocker_stop_reason(blockers: List[Dict[str, Any]]) -> tuple[str, str]:
    levels = {str(item.get("auto_level") or "") for item in blockers}
    codes = {str(item.get("code") or "") for item in blockers}
    if (levels and levels <= {"needs_user_file"}) or (codes and codes <= NEEDS_USER_CODES):
        return "stopped_needs_user_file", "User file/input is required before automatic repair can continue."
    return "stopped_needs_user_input", "User confirmation, dependency repair, or manual inspection is required before automatic repair can continue."


def _apply_safe_repairs(out_path: Path, state: Dict[str, Any]) -> List[Dict[str, Any]]:
    codes = set(state.get("error_codes") or [])
    actions: List[Dict[str, Any]] = []
    if "PLACEHOLDER_TEXT_LEFT" in codes:
        action = _patch_placeholder_cleanup(out_path)
        if action:
            actions.append(action)
    if "STYLE_MISMATCH" in codes:
        action = _patch_reference_east_asia_fonts(out_path, state)
        if action:
            actions.append(action)
    if codes & REBUILD_ONLY_CODES:
        actions.append({"kind": "rebuild", "codes": sorted(codes & REBUILD_ONLY_CODES)})
    return actions


def _patch_placeholder_cleanup(out_path: Path) -> Dict[str, Any] | None:
    build_path = _safe_build_path(out_path)
    text = build_path.read_text(encoding="utf-8")
    if PLACEHOLDER_MARKER in text:
        return {"kind": "script_patch", "code": "PLACEHOLDER_TEXT_LEFT", "status": "already_applied"}
    needle = "if __name__ == '__main__':\n    main()\n"
    if needle not in text:
        return None
    patched = text.replace(needle, _placeholder_cleanup_fragment() + "\n" + needle.replace("main()", "main()\n    _auto_repair_after_main()"))
    build_path.write_text(patched, encoding="utf-8")
    return {
        "kind": "script_patch",
        "code": "PLACEHOLDER_TEXT_LEFT",
        "target": "build_generated.py",
        "summary": "Injected a post-build cleanup that removes obvious placeholder paragraphs from the final DOCX.",
    }


def _placeholder_cleanup_fragment() -> str:
    return r'''
# AUTO_REPAIR_PLACEHOLDER_CLEANUP_V1
def _auto_repair_placeholder_hit(text):
    import re
    t = str(text or '').strip()
    if not t:
        return False
    return bool(re.search(r'(\[[^\]\n]*(?:XX|XXX|TODO|FIXME|待填|待填写|请输入|报名|序号|姓名|学号|学院|专业|班级|题目|指导|教师|日期|编码)[^\]\n]*\])|(\{\{[^}]+\}\}|TODO|FIXME|待填写|待补全|XXXX)', t, re.I))


def _auto_repair_delete_paragraph(paragraph):
    element = paragraph._element
    parent = element.getparent()
    if parent is not None:
        parent.remove(element)


def _auto_repair_clean_placeholders(docx_path):
    import os
    from docx import Document as _AutoRepairDocument
    if not os.path.exists(docx_path):
        return 0
    docx = _AutoRepairDocument(docx_path)
    removed = 0
    for paragraph in list(docx.paragraphs):
        if _auto_repair_placeholder_hit(paragraph.text):
            _auto_repair_delete_paragraph(paragraph)
            removed += 1
    for table in docx.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in list(cell.paragraphs):
                    if _auto_repair_placeholder_hit(paragraph.text):
                        paragraph.clear()
                        removed += 1
    if removed:
        docx.save(docx_path)
    return removed


def _auto_repair_after_main():
    try:
        removed = _auto_repair_clean_placeholders(OUT)
        if removed:
            print(f'Auto repair: removed {removed} placeholder paragraph(s)')
    except Exception as exc:
        print(f'Auto repair warning: placeholder cleanup failed: {exc}')
'''


def _patch_reference_east_asia_fonts(out_path: Path, state: Dict[str, Any]) -> Dict[str, Any] | None:
    details = "\n".join(str(item.get("detail") or "") for item in state.get("errors") or [])
    if "reference" not in details or "eastAsia font" not in details:
        return None
    build_path = _safe_build_path(out_path)
    text = build_path.read_text(encoding="utf-8")
    if REFERENCE_EAST_ASIA_MARKER in text:
        return {"kind": "script_patch", "code": "STYLE_MISMATCH", "status": "already_applied"}
    old = """def add_reference_mixed_runs(p, text, prof):
    # Chinese parts use the role's CJK font; Latin/numeric punctuation uses Times New Roman.
    for seg in re.findall(r'[\\u4e00-\\u9fff]+|[^\\u4e00-\\u9fff]+', text):
        r = p.add_run(seg)
        if has_cjk(seg):
            apply_run_profile(r, prof, seg, force_latin='Times New Roman')
        else:
            p_latin = dict(prof); p_latin['font'] = 'Times New Roman'
            apply_run_profile(r, p_latin, seg, force_latin='Times New Roman')
"""
    new = """def add_reference_mixed_runs(p, text, prof):
    # AUTO_REPAIR_REFERENCE_EAST_ASIA_FONT_V1
    # Keep the template CJK font in w:eastAsia for numeric/Latin label runs.
    for seg in re.findall(r'[\\u4e00-\\u9fff]+|[^\\u4e00-\\u9fff]+', text):
        r = p.add_run(seg)
        apply_run_profile(r, prof, seg, force_latin='Times New Roman')
"""
    if old not in text:
        return None
    build_path.write_text(text.replace(old, new, 1), encoding="utf-8")
    return {
        "kind": "script_patch",
        "code": "STYLE_MISMATCH",
        "target": "build_generated.py",
        "summary": "Patched reference rendering so numeric/Latin runs preserve the role CJK eastAsia font.",
    }


def _safe_build_path(out_path: Path) -> Path:
    build_path = (out_path / "build_generated.py").resolve()
    if build_path.name != "build_generated.py" or build_path.parent != out_path.resolve():
        raise RuntimeError(f"Unsafe build script path: {build_path}")
    if not build_path.exists():
        raise RuntimeError(f"Missing build script: {build_path}")
    return build_path


def _rebuild_current_docx(
    out_path: Path,
    *,
    output_docx_name: str,
    run_generated_script: Callable[..., Any],
    python_executable: str,
) -> Dict[str, Any]:
    build_path = _safe_build_path(out_path)
    result = run_generated_script(str(build_path), str(out_path), python_executable=python_executable)
    return {
        "returncode": int(getattr(result, "returncode", 1)),
        "stdout": _sanitize_log(str(getattr(result, "stdout", "")), out_path)[-2000:],
        "stderr": _sanitize_log(str(getattr(result, "stderr", "")), out_path)[-2000:],
        "docx_exists": (out_path / output_docx_name).exists(),
    }


def _round_snapshot(
    round_no: int,
    phase: str,
    state: Dict[str, Any],
    *,
    actions: List[Dict[str, Any]],
    build: Dict[str, Any] | None = None,
    previous_error_codes: List[str] | None = None,
) -> Dict[str, Any]:
    previous_codes = set(previous_error_codes or [])
    current_codes = set(state.get("error_codes") or [])
    return {
        "round": round_no,
        "phase": phase,
        "passed": state.get("passed"),
        "total_errors": state.get("total_errors", 0),
        "total_warnings": state.get("total_warnings", 0),
        "qa_errors": state.get("qa_errors", 0),
        "qa_warnings": state.get("qa_warnings", 0),
        "error_codes": list(state.get("error_codes") or []),
        "warning_codes": list(state.get("warning_codes") or []),
        "new_error_codes": sorted(current_codes - previous_codes),
        "resolved_error_codes": sorted(previous_codes - current_codes),
        "actions": actions,
        "build": build or {},
    }


def _finish(
    out_path: Path,
    *,
    ok: bool,
    status: str,
    mode: str,
    output_docx_name: str,
    qa_level: str,
    max_rounds: int,
    stop_no_improve: int,
    history: List[Dict[str, Any]],
    stop_detail: str = "",
    blockers: List[Dict[str, Any]] | None = None,
) -> RepairLoopResult:
    final = history[-1] if history else {}
    display_out = _report_display_path(out_path)
    display_docx = _join_report_path(display_out, output_docx_name)
    final_warnings = int(final.get("total_warnings") or 0)
    manual_checks = ["用 Word/WPS 打开最终 DOCX，核对分页、图片、公式、表格和目录。"]
    if final_warnings:
        manual_checks.append("查看 qa_report.md 和 repair_loop_report.md 中的剩余 warning，确认不会影响交付。")
    report = {
        "schema_version": 1,
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "status": status,
        "ok": bool(ok),
        "mode": mode,
        "qa_level": qa_level,
        "output_dir": display_out,
        "final_docx": display_docx,
        "max_rounds": max_rounds,
        "stop_no_improve": stop_no_improve,
        "rounds_run": max(0, len([r for r in history if r.get("round", 0) > 0])),
        "final_errors": int(final.get("total_errors") or 0),
        "final_warnings": final_warnings,
        "final_error_codes": final.get("error_codes") or [],
        "final_warning_codes": final.get("warning_codes") or [],
        "warning_policy": (
            "剩余 warning 不会阻断自动 QA 收敛，但仍可能影响交付质量，交付前必须用 Word/WPS 人工确认。"
            if final_warnings
            else "当前启用的 QA 未报告剩余 warning。"
        ),
        "stop_detail": stop_detail,
        "blockers": blockers or [],
        "manual_check_required": manual_checks,
        "rounds": history,
    }
    report_path = out_path / "repair_loop_report.json"
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    (out_path / "repair_loop_report.md").write_text(_report_to_markdown(report), encoding="utf-8")
    return RepairLoopResult(
        ok=bool(ok),
        status=status,
        report_path=_join_report_path(display_out, "repair_loop_report.json"),
        rounds=int(report["rounds_run"]),
        final_errors=int(report["final_errors"]),
    )


def _report_to_markdown(report: Dict[str, Any]) -> str:
    lines = [
        "# 自动修复闭环报告",
        "",
        f"- 状态：`{report.get('status')}`",
        f"- 结果：`{'已收敛' if report.get('ok') else '已停止'}`",
        f"- 模式：`{report.get('mode')}`",
        f"- QA 等级：`{report.get('qa_level')}`",
        f"- 最终 DOCX：`{report.get('final_docx')}`",
        f"- 最终错误：`{report.get('final_errors')}`",
        f"- 最终警告：`{report.get('final_warnings')}`",
    ]
    if report.get("stop_detail"):
        lines.append(f"- 停止原因：{report.get('stop_detail')}")
    lines.append(f"- warning 处理策略：{report.get('warning_policy')}")
    blockers = report.get("blockers") or []
    if blockers:
        lines.extend(["", "## 需要用户补充", ""])
        for item in blockers:
            lines.append(f"- `{item.get('code')}` ({item.get('auto_level')}): {item.get('user_action')}")
    lines.extend(["", "## 修复轮次", ""])
    for item in report.get("rounds") or []:
        lines.append(
            f"- 第 `{item.get('round')}` 轮 `{item.get('phase')}`："
            f"错误 `{item.get('total_errors')}`，警告 `{item.get('total_warnings')}`，"
            f"错误码 `{', '.join(item.get('error_codes') or []) or '-'}`"
        )
        for action in item.get("actions") or []:
            lines.append(f"  - 动作：`{action.get('kind')}` {action.get('summary') or action.get('code') or action.get('codes')}")
    lines.extend([
        "",
        "## 人工检查",
        "",
        "- 自动 QA 收敛不等于 100% 正确。",
        "- 用 Word/WPS 打开最终 DOCX，核对分页、图片、公式、表格和目录。",
    ])
    if int(report.get("final_warnings") or 0):
        lines.append("- 交付前查看 qa_report.md 和 repair_loop_report.md 中的剩余 warning。")
    else:
        lines.append("- 当前启用的 QA 未报告剩余 warning。")
    lines.append("")
    return "\n".join(lines)


def _report_display_path(path: Path) -> str:
    parts = list(path.parts)
    if "Outputs" in parts:
        return "/".join(parts[parts.index("Outputs"):])
    return path.name


def _join_report_path(parent: str, child: str) -> str:
    return (str(parent).rstrip("/\\") + "/" + str(child).lstrip("/\\")).replace("\\", "/")


def _sanitize_log(text: str, out_path: Path) -> str:
    safe = _report_display_path(out_path)
    raw = str(out_path)
    return str(text or "").replace(raw, safe).replace(raw.replace("\\", "/"), safe)


def _load_json(path: Path, default: Any) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return default
