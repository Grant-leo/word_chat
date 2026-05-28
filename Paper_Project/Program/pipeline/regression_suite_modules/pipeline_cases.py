"""Pipeline runner regression cases."""
from __future__ import annotations

import json
import sys
from pathlib import Path
from types import SimpleNamespace
from typing import Any, Dict, List

from docx import Document
from pipeline_runner.artifacts import build_content_markdown, write_content_artifacts, write_format_artifacts
from pipeline_runner.build_phase import generate_and_build_docx_phase
from pipeline_runner.cli import build_arg_parser, dispatch_cli
from pipeline_runner.contracts import (
    has_contract_errors,
    validate_build_manifest,
    validate_content_data,
    validate_format_data,
    validate_qa_report,
)
from pipeline_runner.context import (
    create_unique_output_dir,
    normalize_qa_level,
    resolve_inputs,
    write_workflow_mode,
)
from pipeline_runner.dependencies import load_optional_dependencies
from pipeline_runner.execution import ScriptExecutionResult, run_generated_script
from pipeline_runner.io import normalize_mode, scan_inputs
from pipeline_runner.qa import QADependencies, run_qa_phases
from pipeline_runner.repair_loop import run_repair_loop
from pipeline_runner.summary import build_completion_summary
from pipeline_runner.template_phase import write_template_profile_phase, write_template_requirements_phase
from pipeline_runner.verification import VerificationError, _merge_content_results, double_verify

from regression_suite_modules.generated_pdf import poppler_available, write_text_pdf
from regression_suite_modules.harness import (
    PNG_1X1,
    assert_true,
    base_content,
    base_format,
    case,
    fail,
    new_workdir,
)

PIPELINE_DIR = Path(__file__).resolve().parents[1]
REPO_ROOT = Path(__file__).resolve().parents[4]


def _write_repair_loop_fixture(work: Path, *, final_docx: str = "final.docx") -> Path:
    work.mkdir(exist_ok=True)
    (work / "format.json").write_text(json.dumps(base_format(), ensure_ascii=False, indent=2), encoding="utf-8")
    (work / "content.json").write_text(
        json.dumps(base_content(["This body paragraph should remain in the final document."]), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    write_workflow_mode(
        str(work),
        mode="user",
        template_path="template.docx",
        content_path="content.docx",
        run_qa=True,
        qa_level="basic",
        golden_dir=None,
        update_golden=False,
        require_wps=False,
        auto_repair=True,
        repair_max_rounds=5,
        repair_stop_no_improve=2,
    )
    build_script = f"""from docx import Document
import json
import os

BASE = os.path.dirname(__file__)
OUT = os.path.join(BASE, {final_docx!r})

def main():
    doc = Document()
    doc.add_paragraph('1 Introduction')
    doc.add_paragraph('This body paragraph should remain in the final document.')
    doc.add_paragraph('TODO')
    doc.save(OUT)
    with open(os.path.join(BASE, 'build_manifest.json'), 'w', encoding='utf-8') as handle:
        json.dump({{'schema_version': 1, 'counts': {{'content_images_rendered': 0, 'content_tables_rendered': 0, 'content_formulas_rendered': 0}}}}, handle)

if __name__ == '__main__':
    main()
"""
    path = work / "build_generated.py"
    path.write_text(build_script, encoding="utf-8")
    return path


def _fake_repair_report(out_dir: str, *, code: str, severity: str = "error", message: str = "synthetic") -> Dict[str, Any]:
    from qa_checker_modules.repair import build_repair_plan
    from qa_checker_modules.reports import write_reports

    report = {
        "schema_version": 1,
        "mode": "user",
        "passed": severity != "error",
        "counts": {},
        "issues": [{"code": code, "severity": severity, "message": message, "detail": ""}],
        "next_action": "repair",
        "output_dir_name": Path(out_dir).name,
    }
    report["repair_plan"] = build_repair_plan(report, out_dir)
    write_reports(report, out_dir)
    return report


def _write_conformance_report(out_dir: str, *, passed: bool) -> Dict[str, Any]:
    issues = [] if passed else [
        {
            "code": "STYLE_MISMATCH",
            "severity": "error",
            "message": "Some final DOCX paragraphs do not match template role styles.",
            "detail": "reference `[1] 作者1.`: eastAsia font Times New Roman != 宋体",
        }
    ]
    report = {
        "schema_version": 1,
        "mode": "user",
        "passed": passed,
        "counts": {},
        "issues": issues,
        "next_action": "Fix Outputs/<run>/build_generated.py and rerun it.",
    }
    Path(out_dir, "conformance_report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return report


@case
def pipeline_contracts_accept_current_handoffs() -> None:
    manifest = {
        "schema_version": 1,
        "counts": {
            "content_images_rendered": 0,
            "content_tables_rendered": 1,
            "content_formulas_rendered": 2,
        },
    }
    qa_report = {
        "schema_version": 1,
        "mode": "developer",
        "passed": True,
        "counts": {},
        "issues": [],
        "next_action": "ok",
    }
    all_issues = (
        validate_format_data(base_format())
        + validate_content_data(base_content(["Body text"]))
        + validate_build_manifest(manifest)
        + validate_qa_report(qa_report)
    )
    assert_true(not has_contract_errors(all_issues), "valid current handoff structures should satisfy contracts")


@case
def pipeline_contracts_report_structural_errors() -> None:
    issues = (
        validate_format_data({"paragraphs": {}})
        + validate_content_data({"sections": {}})
        + validate_build_manifest({"counts": {"content_images_rendered": -1}})
        + validate_qa_report({"passed": "yes", "counts": [], "issues": [{"code": ""}]})
    )
    codes = {issue.code for issue in issues}
    assert_true("FORMAT_SECTIONS_MISSING" in codes, "missing format sections was not reported")
    assert_true("FORMAT_PARAGRAPHS_NOT_LIST" in codes, "invalid format paragraphs was not reported")
    assert_true("CONTENT_SECTIONS_NOT_LIST" in codes, "invalid content sections was not reported")
    assert_true("MANIFEST_COUNT_INVALID" in codes, "invalid manifest count was not reported")
    assert_true("QA_REPORT_PASSED_NOT_BOOL" in codes, "invalid QA passed flag was not reported")
    assert_true("QA_REPORT_ISSUE_CODE_MISSING" in codes, "invalid QA issue code was not reported")
    assert_true(has_contract_errors(issues), "structural contract errors should be marked as errors")


@case
def pipeline_io_helpers_scan_inputs_and_normalize_modes() -> None:
    work = new_workdir("pipeline_io")
    (work / "paper.docx").write_text("x", encoding="utf-8")
    (work / "notes.md").write_text("x", encoding="utf-8")
    (work / "template.PDF").write_text("x", encoding="utf-8")
    (work / "~$paper.docx").write_text("x", encoding="utf-8")
    (work / "ignore.txt").write_text("x", encoding="utf-8")
    assert_true(scan_inputs(str(work)) == ["notes.md", "paper.docx"], "scan_inputs should skip temp and unsupported files")
    assert_true(scan_inputs(str(work), exts=(".docx", ".pdf")) == ["paper.docx", "template.PDF"], "scan_inputs should respect extension filters case-insensitively")
    assert_true(normalize_mode("developer") == "developer", "developer mode should be preserved")
    assert_true(normalize_mode("bad-mode") == "user", "unknown mode should fall back to user")


@case
def pipeline_cli_dispatches_md_parameter_and_single_file_modes() -> None:
    parser = build_arg_parser()
    calls: List[Dict[str, Any]] = []

    def fake_run(template_file, content_file, **kwargs):
        calls.append({"template": template_file, "content": content_file, **kwargs})
        return "ok"

    def expect_exit_zero(args, template_dir="", inputs_dir=""):
        try:
            dispatch_cli(args, run_pipeline=fake_run, template_dir=template_dir, inputs_dir=inputs_dir)
        except SystemExit as exc:
            assert_true(exc.code == 0, f"dispatch returned nonzero exit: {exc.code}")
            return
        fail("dispatch should exit through exit_from_result")

    expect_exit_zero(parser.parse_args(["--md", "paper.md", "--mode", "developer", "--qa-level", "basic", "--no-qa"]))
    assert_true(calls[-1]["template"] is None and calls[-1]["content"] is None, "MD mode should not require template/content")
    assert_true(calls[-1]["md_file"] == "paper.md", "MD mode did not pass md_file")
    assert_true(calls[-1]["mode"] == "developer" and calls[-1]["run_qa"] is False, "MD mode options were not preserved")

    expect_exit_zero(parser.parse_args(["--template", "t.docx", "--content", "c.docx"]))
    assert_true(calls[-1]["template"] == "t.docx" and calls[-1]["content"] == "c.docx", "parameter mode files changed")
    assert_true(calls[-1]["mode"] == "user", "non-interactive auto mode should default to user")

    expect_exit_zero(parser.parse_args(["--template", "t.docx", "--content", "c.docx", "--auto-repair", "--no-qa", "--repair-max-rounds", "4"]))
    assert_true(calls[-1]["run_qa"] is True and calls[-1]["auto_repair"] is True, "auto repair should force QA on")
    assert_true(calls[-1]["repair_max_rounds"] == 4, "auto repair max rounds option was not passed")

    work = new_workdir("pipeline_cli")
    template_dir = work / "Templates"
    inputs_dir = work / "Inputs"
    template_dir.mkdir()
    inputs_dir.mkdir()
    (template_dir / "only.pdf").write_bytes(b"")
    (inputs_dir / "only.md").write_text("# Paper", encoding="utf-8")
    expect_exit_zero(parser.parse_args(["--mode", "user"]), template_dir=str(template_dir), inputs_dir=str(inputs_dir))
    assert_true(calls[-1]["template"] == "only.pdf" and calls[-1]["content"] == "only.md", "single-file interactive dispatch changed")


@case
def pipeline_context_resolves_inputs_outputs_and_workflow() -> None:
    work = new_workdir("pipeline_context")
    template_dir = work / "Templates"
    inputs_dir = work / "Inputs"
    outputs_dir = work / "Outputs"
    template_dir.mkdir()
    inputs_dir.mkdir()
    outputs_dir.mkdir()
    (template_dir / "template.PDF").write_text("x", encoding="utf-8")
    (inputs_dir / "paper.md").write_text("# Paper", encoding="utf-8")

    resolution = resolve_inputs("template.PDF", "paper.md", None, str(template_dir), str(inputs_dir))
    assert_true(resolution.ok, f"expected inputs to resolve: {resolution.error}")
    assert_true(resolution.inputs.use_md_content is True, "markdown content flag was not set")
    assert_true(resolution.inputs.use_md_format is False, "PDF template should use format_extractor, not md_parser")
    assert_true(resolution.inputs.content_name == "paper", "content stem was not preserved")
    assert_true(normalize_qa_level("bad-level") == "strict", "invalid QA level should fall back to strict")

    first_dir, first_name = create_unique_output_dir(str(outputs_dir), "paper", today="2026-05-26")
    second_dir, second_name = create_unique_output_dir(str(outputs_dir), "paper", today="2026-05-26")
    assert_true(first_name == "2026-05-26_paper", "first output folder name changed")
    assert_true(second_name == "2026-05-26_paper_2", "duplicate output folder suffix changed")
    workflow_path = write_workflow_mode(
        first_dir,
        mode="developer",
        template_path=resolution.inputs.template_path,
        content_path=resolution.inputs.content_path,
        run_qa=True,
        qa_level="visual",
        golden_dir="Golden",
        update_golden=False,
        require_wps=True,
    )
    workflow = json.loads(Path(workflow_path).read_text(encoding="utf-8"))
    assert_true(workflow["mode"] == "developer", "workflow mode was not written")
    assert_true(workflow["qa_level"] == "visual", "workflow QA level was not written")
    assert_true(workflow["require_wps"] is True, "workflow require_wps flag was not written")
    assert_true(workflow["auto_repair"] is False, "workflow auto_repair default changed")


@case
def pipeline_auto_repair_patches_build_script_and_reruns_qa() -> None:
    from qa_checker import check_and_write as qa_check_and_write

    work = new_workdir("pipeline_auto_repair_placeholder")
    private_dir = work / "Inputs"
    private_dir.mkdir()
    private_file = private_dir / "private.docx"
    private_file.write_bytes(b"private-user-content")
    private_before = private_file.read_bytes()
    core_file = PIPELINE_DIR / "pipeline_runner" / "qa.py"
    core_before = core_file.read_bytes()

    build_path = _write_repair_loop_fixture(work)
    first_build = run_generated_script(str(build_path), str(work), python_executable=sys.executable)
    assert_true(first_build.returncode == 0, f"synthetic build failed: {first_build.stderr}")

    deps = QADependencies(
        qa_check_and_write=qa_check_and_write,
        conformance_check_and_write=None,
        visual_check_and_write=None,
        optional_import_detail=lambda name: "",
    )
    result = run_repair_loop(
        str(work),
        mode="user",
        output_docx_name="final.docx",
        qa_level="basic",
        project_root=str(REPO_ROOT),
        max_rounds=3,
        stop_no_improve=2,
        deps=deps,
        run_generated_script=run_generated_script,
        python_executable=sys.executable,
    )

    assert_true(result.ok, f"auto repair did not converge: {result.status}")
    assert_true((work / "repair_loop_report.json").exists(), "repair loop JSON report missing")
    assert_true((work / "repair_loop_report.md").exists(), "repair loop markdown report missing")
    assert_true("AUTO_REPAIR_PLACEHOLDER_CLEANUP_V1" in build_path.read_text(encoding="utf-8"), "build script was not patched")
    import zipfile

    with zipfile.ZipFile(work / "final.docx") as zf:
        xml = zf.read("word/document.xml").decode("utf-8", errors="replace")
    assert_true("TODO" not in xml, "placeholder text remained after auto repair")
    qa = json.loads((work / "qa_report.json").read_text(encoding="utf-8"))
    assert_true(qa.get("passed") is True, f"QA was not rerun to passing state: {qa.get('issues')}")
    loop_report = json.loads((work / "repair_loop_report.json").read_text(encoding="utf-8"))
    assert_true(loop_report["rounds_run"] >= 1, "repair loop did not record a repair round")
    assert_true(private_file.read_bytes() == private_before, "auto repair modified a private input file")
    assert_true(core_file.read_bytes() == core_before, "auto repair modified a core engine file")


@case
def pipeline_auto_repair_patches_reference_east_asia_mismatch() -> None:
    work = new_workdir("pipeline_auto_repair_reference_font")
    build_path = _write_repair_loop_fixture(work)
    build_path.write_text(
        build_path.read_text(encoding="utf-8")
        + """
def add_reference_mixed_runs(p, text, prof):
    # Chinese parts use the role's CJK font; Latin/numeric punctuation uses Times New Roman.
    for seg in re.findall(r'[\\u4e00-\\u9fff]+|[^\\u4e00-\\u9fff]+', text):
        r = p.add_run(seg)
        if has_cjk(seg):
            apply_run_profile(r, prof, seg, force_latin='Times New Roman')
        else:
            p_latin = dict(prof); p_latin['font'] = 'Times New Roman'
            apply_run_profile(r, p_latin, seg, force_latin='Times New Roman')
""",
        encoding="utf-8",
    )
    first_build = run_generated_script(str(build_path), str(work), python_executable=sys.executable)
    assert_true(first_build.returncode == 0, f"synthetic build failed: {first_build.stderr}")

    def fake_qa(out_dir, mode="user", output_docx_name="final.docx"):
        return _fake_repair_report(out_dir, code="NO_ISSUE", severity="info")

    def fake_conformance(out_dir, mode="user", output_docx_name="final.docx", project_root=""):
        patched = "AUTO_REPAIR_REFERENCE_EAST_ASIA_FONT_V1" in Path(out_dir, "build_generated.py").read_text(encoding="utf-8")
        return _write_conformance_report(out_dir, passed=patched)

    deps = QADependencies(
        qa_check_and_write=fake_qa,
        conformance_check_and_write=fake_conformance,
        visual_check_and_write=None,
        optional_import_detail=lambda name: "",
    )
    result = run_repair_loop(
        str(work),
        mode="user",
        output_docx_name="final.docx",
        qa_level="strict",
        project_root=str(REPO_ROOT),
        max_rounds=3,
        stop_no_improve=2,
        deps=deps,
        run_generated_script=run_generated_script,
        python_executable=sys.executable,
    )
    assert_true(result.ok, f"reference font auto repair did not converge: {result.status}")
    assert_true("AUTO_REPAIR_REFERENCE_EAST_ASIA_FONT_V1" in build_path.read_text(encoding="utf-8"), "reference font patch was not applied")


@case
def pipeline_auto_repair_stops_after_no_improvement() -> None:
    work = new_workdir("pipeline_auto_repair_no_improve")
    build_path = _write_repair_loop_fixture(work)
    calls = {"qa": 0}

    def fake_qa(out_dir, mode="user", output_docx_name="final.docx"):
        calls["qa"] += 1
        return _fake_repair_report(out_dir, code="MISSING_DOCX", message="synthetic persistent missing docx")

    deps = QADependencies(
        qa_check_and_write=fake_qa,
        conformance_check_and_write=None,
        visual_check_and_write=None,
        optional_import_detail=lambda name: "",
    )
    result = run_repair_loop(
        str(work),
        mode="user",
        output_docx_name="final.docx",
        qa_level="basic",
        project_root=str(REPO_ROOT),
        max_rounds=5,
        stop_no_improve=2,
        deps=deps,
        run_generated_script=run_generated_script,
        python_executable=sys.executable,
    )
    report = json.loads((work / "repair_loop_report.json").read_text(encoding="utf-8"))
    assert_true(not result.ok and result.status == "stopped_no_improvement", f"unexpected stop status: {result.status}")
    assert_true(calls["qa"] >= 3, "QA was not rerun during no-improvement loop")
    assert_true((work / "final.docx").exists(), "rebuild action did not run build_generated.py")
    assert_true(report["stop_detail"], "no-improvement stop detail was not recorded")
    assert_true(build_path.read_text(encoding="utf-8").count("AUTO_REPAIR") == 0, "no-improvement rebuild should not patch script")


@case
def pipeline_auto_repair_stops_for_needs_user_file() -> None:
    work = new_workdir("pipeline_auto_repair_needs_file")
    _write_repair_loop_fixture(work)
    calls = {"qa": 0}

    def fake_qa(out_dir, mode="user", output_docx_name="final.docx"):
        calls["qa"] += 1
        return _fake_repair_report(out_dir, code="CONTENT_IMAGE_MISSING", message="synthetic missing image")

    deps = QADependencies(
        qa_check_and_write=fake_qa,
        conformance_check_and_write=None,
        visual_check_and_write=None,
        optional_import_detail=lambda name: "",
    )
    result = run_repair_loop(
        str(work),
        mode="user",
        output_docx_name="final.docx",
        qa_level="basic",
        project_root=str(REPO_ROOT),
        max_rounds=5,
        stop_no_improve=2,
        deps=deps,
        run_generated_script=run_generated_script,
        python_executable=sys.executable,
    )
    report = json.loads((work / "repair_loop_report.json").read_text(encoding="utf-8"))
    assert_true(not result.ok and result.status == "stopped_needs_user_file", f"unexpected stop status: {result.status}")
    assert_true(calls["qa"] == 1, "needs-user-file blocker should stop before rebuild loops")
    assert_true(report["blockers"] and report["blockers"][0]["code"] == "CONTENT_IMAGE_MISSING", "blocker was not recorded")


@case
def pipeline_auto_repair_strict_requires_conformance_dependency() -> None:
    work = new_workdir("pipeline_auto_repair_missing_conformance")
    _write_repair_loop_fixture(work)
    _fake_repair_report(str(work), code="NO_ISSUE", severity="info")

    deps = QADependencies(
        qa_check_and_write=lambda out_dir, mode="user", output_docx_name="final.docx": _fake_repair_report(out_dir, code="NO_ISSUE", severity="info"),
        conformance_check_and_write=None,
        visual_check_and_write=None,
        optional_import_detail=lambda name: "missing dependency detail",
    )
    result = run_repair_loop(
        str(work),
        mode="user",
        output_docx_name="final.docx",
        qa_level="strict",
        project_root=str(REPO_ROOT),
        max_rounds=1,
        stop_no_improve=1,
        deps=deps,
        run_generated_script=run_generated_script,
        python_executable=sys.executable,
    )
    report = json.loads((work / "repair_loop_report.json").read_text(encoding="utf-8"))
    assert_true(not result.ok, "strict auto repair should not converge without conformance QA")
    assert_true("CONFORMANCE_QA_UNAVAILABLE" in report["final_error_codes"], f"missing conformance error was not recorded: {report}")


@case
def pipeline_auto_repair_visual_preserves_visual_options() -> None:
    work = new_workdir("pipeline_auto_repair_visual_options")
    build_path = _write_repair_loop_fixture(work)
    visual_calls: List[Dict[str, Any]] = []
    qa_calls = {"count": 0}

    def fake_qa(out_dir, mode="user", output_docx_name="final.docx"):
        qa_calls["count"] += 1
        if qa_calls["count"] == 1:
            return _fake_repair_report(out_dir, code="PLACEHOLDER_TEXT_LEFT", message="synthetic placeholder")
        return _fake_repair_report(out_dir, code="NO_ISSUE", severity="info")

    def fake_conformance(out_dir, mode="user", output_docx_name="final.docx", project_root=""):
        return _write_conformance_report(out_dir, passed=True)

    def fake_visual(out_dir, output_docx_name="final.docx", project_root="", render_all_pages=True, require_wps=False, golden_dir=None, update_golden=False):
        visual_calls.append(
            {
                "require_wps": require_wps,
                "golden_dir": golden_dir,
                "update_golden": update_golden,
                "render_all_pages": render_all_pages,
            }
        )
        report = {
            "schema_version": 1,
            "mode": "user",
            "passed": True,
            "counts": {},
            "issues": [],
            "next_action": "ok",
        }
        Path(out_dir, "visual_report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
        return report

    deps = QADependencies(
        qa_check_and_write=fake_qa,
        conformance_check_and_write=fake_conformance,
        visual_check_and_write=fake_visual,
        optional_import_detail=lambda name: "",
    )
    result = run_repair_loop(
        str(work),
        mode="user",
        output_docx_name="final.docx",
        qa_level="visual",
        project_root=str(REPO_ROOT),
        max_rounds=3,
        stop_no_improve=2,
        deps=deps,
        run_generated_script=run_generated_script,
        python_executable=sys.executable,
        golden_dir=str(work / "Golden"),
        update_golden=True,
        require_wps=True,
    )
    assert_true(result.ok, f"visual auto repair did not converge: {result.status}")
    assert_true("AUTO_REPAIR_PLACEHOLDER_CLEANUP_V1" in build_path.read_text(encoding="utf-8"), "placeholder repair was not applied")
    assert_true(len(visual_calls) >= 2, "visual QA was not rerun after repair")
    assert_true(all(call["require_wps"] is True for call in visual_calls), f"require_wps was not preserved: {visual_calls}")
    assert_true(all(call["update_golden"] is True for call in visual_calls), f"update_golden was not preserved: {visual_calls}")
    assert_true(all(str(call["golden_dir"]).endswith("Golden") for call in visual_calls), f"golden_dir was not preserved: {visual_calls}")


@case
def pipeline_auto_repair_report_paths_are_sanitized() -> None:
    work = new_workdir("pipeline_auto_repair_sanitized_paths")
    build_path = _write_repair_loop_fixture(work)
    first_build = run_generated_script(str(build_path), str(work), python_executable=sys.executable)
    assert_true(first_build.returncode == 0, f"synthetic build failed: {first_build.stderr}")

    from qa_checker import check_and_write as qa_check_and_write

    deps = QADependencies(
        qa_check_and_write=qa_check_and_write,
        conformance_check_and_write=None,
        visual_check_and_write=None,
        optional_import_detail=lambda name: "",
    )
    result = run_repair_loop(
        str(work),
        mode="user",
        output_docx_name="final.docx",
        qa_level="basic",
        project_root=str(REPO_ROOT),
        max_rounds=3,
        stop_no_improve=2,
        deps=deps,
        run_generated_script=run_generated_script,
        python_executable=sys.executable,
    )
    report_text = (work / "repair_loop_report.json").read_text(encoding="utf-8")
    report = json.loads(report_text)
    assert_true(result.ok, f"auto repair did not converge: {result.status}")
    assert_true(not Path(report["output_dir"]).is_absolute(), f"output_dir leaked absolute path: {report['output_dir']}")
    assert_true(not Path(report["final_docx"]).is_absolute(), f"final_docx leaked absolute path: {report['final_docx']}")
    assert_true(str(work) not in report_text, "repair loop report leaked the absolute output path")


@case
def pipeline_runs_pdf_template_end_to_end_strict() -> None:
    if not poppler_available():
        return
    if str(REPO_ROOT) not in sys.path:
        sys.path.insert(0, str(REPO_ROOT))
    import run_pipeline as root_runner

    work = new_workdir("pipeline_pdf_e2e")
    template_dir = work / "Templates"
    inputs_dir = work / "Inputs"
    outputs_dir = work / "Outputs"
    template_dir.mkdir()
    inputs_dir.mkdir()
    outputs_dir.mkdir()

    write_text_pdf(
        template_dir / "pdf_template.pdf",
        [
            ("Format requirements", 16, 72, 780),
            ("Page A4 margins top 2.5 cm bottom 2.5 cm left 3.0 cm right 2.5 cm", 11, 72, 742),
            ("Body font Times New Roman 12 pt justified line spacing 1.5", 11, 72, 718),
            ("Heading 1 font SimHei 16 pt bold centered", 11, 72, 694),
            ("Heading 2 font SimHei 14 pt bold left", 11, 72, 670),
            ("References font Times New Roman 12 pt hanging indent", 11, 72, 646),
        ],
    )

    content_doc = Document()
    content_doc.add_heading("PDF Template End-to-End Demo", level=0)
    content_doc.add_paragraph("Abstract")
    content_doc.add_paragraph("This synthetic document verifies that a PDF template can drive the reusable Word pipeline.")
    content_doc.add_heading("1 Introduction", level=1)
    content_doc.add_paragraph("The body includes enough text for extraction, generation, and strict QA to exercise the normal handoff.")
    content_doc.add_heading("1.1 Method", level=2)
    content_doc.add_paragraph("The method section checks heading hierarchy and paragraph rendering.")
    content_doc.add_paragraph("References")
    content_doc.add_paragraph("[1] Synthetic Author. PDF template regression case. 2026.")
    content_doc.save(inputs_dir / "paper.docx")

    old_dirs = (root_runner.TEMPLATE_DIR, root_runner.INPUTS_DIR, root_runner.OUTPUTS_DIR)
    try:
        root_runner.TEMPLATE_DIR = str(template_dir)
        root_runner.INPUTS_DIR = str(inputs_dir)
        root_runner.OUTPUTS_DIR = str(outputs_dir)
        result = root_runner.run("pdf_template.pdf", "paper.docx", mode="developer", qa_level="strict")
    finally:
        root_runner.TEMPLATE_DIR, root_runner.INPUTS_DIR, root_runner.OUTPUTS_DIR = old_dirs

    assert_true(result, "PDF template end-to-end pipeline returned no output directory")
    result_dir = Path(result)
    assert_true((result_dir / "最终论文.docx").exists(), "PDF template E2E did not build final DOCX")
    qa = json.loads((result_dir / "qa_report.json").read_text(encoding="utf-8"))
    assert_true(qa.get("passed") is True, f"PDF template E2E strict QA failed: {qa.get('issues')}")


@case
def pipeline_dependencies_loads_optional_modules_and_reports_missing() -> None:
    def marker(name):
        return lambda *args, **kwargs: name

    modules = {
        "qa_checker": SimpleNamespace(check_and_write=marker("qa")),
        "qa_conformance": SimpleNamespace(check_and_write=marker("conformance"), write_requirements=marker("requirements")),
        "template_profiler": SimpleNamespace(write_profile=marker("profile")),
        "md_parser": SimpleNamespace(extract_format=marker("md_format"), extract_content=marker("md_content")),
    }

    def fake_import_module(name):
        if name == "qa_visual":
            raise ImportError("visual unavailable")
        return modules[name]

    deps = load_optional_dependencies(import_module=fake_import_module)
    assert_true(deps.qa_check_and_write() == "qa", "qa checker dependency was not loaded")
    assert_true(deps.conformance_check_and_write() == "conformance", "conformance dependency was not loaded")
    assert_true(deps.write_template_requirements() == "requirements", "template requirements dependency was not loaded")
    assert_true(deps.write_template_profile() == "profile", "template profiler dependency was not loaded")
    assert_true(deps.extract_md_format() == "md_format" and deps.extract_md_content() == "md_content", "md parser dependencies were not loaded")
    assert_true(deps.visual_check_and_write is None, "missing visual dependency should be None")
    assert_true("visual unavailable" in deps.optional_import_detail("qa_visual"), "missing dependency detail was not preserved")
    assert_true(deps.optional_import_detail("qa_checker") == "", "available dependency should not report an error detail")


@case
def pipeline_artifacts_write_format_and_content_handoffs() -> None:
    work = new_workdir("pipeline_artifacts")
    fmt_json_path, fmt_md_path = write_format_artifacts(base_format(), "# Format", str(work))
    assert_true(Path(fmt_json_path).exists(), "format.json was not written")
    assert_true(Path(fmt_md_path).read_text(encoding="utf-8") == "# Format", "format markdown changed")

    content = base_content(
        [
            {"text": "A paragraph with math", "math": [{"text": "E=mc^2"}]},
            {"role": "figure", "image": "fig2.png", "caption": "Figure 2 Demo"},
            {"role": "table_caption", "text": "表 1 指标对比"},
            {"role": "table", "table_rows": [["Metric", "Value"], ["Accuracy", "98%"]]},
            "Plain paragraph",
        ]
    )
    content["sections"][0]["images"] = ["fig1.png", "fig2.png"]
    cnt_json_path, cnt_md_path = write_content_artifacts(content, str(work), str(work / "paper.docx"))
    summary = Path(cnt_md_path).read_text(encoding="utf-8")
    assert_true(Path(cnt_json_path).exists(), "content.json was not written")
    assert_true("# 内容提取" in summary, "content report title missing")
    assert_true("- [图片] fig1.png" in summary, "content report image line missing")
    assert_true("[图片] fig2.png" in summary, "structured figure was not summarized as an image")
    assert_true(summary.count("[图片] fig2.png") == 1, "structured figure image was duplicated in content report")
    assert_true("[表格] 2行 x 2列" in summary, "structured table was not summarized as a table")
    assert_true("表 1 指标对比" in summary, "table caption was not preserved in content markdown")
    assert_true(summary.count("[公式]") == 0, "non-formula structured content was mislabeled as formula")
    assert_true("(+1公式)" in summary, "content report math count missing")
    assert_true("## 参考文献" in build_content_markdown(content, str(work / "paper.docx")), "references section missing")


@case
def pipeline_execution_runs_generated_script_with_utf8_output() -> None:
    work = new_workdir("pipeline_execution")
    script = work / "build_generated.py"
    script.write_text("print('生成完成')\n", encoding="utf-8")
    result = run_generated_script(str(script), str(work), python_executable=sys.executable)
    assert_true(result.returncode == 0, f"generated script returned {result.returncode}: {result.stderr}")
    assert_true("生成完成" in result.stdout, "UTF-8 stdout was not decoded")


@case
def pipeline_build_phase_generates_and_blocks_on_failure() -> None:
    work = new_workdir("pipeline_build_phase")
    steps: List[str] = []
    calls: List[str] = []

    def fake_step(label):
        steps.append(label)

    def fake_generate(fmt_json_path, cnt_json_path, out_dir, output_docx_name):
        calls.append(f"generate:{Path(out_dir).name}:{output_docx_name}")
        Path(out_dir, "build_generated.py").write_text("print('ok')\n", encoding="utf-8")
        return 11

    def fake_run(gen_py_path, out_dir, python_executable):
        calls.append(f"run:{Path(gen_py_path).name}:{python_executable}")
        return ScriptExecutionResult(returncode=0, stdout="built\n", stderr="")

    ok = generate_and_build_docx_phase(
        "format.json",
        "content.json",
        str(work),
        "folder",
        output_docx_name="out.docx",
        generate_script=fake_generate,
        run_generated_script=fake_run,
        python_executable="python",
        step=fake_step,
    )
    assert_true(ok is True, "build phase should pass when generated script succeeds")
    assert_true(
        steps[:2]
        == [
            "Phase 4/6: 生成构建脚本",
            "Phase 5/6: 构建最终 docx（生成静态目录；可用 Word COM 时写入页码）",
        ],
        f"unexpected build steps: {steps}",
    )
    assert_true(calls == [f"generate:{work.name}:out.docx", "run:build_generated.py:python"], f"unexpected build calls: {calls}")

    def failing_run(gen_py_path, out_dir, python_executable):
        return ScriptExecutionResult(returncode=1, stdout="", stderr="boom")

    failed = generate_and_build_docx_phase(
        "format.json",
        "content.json",
        str(work),
        "folder",
        output_docx_name="out.docx",
        generate_script=fake_generate,
        run_generated_script=failing_run,
        python_executable="python",
        step=fake_step,
    )
    assert_true(failed is False, "build phase should block when generated script fails")


@case
def pipeline_qa_runs_strict_and_blocks_failures() -> None:
    work = new_workdir("pipeline_qa")
    calls: List[str] = []

    def passing_qa(out_dir, mode, output_docx_name):
        calls.append("qa")
        return {"passed": True, "issues": [], "counts": {}, "mode": mode}

    def passing_conformance(out_dir, mode, output_docx_name, project_root):
        calls.append("conformance")
        return {"passed": True, "issues": []}

    deps = QADependencies(
        qa_check_and_write=passing_qa,
        conformance_check_and_write=passing_conformance,
        visual_check_and_write=None,
        optional_import_detail=lambda name: "",
    )
    assert_true(
        run_qa_phases(str(work), mode="developer", output_docx_name="最终论文.docx", qa_level="strict", project_root=str(work), deps=deps),
        "strict QA should pass when structural and conformance QA pass",
    )
    assert_true(calls == ["qa", "conformance"], f"unexpected QA call order: {calls}")

    def failing_qa(out_dir, mode, output_docx_name):
        return {
            "passed": False,
            "counts": {},
            "issues": [{"severity": "error", "code": "TEST_ERROR", "message": "Synthetic failure", "active_owner": "developer"}],
            "repair_plan": {"steps": []},
        }

    failing_deps = QADependencies(
        qa_check_and_write=failing_qa,
        conformance_check_and_write=passing_conformance,
        visual_check_and_write=None,
        optional_import_detail=lambda name: "",
    )
    assert_true(
        not run_qa_phases(str(work), mode="developer", output_docx_name="最终论文.docx", qa_level="basic", project_root=str(work), deps=failing_deps),
        "QA should block the pipeline when the required QA report fails",
    )


@case
def pipeline_summary_mentions_outputs_and_mode() -> None:
    summary = build_completion_summary("2026-05-27_demo", "最终论文.docx", "developer")
    assert_true("Outputs/2026-05-27_demo/" in summary, "output directory missing from completion summary")
    assert_true("最终论文.docx" in summary, "output docx missing from completion summary")
    assert_true("当前模式: 开发者" in summary, "developer mode missing from completion summary")
    assert_true("build_generated.py" in summary, "user fine-tuning target missing from completion summary")


@case
def pipeline_verification_double_verify_arbitrates_format_mismatch() -> None:
    calls: List[int] = []

    def flaky_format_extractor(path):
        calls.append(len(calls) + 1)
        paragraph_count = 1 if len(calls) != 2 else 2
        fmt = {
            "_meta": {"paragraphs": paragraph_count},
            "paragraphs": [{"runs": []} for _ in range(paragraph_count)],
            "tables": [],
            "sections": [],
        }
        return fmt, "# Format"

    result, report = double_verify(flaky_format_extractor, "synthetic.docx", "Format")
    assert_true(len(calls) == 3, "format mismatch should trigger a third arbitration run")
    assert_true(len(result["paragraphs"]) == 1, "third run should arbitrate back to the majority shape")
    assert_true(report == "# Format", "format markdown payload should be preserved")


@case
def pipeline_verification_arbitrates_content_mismatch() -> None:
    work = new_workdir("pipeline_verify_content_majority")
    calls: List[int] = []

    def flaky_content_extractor(path, output_dir=None):
        calls.append(len(calls) + 1)
        references = [{"text": "Synthetic reference"}] if len(calls) == 2 else []
        return {
            "_meta": {"images_extracted": 0, "image_extract_failures": [], "non_body_images": []},
            "sections": [{"heading": "Body", "paragraphs": ["Text"]}],
            "references": references,
        }

    result = double_verify(flaky_content_extractor, "synthetic.docx", "Content", output_dir=str(work))
    assert_true(len(calls) == 3, "content mismatch should trigger a third arbitration run")
    assert_true(result["references"] == [], "third run should arbitrate content back to the majority shape")


@case
def pipeline_verification_converges_incremental_content_and_materializes_images() -> None:
    work = new_workdir("pipeline_verify_content_converge")
    calls: List[int] = []

    def incremental_content_extractor(path, output_dir=None):
        calls.append(len(calls) + 1)
        run_no = len(calls)
        base = Path(path).stem
        fig_dir = Path(output_dir) / base / "figures"
        fig_dir.mkdir(parents=True, exist_ok=True)

        image_names = [f"img{i}.png" for i in range(1, run_no + 1)]
        for image_name in image_names:
            (fig_dir / image_name).write_bytes(PNG_1X1)

        references = [{"text": f"[{i}] Ref {i}"} for i in range(1, run_no + 1)]
        return {
            "_meta": {
                "images_extracted": len(image_names),
                "images_dir": str(fig_dir),
                "image_extract_failures": [{"target": "transient", "error": "miss"}] if run_no == 1 else [],
                "non_body_images": [],
            },
            "sections": [
                {
                    "heading": "Body",
                    "level": 1,
                    "paragraphs": [{"role": "image", "image": image_names[-1]}],
                    "images": [image_names[-1]],
                }
            ],
            "references": references,
        }

    result = double_verify(incremental_content_extractor, "synthetic.docx", "Content", output_dir=str(work))
    assert_true(len(calls) == 5, "content convergence should keep collecting unstable recoverable content up to the cap")
    assert_true(result["_meta"].get("converged_extraction"), "convergence metadata was not recorded")
    assert_true(len(result["references"]) == 5, f"references were not merged: {result['references']}")
    assert_true(result["sections"][0]["images"] == ["img1.png", "img2.png", "img3.png", "img4.png", "img5.png"], "section images were not unioned")

    paragraph_images = sorted(
        item.get("image")
        for item in result["sections"][0]["paragraphs"]
        if isinstance(item, dict) and item.get("image")
    )
    assert_true(paragraph_images == ["img1.png", "img2.png", "img3.png", "img4.png", "img5.png"], f"paragraph images were not merged: {paragraph_images}")
    final_fig_dir = work / "synthetic" / "figures"
    assert_true((final_fig_dir / "img1.png").exists() and (final_fig_dir / "img5.png").exists(), "converged images were not materialized")
    assert_true(not (work / "_extract_verify_runs").exists(), "isolated verification directories were not cleaned")
    assert_true(result["_meta"].get("image_extract_failures") == [], "transient image failures should not survive convergence")


@case
def pipeline_verification_inserts_recovered_items_near_anchors() -> None:
    def content_with(paragraphs, images=None):
        return {
            "_meta": {
                "images_extracted": len(images or []),
                "image_extract_failures": [],
                "non_body_images": [],
            },
            "sections": [
                {
                    "heading": "Body",
                    "level": 1,
                    "role": "body",
                    "paragraphs": paragraphs,
                    "images": images or [],
                }
            ],
            "references": [],
        }

    merged, _reason = _merge_content_results(
        [
            content_with(["before", {"role": "image", "image": "img.png"}, "middle", "after"], ["img.png"]),
            content_with(["before", {"role": "formula", "source": "latex", "latex": "E=mc^2", "text": "E=mc^2"}, "middle", "after"]),
            content_with(["before", {"role": "table", "table_rows": [["A"], ["1"]]}, "middle", "after"]),
        ]
    )
    paragraphs = merged["sections"][0]["paragraphs"]
    middle_index = paragraphs.index("middle")
    formula_index = next(i for i, item in enumerate(paragraphs) if isinstance(item, dict) and item.get("role") == "formula")
    table_index = next(i for i, item in enumerate(paragraphs) if isinstance(item, dict) and item.get("table_rows"))
    assert_true(formula_index < middle_index, f"recovered formula drifted to section tail: {paragraphs}")
    assert_true(table_index < middle_index, f"recovered table drifted to section tail: {paragraphs}")


@case
def pipeline_verification_fails_when_no_majority_signature() -> None:
    calls: List[int] = []

    def unstable_content_extractor(path):
        calls.append(len(calls) + 1)
        return {
            "_meta": {"images_extracted": len(calls), "image_extract_failures": [], "non_body_images": []},
            "sections": [{"heading": f"Body {idx}", "paragraphs": ["Text"]} for idx in range(len(calls))],
            "references": [],
        }

    try:
        double_verify(unstable_content_extractor, "synthetic.docx", "Content")
    except VerificationError as exc:
        assert_true("verification failed" in str(exc), "failure should explain that verification failed")
        assert_true(len(calls) == 3, "unresolved mismatch should stop after the third run")
        return
    fail("unresolved extraction mismatch should raise VerificationError")


@case
def run_pipeline_completion_step_is_named() -> None:
    root = PIPELINE_DIR.parents[2]
    text = (root / "run_pipeline.py").read_text(encoding="utf-8")
    assert_true("step('完成')" in text or 'step("完成")' in text, "completion step should have a user-visible label")
    assert_true("step('??')" not in text and 'step("??")' not in text, "placeholder completion step leaked into run_pipeline.py")


@case
def pipeline_template_phase_writes_profile_and_requirements() -> None:
    work = new_workdir("pipeline_template_phase")
    calls: List[str] = []

    def fake_profile(fmt, out_dir, project_root):
        calls.append(f"profile:{project_root}")
        return {
            "capabilities": {
                "has_cover": True,
                "has_heading_styles": False,
                "has_caption_styles": True,
            },
            "risk_flags": {"mixed_headers": True, "low_sample": False},
        }

    profile = write_template_profile_phase(
        base_format(),
        str(work),
        project_root="ROOT",
        write_template_profile=fake_profile,
    )
    assert_true(profile["capabilities"]["has_cover"] is True, "profile return value was not preserved")

    def fake_requirements(fmt, content, out_dir):
        calls.append("requirements")
        Path(out_dir, "template_requirements.json").write_text("{}", encoding="utf-8")

    ok = write_template_requirements_phase(
        base_format(),
        base_content(["Body"]),
        str(work),
        write_template_requirements=fake_requirements,
        optional_import_detail=lambda name: " detail",
    )
    assert_true(ok is True, "requirements phase should report success")
    assert_true(calls == ["profile:ROOT", "requirements"], f"unexpected template phase calls: {calls}")
    assert_true((work / "template_requirements.json").exists(), "requirements writer was not called")

    skipped = write_template_requirements_phase(
        base_format(),
        base_content(["Body"]),
        str(work),
        write_template_requirements=None,
        optional_import_detail=lambda name: " detail",
    )
    assert_true(skipped is False, "missing requirements writer should report skipped")


