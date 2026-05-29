"""Privacy, visual QA, and CLI regression cases."""
from __future__ import annotations

import json
import subprocess
import sys
import tempfile
from pathlib import Path

from privacy import sanitize_value
from qa_checker import check_output
from regression_suite_modules.harness import assert_true, base_content, base_format, case, new_workdir, write_json

PIPELINE_DIR = Path(__file__).resolve().parents[1]


@case
def privacy_sanitizes_absolute_paths() -> None:
    data = {
        "path": r"X:\workspace\project\Outputs\private\file.docx",
        "tmp": str(Path(tempfile.gettempdir()) / "abc" / "file.pdf"),
    }
    sanitized = sanitize_value(data, project_root=r"X:\workspace\project")
    text = json.dumps(sanitized, ensure_ascii=False)
    assert_true(r"X:\workspace" not in text and "project" not in text, "project path leaked")
    assert_true("<PROJECT>" in text and "<TEMP>" in text, "path labels missing")


@case
def visual_sample_pages_pick_useful_pages() -> None:
    import qa_visual

    pages = ["cover", "contents", "blank", "1. Introduction", "middle"] + ["body"] * 7
    samples = qa_visual._sample_pages(12, pages)
    assert_true(1 in samples and 2 in samples and 4 in samples and 6 in samples, "sample page selection missed key pages")


@case
def visual_qa_fails_closed_without_pdf_tools() -> None:
    import qa_visual

    work = new_workdir("visual_closed")
    (work / "final.docx").write_bytes(b"not a real docx; export is monkeypatched")
    fake_pdf = work / "fake.pdf"
    fake_pdf.write_bytes(b"%PDF-1.4\n%%EOF\n")
    original_export = qa_visual._export_pdf
    original_pdfinfo = qa_visual._pdfinfo
    original_pages_text = qa_visual._pdf_pages_text
    original_render = qa_visual._render_samples
    try:
        qa_visual._export_pdf = lambda _docx, _visual_dir: str(fake_pdf)
        qa_visual._pdfinfo = lambda _pdf: {"available": False}
        qa_visual._pdf_pages_text = lambda _pdf, _visual_dir: []
        qa_visual._render_samples = lambda _pdf, _visual_dir, _pages: []
        report = qa_visual.check_visual(str(work), output_docx_name="final.docx")
    finally:
        qa_visual._export_pdf = original_export
        qa_visual._pdfinfo = original_pdfinfo
        qa_visual._pdf_pages_text = original_pages_text
        qa_visual._render_samples = original_render
    codes = [item["code"] for item in report["issues"]]
    assert_true(report["passed"] is False, "visual QA passed without pdfinfo/text validation")
    assert_true("PDFINFO_UNAVAILABLE" in codes, "missing pdfinfo was not reported")
    assert_true("Poppler" in report.get("next_action", ""), f"visual QA did not guide dependency repair: {report}")
    assert_true("Next action" in qa_visual.report_to_markdown(report), "visual report markdown should include next action")


@case
def sample_pages_empty_when_page_count_unknown() -> None:
    import qa_visual

    assert_true(qa_visual._sample_pages(0, []) == [], "sample pages should be empty for unknown page count")


@case
def run_pipeline_missing_inputs_returns_nonzero() -> None:
    root = PIPELINE_DIR.parents[2]
    result = subprocess.run(
        [
            sys.executable,
            str(root / "run_pipeline.py"),
            "--template",
            "__missing_template__.docx",
            "--content",
            "__missing_content__.docx",
        ],
        cwd=str(root),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=60,
    )
    assert_true(result.returncode != 0, "run_pipeline returned success for missing inputs")


@case
def qa_checker_cli_failure_returns_nonzero() -> None:
    work = new_workdir("qa_cli_nonzero")
    content = base_content([])
    content["_meta"]["missing_images"] = [{"source": "missing.png", "reason": "not_found"}]
    write_json(work / "content.json", content)
    write_json(work / "format.json", base_format())
    write_json(work / "workflow_mode.json", {"mode": "developer"})
    result = subprocess.run(
        [
            sys.executable,
            str(PIPELINE_DIR / "qa_checker.py"),
            str(work),
            "--mode",
            "developer",
            "--docx",
            "out.docx",
        ],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=60,
    )
    assert_true(result.returncode != 0, "qa_checker CLI returned success for a failed report")


@case
def qa_missing_image_detail_is_sanitized() -> None:
    work = new_workdir("qa_missing_image_privacy")
    private_path = str(work / "private" / "missing.png")
    content = base_content([])
    content["_meta"]["missing_images"] = [{"source": private_path, "reason": "not_found"}]
    write_json(work / "content.json", content)
    write_json(work / "format.json", base_format())
    write_json(work / "workflow_mode.json", {"mode": "developer"})
    report = check_output(str(work), mode="developer", output_docx_name="out.docx")
    detail = "\n".join(str(item.get("detail") or "") for item in report["issues"])
    assert_true(private_path not in detail, "QA leaked an absolute missing-image path")
    assert_true("<TEMP>" in detail or "<ABS_PATH>" in detail or "<PROJECT>" in detail, "QA missing-image detail was not sanitized")

