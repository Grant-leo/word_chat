"""Privacy, visual QA, and CLI regression cases."""
from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path

from docx import Document
from privacy import sanitize_value
from qa_checker import check_output
from regression_suite_modules.harness import assert_true, base_content, base_format, case, new_workdir, write_json

PIPELINE_DIR = Path(__file__).resolve().parents[1]


@case
def privacy_sanitizes_absolute_paths() -> None:
    data = {
        "path": r"X:\workspace\project\Outputs\private\file.docx",
        "tmp": str(Path(tempfile.gettempdir()) / "abc" / "file.pdf"),
        "embedded": r"Missing image: X:\workspace\project\Inputs\private\figure.png",
    }
    sanitized = sanitize_value(data, project_root=r"X:\workspace\project")
    text = json.dumps(sanitized, ensure_ascii=False)
    assert_true(r"X:\workspace" not in text and "project" not in text, "project path leaked")
    assert_true("<PROJECT>" in text and "<TEMP>" in text, "path labels missing")
    assert_true("Missing image: <PROJECT>/Inputs/private/figure.png" in text, "embedded absolute path was not sanitized")


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
    assert_true("下一步" in qa_visual.report_to_markdown(report), "visual report markdown should include next action")


@case
def visual_qa_sanitizes_issue_details() -> None:
    import qa_visual

    work = new_workdir("visual_issue_privacy")
    report = qa_visual.check_visual(str(work), output_docx_name="missing.docx", project_root=str(work))
    text = json.dumps(report.get("issues") or [], ensure_ascii=False)
    assert_true(str(work) not in text and str(work).replace("\\", "/") not in text, "visual QA issue leaked output path")
    assert_true("<PROJECT>/missing.docx" in text, f"visual QA issue detail was not sanitized: {text}")


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
def run_pipeline_help_localizes_agent_options() -> None:
    root = PIPELINE_DIR.parents[2]
    result = subprocess.run(
        [sys.executable, str(root / "run_pipeline.py"), "--help"],
        cwd=str(root),
        env={**os.environ, "PYTHONIOENCODING": "utf-8"},
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=60,
    )
    assert_true(result.returncode == 0, f"run_pipeline --help failed: {result.stderr}")
    assert_true("Agent 自动入口" in result.stdout and "自动修复闭环" in result.stdout, "help text lost novice-friendly Chinese option descriptions")
    assert_true("Agent-first mode" not in result.stdout and "Run a bounded" not in result.stdout, "help text still exposes old English option descriptions")


@case
def content_parser_cli_writes_outputs_outside_inputs() -> None:
    work = new_workdir("content_parser_cli_output")
    inputs = work / "Inputs"
    out_dir = work / "Outputs" / "content_cli"
    inputs.mkdir()
    docx = inputs / "paper.docx"
    doc = Document()
    doc.add_paragraph("1 Introduction")
    doc.add_paragraph("Body paragraph.")
    doc.save(docx)

    result = subprocess.run(
        [
            sys.executable,
            str(PIPELINE_DIR / "content_parser.py"),
            str(docx),
            "--output-dir",
            str(out_dir),
        ],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=60,
    )
    assert_true(result.returncode == 0, f"content_parser CLI failed: {result.stdout}\n{result.stderr}")
    assert_true((out_dir / "paper_content.json").exists(), "content_parser CLI did not write JSON to output dir")
    assert_true((out_dir / "paper" / "figures").exists(), "content_parser CLI did not place figures under output dir")
    assert_true(not (inputs / "paper_content.json").exists(), "content_parser CLI wrote JSON beside input")
    assert_true(not (inputs / "paper" / "figures").exists(), "content_parser CLI wrote figures under Inputs")


@case
def content_parser_default_extract_writes_outputs_outside_inputs() -> None:
    from content_parser import extract as extract_docx_content

    work = new_workdir("content_parser_default_output")
    inputs = work / "Inputs"
    inputs.mkdir()
    docx = inputs / "paper.docx"
    doc = Document()
    doc.add_paragraph("1 Introduction")
    doc.add_paragraph("Body paragraph.")
    doc.save(docx)

    old_cwd = os.getcwd()
    try:
        os.chdir(work)
        content = extract_docx_content(str(docx))
    finally:
        os.chdir(old_cwd)

    images_dir = Path(content["_meta"]["images_dir"])
    assert_true("Outputs" in images_dir.parts and "_content_parser_extract" in images_dir.parts, f"default images dir is unsafe: {images_dir}")
    assert_true(not (inputs / "paper" / "figures").exists(), "default content extraction wrote figures under Inputs")


@case
def format_extractor_cli_writes_outputs_outside_templates() -> None:
    work = new_workdir("format_extractor_cli_output")
    templates = work / "Templates"
    templates.mkdir()
    docx = templates / "template.docx"
    doc = Document()
    doc.add_paragraph("Template heading")
    doc.add_paragraph("Template body.")
    doc.save(docx)

    result = subprocess.run(
        [sys.executable, str(PIPELINE_DIR / "format_extractor.py"), str(docx)],
        cwd=str(work),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=60,
    )
    out_dir = work / "Outputs" / "_format_extractor_cli"
    assert_true(result.returncode == 0, f"format_extractor CLI failed: {result.stdout}\n{result.stderr}")
    assert_true((out_dir / "template_format.json").exists(), "format_extractor CLI did not write JSON to Outputs")
    assert_true((out_dir / "template_格式提取.md").exists(), "format_extractor CLI did not write MD report to Outputs")
    assert_true(not (templates / "template_format.json").exists(), "format_extractor CLI wrote JSON beside template")
    assert_true(not (templates / "template_assets").exists(), "format_extractor CLI wrote assets beside template")


@case
def format_extractor_default_assets_stay_outside_templates() -> None:
    from format_extractor import extract as extract_format

    work = new_workdir("format_extractor_default_assets")
    templates = work / "Templates"
    templates.mkdir()
    docx = templates / "template.docx"
    doc = Document()
    doc.add_paragraph("Template heading")
    doc.save(docx)

    old_cwd = os.getcwd()
    try:
        os.chdir(work)
        fmt, _ = extract_format(str(docx))
    finally:
        os.chdir(old_cwd)

    assets_dir = Path(fmt["_meta"]["assets_dir"])
    assert_true("Outputs" in assets_dir.parts and "_format_extractor_extract" in assets_dir.parts, f"default assets dir is unsafe: {assets_dir}")
    assert_true(not (templates / "template_assets").exists(), "default format extraction wrote assets beside template")


@case
def md_parser_cli_writes_outputs_outside_inputs() -> None:
    work = new_workdir("md_parser_cli_output")
    inputs = work / "Inputs"
    inputs.mkdir()
    md = inputs / "paper.md"
    md.write_text("# 格式说明\n\n正文：宋体，小四号。\n\n# 论文标题\n\n正文段落。\n", encoding="utf-8")

    result = subprocess.run(
        [sys.executable, str(PIPELINE_DIR / "md_parser.py"), str(md)],
        cwd=str(work),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=60,
    )
    out_dir = work / "Outputs" / "_md_parser_cli"
    assert_true(result.returncode == 0, f"md_parser CLI failed: {result.stdout}\n{result.stderr}")
    assert_true((out_dir / "paper_format.json").exists(), "md_parser CLI did not write format JSON to Outputs")
    assert_true((out_dir / "paper_content.json").exists(), "md_parser CLI did not write content JSON to Outputs")
    assert_true(not (inputs / "paper_format.json").exists(), "md_parser CLI wrote format JSON beside input")
    assert_true(not (inputs / "paper_content.json").exists(), "md_parser CLI wrote content JSON beside input")


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

