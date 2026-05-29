"""Structural QA and manifest regression cases."""
from __future__ import annotations

import json

from docx import Document

from qa_checker import check_output
from qa_checker_modules.content_metrics import _count_content_formulas, _count_content_images
from qa_checker_modules.registry import OWNER_BY_CODE
from qa_checker_modules.repair import build_repair_plan
from qa_checker_modules.repair_guides import REPAIR_GUIDES
from qa_conformance_modules.content_checks import _expected_paragraphs

from regression_suite_modules.generated_docx import run_generated_case
from regression_suite_modules.harness import (
    assert_true,
    base_content,
    base_format,
    case,
    new_workdir,
    write_json,
    write_sample_png,
)

@case
def table_manifest_matches_structured_body_tables() -> None:
    content = base_content(
        [
            {"role": "table", "table_rows": [["A", "B"], ["1", "2"]]},
        ],
        meta_tables=99,
    )
    result = run_generated_case("table_manifest", content)
    assert_true(result["manifest"]["counts"]["content_tables_rendered"] == 1, "body table was not counted")
    assert_true(result["report"]["counts"]["content_tables"] == 1, "QA used raw doc.tables instead of structured tables")
    assert_true(not result["report"]["issues"], f"unexpected QA issues: {result['report']['issues']}")


@case
def qa_manifest_detects_missing_table_render() -> None:
    content = base_content([{"role": "table", "table_rows": [["A"], ["B"]]}])
    result = run_generated_case("qa_missing_table", content)
    manifest_path = result["work"] / "build_manifest.json"
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    manifest["counts"]["content_tables_rendered"] = 0
    write_json(manifest_path, manifest)
    report = check_output(str(result["work"]), mode="developer", output_docx_name="out.docx")
    codes = [item["code"] for item in report["issues"]]
    assert_true("TABLE_COUNT_MISMATCH" in codes, "QA did not trust manifest for table mismatch")


@case
def code_table_is_not_body_table() -> None:
    content = base_content(
        [
            {
                "role": "code",
                "code": "interface GigabitEthernet0/0/1\nip address 10.0.0.1 255.255.255.0",
                "table_rows": [["interface GigabitEthernet0/0/1\nip address 10.0.0.1 255.255.255.0"]],
            }
        ],
        meta_tables=1,
    )
    result = run_generated_case("code_table", content)
    assert_true(result["manifest"]["counts"]["content_tables_rendered"] == 0, "code box was counted as body table")
    assert_true(result["report"]["counts"]["content_tables"] == 0, "QA counted code table_rows as body table")
    assert_true(not result["report"]["issues"], f"unexpected QA issues: {result['report']['issues']}")


@case
def qa_counts_xml_only_formula_items() -> None:
    content = base_content([
        {"role": "source_omml", "text": "E=mc^2", "xml": "<m:oMath/>"},
    ])
    assert_true(_count_content_formulas(content) == 1, "QA should count source OMML/XML formula items")


@case
def qa_counts_image_fields_even_when_role_varies() -> None:
    content = base_content([
        {"role": "media", "filename": "figure_a.png", "caption": "Figure A"},
    ])
    assert_true(_count_content_images(content) == 1, "QA should count image items by image fields, not only by role")


@case
def conformance_expected_paragraphs_classifies_plain_caption_strings() -> None:
    content = base_content([
        "Table 1 Model results",
        "表 2 变量描述",
        "图 3 系统架构",
        "图 1 展示了系统架构。",
        "Figure 2 shows the workflow.",
        "This is a body paragraph.",
    ])
    expected = _expected_paragraphs(content)
    roles_by_text = {item["text"]: item["role"] for item in expected}
    assert_true(roles_by_text["Table 1 Model results"] == "table_caption", "plain English table caption string should use table_caption style")
    assert_true(roles_by_text["表 2 变量描述"] == "table_caption", "plain Chinese table caption string should use table_caption style")
    assert_true(roles_by_text["图 3 系统架构"] == "figure_caption", "plain Chinese figure caption string should use figure_caption style")
    assert_true(roles_by_text["图 1 展示了系统架构。"] == "body", "plain Chinese figure-reference prose should remain body")
    assert_true(roles_by_text["Figure 2 shows the workflow."] == "body", "plain English figure-reference prose should remain body")
    assert_true(roles_by_text["This is a body paragraph."] == "body", "ordinary strings should remain body paragraphs")


@case
def image_manifest_matches_rendered_body_images() -> None:
    img_src = new_workdir("image_src")
    write_sample_png(img_src / "dot.png")
    content = base_content([
        {"role": "figure", "image": "dot.png", "caption": "Figure 1 sample"}
    ])
    content["_meta"]["images_dir"] = str(img_src)
    content["_meta"]["images_extracted"] = 1
    content["sections"][0]["images"] = ["dot.png"]
    result = run_generated_case("image_manifest", content)
    assert_true(result["manifest"]["counts"]["content_images_rendered"] == 1, "body image was not counted")
    assert_true(result["report"]["counts"]["content_images"] == 1, "QA did not count body image occurrence")
    assert_true(not result["report"]["issues"], f"unexpected QA issues: {result['report']['issues']}")


@case
def qa_counts_mixed_inline_and_section_images() -> None:
    work = new_workdir("mixed_image_count")
    content = base_content([
        {"role": "image", "image": "inline.png"},
    ])
    content["sections"][0]["images"] = ["inline.png", "", "section_only.png"]
    write_json(work / "content.json", content)
    write_json(work / "format.json", base_format())
    write_json(work / "workflow_mode.json", {"mode": "developer"})
    report = check_output(str(work), mode="developer", output_docx_name="out.docx")
    assert_true(report["counts"]["content_images"] == 2, f"mixed image count lost section-only images: {report['counts']}")


@case
def qa_reports_non_body_images_and_raw_latex_text() -> None:
    work = new_workdir("qa_non_body_latex")
    docx = work / "out.docx"
    doc = Document()
    doc.add_paragraph(r"$$x^2+y^2=z^2$$")
    doc.save(docx)
    content = base_content(["Body text"])
    content["_meta"]["non_body_images"] = [{"location": "section_1_header", "target": "media/image1.png"}]
    write_json(work / "content.json", content)
    write_json(work / "format.json", base_format())
    write_json(work / "workflow_mode.json", {"mode": "developer"})
    report = check_output(str(work), mode="developer", output_docx_name="out.docx")
    codes = [item["code"] for item in report["issues"]]
    assert_true("NON_BODY_IMAGE_UNSUPPORTED" in codes, "QA did not flag unsupported header/footer images")
    assert_true("LATEX_DELIMITER_TEXT" in codes, "QA did not flag raw LaTeX delimiters left in final DOCX")
    assert_true(report["passed"] is False, "non-body images and raw LaTeX text should fail QA")


@case
def qa_reports_duplicate_front_matter_headings() -> None:
    work = new_workdir("qa_duplicate_front_heading")
    docx = work / "out.docx"
    doc = Document()
    doc.add_paragraph("摘  要")
    doc.add_paragraph("Synthetic title")
    doc.add_paragraph("摘要")
    doc.add_paragraph("Abstract body.")
    doc.save(docx)
    content = {
        "_meta": {"source": "synthetic.docx", "paragraphs": 1, "tables_count": 0, "images_extracted": 0},
        "title_info": {"title_cn": "Synthetic title"},
        "sections": [
            {"heading": "摘要", "level": 1, "role": "cn_abstract", "paragraphs": ["Abstract body."], "images": []}
        ],
        "references": ["[1] Synthetic reference."],
    }
    write_json(work / "content.json", content)
    write_json(work / "format.json", base_format())
    write_json(work / "workflow_mode.json", {"mode": "developer"})
    (work / "build_generated.py").write_text("# synthetic\n", encoding="utf-8")
    report = check_output(str(work), mode="developer", output_docx_name="out.docx")
    codes = [item["code"] for item in report["issues"]]
    assert_true("DUPLICATE_FRONT_MATTER_HEADING" in codes, "QA did not report duplicate front matter heading")
    assert_true(report["passed"] is False, "duplicate front matter heading should fail QA")


@case
def qa_repair_plan_uses_relative_rebuild_command() -> None:
    work = new_workdir("qa_relative_rebuild_command")
    write_json(work / "workflow_mode.json", {"mode": "user"})
    report = {
        "mode": "user",
        "passed": True,
        "issues": [],
        "counts": {},
    }
    plan = build_repair_plan(report, str(work))
    command = plan["commands"]["rebuild_current_docx"]
    assert_true("build_generated.py" in command, "rebuild command should still point to build_generated.py")
    assert_true(str(work) not in command, f"rebuild command leaked absolute output path: {command}")


@case
def qa_repair_guides_cover_registered_issue_codes() -> None:
    missing = sorted(code for code in OWNER_BY_CODE if code not in REPAIR_GUIDES)
    assert_true(not missing, f"registered QA issue codes missing user-facing repair guides: {missing}")


@case
def qa_repair_plan_preserves_md_visual_workflow_command() -> None:
    work = new_workdir("qa_md_visual_workflow_command")
    write_json(
        work / "workflow_mode.json",
        {
            "mode": "user",
            "template": "demo.md",
            "content": "demo.md",
            "md": "demo.md",
            "qa_level": "visual",
            "golden_dir": "TestData/GoldenBaselines",
            "update_golden": True,
            "require_wps": True,
            "auto_repair": True,
            "repair_max_rounds": 4,
            "repair_stop_no_improve": 2,
        },
    )
    report = {
        "mode": "user",
        "passed": False,
        "issues": [{"code": "PLACEHOLDER_TEXT_LEFT", "severity": "error", "message": "placeholder"}],
        "counts": {},
    }
    plan = build_repair_plan(report, str(work))
    command = plan["commands"]["rerun_current_pipeline"]
    assert_true("--md demo.md" in command, f"MD workflow should rerun with --md: {command}")
    assert_true("--template" not in command and "--content" not in command, f"MD workflow should not be rewritten as template/content: {command}")
    assert_true("--qa-level visual" in command, f"QA level was not preserved: {command}")
    assert_true("--auto-repair" in command and "--repair-max-rounds 4" in command, f"auto repair options were not preserved: {command}")
    assert_true("--require-wps" in command and "--update-golden" in command, f"visual options were not preserved: {command}")
    assert_true("--golden-dir TestData/GoldenBaselines" in command, f"golden dir was not preserved: {command}")


@case
def qa_routes_user_file_errors_to_input_fix() -> None:
    work = new_workdir("qa_user_file_routing")
    write_json(
        work / "format.json",
        {
            "_meta": {
                "source": "blank_scan.pdf",
                "pdf_template": {
                    "type": "scanned_or_unsupported_pdf",
                    "errors": ["PDF_TEMPLATE_NO_TEXT"],
                    "confidence": 0.0,
                    "text_chars": 0,
                },
            },
            "paragraphs": [],
            "tables": [],
            "sections": [{"page_width_cm": 21.0, "page_height_cm": 29.7}],
            "cover": [],
            "style_profiles": {},
        },
    )
    write_json(work / "content.json", base_content(["Body text"]))
    write_json(work / "workflow_mode.json", {"mode": "user"})
    (work / "build_generated.py").write_text("# synthetic\n", encoding="utf-8")
    doc = Document()
    doc.add_paragraph("1 Introduction")
    doc.add_paragraph("Body text")
    doc.save(work / "out.docx")
    write_json(work / "build_manifest.json", {"schema_version": 1, "counts": {}})
    report = check_output(str(work), mode="user", output_docx_name="out.docx")
    issue = next(item for item in report["issues"] if item["code"] == "PDF_TEMPLATE_UNSUPPORTED")
    assert_true(issue["active_owner"] == "User input/template file", f"PDF template issue target was misleading: {issue}")
    assert_true("用户确认或补充输入文件" in report["next_action"], f"next action did not route to user file fix: {report['next_action']}")
    step = next(item for item in report["repair_plan"]["steps"] if item["code"] == "PDF_TEMPLATE_UNSUPPORTED")
    assert_true(step["target"] == "User input/template file", f"repair target was misleading: {step}")
    assert_true(not report["repair_plan"]["commands"].get("rebuild_current_docx"), "user-file-only error should not suggest rebuilding build_generated.py")
    assert_true("不要只修改 `build_generated.py`" in report["repair_plan"]["copy_to_ai_prompt"], "AI prompt should not route user-file-only errors to generated-script edits")

    work2 = new_workdir("qa_user_confirmation_routing")
    content = base_content(["Body text"])
    content["_meta"]["non_body_images"] = [{"location": "section_1_header", "target": "media/image1.png"}]
    write_json(work2 / "content.json", content)
    write_json(work2 / "format.json", base_format())
    write_json(work2 / "workflow_mode.json", {"mode": "user"})
    (work2 / "build_generated.py").write_text("# synthetic\n", encoding="utf-8")
    doc2 = Document()
    doc2.add_paragraph("1 Introduction")
    doc2.add_paragraph("Body text")
    doc2.save(work2 / "out.docx")
    write_json(work2 / "build_manifest.json", {"schema_version": 1, "counts": {}})
    report2 = check_output(str(work2), mode="user", output_docx_name="out.docx")
    issue2 = next(item for item in report2["issues"] if item["code"] == "NON_BODY_IMAGE_UNSUPPORTED")
    assert_true(issue2["active_owner"] == "User input/template file", f"user-confirmation issue target was misleading: {issue2}")
