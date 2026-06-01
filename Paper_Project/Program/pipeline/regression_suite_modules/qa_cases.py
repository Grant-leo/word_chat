"""Structural QA and manifest regression cases."""
from __future__ import annotations

import json
from xml.etree import ElementTree as ET

from docx import Document

from qa_checker import check_output
from qa_checker_modules.content_samples import _content_toc_pollution_samples
from qa_checker_modules.content_metrics import _count_content_formulas, _count_content_images
from qa_checker_modules.registry import OWNER_BY_CODE
from qa_checker_modules.report_phase import build_report
from qa_checker_modules.repair import build_repair_plan
from qa_checker_modules.repair_guides import REPAIR_GUIDES
from qa_checker_modules.reports import repair_plan_to_markdown
from qa_conformance_modules.reports import build_report as build_conformance_report
from qa_conformance_modules.content_checks import _expected_paragraphs, _find_body_start_index, _find_para_by_text
from qa_visual_modules.checks import check_visual

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
def qa_toc_pollution_allows_numbered_multilevel_headings() -> None:
    content = {
        "sections": [
            {"heading": "1 Chapter 1", "level": 1, "paragraphs": ["Opening paragraph."]},
            {"heading": "1.1 Section 1.1", "level": 2, "paragraphs": ["Nested body paragraph."]},
            {"heading": "1.1.1 Detail", "level": 3, "paragraphs": ["Third-level content."]},
        ]
    }
    assert_true(not _content_toc_pollution_samples(content), "legitimate numbered multilevel headings were flagged as TOC pollution")


@case
def conformance_finds_inline_math_paragraphs_by_linearized_text() -> None:
    xml = """
    <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
         xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
      <w:r><w:t>Inline formula </w:t></w:r>
      <m:oMath>
        <m:r><m:t>E</m:t></m:r>
        <m:r><m:t>=</m:t></m:r>
        <m:r><m:t>m</m:t></m:r>
        <m:sSup><m:e><m:r><m:t>c</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>
      </m:oMath>
      <w:r><w:t> should remain editable.</w:t></w:r>
    </w:p>
    """
    para = ET.fromstring(xml)
    found = _find_para_by_text([para], "Inline formula E=mc^2 should remain editable.")
    assert_true(found is para, "conformance QA did not match paragraph text containing editable inline math")


@case
def conformance_body_start_keeps_default_body_before_first_heading() -> None:
    def para(text: str) -> ET.Element:
        return ET.fromstring(
            '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f"<w:r><w:t>{text}</w:t></w:r>"
            "</w:p>"
        )

    content = {
        "sections": [
            {
                "heading": "正文",
                "level": 1,
                "role": "body",
                "paragraphs": [
                    "This paragraph appears before the first explicit heading.",
                    "It must still be checked by strict QA.",
                ],
            },
            {
                "heading": "Methods",
                "level": 2,
                "role": "",
                "paragraphs": ["Method body text."],
            },
        ],
        "references": [],
    }
    expected = _expected_paragraphs(content)
    paragraphs = [
        para("Cover Title"),
        para("目 录"),
        para("Methods\t1"),
        para("This paragraph appears before the first explicit heading."),
        para("It must still be checked by strict QA."),
        para("Methods"),
        para("Method body text."),
    ]
    start = _find_body_start_index(paragraphs, expected)
    used: set[int] = set()
    assert_true(
        _find_para_by_text(paragraphs[start:], "This paragraph appears before the first explicit heading.", used) is not None,
        f"strict QA body start skipped default body paragraphs before the first explicit heading; start={start}",
    )


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
def qa_repair_plan_surfaces_next_action_and_resume_route() -> None:
    work = new_workdir("qa_repair_next_action_user_file")
    write_json(
        work / "workflow_mode.json",
        {
            "mode": "user",
            "template": "demo_template.docx",
            "content": "paper.md",
            "qa_level": "strict",
        },
    )
    user_file_report = {
        "mode": "user",
        "passed": False,
        "issues": [{"code": "CONTENT_IMAGE_MISSING", "severity": "error", "message": "missing image"}],
        "counts": {},
    }
    user_file_plan = build_repair_plan(user_file_report, str(work))
    assert_true("CONTENT_IMAGE_MISSING" in user_file_plan.get("next_action", ""), f"repair plan lost leading code: {user_file_plan}")
    assert_true("把缺失图片放回" in user_file_plan.get("next_action", ""), f"repair plan lost concrete user action: {user_file_plan}")
    assert_true(user_file_plan.get("resume_scope") == "input_files", f"user-file blocker should route to input files: {user_file_plan}")
    assert_true(
        user_file_plan.get("resume_command") == user_file_plan["commands"]["rerun_current_pipeline"],
        f"user-file blocker should resume with full pipeline rerun: {user_file_plan}",
    )
    assert_true(not user_file_plan["commands"].get("rebuild_current_docx"), f"user-file blocker should not suggest rebuild-only: {user_file_plan}")
    user_file_markdown = repair_plan_to_markdown(user_file_plan)
    assert_true("下一步" in user_file_markdown and "修复后运行" in user_file_markdown, f"repair markdown should show next action and resume command: {user_file_markdown}")

    work2 = new_workdir("qa_repair_next_action_generated_script")
    write_json(work2 / "workflow_mode.json", {"mode": "user"})
    generated_report = {
        "mode": "user",
        "passed": False,
        "issues": [{"code": "MISSING_DOCX", "severity": "error", "message": "missing docx"}],
        "counts": {},
    }
    generated_plan = build_repair_plan(generated_report, str(work2))
    assert_true("MISSING_DOCX" in generated_plan.get("next_action", ""), f"generated-script repair lost leading code: {generated_plan}")
    assert_true(generated_plan.get("resume_scope") == "current_docx", f"generated-script repair should route to current DOCX rebuild: {generated_plan}")
    assert_true(
        generated_plan.get("resume_command") == generated_plan["commands"]["rebuild_current_docx"],
        f"generated-script repair should resume with build_generated.py rebuild: {generated_plan}",
    )
    assert_true("修复后运行" in generated_plan.get("copy_to_ai_prompt", ""), f"AI prompt should include the resume command: {generated_plan}")


@case
def qa_repair_plan_warning_only_is_not_plain_pass() -> None:
    work = new_workdir("qa_repair_warning_only")
    write_json(
        work / "workflow_mode.json",
        {
            "mode": "user",
            "template": "demo_template.docx",
            "content": "paper.docx",
            "qa_level": "basic",
        },
    )
    report = {
        "mode": "user",
        "passed": True,
        "issues": [{"code": "REFERENCES_MISSING", "severity": "warning", "message": "references missing"}],
        "counts": {},
    }
    plan = build_repair_plan(report, str(work))
    markdown = repair_plan_to_markdown(plan)
    assert_true(plan["passed"] is True, f"warning-only QA should remain non-blocking: {plan}")
    assert_true(plan["warnings"] == 1, f"warning count should be visible: {plan}")
    assert_true("REFERENCES_MISSING" in plan["next_action"], f"repair plan lost warning code: {plan['next_action']}")
    assert_true("参考文献" in plan["next_action"], f"repair plan lost warning action: {plan['next_action']}")
    assert_true("警告" in plan["summary"], f"warning-only summary should not sound like a plain pass: {plan['summary']}")
    assert_true(plan["resume_scope"] == "warning_review", f"warning-only plan should route to warning review: {plan}")
    assert_true(
        plan["resume_command"] == plan["commands"]["rerun_current_pipeline"],
        f"warning-only plan should preserve a rerun command for fixes: {plan}",
    )
    assert_true("QA 已通过，仍建议" not in markdown, f"repair markdown should not hide warning-only issues: {markdown}")
    assert_true("REFERENCES_MISSING" in markdown and "警告" in markdown, f"repair markdown should surface warning details: {markdown}")


@case
def qa_report_markdown_lists_repair_plan_open_first_files() -> None:
    from qa_checker_modules.reports import report_to_markdown

    report = {
        "mode": "user",
        "passed": False,
        "output_dir_name": "demo",
        "next_action": "优先处理 `CONTENT_IMAGE_MISSING`。",
        "counts": {},
        "issues": [
            {
                "code": "CONTENT_IMAGE_MISSING",
                "severity": "error",
                "message": "missing image",
                "active_owner": "User input/template file",
            }
        ],
        "repair_plan": {
            "summary": "QA 发现 1 个阻断错误。",
            "output_dir": "Outputs/demo",
            "open_first": [
                "qa_repair_plan.md",
                "qa_report.md",
                "内容提取.md",
                "build_manifest.json",
                "最终论文.docx",
            ],
            "commands": {
                "rerun_current_pipeline": "python run_pipeline.py --mode user --template t.docx --content c.docx",
                "rebuild_current_docx": "python Outputs/demo/build_generated.py",
            },
            "steps": [],
        },
    }
    markdown = report_to_markdown(report)
    assert_true("## 先打开这些文件" in markdown, f"QA report should surface repair-plan review files: {markdown}")
    assert_true("Outputs/demo/qa_repair_plan.md" in markdown, f"QA report should point to repair plan: {markdown}")
    assert_true("Outputs/demo/内容提取.md" in markdown, f"QA report should point to content summary: {markdown}")
    assert_true("Outputs/demo/build_manifest.json" in markdown, f"QA report should point to build manifest: {markdown}")
    assert_true("Outputs/demo/最终论文.docx" in markdown, f"QA report should point to final DOCX: {markdown}")
    assert_true("## 可执行命令" in markdown, f"QA report should keep commands separate from open-first files: {markdown}")
    assert_true(markdown.index("## 先打开这些文件") < markdown.index("## 可执行命令"), f"open-first files should be listed before commands: {markdown}")


@case
def qa_report_next_action_names_first_repair_step() -> None:
    work = new_workdir("qa_next_action_first_step")
    write_json(work / "workflow_mode.json", {"mode": "user"})
    report = build_report(
        str(work),
        "user",
        {},
        [
            {
                "code": "CONTENT_IMAGE_MISSING",
                "severity": "error",
                "message": "missing image",
                "active_owner": "User input/template file",
            }
        ],
    )
    assert_true("CONTENT_IMAGE_MISSING" in report["next_action"], f"next_action lost the issue code: {report['next_action']}")
    assert_true("把缺失图片放回" in report["next_action"], f"next_action lost the beginner repair action: {report['next_action']}")
    assert_true("用户确认或补充输入文件" in report["next_action"], f"user-file routing disappeared: {report['next_action']}")


@case
def qa_report_next_action_names_warning_step() -> None:
    work = new_workdir("qa_next_action_warning_step")
    write_json(work / "workflow_mode.json", {"mode": "user"})
    report = build_report(
        str(work),
        "user",
        {},
        [
            {
                "code": "REFERENCES_MISSING",
                "severity": "warning",
                "message": "references missing",
            }
        ],
    )
    assert_true(report["passed"] is True, f"warning-only QA should remain non-blocking: {report}")
    assert_true("REFERENCES_MISSING" in report["next_action"], f"warning next_action lost the issue code: {report['next_action']}")
    assert_true("参考文献" in report["next_action"], f"warning next_action lost the beginner action: {report['next_action']}")
    assert_true("警告" in report["next_action"] or "warning" in report["next_action"], f"warning next_action should not sound like plain pass: {report['next_action']}")


@case
def qa_json_reports_expose_explicit_status_labels() -> None:
    work = new_workdir("qa_json_status_labels")
    write_json(work / "workflow_mode.json", {"mode": "developer"})
    failed = build_report(
        str(work),
        "developer",
        {},
        [{"code": "CONTENT_IMAGE_MISSING", "severity": "error", "message": "missing image"}],
    )
    assert_true(failed["status"] == "failed", f"structural error status should be failed: {failed}")
    assert_true(failed["result_label"] == "未通过", f"structural error label should be explicit: {failed}")

    warning = build_report(
        str(work),
        "developer",
        {},
        [{"code": "REFERENCES_MISSING", "severity": "warning", "message": "missing refs"}],
    )
    assert_true(warning["status"] == "passed_with_warnings", f"structural warning status should be explicit: {warning}")
    assert_true(warning["result_label"] == "通过但有警告", f"structural warning label should be explicit: {warning}")

    clean = build_report(str(work), "developer", {}, [])
    assert_true(clean["status"] == "passed", f"structural clean status should be explicit: {clean}")
    assert_true(clean["result_label"] == "通过", f"structural clean label should be explicit: {clean}")

    conformance = build_conformance_report(
        str(work),
        "developer",
        {},
        [{"code": "STYLE_MISMATCH", "severity": "warning", "message": "style differs"}],
        project_root=str(work),
    )
    assert_true(conformance["status"] == "passed_with_warnings", f"strict QA status should expose warnings: {conformance}")
    assert_true(conformance["result_label"] == "通过但有警告", f"strict QA label should expose warnings: {conformance}")

    visual = check_visual(str(work), output_docx_name="missing.docx", project_root=str(work))
    assert_true(visual["status"] == "failed", f"visual missing-DOCX status should fail explicitly: {visual}")
    assert_true(visual["result_label"] == "未通过", f"visual missing-DOCX label should fail explicitly: {visual}")


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
