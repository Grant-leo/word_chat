"""Content parser regression cases."""
from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any, Dict, List

from docx import Document
from docx.enum.section import WD_SECTION
from docx.shared import Pt
from lxml import etree

from content_parser import extract as extract_docx_content
from content_parser_modules.caption_flow import (
    is_figure_caption,
    is_table_caption,
    normalize_caption_spacing,
    pair_figure_blocks,
)
from content_parser_modules.body_dispatcher import append_text_or_code, parse_body_sections
from content_parser_modules.formula_extractor import _strip_trailing_formula_labels_from_xml
from content_parser_modules.image_extractor import ImageRegistry
from content_parser_modules.paragraph_stream import append_stream_run_group
from content_parser_modules.reference_collector import ReferenceCollector, is_reference_heading
from content_parser_modules.section_builder import (
    filter_content_sections,
    make_body_section,
    mark_first_body_page_break,
    postprocess_section_paragraphs,
)
from content_parser_modules.text_cleaner import clean_code_text, clean_text_artifacts
from formula_semantics import (
    CATEGORY_CONTAMINATED,
    CATEGORY_DISPLAY_MATH,
    CATEGORY_QUANTITY_TEXT,
    classify_formula_text,
    is_formula_problem_text,
    looks_like_formula_text as semantic_looks_like_formula_text,
    split_inline_math_spans,
)
from latex_omath import latex_to_omath
from qa_checker import check_output, write_reports

from regression_suite_modules.generated_docx import make_vml_picture_docx, omath_count, omath_para_count, run_generated_case
from regression_suite_modules.harness import (
    PNG_1X1,
    assert_true,
    base_content,
    base_format,
    case,
    fail,
    new_workdir,
    write_json,
    write_sample_png,
)

@case
def caption_flow_pairs_images_with_nearby_captions() -> None:
    paragraphs = [
        {"role": "image", "image": "a.png"},
        "intervening prose",
        {"role": "figure_caption", "text": "Fig. 1 Demo"},
        {"role": "image", "image": "b.png"},
        {"role": "image", "image": "c.png"},
        {"role": "figure_caption", "text": "Figure 2 First"},
        {"role": "figure_caption", "text": "Figure 3 Second"},
    ]
    paired = pair_figure_blocks(paragraphs)
    assert_true(paired[0] == {"role": "figure", "image": "a.png", "caption": "Fig. 1 Demo"}, "image/body/caption was not paired")
    assert_true(paired[1] == "intervening prose", "intervening prose was not preserved after pairing")
    assert_true(paired[2] == {"role": "figure", "image": "b.png", "caption": "Figure 2 First"}, "first stacked image was not paired")
    assert_true(paired[3] == {"role": "figure", "image": "c.png", "caption": "Figure 3 Second"}, "second stacked image was not paired")
    assert_true(is_figure_caption("\u56fe1\u7ed3\u679c\u5bf9\u6bd4"), "Chinese figure caption was not detected")
    assert_true(is_figure_caption("图 1 机器学习研究流程示意图"), "Chinese noun-phrase figure caption was not detected")
    assert_true(not is_figure_caption("图 1 展示了从数据到决策的机器学习研究流程。"), "figure reference prose was misclassified as caption")
    assert_true(is_table_caption("表 1 不同模型在验证集上的表现"), "Chinese table caption was not detected")
    assert_true(not is_table_caption("表 1 显示 XGBoost 在多数指标上更高。"), "table reference prose was misclassified as caption")
    assert_true(normalize_caption_spacing("\u56fe1\u7ed3\u679c") == "\u56fe 1 \u7ed3\u679c", "Chinese figure caption spacing was not normalized")


@case
def text_cleaner_removes_editor_noise_and_preserves_code_lines() -> None:
    assert_true(clean_text_artifacts("[label](https://example.test)") == "label", "markdown link label was not preserved")
    assert_true(clean_text_artifacts("Plain Text") == "", "editor noise was not removed")
    assert_true(clean_text_artifacts("\u590d\u5236") == "", "Chinese editor noise was not removed")
    code = clean_code_text(" interface  Gi0/0/1 \nPlain Text\n ip   address  10.0.0.1 ")
    assert_true(code == "interface Gi0/0/1\nip address 10.0.0.1", f"code cleanup changed unexpectedly: {code!r}")


@case
def paragraph_stream_run_group_preserves_text_and_math_semantics() -> None:
    section = {"paragraphs": []}

    def append_text(section_obj: Dict[str, Any], text: str, in_appendix: bool = False) -> None:
        section_obj["paragraphs"].append({"role": "text", "text": text, "appendix": in_appendix})

    append_stream_run_group(
        section,
        [{"type": "text", "text": "plain text"}],
        append_text_or_code_func=append_text,
        in_appendix=True,
    )
    assert_true(section["paragraphs"][0] == {"role": "text", "text": "plain text", "appendix": True}, "plain run group bypassed text callback")

    append_stream_run_group(
        section,
        [{"type": "math", "text": "a=b", "math": [{"text": "a=b", "had_number_label": True}]}],
        append_text_or_code_func=append_text,
    )
    formula = section["paragraphs"][1]
    assert_true(formula["role"] == "formula", "math-only run group did not become a formula")
    assert_true(formula["numbered"] is True, "math had_number_label was not preserved")


@case
def reference_collector_exits_before_backmatter_and_keeps_tables() -> None:
    collector = ReferenceCollector(
        clean_text_func=clean_text_artifacts,
        is_backmatter_heading_func=lambda text: text == "Appendix A",
        normalize_heading_spacing_func=lambda text: text,
        classify_section_role_func=lambda heading, level: "appendix",
        table_rows_look_like_code_func=lambda rows: rows and rows[0] and rows[0][0].startswith("interface"),
        code_text_from_table_rows_func=lambda rows, clean_code_func=None: clean_code_func(rows[0][0]) if clean_code_func else rows[0][0],
        clean_code_func=clean_code_text,
    )
    assert_true(is_reference_heading("\u53c2\u8003\u6587\u732e"), "Chinese reference heading was not detected")
    assert_true(collector.start_if_heading("References"), "English reference heading was not accepted")
    assert_true(collector.consume_text("[1]  Example  Paper"), "active collector did not consume reference text")
    assert_true(collector.consume_table_rows([["interface  Gi0/0/1\n ip  address  10.0.0.1"]]), "active collector did not consume reference table")
    backmatter = collector.exit_to_backmatter_section("Appendix A", 1)
    assert_true(backmatter and backmatter["role"] == "appendix", "collector did not exit on backmatter")
    refs = collector.finish()
    assert_true(refs[0] == "[1] Example Paper", f"reference text was not cleaned: {refs[0]!r}")
    assert_true(refs[1]["role"] == "code", "reference code table was not preserved as code")
    assert_true(refs[1]["code"] == "interface Gi0/0/1\nip address 10.0.0.1", "reference code table was not cleaned")


@case
def section_builder_filters_placeholder_and_finalizes_blocks() -> None:
    sections = [
        make_body_section(),
        {"heading": "Abstract", "level": 1, "role": "en_abstract", "paragraphs": ["Summary"], "images": []},
        {"heading": "1 Introduction", "level": 1, "role": "body", "paragraphs": [], "images": []},
        {
            "heading": "2 Results",
            "level": 1,
            "role": "body",
            "paragraphs": [
                {"role": "image", "image": "fig1.png"},
                {"role": "figure_caption", "text": "Fig. 1. Result overview"},
            ],
            "images": ["fig1.png"],
        },
    ]
    filtered = filter_content_sections(sections)
    assert_true(filtered[0]["heading"] == "Abstract", "initial body placeholder was not filtered")
    assert_true(filtered[1]["heading"] == "1 Introduction", "empty structural heading was dropped")
    mark_first_body_page_break(filtered)
    assert_true(not filtered[0].get("page_break_before"), "front matter was marked as first body page")
    assert_true(filtered[1].get("page_break_before") is True, "first body section was not marked for page break")
    postprocess_section_paragraphs(filtered)
    assert_true(filtered[2]["paragraphs"][0]["role"] == "figure", "image and figure caption were not paired")
    assert_true(filtered[2]["paragraphs"][0]["caption"] == "Fig. 1. Result overview", "figure caption text changed")


@case
def body_dispatcher_routes_sections_references_tables_and_appendix_code() -> None:
    work = new_workdir("body_dispatcher")
    fig_dir = work / "figures"
    fig_dir.mkdir(parents=True, exist_ok=True)
    doc = Document()
    doc.add_paragraph("1 Introduction")
    doc.add_paragraph("报名序号: [报名序号]")
    doc.add_paragraph("图 1 系统结构示意")
    doc.add_paragraph("参考文献")
    table = doc.add_table(rows=1, cols=1)
    table.cell(0, 0).text = "[1] Zhang S. Synthetic reference."
    doc.add_paragraph("附录 A")
    doc.add_paragraph("interface Gi0/0/1\n ip address 10.0.0.1")

    scratch = {"paragraphs": []}
    append_text_or_code(scratch, r"$$a=b$$")
    assert_true(scratch["paragraphs"][0]["role"] == "formula", "dispatcher text append did not preserve display math")

    result = parse_body_sections(doc, 0, ImageRegistry(str(fig_dir), "dispatch_img"))
    sections = filter_content_sections(result.sections)
    assert_true(sections[0]["heading"] == "1 Introduction", "body heading was not routed into a section")
    assert_true(result.placeholders_removed == 1, "unfilled placeholder paragraph was not filtered")
    assert_true(sections[0]["paragraphs"][0]["role"] == "figure_caption", "figure caption was not classified")
    assert_true(result.references and result.references[0]["role"] == "table", "reference table was not captured before appendix")
    assert_true(sections[1]["role"] == "appendix", "back matter did not exit reference collection")
    assert_true(sections[1]["paragraphs"][0]["role"] == "code", "appendix command text was not routed as code")


@case
def content_parser_preserves_table_at_start_of_body() -> None:
    work = new_workdir("parser_table_first")

    docx = work / "table_first_then_text.docx"
    doc = Document()
    table = doc.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "Metric"
    table.cell(0, 1).text = "Value"
    table.cell(1, 0).text = "Accuracy"
    table.cell(1, 1).text = "98%"
    doc.add_paragraph("Body paragraph after the opening table.")
    doc.save(docx)

    content = extract_docx_content(str(docx), output_dir=str(work / "out"))
    items = [item for sec in content.get("sections") or [] for item in sec.get("paragraphs") or []]
    table_items = [item for item in items if isinstance(item, dict) and item.get("table_rows")]
    assert_true(len(table_items) == 1, f"opening table was not preserved in content stream: {items}")
    assert_true(table_items[0]["table_rows"][1] == ["Accuracy", "98%"], "opening table rows changed")

    table_only_docx = work / "table_only.docx"
    doc = Document()
    table = doc.add_table(rows=1, cols=2)
    table.cell(0, 0).text = "Only"
    table.cell(0, 1).text = "Table"
    doc.save(table_only_docx)

    table_only = extract_docx_content(str(table_only_docx), output_dir=str(work / "only_out"))
    only_items = [item for sec in table_only.get("sections") or [] for item in sec.get("paragraphs") or []]
    assert_true(
        any(isinstance(item, dict) and item.get("table_rows") for item in only_items),
        f"table-only document became empty content: {table_only}",
    )


@case
def content_parser_does_not_count_cover_table_as_body_table() -> None:
    work = new_workdir("parser_cover_table_count")
    docx = work / "cover_table_count.docx"
    doc = Document()
    cover = doc.add_table(rows=1, cols=2)
    cover.cell(0, 0).text = "\u5b66\u6821\u7f16\u7801"
    cover.cell(0, 1).text = "10001"
    title = "\u9762\u5411\u7eff\u7535\u6d88\u7eb3\u7684\u591a\u6e90\u80fd\u6e90\u8c03\u5ea6\u7814\u7a76"
    p = doc.add_paragraph(title)
    p.runs[0].font.size = Pt(16)
    doc.add_paragraph("\u6458\u8981\uff1a\u672c\u6587\u6458\u8981\u7528\u4e8e\u9a8c\u8bc1\u5c01\u9762\u8868\u683c\u4e0d\u5e94\u8ba1\u5165\u6b63\u6587\u8868\u683c\u3002")
    doc.add_paragraph("\u5173\u952e\u8bcd\uff1a\u7eff\u7535\uff1b\u8c03\u5ea6")
    doc.add_heading("\u7b2c1\u7ae0 \u7eea\u8bba", 1)
    doc.add_paragraph("\u6b63\u6587\u6bb5\u843d\u3002")
    doc.save(docx)

    content = extract_docx_content(str(docx), output_dir=str(work / "out"))
    meta = content.get("_meta") or {}
    assert_true(meta.get("source_tables_count") == 1, f"source table count was not recorded: {meta}")
    assert_true(meta.get("tables_count") == 0, f"cover table was counted as a body table: {meta}")


@case
def low_res_image_fragment_is_reported_and_not_upscaled() -> None:
    img_src = new_workdir("tiny_image_src")
    write_sample_png(img_src / "tiny.png", width=76, height=18)
    content = base_content([
        {"role": "figure", "image": "tiny.png", "caption": "Figure 1 broken label shard"}
    ])
    content["_meta"]["images_dir"] = str(img_src)
    content["_meta"]["images_extracted"] = 1
    content["sections"][0]["images"] = ["tiny.png"]
    result = run_generated_case("tiny_image_fragment", content)
    codes = [item["code"] for item in result["report"]["issues"]]
    assert_true("LOW_RES_IMAGE_FRAGMENT" in codes, "QA did not report a low-resolution image fragment")
    issue = next(item for item in result["report"]["issues"] if item["code"] == "LOW_RES_IMAGE_FRAGMENT")
    assert_true(issue["severity"] == "warning", f"contained image fragment should be a warning: {issue}")
    assert_true(result["manifest"]["counts"].get("content_image_fragments_contained") == 1, "contained image fragment was not counted")
    m = re.search(r"<wp:extent[^>]+cx=\"(\d+)\"[^>]+cy=\"(\d+)\"", result["xml"])
    assert_true(bool(m), "generated DOCX did not contain an image extent")
    width_inches = int(m.group(1)) / 914400
    assert_true(width_inches < 2.0, f"tiny image was upscaled to page width: {width_inches:.2f}in")


@case
def small_square_icon_is_not_reported_as_low_res_fragment() -> None:
    img_src = new_workdir("small_icon_src")
    write_sample_png(img_src / "icon.png", width=64, height=64)
    content = base_content([
        {"role": "figure", "image": "icon.png", "caption": "QR code icon"}
    ])
    content["_meta"]["images_dir"] = str(img_src)
    content["_meta"]["images_extracted"] = 1
    content["sections"][0]["images"] = ["icon.png"]
    result = run_generated_case("small_square_icon", content)
    codes = [item["code"] for item in result["report"]["issues"]]
    assert_true("LOW_RES_IMAGE_FRAGMENT" not in codes, f"legitimate small square image was reported as fragment: {result['report']['issues']}")


@case
def qa_manifest_detects_missing_image_render() -> None:
    img_src = new_workdir("image_missing_src")
    write_sample_png(img_src / "dot.png")
    content = base_content([
        {"role": "figure", "image": "dot.png", "caption": "Figure 1 sample"}
    ])
    content["_meta"]["images_dir"] = str(img_src)
    content["_meta"]["images_extracted"] = 1
    content["sections"][0]["images"] = ["dot.png"]
    result = run_generated_case("qa_missing_image", content)
    manifest_path = result["work"] / "build_manifest.json"
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    manifest["counts"]["content_images_rendered"] = 0
    write_json(manifest_path, manifest)
    report = check_output(str(result["work"]), mode="developer", output_docx_name="out.docx")
    codes = [item["code"] for item in report["issues"]]
    assert_true("IMAGE_COUNT_MISMATCH" in codes, "QA did not trust manifest for image mismatch")
    repair = report.get("repair_plan") or {}
    assert_true(repair.get("blocking_errors", 0) >= 1, "repair plan did not count blocking errors")
    steps = repair.get("steps") or []
    assert_true(steps and steps[0].get("severity") == "error", "repair plan did not prioritize errors")
    assert_true(any(step.get("code") == "IMAGE_COUNT_MISMATCH" for step in steps), "repair plan omitted image repair guidance")
    write_reports(report, str(result["work"]))
    assert_true((result["work"] / "qa_repair_plan.md").exists(), "repair plan markdown was not written")
    assert_true((result["work"] / "qa_fix_prompt.txt").exists(), "AI fix prompt was not written")


@case
def qa_manifest_detects_missing_formula_render() -> None:
    content = base_content([
        {"role": "formula", "latex": "a=b+c", "text": "a=b+c", "numbered": False}
    ])
    result = run_generated_case("qa_missing_formula", content)
    manifest_path = result["work"] / "build_manifest.json"
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    manifest["counts"]["content_formulas_rendered"] = 0
    write_json(manifest_path, manifest)
    report = check_output(str(result["work"]), mode="developer", output_docx_name="out.docx")
    codes = [item["code"] for item in report["issues"]]
    assert_true("FORMULA_COUNT_MISMATCH" in codes, "QA did not trust manifest for formula mismatch")


@case
def omml_display_item_builds_display_formula() -> None:
    xml = latex_to_omath("a=b+c", display=True)
    content = base_content([
        {
            "role": "formula",
            "source": "omml",
            "text": "a=b+c",
            "math": [{"type": "display", "xml": xml, "text": "a=b+c"}],
            "numbered": False,
        }
    ])
    result = run_generated_case("omml_display", content)
    assert_true(result["manifest"]["counts"]["display_formulas_rendered"] == 1, "OMML display formula was not counted")
    assert_true(omath_para_count(result["xml"]) == 1, "OMML display formula did not remain display")


@case
def content_parser_preserves_mixed_docx_tokens() -> None:
    work = new_workdir("parser_mixed")
    img = work / "dot.png"
    img.write_bytes(PNG_1X1)
    docx = work / "mixed.docx"
    doc = Document()
    p = doc.add_paragraph()
    p.add_run("before ")
    p._element.append(etree.fromstring(latex_to_omath("x^2", display=False).encode("utf-8")))
    p.add_run(" after")
    p_img = doc.add_paragraph()
    p_img.add_run("image before ")
    p_img.add_run().add_picture(str(img))
    p_img._element.append(etree.fromstring(latex_to_omath("z^2", display=False).encode("utf-8")))
    p_img.add_run(" image after")
    p_disp = doc.add_paragraph()
    p_disp._element.append(etree.fromstring(latex_to_omath("a=b+c", display=True).encode("utf-8")))
    doc.save(docx)

    content = extract_docx_content(str(docx), output_dir=str(work))
    paragraphs = [p for sec in content["sections"] for p in sec.get("paragraphs", [])]
    rich_items = [p for p in paragraphs if isinstance(p, dict) and p.get("role") == "rich_text"]
    formula_items = [p for p in paragraphs if isinstance(p, dict) and p.get("role") == "formula"]
    image_items = [p for p in paragraphs if isinstance(p, dict) and p.get("role") == "image"]
    assert_true(any([r.get("type") for r in item.get("runs", [])] == ["text", "math", "text"] for item in rich_items), "mixed text/math run order was not preserved")
    assert_true(any((item.get("runs") or [])[0].get("text") == "before " and (item.get("runs") or [])[-1].get("text") == " after" for item in rich_items), "inline formula surrounding spaces were not preserved")
    assert_true(image_items, "image token disappeared from mixed paragraph")
    assert_true(any((item.get("math") or [{}])[0].get("type") == "display" for item in formula_items), "display OMML was not extracted as formula")


@case
def content_parser_extracts_vml_pictures() -> None:
    work = new_workdir("parser_vml")
    img = work / "dot.png"
    img.write_bytes(PNG_1X1)
    src_docx = work / "drawing.docx"
    vml_docx = work / "vml.docx"
    doc = Document()
    doc.add_paragraph().add_run().add_picture(str(img))
    doc.save(src_docx)
    make_vml_picture_docx(src_docx, vml_docx)
    content = extract_docx_content(str(vml_docx), output_dir=str(work))
    paragraphs = [p for sec in content["sections"] for p in sec.get("paragraphs", [])]
    assert_true(any(isinstance(p, dict) and p.get("role") == "image" for p in paragraphs), "VML picture was not extracted")


@case
def content_parser_detects_latex_delimited_formula_paragraphs() -> None:
    work = new_workdir("parser_latex_delimited")
    docx = work / "latex_delimited.docx"
    doc = Document()
    doc.add_paragraph("1 Formula Cases")
    doc.add_paragraph(r"$$L=\lim_{n\to\infty}\frac{1}{n}\sum_{i=1}^{n}x_i$$")
    doc.add_paragraph(r"$$M=\begin{matrix}a&b\\c&d\end{matrix}$$")
    doc.add_paragraph(r"$$I=\int_0^T f(t)\,dt$$")
    doc.add_paragraph(r"$$E_{total}=\sum_{t=1}^{24}P(t)\Delta t$$ (1.1)")
    doc.add_paragraph(r"$$R_{green}=\frac{E_{renew}}{E_{total}}\times100\%$$" + " \uff08A.1\uff09")
    doc.save(docx)

    content = extract_docx_content(str(docx), output_dir=str(work))
    paragraphs = [p for sec in content["sections"] for p in sec.get("paragraphs", [])]
    formulas = [p for p in paragraphs if isinstance(p, dict) and p.get("role") == "formula"]
    assert_true(len(formulas) == 5, f"LaTeX-delimited paragraphs were not all formulas: {paragraphs}")
    assert_true(all(f.get("source") == "latex" and f.get("latex") for f in formulas), "LaTeX delimiters were not stripped into latex fields")
    assert_true(all("$$" not in (f.get("latex") or "") for f in formulas), "LaTeX delimiters leaked into formula latex")
    assert_true(any(f.get("text", "").endswith("\uff08A.1\uff09") and f.get("latex", "").startswith("R_{green}") for f in formulas), "appendix formula label was not handled")


@case
def formula_semantics_classifies_quantities_and_contamination() -> None:
    quantity = "全年发电量为 172.04 MWh"
    equation = "P_total(t)=P_AK+P_PM+P_NH3+P_L(t)=20.75+6*p_L(t)"
    contaminated = "当PRE(t)-PL(t)<0.1*41.5=4.15MW时,该时段无法开机;产量约束为="

    quantity_result = classify_formula_text(quantity)
    equation_result = classify_formula_text(equation)
    contaminated_result = classify_formula_text(contaminated)

    assert_true(quantity_result.category == CATEGORY_QUANTITY_TEXT, f"quantity became {quantity_result}")
    assert_true(not semantic_looks_like_formula_text(quantity), "quantity/unit text was mistaken for a formula")
    assert_true(equation_result.category == CATEGORY_DISPLAY_MATH, f"equation became {equation_result}")
    assert_true(semantic_looks_like_formula_text(equation), "standalone equation was not recognized")
    assert_true(contaminated_result.category == CATEGORY_CONTAMINATED, f"contaminated formula became {contaminated_result}")
    assert_true(not semantic_looks_like_formula_text(contaminated), "contaminated narrative was accepted as a formula")


@case
def content_parser_marks_formula_semantic_problems() -> None:
    work = new_workdir("parser_formula_semantics")
    docx = work / "formula_semantics.docx"
    quantity = "全年发电量为 172.04 MWh"
    equation = "P_total(t)=P_AK+P_PM+P_NH3+P_L(t)=20.75+6*p_L(t)"
    contaminated = "当PRE(t)-PL(t)<0.1*41.5=4.15MW时,该时段无法开机;产量约束为="
    doc = Document()
    doc.add_paragraph("1 Formula Semantics")
    doc.add_paragraph(quantity)
    doc.add_paragraph(equation)
    doc.add_paragraph(contaminated)
    doc.save(docx)

    content = extract_docx_content(str(docx), output_dir=str(work))
    paragraphs = [p for sec in content["sections"] for p in sec.get("paragraphs", [])]
    formulas = [p for p in paragraphs if isinstance(p, dict) and p.get("role") == "formula"]
    problems = [p for p in paragraphs if isinstance(p, dict) and p.get("role") == "formula_problem"]

    assert_true(quantity in paragraphs, "quantity text should remain ordinary text")
    assert_true(any(p.get("text") == equation for p in formulas), "standalone equation was not extracted as formula")
    assert_true(any(p.get("text") == contaminated for p in problems), "contaminated formula text was not marked")


@case
def formula_semantics_splits_inline_math_without_units() -> None:
    text = "模型中 P_total(t)=20.75+6*p_L(t)，全年电量为 172.04 MWh。"
    spans = split_inline_math_spans(text)
    assert_true(len(spans) == 1, f"expected one inline formula span, got {spans}")
    assert_true(spans[0]["text"] == "P_total(t)=20.75+6*p_L(t)", f"wrong inline span: {spans}")
    assert_true("172.04" not in spans[0]["text"], "quantity text was absorbed into inline formula")
    assert_true(not split_inline_math_spans("$100$ cost"), "plain dollar amount was mistaken for inline math")
    assert_true(not split_inline_math_spans("$$x=1$$"), "display math delimiter was mistaken for inline math")
    dollar_spans = split_inline_math_spans("变量 $x_i$ 表示第 i 个样本。")
    assert_true(dollar_spans and dollar_spans[0]["text"] == "x_i", f"valid dollar inline math was lost: {dollar_spans}")
    mixed = "由于约束仅为 u(t) = n 且 u(t) ∈ {0, 1}，这是标准选择问题。"
    mixed_spans = split_inline_math_spans(mixed)
    assert_true(not any("且" in str(s.get("text")) or "{" in str(s.get("text")) for s in mixed_spans), f"unsafe inline span accepted: {mixed_spans}")


@case
def prose_with_inline_formula_is_not_formula_problem() -> None:
    text = "假设三：每小时为一个调度时段。题目明确园区运行分析时段为 1 小时，即 ∆t = 1。"
    spans = split_inline_math_spans(text)
    assert_true(any(s.get("text") == "t = 1" for s in spans), f"inline formula was not found: {spans}")
    assert_true(not is_formula_problem_text(text), "ordinary prose with inline math should not block as formula_problem")


@case
def content_parser_builds_rich_text_for_plain_inline_formula() -> None:
    work = new_workdir("parser_plain_inline_formula")
    docx = work / "plain_inline_formula.docx"
    paragraph = "模型中 P_total(t)=20.75+6*p_L(t)，全年电量为 172.04 MWh。"
    doc = Document()
    doc.add_paragraph("1 Inline Formula")
    doc.add_paragraph(paragraph)
    doc.save(docx)

    content = extract_docx_content(str(docx), output_dir=str(work))
    paragraphs = [p for sec in content["sections"] for p in sec.get("paragraphs", [])]
    rich_items = [p for p in paragraphs if isinstance(p, dict) and p.get("role") == "rich_text"]
    assert_true(rich_items, f"inline formula paragraph was not converted to rich_text: {paragraphs}")
    runs = rich_items[0].get("runs") or []
    assert_true([r.get("type") for r in runs] == ["text", "math", "text"], f"inline formula runs are wrong: {runs}")
    result = run_generated_case("plain_inline_formula_render", content)
    assert_true(result["manifest"]["counts"]["inline_formulas_rendered"] == 1, "plain inline formula did not render as native inline math")
    assert_true(omath_count(result["xml"]) >= 1 and omath_para_count(result["xml"]) == 0, "plain inline formula rendered with wrong OMML shape")


@case
def content_parser_keeps_prose_inline_formula_as_rich_text() -> None:
    work = new_workdir("parser_prose_inline_formula")
    docx = work / "prose_inline_formula.docx"
    doc = Document()
    doc.add_paragraph("1 Formula Prose")
    doc.add_paragraph("假设三：每小时为一个调度时段。题目明确园区运行分析时段为 1 小时，即 ∆t = 1。")
    doc.save(docx)

    content = extract_docx_content(str(docx), output_dir=str(work))
    paragraphs = [p for sec in content["sections"] for p in sec.get("paragraphs", [])]
    assert_true(not any(isinstance(p, dict) and p.get("role") == "formula_problem" for p in paragraphs), f"prose inline formula became formula_problem: {paragraphs}")
    rich = [p for p in paragraphs if isinstance(p, dict) and p.get("role") == "rich_text"]
    assert_true(rich and any(run.get("type") == "math" for run in rich[0].get("runs") or []), f"prose inline formula was not rich_text math: {paragraphs}")


@case
def numeric_formula_denominator_is_not_heading() -> None:
    work = new_workdir("parser_numeric_fragment_heading")
    docx = work / "numeric_fragment_heading.docx"
    doc = Document()
    doc.add_paragraph("论文题目: Numeric Fragment")
    doc.add_paragraph("8 离网运行分析")
    p = doc.add_paragraph("41 .5")
    run = p.runs[0]
    run.font.size = Pt(16)
    run.bold = True
    doc.add_paragraph("This text should stay under the previous heading.")
    doc.save(docx)

    content = extract_docx_content(str(docx), output_dir=str(work))
    headings = [sec.get("heading") for sec in content.get("sections") or []]
    body_text = "\n".join(
        p if isinstance(p, str) else str(p.get("text") or "")
        for sec in content.get("sections") or []
        for p in sec.get("paragraphs", [])
    )
    assert_true("41 .5" not in headings, f"numeric formula denominator became heading: {headings}")
    assert_true("41 .5" in body_text and "previous heading" in body_text, f"numeric fragment/body text was lost: {body_text}")


@case
def content_parser_repairs_split_ratio_formula_layout() -> None:
    work = new_workdir("parser_split_ratio_layout")
    docx = work / "split_ratio_layout.docx"
    doc = Document()
    doc.add_paragraph("1 Split Formula Layout")
    doc.add_paragraph("日总用电量和新能源发电量分别为：")
    doc.add_paragraph("24 24")
    doc.add_paragraph("Etotal = ∑ Ptotal(t) · ∆t, ERE = ∑ PRE(t) · ∆t")
    doc.add_paragraph("t=1")
    doc.add_paragraph("三项绿电直连指标定义为：")
    doc.add_paragraph("t=1")
    doc.add_paragraph("rself")
    p1 = doc.add_paragraph()
    p1._element.append(etree.fromstring(latex_to_omath(r"=Etotal-Esell-Ebuy\times100\%(6)(3)", display=True).encode("utf-8")))
    doc.add_paragraph("E")
    doc.add_paragraph("rgreen")
    doc.add_paragraph("RE")
    p2 = doc.add_paragraph()
    p2._element.append(etree.fromstring(latex_to_omath(r"=ERE-Esell\times100\%(7)(4)", display=True).encode("utf-8")))
    doc.add_paragraph("E")
    doc.add_paragraph("rup")
    doc.add_paragraph("total")
    p3 = doc.add_paragraph()
    p3._element.append(etree.fromstring(latex_to_omath(r"=Esell\times100\%(8)(5)", display=True).encode("utf-8")))
    doc.add_paragraph("E")
    doc.add_paragraph("RE")
    doc.add_paragraph("吨氨成本模型")
    doc.add_paragraph("吨氨成本由购电成本、风光度电成本和运维成本三部分构成，扣除售电收入：")
    doc.add_paragraph("∑24")
    doc.add_paragraph("[Pbuy(t) · λbuy(t) − Psell(t) · λsell] · 1000 + CRE + COM")
    doc.add_paragraph("其中风光度电成本 CRE = ∑24")
    doc.add_paragraph("[Pw(t) · 0.15 + Ppv(t) · 0.12] · 1000（元），运维成本")
    doc.add_paragraph("COM = ∑24")
    doc.add_paragraph("u(t)·(10×0.1+10×0.15+0.75×0.002)·1000(元),日产氨量QNH3=1.5×24=36")
    doc.add_paragraph("24")
    doc.add_paragraph("u(t)=n=Qtarget/3")
    doc.add_paragraph("t=1")
    doc.add_paragraph("Cton =")
    doc.add_paragraph("1")
    doc.add_paragraph("Qtarget")
    doc.add_paragraph("24")
    doc.add_paragraph("t=1")
    doc.add_paragraph("∆c(t) · u(t) + Cfixed]")
    doc.add_paragraph("(13)")
    doc.add_paragraph("max")
    doc.add_paragraph("α")
    doc.add_paragraph("24")
    doc.add_paragraph("QNH3 = 3 α(t) (20)")
    doc.add_paragraph("t=1")
    doc.add_paragraph("CRE=∑_{t=1}^{24} [Pw(t) · 0.15 + Ppv(t) · 0.12] · 1000")
    doc.add_paragraph("COM=∑24 (10×0.1+10×0.15+0.75×0.002)·1000")
    doc.add_paragraph("S = ∑10")
    doc.add_paragraph("x_i")
    doc.add_paragraph("K = ∑10")
    doc.add_paragraph("42")
    doc.save(docx)

    content = extract_docx_content(str(docx), output_dir=str(work))
    paragraphs = [p for sec in content["sections"] for p in sec.get("paragraphs", [])]
    repaired = [p for p in paragraphs if isinstance(p, dict) and str(p.get("source") or "").startswith("repaired_")]
    repaired_text = "\n".join(p.get("text") or "" for p in repaired)
    residue = [p for p in paragraphs if p in ("rself", "rgreen", "rup", "RE") or (isinstance(p, str) and p.strip() == "E")]
    split_sum_problems = [p for p in paragraphs if isinstance(p, dict) and p.get("problem") == "split_sum_index_unknown"]

    assert_true(any(p.get("source") == "repaired_sum_bounds" for p in repaired), f"split sum bounds were not repaired: {paragraphs}")
    assert_true(sum(1 for p in repaired if p.get("source") == "repaired_ratio_cluster") == 3, f"ratio formulas were not repaired: {repaired}")
    assert_true(any(p.get("source") == "repaired_sum_prefix" for p in repaired), f"split leading sum prefix was not repaired: {paragraphs}")
    assert_true(sum(1 for p in repaired if p.get("source") == "repaired_labeled_sum_continuation") >= 2, f"labeled sum continuations were not repaired: {repaired}")
    assert_true(any(p.get("source") == "repaired_missing_sum_symbol" for p in repaired), f"missing sum symbol layout was not repaired: {repaired}")
    assert_true(any(p.get("source") == "repaired_fraction_sum_layout" for p in repaired), f"fraction plus sum layout was not repaired: {repaired}")
    assert_true(any(p.get("source") == "repaired_max_sum_layout" for p in repaired), f"max/min sum layout was not repaired: {repaired}")
    assert_true(any(p.get("source") == "repaired_inline_sum_missing_lower" for p in repaired), f"inline sum missing lower was not repaired from context: {repaired}")
    assert_true("rself=" in repaired_text and "rgreen=" in repaired_text and "rup=" in repaired_text, f"ratio variables missing: {repaired_text}")
    assert_true(all(r"\mathrm" in (p.get("latex") or "") for p in repaired if p.get("source") == "repaired_ratio_cluster"), "multi-letter repaired variables should use roman subscripts")
    latex_blob = "\n".join(p.get("latex") or "" for p in repaired)
    assert_true(r"\Deltat" not in latex_blob and r"\lambdabuy" not in latex_blob, f"greek-letter suffixes collapsed into invalid commands: {latex_blob}")
    assert_true(r"\Delta t" in latex_blob and r"\lambda_{\mathrm{buy}}" in latex_blob, f"greek-letter suffixes were not preserved: {latex_blob}")
    assert_true(r"\sum_{i=1}^{10}" in latex_blob, f"split sum index was not inferred from subscript variable: {latex_blob}")
    assert_true(r"\sum_{t=1}^{24}" in latex_blob and r"\frac{1}{Q_{\mathrm{target}}}" in latex_blob, f"time-indexed sum/fraction layout was not rendered: {latex_blob}")
    assert_true(r"\max_{\alpha}" in latex_blob, f"max sum layout did not preserve optimization variable: {latex_blob}")
    assert_true(not any((p.get("text") or "").startswith("K=") for p in repaired), f"index-less split sum was repaired by guessing: {repaired}")
    assert_true(split_sum_problems, "index-less split sum should be flagged instead of repaired by guessing")
    assert_true("(5)" not in latex_blob, f"stale source equation label leaked into repaired latex: {latex_blob}")
    assert_true(not residue, f"split ratio fragments leaked as body text: {residue}")
    result = run_generated_case("split_ratio_layout_render", content)
    assert_true(result["manifest"]["counts"]["display_formulas_rendered"] >= 7, "repaired split formulas did not render as display math")


@case
def content_parser_extracts_table_cell_images_and_flags_header_images() -> None:
    work = new_workdir("parser_table_header_images")
    img = work / "dot.png"
    img.write_bytes(PNG_1X1)
    docx = work / "table_header_images.docx"
    doc = Document()
    doc.sections[0].header.paragraphs[0].add_run().add_picture(str(img))
    doc.add_paragraph("1 Images")
    table = doc.add_table(rows=1, cols=2)
    table.cell(0, 0).paragraphs[0].add_run().add_picture(str(img))
    table.cell(0, 1).text = "image in table cell"
    doc.save(docx)

    content = extract_docx_content(str(docx), output_dir=str(work))
    paragraphs = [p for sec in content["sections"] for p in sec.get("paragraphs", [])]
    table_images = [p for p in paragraphs if isinstance(p, dict) and p.get("role") == "image" and p.get("location") == "table_cell"]
    assert_true(table_images, "table-cell image was not promoted into the content image stream")
    assert_true(content["_meta"]["images_extracted"] == 1, "header image should not be counted as a body image")
    assert_true(content["_meta"].get("non_body_images"), "header image was not recorded as a non-body image")


@case
def content_parser_keeps_references_before_english_appendix() -> None:
    work = new_workdir("parser_refs_appendix")
    docx = work / "refs_appendix.docx"
    doc = Document()
    doc.add_paragraph("1 Introduction")
    doc.add_paragraph("Body paragraph before references.")
    doc.add_paragraph("References")
    doc.add_paragraph("[1] Synthetic reference one.")
    doc.add_paragraph("[2] Synthetic reference two.")
    doc.add_paragraph("Appendix A Reproducible Commands")
    doc.add_paragraph("python run_pipeline.py --mode developer")
    doc.save(docx)

    content = extract_docx_content(str(docx), output_dir=str(work))
    assert_true(len(content.get("references") or []) == 2, f"reference entries were not preserved: {content.get('references')}")
    assert_true(all("python run_pipeline" not in str(ref) for ref in content.get("references") or []), "appendix command leaked into references")
    appendix = [sec for sec in content.get("sections") or [] if sec.get("role") == "appendix"]
    assert_true(appendix and any("python run_pipeline" in str(p) for p in appendix[0].get("paragraphs") or []), "English appendix section was not preserved")


@case
def content_parser_extracts_english_title_before_abstract() -> None:
    work = new_workdir("parser_english_title")
    docx = work / "english_title.docx"
    doc = Document()
    p = doc.add_paragraph("中文论文标题示例用于测试")
    p.runs[0].font.size = Pt(16)
    doc.add_paragraph("Research on Template Driven Document Quality Assessment")
    doc.add_paragraph("Abstract: This is the abstract body.")
    doc.add_paragraph("Key words: document automation")
    doc.add_paragraph("1 Introduction")
    doc.add_paragraph("Body text.")
    doc.save(docx)

    content = extract_docx_content(str(docx), output_dir=str(work))
    assert_true(content.get("title_info", {}).get("title_en") == "Research on Template Driven Document Quality Assessment", "English title before abstract was not extracted")
    assert_true(not any((sec.get("heading") or "").startswith("Research on") for sec in content.get("sections") or []), "English title leaked into body sections")


@case
def content_parser_extracts_docx_title_style_without_explicit_font_size() -> None:
    work = new_workdir("parser_title_style")
    docx = work / "title_style.docx"
    title = "\u9762\u5411\u4e2d\u6587\u6bd5\u4e1a\u8bba\u6587\u7684\u673a\u5668\u5b66\u4e60\u6392\u7248\u6d41\u6c34\u7ebf\u7814\u7a76"
    doc = Document()
    doc.add_heading(title, 0)
    doc.add_heading("\u6458\u8981", 1)
    doc.add_paragraph("\u672c\u6587\u6458\u8981\u7528\u4e8e\u9a8c\u8bc1\u9898\u540d\u6837\u5f0f\u63d0\u53d6\u3002")
    doc.add_heading("1 \u7eea\u8bba", 1)
    doc.add_paragraph("Body text.")
    doc.save(docx)

    content = extract_docx_content(str(docx), output_dir=str(work))
    headings = [sec.get("heading") for sec in content.get("sections") or []]
    assert_true(content.get("title_info", {}).get("title_cn") == title, f"title style was not extracted: {content.get('title_info')}")
    assert_true(title not in headings, f"title style leaked into body sections: {headings}")
    assert_true("\u6458\u8981" in headings, f"abstract heading was lost after title extraction: {headings}")


@case
def content_parser_extracts_plain_title_before_abstract() -> None:
    work = new_workdir("parser_plain_front_title")
    docx = work / "plain_front_title.docx"
    title = "\u9762\u5411\u7eff\u7535\u6d88\u7eb3\u7684\u591a\u6e90\u80fd\u6e90\u8c03\u5ea6\u4e0e\u7ecf\u6d4e\u6027\u8bc4\u4f30\u7814\u7a76"
    doc = Document()
    doc.add_paragraph(title)
    doc.add_paragraph("\u6458\u8981\uff1a\u672c\u6587\u6458\u8981\u7528\u4e8e\u9a8c\u8bc1\u65e0\u663e\u5f0f\u5b57\u53f7\u7684\u524d\u7f6e\u9898\u540d\u3002")
    doc.add_paragraph("\u5173\u952e\u8bcd\uff1a\u7eff\u7535\uff1b\u8c03\u5ea6")
    doc.add_heading("\u7b2c1\u7ae0 \u7eea\u8bba", 1)
    doc.add_paragraph("\u6b63\u6587\u6bb5\u843d\u3002")
    doc.save(docx)

    content = extract_docx_content(str(docx), output_dir=str(work))
    headings = [sec.get("heading") for sec in content.get("sections") or []]
    assert_true(content.get("title_info", {}).get("title_cn") == title, f"plain front title was not extracted: {content.get('title_info')}")
    assert_true(title not in headings, f"plain title leaked into body sections: {headings}")


@case
def content_parser_does_not_treat_chapter_heading_as_title() -> None:
    work = new_workdir("parser_chapter_not_title")
    docx = work / "chapter_not_title.docx"
    heading = "\u7b2c\u4e00\u7ae0 \u7eea\u8bba\u4e0e\u7814\u7a76\u80cc\u666f"
    doc = Document()
    doc.add_heading(heading, 1)
    doc.add_paragraph("\u8fd9\u662f\u6b63\u6587\u7b2c\u4e00\u6bb5\uff0c\u7528\u4e8e\u9a8c\u8bc1\u76f4\u63a5\u4ece\u7ae0\u8282\u5f00\u59cb\u7684\u6587\u6863\u3002")
    doc.save(docx)

    content = extract_docx_content(str(docx), output_dir=str(work))
    headings = [sec.get("heading") for sec in content.get("sections") or []]
    assert_true(content.get("title_info", {}).get("title_cn") != heading, f"chapter heading was misdetected as title: {content.get('title_info')}")
    assert_true(heading in headings, f"chapter heading was not preserved as a body section: {headings}")


@case
def content_parser_splits_chinese_enumerated_headings_after_keywords() -> None:
    work = new_workdir("parser_cn_enum")
    img = work / "dot.png"
    img.write_bytes(PNG_1X1)
    docx = work / "cn_enum.docx"
    doc = Document()
    doc.add_paragraph("摘要")
    doc.add_paragraph("本文用于测试。")
    doc.add_paragraph("关键词：")
    doc.add_paragraph("绿电直连；优化运行")
    doc.add_paragraph("一、问题重述")
    doc.add_paragraph("正文段落。")
    doc.add_paragraph().add_run().add_picture(str(img))
    doc.add_paragraph("图 1-1 示例图片")
    doc.save(docx)

    content = extract_docx_content(str(docx), output_dir=str(work))
    headings = [sec.get("heading") for sec in content.get("sections") or []]
    assert_true("一、问题重述" in headings, "Chinese enumerated heading was not split into a body section")
    body_sec = next(sec for sec in content.get("sections") or [] if sec.get("heading") == "一、问题重述")
    assert_true(body_sec.get("images"), "Image after Chinese enumerated heading was not assigned to body section")
    kw_sec = next((sec for sec in content.get("sections") or [] if sec.get("role") == "cn_keywords"), {})
    assert_true(not kw_sec.get("images"), "Image after body heading leaked into keyword front matter")



@case
def content_parser_skips_source_toc_and_records_placeholders() -> None:
    work = new_workdir("source_toc_placeholder")
    docx = work / "source_toc.docx"
    doc = Document()
    doc.add_paragraph("报名序号: [报名序号]")
    doc.add_paragraph("论文题目: Synthetic Energy Paper")
    doc.add_paragraph("目 录")
    doc.add_paragraph("一、 问题重述")
    doc.add_paragraph("172.04 MWh should not become a heading")
    doc.add_section(WD_SECTION.NEW_PAGE)
    doc.add_paragraph("一、 问题重述")
    doc.add_paragraph("This is the real body paragraph.")
    doc.save(docx)

    content = extract_docx_content(str(docx), output_dir=str(work))
    headings = [sec.get("heading") for sec in content.get("sections") or []]
    assert_true(content.get("title_info", {}).get("title_cn") == "Synthetic Energy Paper", "labeled title was not extracted")
    assert_true(content.get("_meta", {}).get("source_toc_skipped_paragraphs", 0) >= 3, "source TOC block was not skipped")
    assert_true(content.get("_meta", {}).get("source_placeholders"), "source placeholders were not recorded")
    assert_true("172.04 MWh should not become a heading" not in headings, "TOC/body sentence became a heading")
    assert_true("一、 问题重述" in headings, "real body heading was lost after TOC skip")


@case
def content_parser_skips_source_toc_with_roman_page_numbers() -> None:
    work = new_workdir("source_toc_roman_pages")
    docx = work / "source_toc_roman_pages.docx"
    doc = Document()
    doc.add_paragraph("\u8bba\u6587\u9898\u76ee: Roman Page TOC")
    doc.add_paragraph("\u76ee\u5f55")
    doc.add_paragraph("\u6458\u8981........................................ I")
    doc.add_paragraph("Abstract.................................... II")
    doc.add_paragraph("1 \u7eea\u8bba..................................... 1")
    doc.add_paragraph("2 \u7cfb\u7edf\u8bbe\u8ba1................................. 5")
    doc.add_paragraph("\u53c2\u8003\u6587\u732e................................... 32")
    doc.add_paragraph("1 \u7eea\u8bba")
    doc.add_paragraph("This real body paragraph must not be collected as a reference.")
    doc.save(docx)

    content = extract_docx_content(str(docx), output_dir=str(work))
    headings = [sec.get("heading") for sec in content.get("sections") or []]
    body_text = "\n".join(
        p if isinstance(p, str) else str(p.get("text") or "")
        for sec in content.get("sections") or []
        for p in sec.get("paragraphs", [])
    )
    assert_true(content.get("_meta", {}).get("source_toc_skipped_paragraphs", 0) >= 5, "roman-page source TOC entries were not skipped")
    assert_true(not content.get("references"), f"source TOC reference line activated reference collection: {content.get('references')}")
    assert_true("1 \u7eea\u8bba" in headings, f"real body heading after roman TOC was lost: {headings}")
    assert_true("real body paragraph" in body_text, f"real body after roman TOC was lost: {body_text}")


@case
def content_parser_keeps_numbered_long_body_sentence_as_body() -> None:
    work = new_workdir("numbered_long_body_sentence")
    docx = work / "numbered_long_body_sentence.docx"
    long_sentence = (
        "1 绪论的第1个分析段落用于模拟完整论文正文，"
        "研究对象既包含宏观制度因素，也包含微观行为数据，"
        "用于检验文本抽取、分页和样式迁移的稳定性。"
    )
    doc = Document()
    doc.add_paragraph("论文题目: Numbered Sentence")
    doc.add_paragraph("1 绪论")
    doc.add_paragraph(long_sentence)
    doc.save(docx)

    content = extract_docx_content(str(docx), output_dir=str(work))
    headings = [sec.get("heading") for sec in content.get("sections") or []]
    body_text = "\n".join(
        p if isinstance(p, str) else str(p.get("text") or "")
        for sec in content.get("sections") or []
        for p in sec.get("paragraphs", [])
    )
    assert_true(long_sentence not in headings, f"numbered body sentence became heading: {headings}")
    assert_true(long_sentence in body_text, f"numbered body sentence was not preserved: {body_text}")


@case
def content_parser_exits_source_toc_without_page_break() -> None:
    work = new_workdir("source_toc_no_break")
    docx = work / "source_toc_no_break.docx"
    doc = Document()
    doc.add_paragraph("论文题目: No Break TOC")
    doc.add_paragraph("目录")
    doc.add_paragraph("1 绪论 1")
    doc.add_paragraph("1 正文第一章")
    doc.add_paragraph("This paragraph must survive even though no page break follows the source TOC.")
    doc.save(docx)

    content = extract_docx_content(str(docx), output_dir=str(work))
    headings = [sec.get("heading") for sec in content.get("sections") or []]
    body_text = "\n".join(
        p if isinstance(p, str) else str(p.get("text") or "")
        for sec in content.get("sections") or []
        for p in sec.get("paragraphs", [])
    )
    assert_true(content.get("_meta", {}).get("source_toc_skipped_paragraphs", 0) >= 2, "source TOC entries were not skipped")
    assert_true("1 正文第一章" in headings, f"real heading after source TOC was lost: {headings}")
    assert_true("must survive" in body_text, f"body text after source TOC was lost: {body_text}")


@case
def content_parser_skips_unpaged_source_toc_without_boundary() -> None:
    work = new_workdir("source_toc_unpaged_no_boundary")
    docx = work / "source_toc_unpaged_no_boundary.docx"
    doc = Document()
    doc.add_paragraph("论文题目: Unpaged TOC")
    doc.add_paragraph("目录")
    doc.add_paragraph("一、 问题重述")
    doc.add_paragraph("二、 问题分析")
    doc.add_paragraph("三、 模型假设")
    doc.add_paragraph("一、 问题重述")
    doc.add_paragraph("This is the real body paragraph after the repeated first heading.")
    doc.save(docx)

    content = extract_docx_content(str(docx), output_dir=str(work))
    headings = [sec.get("heading") for sec in content.get("sections") or []]
    body_text = "\n".join(
        p if isinstance(p, str) else str(p.get("text") or "")
        for sec in content.get("sections") or []
        for p in sec.get("paragraphs", [])
    )
    assert_true(content.get("_meta", {}).get("source_toc_skipped_paragraphs", 0) >= 4, "unpaged source TOC block was not skipped")
    assert_true("二、 问题分析" not in headings and "三、 模型假设" not in headings, f"unpaged TOC entries leaked as sections: {headings}")
    assert_true("一、 问题重述" in headings, f"real repeated heading was lost: {headings}")
    assert_true("real body paragraph" in body_text, f"body text after repeated heading was lost: {body_text}")


@case
def content_parser_skips_source_toc_with_duplicate_formula_fragments_before_boundary() -> None:
    work = new_workdir("source_toc_duplicate_fragments")
    docx = work / "source_toc_duplicate_fragments.docx"
    doc = Document()
    doc.add_paragraph("论文题目: Duplicate Fragment TOC")
    doc.add_paragraph("目 录")
    doc.add_paragraph("一、 问题重述")
    doc.add_paragraph("二、 问题分析")
    doc.add_paragraph("41 .5")
    doc.add_paragraph("41 .5")
    doc.add_paragraph("九、 问题五")
    doc.add_paragraph("十、 灵敏度分析")
    doc.add_section(WD_SECTION.NEW_PAGE)
    doc.add_paragraph("一、 问题重述")
    doc.add_paragraph("This is the real body paragraph.")
    doc.save(docx)

    content = extract_docx_content(str(docx), output_dir=str(work))
    headings = [sec.get("heading") for sec in content.get("sections") or []]
    body_text = "\n".join(
        p if isinstance(p, str) else str(p.get("text") or "")
        for sec in content.get("sections") or []
        for p in sec.get("paragraphs", [])
    )
    assert_true(content.get("_meta", {}).get("source_toc_skipped_paragraphs", 0) >= 8, "source TOC with duplicate fragments was not skipped to boundary")
    assert_true("九、 问题五" not in headings and "十、 灵敏度分析" not in headings, f"late TOC entries leaked: {headings}")
    assert_true("一、 问题重述" in headings and "real body paragraph" in body_text, f"real body after source TOC was lost: {headings} / {body_text}")


@case
def content_parser_preserves_real_directory_section_content() -> None:
    work = new_workdir("real_directory_section")
    docx = work / "real_directory_section.docx"
    doc = Document()
    doc.add_paragraph("论文题目: Directory Section")
    directory_heading = doc.add_paragraph("目录")
    directory_heading.style = doc.styles["Heading 1"]
    doc.add_paragraph("本节介绍产品目录与数据字典，不是源文档自动目录。")
    doc.add_paragraph("1 正文")
    doc.add_paragraph("This paragraph must remain in the document.")
    doc.save(docx)

    content = extract_docx_content(str(docx), output_dir=str(work))
    body_text = "\n".join(
        p if isinstance(p, str) else str(p.get("text") or "")
        for sec in content.get("sections") or []
        for p in sec.get("paragraphs", [])
    )
    assert_true(not content.get("_meta", {}).get("source_toc_skipped_paragraphs"), "real directory section was treated as a source TOC")
    assert_true("产品目录与数据字典" in body_text, f"real directory section content was lost: {body_text}")
    assert_true("must remain" in body_text, f"body text after real directory section was lost: {body_text}")


@case
def content_parser_preserves_real_directory_section_before_page_break() -> None:
    work = new_workdir("real_directory_section_page_break")
    docx = work / "real_directory_section_page_break.docx"
    doc = Document()
    doc.add_paragraph("论文题目: Directory Section With Page Break")
    directory_heading = doc.add_paragraph("目录")
    directory_heading.style = doc.styles["Heading 1"]
    subheading = doc.add_paragraph("一、 产品目录")
    subheading.style = doc.styles["Heading 2"]
    doc.add_paragraph("本节介绍产品目录与数据字典，不是源文档自动目录。")
    doc.add_section(WD_SECTION.NEW_PAGE)
    doc.add_paragraph("1 正文")
    doc.add_paragraph("This paragraph must remain after the page break.")
    doc.save(docx)

    content = extract_docx_content(str(docx), output_dir=str(work))
    body_text = "\n".join(
        p if isinstance(p, str) else str(p.get("text") or "")
        for sec in content.get("sections") or []
        for p in sec.get("paragraphs", [])
    )
    headings = [sec.get("heading") for sec in content.get("sections") or []]
    assert_true(not content.get("_meta", {}).get("source_toc_skipped_paragraphs"), "real directory section before a page break was treated as source TOC")
    assert_true("目录" in headings and "一、 产品目录" in headings, f"real directory headings were lost: {headings}")
    assert_true("产品目录与数据字典" in body_text and "must remain" in body_text, f"real directory/page-break content was lost: {body_text}")


@case
def qa_reports_toc_pollution_placeholders_and_formula_fragments() -> None:
    work = new_workdir("qa_semantic_guards")
    docx = work / "out.docx"
    doc = Document()
    doc.add_paragraph("报名序号: [报名序号]")
    doc.add_paragraph("172.04 MWh should not be a heading")
    doc.save(docx)
    content = base_content(["E", "rgreen", "RE"])
    content["_meta"]["source_placeholders"] = [{"paragraph": 1, "text": "报名序号: [报名序号]"}]
    content["sections"][0]["heading"] = "172.04 MWh should not be a heading"
    content["sections"][0]["paragraphs"] = [
        {"role": "formula", "text": "x=1(1)(1)", "numbered": True},
        {"role": "formula", "text": "吨/日对应的开机时段数n=Q/qrate∈24,21,18,15,12.产量约束为=", "numbered": True},
        "E",
        "rgreen",
        "RE",
    ]
    write_json(work / "content.json", content)
    write_json(work / "format.json", base_format())
    write_json(work / "workflow_mode.json", {"mode": "developer"})
    (work / "build_generated.py").write_text("# synthetic\n", encoding="utf-8")
    report = check_output(str(work), mode="developer", output_docx_name="out.docx")
    codes = [item["code"] for item in report["issues"]]
    for code in ["CONTENT_TOC_POLLUTION", "UNFILLED_PLACEHOLDER_TEXT", "FORMULA_NUMBER_CONFLICT", "FORMULA_TEXT_FRAGMENTED", "PLACEHOLDER_TEXT_LEFT"]:
        assert_true(code in codes, f"QA did not report {code}")
    formula_issue = next(item for item in report["issues"] if item["code"] == "FORMULA_TEXT_FRAGMENTED")
    assert_true("contaminated formula text" in formula_issue.get("detail", ""), "QA did not explain narrative text inside a formula")
    assert_true(formula_issue["severity"] == "warning", "fragmented formula text should be downgraded to warning")
    assert_true(report["passed"] is False, "semantic guard issues should fail QA")


