"""
regression_suite.py - synthetic regression checks for the Word pipeline.

The suite creates temporary, non-private DOCX/MD fixtures and verifies the
engine behavior that is hard to cover by manual inspection alone:

- inline math stays in the current paragraph
- display math remains native OMML display math
- mixed text/image/math paragraphs do not drop tokens
- markdown tables/code/math keep structure
- build_manifest.json drives body element QA counts
- template profiles avoid private source filenames
- visual QA fails closed when required render tools are unavailable
"""
from __future__ import annotations

import argparse
import base64
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from typing import Any, Callable, Dict, List

from docx import Document
from docx.enum.section import WD_SECTION
from docx.shared import Pt
from lxml import etree

PIPELINE_DIR = Path(__file__).resolve().parent
if str(PIPELINE_DIR) not in sys.path:
    sys.path.insert(0, str(PIPELINE_DIR))

from content_parser import extract as extract_docx_content
from content_parser import _strip_trailing_formula_labels_from_xml
from formula_semantics import (
    CATEGORY_CONTAMINATED,
    CATEGORY_DISPLAY_MATH,
    CATEGORY_QUANTITY_TEXT,
    classify_formula_text,
    is_formula_problem_text,
    looks_like_formula_text as semantic_looks_like_formula_text,
    split_inline_math_spans,
)
from format_extractor import extract as extract_docx_format
from latex_omath import latex_to_omath
from md_parser import extract_content as extract_md_content
from qa_conformance import check_conformance
from qa_checker import check_output, write_reports
from script_generator import generate
from script_generator import _front_matter_sections
from template_profiler import profile_format
from privacy import sanitize_value


PNG_1X1 = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAFgwJ/lc0ndwAAAABJRU5ErkJggg=="
)

CASES: List[Callable[[], None]] = []
TEMP_DIRS: List[Path] = []
KEEP_ARTIFACTS = False


def case(fn: Callable[[], None]) -> Callable[[], None]:
    CASES.append(fn)
    return fn


def fail(message: str) -> None:
    raise AssertionError(message)


def assert_true(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def new_workdir(name: str) -> Path:
    work = Path(tempfile.mkdtemp(prefix=f"wordchat_{name}_"))
    TEMP_DIRS.append(work)
    return work


def write_json(path: Path, value: Dict[str, Any]) -> None:
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")


def write_sample_png(path: Path, width: int = 480, height: int = 270) -> None:
    from PIL import Image, ImageDraw

    image = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(image)
    draw.rectangle((8, 8, width - 8, height - 8), outline=(30, 80, 160), width=3)
    draw.line((24, height - 40, width - 24, 40), fill=(180, 50, 50), width=4)
    image.save(path)


def base_format(source: str = "synthetic.md") -> Dict[str, Any]:
    return {
        "_meta": {
            "source": source,
            "sha256": "synthetic",
            "paragraphs": 1,
            "tables": 0,
            "sections": 1,
        },
        "paragraphs": [
            {
                "text": "Synthetic body paragraph used for style inference.",
                "runs": [{"font": "Times New Roman", "size_pt": 12}],
                "align": "JUSTIFY",
            }
        ],
        "tables": [],
        "sections": [
            {
                "page_width_cm": 21.0,
                "page_height_cm": 29.7,
                "margin_top_cm": 2.54,
                "margin_bottom_cm": 2.54,
                "margin_left_cm": 3.17,
                "margin_right_cm": 3.17,
            }
        ],
        "cover": [],
        "style_profiles": {},
    }


def base_content(paragraphs: List[Any], meta_tables: int = 0) -> Dict[str, Any]:
    return {
        "_meta": {
            "source": "synthetic.docx",
            "sha256": "synthetic",
            "paragraphs": max(1, len(paragraphs)),
            "tables_count": meta_tables,
            "images_extracted": 0,
        },
        "title_info": {"title_cn": "Synthetic Thesis"},
        "sections": [
            {
                "heading": "1 Introduction",
                "level": 1,
                "role": "body",
                "paragraphs": paragraphs,
                "images": [],
            }
        ],
        "references": ["[1] Synthetic reference."],
    }


def run_generated_case(name: str, content: Dict[str, Any], fmt: Dict[str, Any] | None = None) -> Dict[str, Any]:
    work = new_workdir(name)
    fmt_path = work / "format.json"
    cnt_path = work / "content.json"
    out_docx = work / "out.docx"
    write_json(fmt_path, fmt or base_format())
    write_json(cnt_path, content)
    write_json(work / "workflow_mode.json", {"mode": "developer"})

    generate(str(fmt_path), str(cnt_path), str(work), "out.docx")
    build_py = work / "build_generated.py"
    subprocess.run([sys.executable, "-m", "py_compile", str(build_py)], check=True, timeout=120)
    result = subprocess.run(
        [sys.executable, str(build_py)],
        cwd=str(work),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=180,
    )
    if result.returncode != 0:
        fail(f"{name}: generated build failed: {result.stderr[:1000] or result.stdout[:1000]}")
    assert_true(out_docx.exists(), f"{name}: output docx was not created")

    with zipfile.ZipFile(out_docx) as zf:
        xml = zf.read("word/document.xml").decode("utf-8", errors="replace")
    manifest = json.loads((work / "build_manifest.json").read_text(encoding="utf-8"))
    report = check_output(str(work), mode="developer", output_docx_name="out.docx")
    return {"work": work, "xml": xml, "manifest": manifest, "report": report}


def omath_count(xml: str) -> int:
    return len(re.findall(r"<(?:[A-Za-z_][\w.-]*:)?oMath\b", xml))


def omath_para_count(xml: str) -> int:
    return len(re.findall(r"<(?:[A-Za-z_][\w.-]*:)?oMathPara\b", xml))


def make_vml_picture_docx(src_docx: Path, dst_docx: Path) -> None:
    tmp = Path(tempfile.mkdtemp(prefix="wordchat_vml_zip_"))
    try:
        with zipfile.ZipFile(src_docx) as zf:
            zf.extractall(tmp)
        document_xml = tmp / "word" / "document.xml"
        xml = document_xml.read_text(encoding="utf-8")
        m = re.search(r'r:embed="([^"]+)"', xml)
        if not m:
            fail("source docx did not contain a drawing relationship")
        rid = m.group(1)
        xml = re.sub(
            r"<w:drawing>.*?</w:drawing>",
            f'<w:pict><v:shape><v:imagedata r:id="{rid}"/></v:shape></w:pict>',
            xml,
            count=1,
            flags=re.S,
        )
        document_xml.write_text(xml, encoding="utf-8")
        with zipfile.ZipFile(dst_docx, "w", zipfile.ZIP_DEFLATED) as zf:
            for path in tmp.rglob("*"):
                if path.is_file():
                    zf.write(path, path.relative_to(tmp).as_posix())
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


@case
def inline_rich_text_stays_in_paragraph() -> None:
    content = base_content(
        [
            {
                "role": "rich_text",
                "text": "before x2 after",
                "runs": [
                    {"type": "text", "text": "before "},
                    {
                        "type": "math",
                        "text": "x2",
                        "math": [{"type": "inline", "latex": "x^2", "text": "x2"}],
                    },
                    {"type": "text", "text": " after"},
                ],
                "math": [{"type": "inline", "latex": "x^2", "text": "x2"}],
            }
        ]
    )
    result = run_generated_case("inline_rich", content)
    counts = result["manifest"]["counts"]
    xml = result["xml"]
    assert_true(counts["inline_formulas_rendered"] == 1, "inline formula was not counted")
    assert_true(counts["display_formulas_rendered"] == 0, "inline formula became display math")
    assert_true(omath_count(xml) == 1, "expected one native oMath")
    assert_true(omath_para_count(xml) == 0, "inline case should not create oMathPara")
    before_idx = xml.find("before ")
    math_idx = xml.find("<m:oMath", before_idx)
    after_idx = xml.find(" after", math_idx)
    assert_true(before_idx >= 0 and before_idx < math_idx < after_idx, "inline math order drifted")
    assert_true(not result["report"]["issues"], f"unexpected QA issues: {result['report']['issues']}")


@case
def legacy_text_math_item_is_inline() -> None:
    content = base_content(
        [
            {
                "text": "legacy inline math",
                "math": [{"type": "inline", "latex": "y^2", "text": "y2"}],
            }
        ]
    )
    result = run_generated_case("legacy_inline", content)
    counts = result["manifest"]["counts"]
    assert_true(counts["inline_formulas_rendered"] == 1, "legacy inline math was not rendered inline")
    assert_true(counts["display_formulas_rendered"] == 0, "legacy inline math became display math")
    assert_true(not result["report"]["issues"], f"unexpected QA issues: {result['report']['issues']}")


@case
def display_formula_remains_display() -> None:
    content = base_content([
        {"role": "formula", "latex": "a=b+c", "text": "a=b+c", "numbered": False}
    ])
    result = run_generated_case("display_formula", content)
    counts = result["manifest"]["counts"]
    assert_true(counts["display_formulas_rendered"] == 1, "display formula was not counted")
    assert_true(omath_para_count(result["xml"]) == 1, "display formula did not create oMathPara")
    assert_true(not result["report"]["issues"], f"unexpected QA issues: {result['report']['issues']}")


@case
def source_formula_label_cleanup_preserves_arguments() -> None:
    for expr in ["f(1)", "x^{(1)}", "P=f(1)"]:
        _xml, text, had_label = _strip_trailing_formula_labels_from_xml(latex_to_omath(expr, display=True))
        assert_true(not had_label, f"formula argument was mistaken for an equation label: {expr} -> {text}")
    _xml, text, had_label = _strip_trailing_formula_labels_from_xml(latex_to_omath("E=mc^2(1.1)", display=True))
    assert_true(had_label and "(1.1)" not in text, f"equation label was not stripped: {text}")


@case
def multiple_display_omml_entries_render_all() -> None:
    content = base_content([
        {
            "role": "formula",
            "source": "omml",
            "text": "",
            "math": [
                {"type": "display", "xml": latex_to_omath("x=1", display=True), "text": "x=1"},
                {"type": "display", "xml": latex_to_omath("y=2", display=True), "text": "y=2"},
            ],
            "numbered": False,
        }
    ])
    result = run_generated_case("multi_display_omml", content)
    counts = result["manifest"]["counts"]
    assert_true(counts["display_formulas_rendered"] == 2, f"multiple OMML formulas were not all rendered: {counts}")
    assert_true(omath_para_count(result["xml"]) == 2, "expected two display OMML paragraphs")
    assert_true(not result["report"]["issues"], f"unexpected QA issues: {result['report']['issues']}")


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
    doc.save(docx)

    content = extract_docx_content(str(docx), output_dir=str(work))
    paragraphs = [p for sec in content["sections"] for p in sec.get("paragraphs", [])]
    formulas = [p for p in paragraphs if isinstance(p, dict) and p.get("role") == "formula"]
    assert_true(len(formulas) == 3, f"LaTeX-delimited paragraphs were not all formulas: {paragraphs}")
    assert_true(all(f.get("source") == "latex" and f.get("latex") for f in formulas), "LaTeX delimiters were not stripped into latex fields")


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
def md_parser_keeps_table_code_and_rich_math() -> None:
    work = new_workdir("md_parser")
    md = work / "sample.md"
    md.write_text(
        "\n".join(
            [
                "# Sample",
                "",
                "Alpha $x^2$ beta.",
                "",
                "$$a=b+c$$",
                "",
                "| A | B |",
                "| --- | --- |",
                "| 1 | 2 |",
                "",
                "```text",
                "interface G0/0/1",
                "ip address 10.0.0.1 255.255.255.0",
                "```",
            ]
        ),
        encoding="utf-8",
    )
    content = extract_md_content(str(md), output_dir=str(work))
    paragraphs = [p for sec in content["sections"] for p in sec.get("paragraphs", [])]
    assert_true(any(isinstance(p, dict) and p.get("role") == "rich_text" and p.get("runs") for p in paragraphs), "MD inline math did not become rich_text runs")
    rich = next(p for p in paragraphs if isinstance(p, dict) and p.get("role") == "rich_text")
    assert_true(rich["runs"][0]["text"] == "Alpha " and rich["runs"][-1]["text"] == " beta.", "MD inline math surrounding spaces were not preserved")
    assert_true(any(isinstance(p, dict) and p.get("role") == "formula" for p in paragraphs), "MD display formula missing")
    assert_true(any(isinstance(p, dict) and p.get("table_rows") and p.get("role") == "table" for p in paragraphs), "MD table missing")
    assert_true(any(isinstance(p, dict) and p.get("role") == "code" for p in paragraphs), "MD code block missing")
    assert_true(content["_meta"]["tables_count"] == 1, "MD table count should be one")


@case
def md_title_only_document_keeps_body_section() -> None:
    work = new_workdir("md_title_only")
    md = work / "title_only.md"
    md.write_text("# Only Title\n\nBody text under the title without another heading.", encoding="utf-8")
    content = extract_md_content(str(md), output_dir=str(work))
    assert_true(content["sections"], "MD body after title-only H1 was dropped")
    assert_true(content["sections"][0]["paragraphs"], "MD title-only body paragraph missing")


@case
def md_image_keeps_paragraph_position() -> None:
    work = new_workdir("md_image_order")
    img = work / "dot.png"
    img.write_bytes(PNG_1X1)
    md = work / "image_order.md"
    md.write_text("# Image Order\n\nBefore image ![dot](dot.png) after image.", encoding="utf-8")
    content = extract_md_content(str(md), output_dir=str(work))
    paragraphs = [p for sec in content["sections"] for p in sec.get("paragraphs", [])]
    roles = [p.get("role") if isinstance(p, dict) else "text" for p in paragraphs]
    assert_true(roles == ["text", "image", "text"], f"MD image order drifted: {roles}")
    assert_true(content["_meta"]["images_extracted"] == 1, "MD image was not copied")


@case
def md_missing_images_are_reported_to_qa() -> None:
    work = new_workdir("md_missing_images")
    md = work / "missing_images.md"
    md.write_text(
        "# Missing Images\n\nLocal ![local](missing.png) and remote ![remote](https://example.com/a.png).",
        encoding="utf-8",
    )
    content = extract_md_content(str(md), output_dir=str(work))
    missing = content["_meta"].get("missing_images") or []
    paragraphs = [p for sec in content["sections"] for p in sec.get("paragraphs", [])]
    assert_true(len(missing) == 2, f"missing image references were not recorded: {missing}")
    assert_true(any(isinstance(p, dict) and p.get("role") == "missing_image" for p in paragraphs), "missing image marker not preserved in content stream")

    write_json(work / "content.json", content)
    write_json(work / "format.json", base_format())
    write_json(work / "workflow_mode.json", {"mode": "developer"})
    report = check_output(str(work), mode="developer", output_docx_name="out.docx")
    codes = [item["code"] for item in report["issues"]]
    assert_true("CONTENT_IMAGE_MISSING" in codes, "QA did not report missing Markdown images")
    assert_true(report["passed"] is False, "missing images should fail QA")


@case
def md_parser_keeps_references_before_english_appendix() -> None:
    work = new_workdir("md_refs_appendix")
    md = work / "refs_appendix.md"
    md.write_text(
        "\n".join([
            "# Demo",
            "",
            "Body text.",
            "",
            "## References",
            "[1] Synthetic Markdown reference one.",
            "[2] Synthetic Markdown reference two.",
            "",
            "## Appendix A Commands",
            "python run_pipeline.py --mode developer",
        ]),
        encoding="utf-8",
    )
    content = extract_md_content(str(md), output_dir=str(work))
    assert_true(len(content.get("references") or []) == 2, f"Markdown references were not preserved: {content.get('references')}")
    assert_true(all("python run_pipeline" not in str(ref) for ref in content.get("references") or []), "Markdown appendix command leaked into references")
    assert_true(any((sec.get("heading") or "").startswith("Appendix") for sec in content.get("sections") or []), "Markdown appendix section was not preserved")


@case
def qa_counts_mixed_inline_and_section_images() -> None:
    work = new_workdir("mixed_image_count")
    content = base_content([
        {"role": "image", "image": "inline.png"},
    ])
    content["sections"][0]["images"] = ["inline.png", "section_only.png"]
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
    assert_true(report["passed"] is False, "semantic guard issues should fail QA")


@case
def source_omml_is_made_wps_compatible() -> None:
    xml = latex_to_omath("x=1", display=True)
    root = etree.fromstring(xml.encode("utf-8"))
    M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
    for mr in root.iter(f"{{{M}}}r"):
        for rpr in list(mr.findall(f"{{{M}}}rPr")):
            mr.remove(rpr)
    raw_without_rpr = etree.tostring(root, encoding="unicode")
    content = base_content([
        {
            "role": "formula",
            "source": "omml",
            "text": "x=1",
            "math": [{"type": "display", "xml": raw_without_rpr, "text": "x=1"}],
            "numbered": True,
        }
    ])
    result = run_generated_case("source_omml_wps", content)
    conf = check_conformance(str(result["work"]), mode="developer", output_docx_name="out.docx")
    codes = [item["code"] for item in conf["issues"]]
    assert_true("OMML_WPS_COMPAT" not in codes, f"source OMML was not normalized for WPS: {conf['issues']}")


@case
def conformance_style_check_ignores_static_toc_lines() -> None:
    content = base_content(
        [
            "Body paragraph after heading.",
        ]
    )
    content["sections"] = [
        {
            "heading": "2 Custom Chapter",
            "level": 1,
            "role": "body",
            "paragraphs": ["Body paragraph after heading."],
            "images": [],
        }
    ]
    result = run_generated_case("conformance_toc", content)
    conf = check_conformance(str(result["work"]), mode="developer", output_docx_name="out.docx")
    codes = [item["code"] for item in conf["issues"]]
    assert_true("STYLE_MISMATCH" not in codes, f"conformance matched TOC lines instead of body headings: {conf['issues']}")


@case
def md_rich_math_builds_inline_omml() -> None:
    content = base_content(
        [
            {
                "role": "rich_text",
                "text": "Alpha x^2 beta.",
                "runs": [
                    {"type": "text", "text": "Alpha "},
                    {
                        "type": "math",
                        "text": "x^2",
                        "math": [{"type": "inline", "latex": "x^2", "text": "x^2"}],
                    },
                    {"type": "text", "text": " beta."},
                ],
                "math": [{"type": "inline", "latex": "x^2", "text": "x^2"}],
            }
        ]
    )
    result = run_generated_case("md_rich_build", content)
    assert_true(result["manifest"]["counts"]["inline_formulas_rendered"] == 1, "MD rich inline math not rendered inline")
    assert_true(omath_para_count(result["xml"]) == 0, "MD rich inline math created display formula")


@case
def latex_omath_display_flag_is_honored() -> None:
    inline_xml = latex_to_omath("x^2", display=False)
    display_xml = latex_to_omath("x^2", display=True)
    assert_true("oMathPara" not in inline_xml, "inline latex_to_omath wrapped in oMathPara")
    assert_true("oMathPara" in display_xml, "display latex_to_omath did not wrap in oMathPara")


@case
def latex_omath_limit_accepts_multitoken_subscript() -> None:
    xml = latex_to_omath(r"L=\lim_{n\to\infty}\frac{1}{n}\sum_{i=1}^{n}x_i", display=True)
    assert_true("[LaTeX error" not in xml, "limit with n\\to\\infty subscript produced a LaTeX error")
    assert_true("oMathPara" in xml and "lim" in xml, "limit formula did not render as display OMML")


@case
def latex_omath_invisible_delimiters_hide_separators() -> None:
    xml = latex_to_omath(r"\frac{E_{\mathrm{total}}-E_{\mathrm{sell}}-E_{\mathrm{buy}}}{E_{\mathrm{RE}}}+\sum_{t=1}^{24}x_t", display=True)
    root = etree.fromstring(xml.encode("utf-8"))
    ns = {"m": "http://schemas.openxmlformats.org/officeDocument/2006/math"}
    delimiters = root.findall(".//m:d", ns)
    assert_true(delimiters, "complex formula did not create grouped delimiter elements")
    for delim in delimiters:
        entries = delim.findall("./m:e", ns)
        if len(entries) <= 1:
            continue
        dpr = delim.find("./m:dPr", ns)
        sep = dpr.find("./m:sepChr", ns) if dpr is not None else None
        assert_true(sep is not None and sep.get(f"{{{ns['m']}}}val") == "", "invisible delimiter can render visible vertical separators")
    texts = "".join(t.text or "" for t in root.findall(".//m:t", ns))
    assert_true("t=1" in texts and "24" in texts, f"plain multi-character scripts were split incorrectly: {texts}")
    styled_xml = latex_to_omath(r"\frac{\mathrm{abc}\mathrm{def}}{x}+\frac{\mathrm{abc}+\mathrm{def}}{x}", display=True)
    styled_root = etree.fromstring(styled_xml.encode("utf-8"))
    styled_runs = []
    for run in styled_root.findall(".//m:r", ns):
        text = "".join(t.text or "" for t in run.findall("./m:t", ns))
        sty = run.find("./m:rPr/m:sty", ns)
        styled_runs.append((text, sty.get(f"{{{ns['m']}}}val") if sty is not None else None))
    assert_true(("abcdef", "p") in styled_runs, f"merged \\mathrm runs lost upright style: {styled_runs}")
    assert_true(("abc", "p") in styled_runs and ("def", "p") in styled_runs, f"mixed-style grouped runs lost \\mathrm style: {styled_runs}")


@case
def template_profile_sanitizes_private_source() -> None:
    fmt = base_format(source="private_school_template.docx")
    fmt["_meta"]["assets_dir"] = r"E:\private\Templates\private_school_template_assets"
    profile = profile_format(fmt, project_root=r"E:\private")
    text = json.dumps(profile, ensure_ascii=False)
    assert_true("private_school_template" not in text, "template profile leaked source filename")
    assert_true(profile["source"]["source_ext"] == ".docx", "template profile lost source extension")
    assert_true(profile["source"]["has_assets_dir"] is True, "template profile lost assets flag")


@case
def format_extractor_stops_cover_before_spaced_abstract_heading() -> None:
    work = new_workdir("format_cover_stop_abstract")
    docx = work / "cover_stop.docx"
    doc = Document()
    doc.add_paragraph("Cover title")
    doc.add_paragraph("摘  要")
    doc.add_paragraph("Template abstract sample paragraph should not be replayed as cover.")
    doc.save(docx)

    fmt, _ = extract_docx_format(str(docx))
    cover_text = "\n".join(
        "".join(run.get("t", "") for run in el.get("r", []))
        for el in fmt.get("cover") or []
        if isinstance(el, dict)
    )
    assert_true("Cover title" in cover_text, "cover text before abstract was not extracted")
    assert_true("摘" not in cover_text and "Template abstract sample" not in cover_text, "abstract page leaked into cover extraction")


@case
def numbered_english_heading_is_not_front_title() -> None:
    cnt = {
        "sections": [
            {
                "heading": "1 Introduction",
                "level": 1,
                "role": "body",
                "paragraphs": ["Body"],
                "images": [],
            }
        ]
    }
    front = _front_matter_sections(cnt)
    assert_true(front.get("en_title") in ("", None), "numbered English body heading became English title")
    assert_true(0 not in front.get("front_indices", set()), "numbered English body heading was marked as front matter")


@case
def english_appendix_heading_is_not_front_title() -> None:
    cnt = {
        "sections": [
            {"heading": "1 Introduction", "level": 1, "role": "body", "paragraphs": ["Body"], "images": []},
            {"heading": "Appendix A Commands", "level": 1, "role": "appendix", "paragraphs": ["python run_pipeline.py"], "images": []},
        ]
    }
    front = _front_matter_sections(cnt)
    assert_true(front.get("en_title") in ("", None), "English appendix heading became English title")
    assert_true(1 not in front.get("front_indices", set()), "English appendix was marked as front matter")


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


def run_cases(selected: str | None = None) -> int:
    passed = 0
    failed = 0
    matched = 0
    for fn in CASES:
        if selected and selected not in fn.__name__:
            continue
        matched += 1
        try:
            fn()
            print(f"PASS {fn.__name__}")
            passed += 1
        except Exception as exc:
            failed += 1
            print(f"FAIL {fn.__name__}: {exc}")
    if selected and matched == 0:
        print(f"RESULT passed=0 failed=1")
        print(f"FAIL no tests matched filter: {selected}")
        return 1
    print(f"RESULT passed={passed} failed={failed}")
    return 1 if failed else 0


def main() -> None:
    global KEEP_ARTIFACTS
    parser = argparse.ArgumentParser(description="Run synthetic regression checks for the Word pipeline.")
    parser.add_argument("--keep", action="store_true", help="Keep temporary test artifacts.")
    parser.add_argument("--filter", default=None, help="Run cases whose function name contains this text.")
    args = parser.parse_args()
    KEEP_ARTIFACTS = bool(args.keep)
    try:
        code = run_cases(args.filter)
    finally:
        if not KEEP_ARTIFACTS and "code" in locals() and code == 0:
            for path in TEMP_DIRS:
                shutil.rmtree(path, ignore_errors=True)
        elif TEMP_DIRS:
            print("ARTIFACTS")
            for path in TEMP_DIRS:
                print(path)
    raise SystemExit(code)


if __name__ == "__main__":
    main()
