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
from docx.shared import Pt
from lxml import etree

PIPELINE_DIR = Path(__file__).resolve().parent
if str(PIPELINE_DIR) not in sys.path:
    sys.path.insert(0, str(PIPELINE_DIR))

from content_parser import extract as extract_docx_content
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
    (img_src / "dot.png").write_bytes(PNG_1X1)
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
def qa_manifest_detects_missing_image_render() -> None:
    img_src = new_workdir("image_missing_src")
    (img_src / "dot.png").write_bytes(PNG_1X1)
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
def template_profile_sanitizes_private_source() -> None:
    fmt = base_format(source="private_school_template.docx")
    fmt["_meta"]["assets_dir"] = r"E:\private\Templates\private_school_template_assets"
    profile = profile_format(fmt, project_root=r"E:\private")
    text = json.dumps(profile, ensure_ascii=False)
    assert_true("private_school_template" not in text, "template profile leaked source filename")
    assert_true(profile["source"]["source_ext"] == ".docx", "template profile lost source extension")
    assert_true(profile["source"]["has_assets_dir"] is True, "template profile lost assets flag")


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
