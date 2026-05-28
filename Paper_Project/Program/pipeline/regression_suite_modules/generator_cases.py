"""Script generator and generated-runtime regression cases."""
from __future__ import annotations

import re

from qa_conformance import check_conformance
from regression_suite_modules.generated_docx import run_generated_case
from regression_suite_modules.harness import PNG_1X1, assert_true, base_content, base_format, case, new_workdir
from script_generator import (
    RUNTIME_TEMPLATE,
    _extract_page_and_header,
    _front_matter_sections,
    _infer_style_profiles,
    _infer_template_rules,
    _normalize_numbered_section_order,
)


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
    assert_true("docx_sections" in conf["counts"], "conformance report should name Word section count as docx_sections")
    assert_true("sections" not in conf["counts"], "ambiguous conformance count key 'sections' should not be emitted")


@case
def conformance_reference_labels_keep_template_cjk_font() -> None:
    content = base_content(["Body paragraph before references."])
    content["references"] = ["[1] 作者1. 中文参考文献与自动排版测试[J]. 综合研究评论, 2026."]
    result = run_generated_case("reference_cjk_font", content)
    conf = check_conformance(str(result["work"]), mode="developer", output_docx_name="out.docx")
    codes = [item["code"] for item in conf["issues"]]
    assert_true("STYLE_MISMATCH" not in codes, f"reference label triggered CJK font mismatch: {conf['issues']}")


@case
def script_generator_keeps_figure_reference_prose_as_body() -> None:
    fmt = base_format()
    fmt["style_profiles"] = {
        "body": {
            "font": "宋体",
            "size": 12.0,
            "align": "JUSTIFY",
            "line_spacing_fixed_pt": 28.0,
            "first_indent_cm": 0.74,
            "space_before_pt": 0.0,
            "space_after_pt": 0.0,
        },
        "figure_caption": {
            "font": "宋体",
            "size": 10.5,
            "align": "CENTER",
            "line_spacing_fixed_pt": 28.0,
            "first_indent_cm": 0.0,
            "space_before_pt": 6.0,
            "space_after_pt": 6.0,
        },
    }
    content = base_content(["图 1 展示了从数据到决策的机器学习研究流程。"])
    result = run_generated_case("figure_reference_prose_body", content, fmt=fmt)
    conf = check_conformance(str(result["work"]), mode="developer", output_docx_name="out.docx")
    codes = [item["code"] for item in conf["issues"]]
    assert_true("STYLE_MISMATCH" not in codes, f"figure reference prose was rendered as caption: {conf['issues']}")


@case
def script_generator_honors_body_page_break_before() -> None:
    content = base_content([])
    content["sections"] = [
        {
            "heading": "1 First",
            "level": 1,
            "role": "body",
            "paragraphs": ["First body paragraph."],
            "images": [],
            "page_break_before": True,
        },
        {
            "heading": "2 Second",
            "level": 1,
            "role": "body",
            "paragraphs": ["Second body paragraph."],
            "images": [],
            "page_break_before": True,
        },
    ]
    result = run_generated_case("body_page_break_before", content)
    page_breaks = len(re.findall(r'<w:br[^>]+w:type="page"', result["xml"]))
    assert_true(page_breaks >= 1, "body page_break_before was not rendered as a page break")


@case
def script_generator_renders_chinese_backmatter_headings() -> None:
    content = base_content([])
    content["sections"] = [
        {
            "heading": "1 Body",
            "level": 1,
            "role": "body",
            "paragraphs": ["Body paragraph."],
            "images": [],
        },
        {
            "heading": "\u81f4\u8c22",
            "level": 1,
            "role": "acknowledgement",
            "paragraphs": ["Thanks paragraph."],
            "images": [],
        },
        {
            "heading": "\u9644\u5f55",
            "level": 1,
            "role": "appendix",
            "paragraphs": ["Appendix paragraph."],
            "images": [],
        },
    ]
    result = run_generated_case("backmatter_unicode_headings", content)
    xml_compact = re.sub(r"\s+", "", result["xml"])
    assert_true("\u81f4\u8c22" in xml_compact, "acknowledgement heading was not rendered as Chinese text")
    assert_true("\u9644\u5f55" in xml_compact, "appendix heading was not rendered as Chinese text")
    codes = [item["code"] for item in result["report"]["issues"]]
    assert_true("CONTENT_HEADING_MISSING" not in codes, f"backmatter headings were still reported missing: {codes}")


@case
def script_generator_renders_backmatter_rich_items() -> None:
    work = new_workdir("backmatter_item_assets")
    img_dir = work / "figures"
    img_dir.mkdir()
    (img_dir / "dot.png").write_bytes(PNG_1X1)

    content = base_content([])
    content["_meta"]["images_dir"] = str(img_dir)
    content["_meta"]["images_extracted"] = 1
    content["_meta"]["tables_count"] = 1
    content["sections"] = [
        {
            "heading": "1 Body",
            "level": 1,
            "role": "body",
            "paragraphs": ["Body paragraph."],
            "images": [],
        },
        {
            "heading": "\u9644\u5f55",
            "level": 1,
            "role": "appendix",
            "paragraphs": [
                {"role": "table", "table_rows": [["A", "B"], ["1", "2"]]},
                {"role": "formula", "source": "latex", "latex": "E=mc^2", "text": "E=mc^2", "numbered": False},
                {"role": "image", "image": "dot.png"},
            ],
            "images": ["dot.png"],
        },
    ]
    result = run_generated_case("backmatter_rich_items", content)
    counts = result["manifest"]["counts"]
    assert_true(counts["content_tables_rendered"] == 1, f"appendix table was not rendered: {counts}")
    assert_true(counts["content_formulas_rendered"] == 1, f"appendix formula was not rendered: {counts}")
    assert_true(counts["content_images_rendered"] == 1, f"appendix image was not rendered: {counts}")


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
def script_generator_section_order_preserves_content_while_sorting_subsections() -> None:
    sections = [
        {"heading": "第1章 Intro", "level": 1, "paragraphs": ["chapter"], "images": []},
        {"heading": "1.2 Later", "level": 2, "paragraphs": ["later"], "images": []},
        {"heading": "1.2.2 Detail B", "level": 3, "paragraphs": ["b"], "images": []},
        {"heading": "1.2.1 Detail A", "level": 3, "paragraphs": ["a"], "images": []},
        {"heading": "1.1 Earlier", "level": 2, "paragraphs": ["earlier"], "images": []},
        {"heading": "第2章 Methods", "level": 1, "paragraphs": ["next"], "images": []},
    ]
    ordered = _normalize_numbered_section_order(sections)
    headings = [sec["heading"] for sec in ordered]
    assert_true(headings[:5] == ["第1章 Intro", "1.1 Earlier", "1.2 Later", "1.2.1 Detail A", "1.2.2 Detail B"], "numbered sections were not sorted safely")
    assert_true(ordered[1]["paragraphs"] == ["earlier"], "section content moved away from its heading")
    assert_true(headings[-1] == "第2章 Methods", "next numbered chapter was not preserved")


@case
def script_generator_template_rules_extract_page_header_and_flags() -> None:
    fmt = {
        "sections": [
            {
                "page_width_cm": 21.0,
                "page_height_cm": 29.7,
                "margin_top_cm": 2.5,
                "margin_bottom_cm": 2.4,
                "margin_left_cm": 2.8,
                "margin_right_cm": 2.2,
                "header": [
                    {
                        "text": "Thesis Header",
                        "alignment": "DEFAULT",
                        "runs": [{"text": "Thesis Header", "font": "Times New Roman", "size_pt": 9, "italic": True}],
                    }
                ],
            }
        ],
        "paragraphs": [
            {"text": "图1 系统结构"},
            {"text": "中文摘要不分自然段，英文题目要求大写字母，公式居中并编号，英文参考文献左对齐，参考文献悬挂缩进2字符。"},
        ],
    }
    page = _extract_page_and_header(fmt)
    rules = _infer_template_rules(fmt)
    assert_true(page["mt"] == 2.5 and page["mr"] == 2.2, "page margin extraction changed")
    assert_true(page["header"]["align"] == "CENTER", "DEFAULT header alignment was not normalized")
    assert_true(rules["cn_abstract_single_paragraph"] is True, "Chinese abstract rule was not detected")
    assert_true(rules["formula_center"] is True and rules["formula_numbered"] is True, "formula rules were not detected")
    assert_true(rules["reference_hanging_chars"] == 2.0, "reference hanging indent rule changed")


@case
def script_generator_style_profiles_apply_template_text_rules() -> None:
    body_text = "这是一个足够长的中文正文样本，用于模拟模板正文段落的自然排版特征，确保样式推断不会只依赖说明文字。"
    fmt = {
        "style_profiles": {},
        "paragraphs": [
            {
                "text": "论文正文使用宋体小四号，固定值28磅，首行缩进2字符，两端对齐。",
                "runs": [{"text": "论文正文使用宋体小四号，固定值28磅，首行缩进2字符，两端对齐。", "font": "宋体", "size_pt": 12}],
                "align": "LEFT",
            },
            {
                "text": "第1章 一级标题黑体三号居中加粗",
                "runs": [{"text": "第1章 一级标题黑体三号居中加粗", "font": "黑体", "size_pt": 16, "bold": True}],
                "align": "CENTER",
            },
            {
                "text": "图标题宋体五号居中",
                "runs": [{"text": "图标题宋体五号居中", "font": "宋体", "size_pt": 10.5}],
                "align": "CENTER",
            },
            {
                "text": "参考文献中文使用宋体小四号，悬挂缩进2字符。",
                "runs": [{"text": "参考文献中文使用宋体小四号，悬挂缩进2字符。", "font": "宋体", "size_pt": 12}],
                "align": "LEFT",
            },
            {
                "text": body_text,
                "runs": [{"text": body_text, "font": "宋体", "size_pt": 12}],
                "align": "JUSTIFY",
                "ls": 1.5,
                "indent": 0.74,
            },
        ],
    }
    profiles = _infer_style_profiles(fmt)
    assert_true(profiles["body"]["font"] == "宋体", "body font rule was not applied")
    assert_true(profiles["body"]["line_spacing_fixed_pt"] == 28.0, "fixed body line spacing rule was not applied")
    assert_true(profiles["body"]["align"] == "JUSTIFY", "body alignment rule was not normalized")
    assert_true(profiles["body"]["first_indent_cm"] > 0, "body first-line indent rule was not applied")
    assert_true(profiles["h1"]["font"] == "黑体" and profiles["h1"]["align"] == "CENTER", "h1 style inference changed")
    assert_true(profiles["figure_caption"]["size"] == 10.5 and profiles["figure_caption"]["align"] == "CENTER", "figure caption style rule changed")
    assert_true(profiles["reference"]["font"] == "宋体", "reference style rule changed")


@case
def script_generator_runtime_base_fragment_is_injected() -> None:
    assert_true("__BASE_RUNTIME__" not in RUNTIME_TEMPLATE, "base runtime placeholder leaked into generated template")
    assert_true(
        RUNTIME_TEMPLATE.lstrip().startswith("# -*- coding: utf-8 -*-"),
        "generated script header should come from base runtime",
    )
    assert_true("DATA = json.loads(__DATA_BLOB__)" in RUNTIME_TEMPLATE, "base runtime data bootstrap was not injected")
    assert_true("sys.dont_write_bytecode = True" in RUNTIME_TEMPLATE, "generated scripts should suppress __pycache__ clutter")
    assert_true(
        RUNTIME_TEMPLATE.index("sys.dont_write_bytecode = True") < RUNTIME_TEMPLATE.index("from latex_omath import latex_to_omath"),
        "bytecode suppression must run before local latex_omath imports",
    )
    assert_true("def apply_run_profile" in RUNTIME_TEMPLATE, "base run styling helper was not injected")
    assert_true("def add_text" in RUNTIME_TEMPLATE, "base text helper was not injected")
    assert_true("def setup_section" in RUNTIME_TEMPLATE, "base section setup helper was not injected")
    assert_true("def force_cover_headerless" in RUNTIME_TEMPLATE, "base cover cleanup helper was not injected")
    assert_true("def is_figure_caption_text" in RUNTIME_TEMPLATE, "caption prose classifier was not injected")
    assert_true("elif is_figure_caption_text(text):" in RUNTIME_TEMPLATE, "body renderer should use the safe caption classifier")
    assert_true(
        RUNTIME_TEMPLATE.index("def add_text") < RUNTIME_TEMPLATE.index("def render_cover_and_declarations"),
        "base text helpers should be defined before cover rendering",
    )
    assert_true(
        RUNTIME_TEMPLATE.index("def setup_section") < RUNTIME_TEMPLATE.index("def build_document"),
        "base page helpers should be defined before build orchestration",
    )


@case
def script_generator_runtime_content_helpers_fragment_is_injected() -> None:
    assert_true(
        "__CONTENT_HELPERS_RUNTIME__" not in RUNTIME_TEMPLATE,
        "content helper runtime placeholder leaked into generated template",
    )
    assert_true("def is_front_section_index" in RUNTIME_TEMPLATE, "front-section predicate was not injected")
    assert_true("def normalize_caption" in RUNTIME_TEMPLATE, "caption normalizer was not injected")
    assert_true("def clean_text_artifacts" in RUNTIME_TEMPLATE, "text cleanup helper was not injected")
    assert_true("def clean_formula_text" in RUNTIME_TEMPLATE, "formula cleanup helper was not injected")
    assert_true("def add_caption" in RUNTIME_TEMPLATE, "caption renderer helper was not injected")
    assert_true(RUNTIME_TEMPLATE.count("def clean_text_artifacts") == 1, "text cleanup helper should be injected exactly once")
    assert_true(
        RUNTIME_TEMPLATE.index("def clean_formula_text") < RUNTIME_TEMPLATE.index("def text_formula_to_latex"),
        "formula cleanup should be defined before plain-text formula conversion",
    )
    assert_true(
        RUNTIME_TEMPLATE.index("def add_caption") < RUNTIME_TEMPLATE.index("def render_table"),
        "caption helper should be defined before table/image rendering",
    )
    assert_true(
        RUNTIME_TEMPLATE.index("def is_front_section_index") < RUNTIME_TEMPLATE.index("def render_body"),
        "front-section predicate should be defined before body rendering",
    )


@case
def script_generator_runtime_formula_fragment_is_injected() -> None:
    assert_true("__FORMULA_RUNTIME__" not in RUNTIME_TEMPLATE, "formula runtime placeholder leaked into generated template")
    assert_true("def append_inline_formula" in RUNTIME_TEMPLATE, "inline formula runtime fragment was not injected")
    assert_true("def add_rich_text_runs" in RUNTIME_TEMPLATE, "rich_text runtime fragment was not injected")


@case
def script_generator_runtime_formula_text_fragment_is_injected() -> None:
    assert_true("__FORMULA_TEXT_RUNTIME__" not in RUNTIME_TEMPLATE, "formula text runtime placeholder leaked into generated template")
    assert_true("def chapter_number_from_heading" in RUNTIME_TEMPLATE, "chapter-number helper was not injected")
    assert_true("def text_formula_to_latex" in RUNTIME_TEMPLATE, "plain-text formula conversion helper was not injected")
    assert_true("def split_formula_number" in RUNTIME_TEMPLATE, "formula-number parsing helper was not injected")
    assert_true("def formula_has_number" in RUNTIME_TEMPLATE, "formula-number predicate was not injected")
    assert_true(RUNTIME_TEMPLATE.count("def text_formula_to_latex") == 1, "plain-text formula conversion helper should be injected exactly once")
    assert_true(
        RUNTIME_TEMPLATE.index("def clean_formula_text") < RUNTIME_TEMPLATE.index("def text_formula_to_latex"),
        "plain-text formula conversion should be defined after formula text cleanup",
    )
    assert_true(
        RUNTIME_TEMPLATE.index("def text_formula_to_latex") < RUNTIME_TEMPLATE.index("def render_formula"),
        "plain-text formula conversion should be defined before formula rendering",
    )
    assert_true(
        RUNTIME_TEMPLATE.index("def chapter_number_from_heading") < RUNTIME_TEMPLATE.index("def render_body"),
        "chapter-number helper should be defined before body rendering",
    )


@case
def script_generator_runtime_formula_render_fragment_is_injected() -> None:
    assert_true("__FORMULA_RENDER_RUNTIME__" not in RUNTIME_TEMPLATE, "formula render runtime placeholder leaked into generated template")
    assert_true("def next_formula_label" in RUNTIME_TEMPLATE, "formula label counter helper was not injected")
    assert_true("def render_plain_formula" in RUNTIME_TEMPLATE, "plain formula renderer was not injected")
    assert_true("def render_formula" in RUNTIME_TEMPLATE, "formula renderer was not injected")
    assert_true(RUNTIME_TEMPLATE.count("def render_formula") == 1, "formula renderer should be injected exactly once")
    assert_true(
        RUNTIME_TEMPLATE.index("def text_formula_to_latex") < RUNTIME_TEMPLATE.index("def render_formula"),
        "formula renderer should be defined after plain-text formula conversion",
    )
    assert_true(
        RUNTIME_TEMPLATE.index("def render_formula") < RUNTIME_TEMPLATE.index("def render_paragraph_item"),
        "formula renderer should be defined before body paragraph dispatch",
    )
    assert_true(
        RUNTIME_TEMPLATE.index("FORMULA_COUNTERS = {}") < RUNTIME_TEMPLATE.index("if __name__ == '__main__'"),
        "formula counters should be initialized before the generated-script entrypoint runs",
    )


@case
def script_generator_runtime_media_table_fragment_is_injected() -> None:
    assert_true("__MEDIA_TABLE_RUNTIME__" not in RUNTIME_TEMPLATE, "media/table runtime placeholder leaked into generated template")
    assert_true("def render_table" in RUNTIME_TEMPLATE, "table runtime fragment was not injected")
    assert_true("def render_image" in RUNTIME_TEMPLATE, "image runtime fragment was not injected")
    assert_true("def add_code_block" in RUNTIME_TEMPLATE, "code-block runtime fragment was not injected")


@case
def script_generator_runtime_references_fragment_is_injected() -> None:
    assert_true("__REFERENCES_RUNTIME__" not in RUNTIME_TEMPLATE, "references runtime placeholder leaked into generated template")
    assert_true("def render_reference_entries" in RUNTIME_TEMPLATE, "reference runtime fragment was not injected")
    assert_true("def render_backmatter_section" in RUNTIME_TEMPLATE, "backmatter runtime fragment was not injected")
    assert_true("def collect_structural_backmatter" in RUNTIME_TEMPLATE, "structural backmatter runtime fragment was not injected")


@case
def script_generator_runtime_toc_fragment_is_injected() -> None:
    assert_true("__TOC_RUNTIME__" not in RUNTIME_TEMPLATE, "TOC runtime placeholder leaked into generated template")
    assert_true("def collect_toc_entries" in RUNTIME_TEMPLATE, "TOC entry collection runtime fragment was not injected")
    assert_true("def add_toc" in RUNTIME_TEMPLATE, "TOC rendering runtime fragment was not injected")
    assert_true("def _infer_heading_pages_from_word_com" in RUNTIME_TEMPLATE, "Word COM TOC page-resolution fragment was not injected")


@case
def script_generator_runtime_build_fragment_is_injected() -> None:
    assert_true("__BUILD_RUNTIME__" not in RUNTIME_TEMPLATE, "build runtime placeholder leaked into generated template")
    assert_true("def reset_build_stats" in RUNTIME_TEMPLATE, "build stats runtime fragment was not injected")
    assert_true("def write_build_manifest" in RUNTIME_TEMPLATE, "manifest writer runtime fragment was not injected")
    assert_true("def build_document" in RUNTIME_TEMPLATE, "document build runtime fragment was not injected")
    assert_true("def main" in RUNTIME_TEMPLATE, "generated-script main runtime fragment was not injected")


@case
def script_generator_runtime_cover_fragment_is_injected() -> None:
    assert_true("__COVER_RUNTIME__" not in RUNTIME_TEMPLATE, "cover runtime placeholder leaked into generated template")
    assert_true("def render_cover_and_declarations" in RUNTIME_TEMPLATE, "cover renderer runtime fragment was not injected")
    assert_true("def render_cover_table" in RUNTIME_TEMPLATE, "cover table runtime fragment was not injected")
    assert_true("def compute_cover_skip_indices" in RUNTIME_TEMPLATE, "cover spacer runtime fragment was not injected")
    assert_true("def set_cover_cell_margins" in RUNTIME_TEMPLATE, "cover cell-margin helper should not collide with body table helper")
    assert_true(
        RUNTIME_TEMPLATE.index("def render_cover_and_declarations") < RUNTIME_TEMPLATE.index("def render_front_matter"),
        "cover runtime should be defined before front matter rendering",
    )
    assert_true(
        RUNTIME_TEMPLATE.index("def set_cell_borders") < RUNTIME_TEMPLATE.index("def add_code_block"),
        "shared border helper should be defined before media/table runtime uses it",
    )


@case
def script_generator_runtime_front_matter_fragment_is_injected() -> None:
    assert_true("__FRONT_MATTER_RUNTIME__" not in RUNTIME_TEMPLATE, "front matter runtime placeholder leaked into generated template")
    assert_true("def section_text" in RUNTIME_TEMPLATE, "front matter section_text helper was not injected")
    assert_true("def add_keywords" in RUNTIME_TEMPLATE, "front matter keyword helper was not injected")
    assert_true("def render_front_matter" in RUNTIME_TEMPLATE, "front matter renderer runtime fragment was not injected")
    assert_true(RUNTIME_TEMPLATE.count("def render_front_matter") == 1, "front matter renderer should be injected exactly once")
    assert_true(
        RUNTIME_TEMPLATE.index("def add_rich_text_item") < RUNTIME_TEMPLATE.index("def render_front_matter"),
        "front matter renderer should be defined after rich-text formula helpers",
    )
    assert_true(
        RUNTIME_TEMPLATE.index("def add_toc") < RUNTIME_TEMPLATE.index("def render_front_matter"),
        "front matter renderer should be defined after TOC helpers",
    )
    assert_true(
        RUNTIME_TEMPLATE.index("def render_front_matter") < RUNTIME_TEMPLATE.index("def is_front_section_index"),
        "front matter renderer should stay before body/front-index helpers",
    )


@case
def script_generator_runtime_body_fragment_is_injected() -> None:
    assert_true("__BODY_RUNTIME__" not in RUNTIME_TEMPLATE, "body runtime placeholder leaked into generated template")
    assert_true("def render_paragraph_item" in RUNTIME_TEMPLATE, "body paragraph dispatcher runtime fragment was not injected")
    assert_true("def render_body" in RUNTIME_TEMPLATE, "body renderer runtime fragment was not injected")
    assert_true(RUNTIME_TEMPLATE.count("def render_body") == 1, "body renderer should be injected exactly once")
    assert_true(
        RUNTIME_TEMPLATE.index("def render_formula") < RUNTIME_TEMPLATE.index("def render_paragraph_item"),
        "body renderer should be defined after formula rendering helpers",
    )
    assert_true(
        RUNTIME_TEMPLATE.index("def render_table") < RUNTIME_TEMPLATE.index("def render_paragraph_item"),
        "body renderer should be defined after table/image helpers",
    )
    assert_true(
        RUNTIME_TEMPLATE.index("def render_reference_entries") < RUNTIME_TEMPLATE.index("def render_body"),
        "body renderer should be defined after reference/backmatter helpers",
    )
    assert_true(
        RUNTIME_TEMPLATE.index("def render_body") < RUNTIME_TEMPLATE.index("def build_document"),
        "body renderer should be defined before document build orchestration",
    )

