"""Markdown parser regression cases."""
from __future__ import annotations

import base64
import re
from io import BytesIO

from PIL import Image
from md_parser import extract_content as extract_md_content
from md_parser_modules.format_extractor import extract_format as extract_md_format

from regression_suite_modules.generated_docx import run_generated_case
from regression_suite_modules.harness import (
    PNG_1X1,
    assert_true,
    base_format,
    case,
    new_workdir,
)


GIF_1X1 = (
    b"GIF89a\x01\x00\x01\x00\x80\x00\x00\x00\x00\x00\xff\xff\xff"
    b"!\xf9\x04\x01\x00\x00\x00\x00,\x00\x00\x00\x00\x01\x00\x01"
    b"\x00\x00\x02\x02D\x01\x00;"
)


def _jpeg_1x1() -> bytes:
    buffer = BytesIO()
    Image.new("RGB", (1, 1), "white").save(buffer, format="JPEG")
    return buffer.getvalue()


def _table_cell_media_items(paragraphs):
    items = []
    for paragraph in paragraphs:
        if not isinstance(paragraph, dict):
            continue
        for cell in paragraph.get("table_cell_items") or []:
            items.extend(cell.get("items") or [])
    return items


@case
def md_parser_reads_gb18030_chinese_markdown_without_escape_decode() -> None:
    work = new_workdir("md_gb18030_chinese")
    md = work / "gb18030_sample.md"
    source = "\n".join(
        [
            "# 中文编码测试",
            "",
            "# 格式说明",
            "正文：Times New Roman，小四号(12pt)，两端对齐，首行缩进2字符。",
            "---",
            "## 第一章 绪论",
            "正文包含中文字符：硕士论文排版，不应乱码。",
            "字面转义保留：\\u4e2d\\u6587 不能被二次解码。",
        ]
    )
    md.write_bytes(source.encode("gb18030"))

    content = extract_md_content(str(md), output_dir=str(work))
    assert_true(content["title_info"]["title_cn"] == "中文编码测试", f"GB18030 title was not decoded safely: {content['title_info']}")
    paragraphs = [p for sec in content["sections"] for p in sec.get("paragraphs", [])]
    joined = "\n".join(str(p.get("text") if isinstance(p, dict) else p) for p in paragraphs)
    assert_true("正文包含中文字符：硕士论文排版，不应乱码。" in joined, f"GB18030 Chinese body text was corrupted: {joined!r}")
    assert_true("\\u4e2d\\u6587" in joined, f"literal unicode escape text was incorrectly decoded: {joined!r}")
    assert_true("中文" not in joined.split("字面转义保留：", 1)[1].split("不能被二次解码", 1)[0],
                f"literal unicode escape became Chinese text: {joined!r}")

    fmt, raw = extract_md_format(str(md))
    assert_true("格式说明" in raw and "正文：Times New Roman" in raw, f"GB18030 raw markdown was not preserved: {raw!r}")
    fmt_text = "\n".join(p.get("text", "") for p in fmt.get("paragraphs", []))
    assert_true("首行缩进2字符" in fmt_text, f"GB18030 format instruction was not extracted: {fmt_text!r}")


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
def md_parser_classifies_front_and_back_matter_headings() -> None:
    work = new_workdir("md_heading_roles")
    md = work / "roles.md"
    md.write_text(
        "\n".join(
            [
                "# Demo",
                "",
                "## Abstract",
                "English abstract paragraph.",
                "",
                "## Key words",
                "template; parser",
                "",
                "## Appendix A",
                "Appendix paragraph.",
            ]
        ),
        encoding="utf-8",
    )
    content = extract_md_content(str(md), output_dir=str(work))
    roles = {sec.get("heading"): sec.get("role") for sec in content.get("sections") or []}
    assert_true(roles.get("Abstract") == "en_abstract", f"Markdown Abstract role was not classified: {roles}")
    assert_true(roles.get("Key words") == "en_keywords", f"Markdown keywords role was not classified: {roles}")
    assert_true(roles.get("Appendix A") == "appendix", f"Markdown appendix role was not classified: {roles}")


@case
def md_parser_skips_front_format_rules_until_delimiter() -> None:
    work = new_workdir("md_format_skip")
    md = work / "format_skip.md"
    md.write_text(
        "\n".join(
            [
                "---",
                "title: Demo",
                "---",
                "# ????",
                "???Times New Roman?????1.5????",
                "---",
                "# Paper Title",
                "",
                "## Abstract",
                "Real abstract paragraph.",
                "",
                "## 1 Body",
                "Real body paragraph.",
            ]
        ),
        encoding="utf-8",
    )
    content = extract_md_content(str(md), output_dir=str(work))
    assert_true(content.get("title_info", {}).get("title_en") == "Paper Title", f"format heading leaked into title: {content.get('title_info')}")
    paragraphs = [p for sec in content["sections"] for p in sec.get("paragraphs", [])]
    joined = "\n".join(str(p) for p in paragraphs)
    assert_true("Times New Roman" not in joined and "1.5" not in joined, f"format rule leaked into content paragraphs: {paragraphs}")
    assert_true("---" not in joined, f"format delimiter leaked into content paragraphs: {paragraphs}")


@case
def md_bom_yaml_format_formula_missing_image_boundary() -> None:
    work = new_workdir("md_bom_yaml_combo")
    figures = work / "figures"
    figures.mkdir()
    (figures / "ok.png").write_bytes(PNG_1X1)
    md = work / "combo.md"
    md.write_text(
        "\ufeff---\n"
        "title: Demo\n"
        "---\n"
        "# ????\n"
        "正文：Times New Roman，小四号，1.5倍行距。\n"
        "---\n"
        "Mixed Markdown Boundary Demo\n"
        "============================\n\n"
        "## Abstract\n"
        "This abstract keeps inline math $E=mc^2$ and reports a missing image ![missing](missing.png).\n\n"
        "## 1 Introduction\n"
        "Before formula.\n\n"
        "$$a=b+c$$\n\n"
        "Existing image ![ok](figures/ok.png).\n\n"
        "## References\n"
        "[1] Synthetic reference.",
        encoding="utf-8",
    )
    content = extract_md_content(str(md), output_dir=str(work))
    title_info = content.get("title_info") or {}
    paragraphs = [p for sec in content.get("sections") or [] for p in sec.get("paragraphs", [])]
    joined = "\n".join(str(p) for p in paragraphs)
    missing = content["_meta"].get("missing_images") or []

    assert_true(title_info.get("title_en") == "Mixed Markdown Boundary Demo", f"BOM YAML/front-format block polluted title detection: {title_info}")
    assert_true("Times New Roman" not in joined and "1.5" not in joined and "---" not in joined, f"front format block leaked into content: {paragraphs}")
    assert_true(any(isinstance(p, dict) and p.get("role") == "rich_text" for p in paragraphs), f"inline math was not preserved as rich_text: {paragraphs}")
    assert_true(any(isinstance(p, dict) and p.get("role") == "formula" for p in paragraphs), f"display math was not preserved as a formula: {paragraphs}")
    assert_true(content["_meta"].get("images_extracted") == 1, f"existing Markdown image was not copied: {content['_meta']}")
    assert_true(len(missing) == 1 and missing[0].get("reason") == "not_found", f"missing Markdown image was not recorded clearly: {missing}")

    result = run_generated_case("md_bom_yaml_combo_generated", content, base_format())
    report = result["report"]
    plan = report.get("repair_plan") or {}
    codes = [item["code"] for item in report["issues"]]
    assert_true("CONTENT_IMAGE_MISSING" in codes, f"QA did not report missing Markdown image in combo boundary: {codes}")
    assert_true(report["passed"] is False, "missing image combo boundary should fail structural QA")
    assert_true(plan.get("resume_scope") == "input_files", f"missing image combo should route users to input files: {plan}")
    assert_true("CONTENT_IMAGE_MISSING" in str(report.get("next_action") or ""), f"qa_report next_action did not name CONTENT_IMAGE_MISSING: {report.get('next_action')}")


@case
def md_utf8_bom_h1_title_populates_english_title_info() -> None:
    work = new_workdir("md_bom_h1_title")
    md = work / "bom_title.md"
    md.write_text(
        "\ufeff# Markdown Path Variant Demo\n\n"
        "## Abstract\n"
        "This paper checks beginner-friendly Markdown title extraction.\n\n"
        "## 1 Introduction\n"
        "Body text.\n\n"
        "## References\n"
        "[1] Synthetic reference.",
        encoding="utf-8",
    )
    content = extract_md_content(str(md), output_dir=str(work))
    title_info = content.get("title_info") or {}
    assert_true(title_info.get("title_en") == "Markdown Path Variant Demo", f"UTF-8 BOM H1 English title was not extracted as title_en: {title_info}")


@case
def md_setext_h1_title_populates_english_title_info() -> None:
    work = new_workdir("md_setext_h1_title")
    md = work / "setext_title.md"
    md.write_text(
        "\n".join(
            [
                "Markdown Setext Title Demo",
                "===========================",
                "",
                "## Abstract",
                "This abstract should remain content, not title text.",
                "",
                "## 1 Introduction",
                "Body text.",
                "",
                "## References",
                "[1] Synthetic reference.",
            ]
        ),
        encoding="utf-8",
    )
    content = extract_md_content(str(md), output_dir=str(work))
    title_info = content.get("title_info") or {}
    paragraphs = [p for sec in content.get("sections") or [] for p in sec.get("paragraphs", [])]
    joined = "\n".join(str(p) for p in paragraphs)
    assert_true(title_info.get("title_en") == "Markdown Setext Title Demo", f"Setext H1 title was not extracted as title_en: {title_info}")
    assert_true("===" not in joined, f"Setext underline leaked into content paragraphs: {paragraphs}")


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
def md_image_resolves_wrapped_and_percent_encoded_paths() -> None:
    work = new_workdir("md_image_path_variants")
    figures = work / "figures"
    figures.mkdir()
    (figures / "my figure.png").write_bytes(PNG_1X1)
    (figures / "angle figure.png").write_bytes(PNG_1X1)
    md = work / "image_paths.md"
    md.write_text(
        "# Image Paths\n\nEncoded ![one](figures/my%20figure.png)\n\nWrapped ![two](<figures/angle figure.png>).",
        encoding="utf-8",
    )
    content = extract_md_content(str(md), output_dir=str(work))
    paragraphs = [p for sec in content["sections"] for p in sec.get("paragraphs", [])]
    images = [p for p in paragraphs if isinstance(p, dict) and p.get("role") == "image"]
    missing = content["_meta"].get("missing_images") or []
    assert_true(content["_meta"]["images_extracted"] == 2, f"common Markdown image path variants were not copied: {content['_meta']}")
    assert_true(len(images) == 2, f"image path variants were not preserved in content stream: {paragraphs}")
    assert_true(not missing, f"existing wrapped/percent-encoded local images were reported missing: {missing}")


@case
def md_image_resolves_local_uri_suffixes() -> None:
    work = new_workdir("md_image_uri_suffixes")
    figures = work / "figures"
    figures.mkdir()
    (figures / "panel one.png").write_bytes(PNG_1X1)
    (figures / "panel two.png").write_bytes(PNG_1X1)
    md = work / "image_uri_suffixes.md"
    md.write_text(
        "# Image URI Suffixes\n\n"
        "Query ![one](figures/panel%20one.png?raw=true#panel-a)\n\n"
        "Fragment ![two](figures\\panel%20two.png#caption).",
        encoding="utf-8",
    )
    content = extract_md_content(str(md), output_dir=str(work))
    paragraphs = [p for sec in content["sections"] for p in sec.get("paragraphs", [])]
    images = [p for p in paragraphs if isinstance(p, dict) and p.get("role") == "image"]
    missing = content["_meta"].get("missing_images") or []
    assert_true(content["_meta"]["images_extracted"] == 2, f"local Markdown image URI suffixes were not copied: {content['_meta']}")
    assert_true(len(images) == 2, f"local Markdown image URI suffixes were not preserved in content stream: {paragraphs}")
    assert_true(not missing, f"existing local images with URI suffixes were reported missing: {missing}")


@case
def md_image_resolves_parenthesized_inline_paths() -> None:
    work = new_workdir("md_image_parenthesized_paths")
    figures = work / "figures"
    figures.mkdir()
    (figures / "plot (1).png").write_bytes(PNG_1X1)
    md = work / "image_parentheses.md"
    md.write_text(
        "# Image Parentheses\n\n"
        "Existing image ![plot](figures/plot%20(1).png?raw=true#panel-a) should not be truncated.",
        encoding="utf-8",
    )
    content = extract_md_content(str(md), output_dir=str(work))
    paragraphs = [p for sec in content["sections"] for p in sec.get("paragraphs", [])]
    images = [p for p in paragraphs if isinstance(p, dict) and p.get("role") == "image"]
    missing = content["_meta"].get("missing_images") or []
    assert_true(content["_meta"]["images_extracted"] == 1, f"parenthesized inline Markdown image was not copied: {content['_meta']}")
    assert_true(len(images) == 1, f"parenthesized inline Markdown image was not preserved in content stream: {paragraphs}")
    assert_true(not missing, f"existing parenthesized image path was reported missing: {missing}")
    assert_true(
        all(".png?raw=true#panel-a)" not in str(p) for p in paragraphs),
        f"trailing parenthesized image path leaked into body text: {paragraphs}",
    )


@case
def md_image_resolves_inline_paths_with_titles() -> None:
    work = new_workdir("md_image_inline_titles")
    figures = work / "figures"
    figures.mkdir()
    (figures / "inline panel.png").write_bytes(PNG_1X1)
    (figures / "angle panel.png").write_bytes(PNG_1X1)
    md = work / "image_inline_titles.md"
    md.write_text(
        "# Image Inline Titles\n\n"
        "Bare image ![one](figures/inline%20panel.png \"Figure caption title\") should resolve.\n\n"
        "Wrapped image ![two](<figures/angle panel.png> 'Wrapped caption title') should resolve.",
        encoding="utf-8",
    )
    content = extract_md_content(str(md), output_dir=str(work))
    paragraphs = [p for sec in content["sections"] for p in sec.get("paragraphs", [])]
    images = [p for p in paragraphs if isinstance(p, dict) and p.get("role") == "image"]
    missing = content["_meta"].get("missing_images") or []
    joined = "\n".join(str(p) for p in paragraphs)
    assert_true(content["_meta"]["images_extracted"] == 2, f"inline Markdown image titles were not copied: {content['_meta']}")
    assert_true(len(images) == 2, f"inline Markdown image titles were not preserved in content stream: {paragraphs}")
    assert_true(not missing, f"existing inline-title images were reported missing: {missing}")
    assert_true(
        "Figure caption title" not in joined and "Wrapped caption title" not in joined,
        f"Markdown image titles leaked into body text: {paragraphs}",
    )


@case
def md_html_images_are_extracted_or_reported() -> None:
    work = new_workdir("md_html_images")
    figures = work / "figures"
    figures.mkdir()
    (figures / "html panel.png").write_bytes(PNG_1X1)
    md = work / "html_images.md"
    md.write_text(
        "# HTML Images\n\n"
        "Before <img alt=\"panel\" src=\"figures/html%20panel.png?raw=true#panel-a\"> after.\n\n"
        "Missing <img src='figures/missing.png' alt='missing panel'> should route to QA.\n\n"
        "Remote <img src=\"https://example.com/remote.png\" alt=\"remote panel\"> should not be downloaded.",
        encoding="utf-8",
    )
    content = extract_md_content(str(md), output_dir=str(work))
    paragraphs = [p for sec in content["sections"] for p in sec.get("paragraphs", [])]
    images = [p for p in paragraphs if isinstance(p, dict) and p.get("role") == "image"]
    missing = content["_meta"].get("missing_images") or []
    joined = "\n".join(str(p) for p in paragraphs)
    reasons = sorted(item.get("reason") for item in missing)
    assert_true(content["_meta"]["images_extracted"] == 1, f"HTML Markdown image was not copied: {content['_meta']}")
    assert_true(len(images) == 1, f"HTML image was not preserved in content stream: {paragraphs}")
    assert_true(reasons == ["not_found", "remote"], f"HTML missing/remote image reasons were not preserved: {missing}")
    assert_true("<img" not in joined, f"raw HTML image tag leaked into body content: {paragraphs}")

    result = run_generated_case("md_html_images_generated", content, base_format())
    codes = [item["code"] for item in result["report"]["issues"]]
    assert_true("CONTENT_IMAGE_MISSING" in codes, f"QA did not report missing HTML image: {codes}")
    assert_true("CONTENT_IMAGE_REMOTE_UNSUPPORTED" in codes, f"QA did not report remote HTML image: {codes}")
    assert_true(result["report"]["passed"] is False, "HTML missing/remote images should fail structural QA")


@case
def md_html_table_cell_images_render_inside_generated_table() -> None:
    work = new_workdir("md_html_table_cell_images")
    figures = work / "figures"
    figures.mkdir()
    (figures / "html table panel.png").write_bytes(PNG_1X1)
    md = work / "html_table_cell_images.md"
    md.write_text(
        "# HTML Table Cell Images\n\n"
        "| Item | Evidence |\n"
        "| --- | --- |\n"
        "| Existing | <img src=\"figures/html%20table%20panel.png\" alt=\"panel\"> |\n",
        encoding="utf-8",
    )
    content = extract_md_content(str(md), output_dir=str(work))
    paragraphs = [p for sec in content["sections"] for p in sec.get("paragraphs", [])]
    images = [p for p in _table_cell_media_items(paragraphs) if isinstance(p, dict) and p.get("role") == "image"]
    assert_true(len(images) == 1, f"HTML table-cell image was not attached to table cell content: {paragraphs}")

    result = run_generated_case("md_html_table_cell_images_generated", content, base_format())
    table_xmls = re.findall(r"<w:tbl\b.*?</w:tbl>", result["xml"], flags=re.S)
    table_drawings = sum(xml.count("<w:drawing>") for xml in table_xmls)
    total_drawings = result["xml"].count("<w:drawing>")
    assert_true(total_drawings == 1, f"expected one rendered HTML table-cell image, saw {total_drawings}")
    assert_true(table_drawings == 1, "HTML table-cell image rendered outside the generated Word table")
    assert_true("img src" not in result["xml"], "raw HTML image tag leaked into generated Word table text")
    assert_true(result["report"]["passed"] is True, f"HTML table-cell image render should pass QA: {result['report']}")


@case
def md_html_lazy_srcset_and_data_uri_images_are_extracted() -> None:
    work = new_workdir("md_html_lazy_srcset_data_uri_images")
    figures = work / "figures"
    figures.mkdir()
    (figures / "lazy panel.png").write_bytes(PNG_1X1)
    (figures / "srcset panel.png").write_bytes(PNG_1X1)
    data_uri = "data:image/png;base64," + base64.b64encode(PNG_1X1).decode("ascii")
    md = work / "html_advanced_images.md"
    md.write_text(
        "# HTML Advanced Images\n\n"
        "Lazy <img alt=\"lazy\" data-src=\"figures/lazy%20panel.png\"> image.\n\n"
        "Srcset <img alt=\"srcset\" srcset=\"figures/srcset%20panel.png 1x, figures/missing@2x.png 2x\"> image.\n\n"
        f"Inline <img alt=\"inline\" src=\"{data_uri}\"> image.",
        encoding="utf-8",
    )
    content = extract_md_content(str(md), output_dir=str(work))
    paragraphs = [p for sec in content["sections"] for p in sec.get("paragraphs", [])]
    images = [p for p in paragraphs if isinstance(p, dict) and p.get("role") == "image"]
    joined = "\n".join(str(p) for p in paragraphs)
    assert_true(content["_meta"]["images_extracted"] == 3, f"advanced HTML images were not copied: {content['_meta']}")
    assert_true(len(images) == 3, f"advanced HTML images were not preserved in content stream: {paragraphs}")
    assert_true(not content["_meta"].get("missing_images"), f"advanced HTML images should not be missing: {content['_meta'].get('missing_images')}")
    assert_true("<img" not in joined and "data:image" not in joined, f"raw advanced HTML image tag leaked into body content: {paragraphs}")

    result = run_generated_case("md_html_advanced_images_generated", content, base_format())
    assert_true(result["xml"].count("<w:drawing>") == 3, "advanced HTML images did not render as Word drawings")
    assert_true("data:image" not in result["xml"] and "img src" not in result["xml"], "raw advanced HTML image data leaked into DOCX XML")
    assert_true(result["report"]["passed"] is True, f"advanced HTML images should pass QA: {result['report']}")


@case
def md_html_bad_data_uri_images_are_reported_as_unreadable() -> None:
    work = new_workdir("md_html_bad_data_uri_images")
    md = work / "html_bad_data_uri.md"
    md.write_text(
        "# HTML Bad Data URI Images\n\n"
        "Broken <img alt=\"broken\" src=\"data:image/png;base64,not-valid-base64\"> image.",
        encoding="utf-8",
    )
    content = extract_md_content(str(md), output_dir=str(work))
    paragraphs = [p for sec in content["sections"] for p in sec.get("paragraphs", [])]
    missing = content["_meta"].get("missing_images") or []
    joined = "\n".join(str(p) for p in paragraphs)
    assert_true(content["_meta"]["images_extracted"] == 0, f"broken data URI should not be copied: {content['_meta']}")
    assert_true(len(missing) == 1 and missing[0].get("reason") == "unreadable", f"broken data URI was not reported as unreadable: {missing}")
    assert_true("not-valid-base64" not in str(missing), f"raw base64 payload leaked into missing-image metadata: {missing}")
    assert_true("<img" not in joined and "data:image" not in joined, f"raw broken HTML image tag leaked into body content: {paragraphs}")

    result = run_generated_case("md_html_bad_data_uri_images_generated", content, base_format())
    codes = [item["code"] for item in result["report"]["issues"]]
    assert_true("CONTENT_IMAGE_UNREADABLE" in codes, f"QA did not report broken data URI as unreadable: {codes}")
    assert_true("CONTENT_IMAGE_REMOTE_UNSUPPORTED" not in codes, f"data URI should not be reported as a remote image: {codes}")
    assert_true(result["report"]["passed"] is False, "broken data URI images should fail structural QA")


@case
def md_html_data_uri_mime_mismatch_is_reported_as_unreadable() -> None:
    work = new_workdir("md_html_data_uri_mime_mismatch")
    mismatched_uri = "data:image/png;base64," + base64.b64encode(_jpeg_1x1()).decode("ascii")
    md = work / "html_data_uri_mismatch.md"
    md.write_text(
        "# HTML Data URI MIME Mismatch\n\n"
        "## Abstract\n"
        "This abstract keeps the document in normal paper shape.\n\n"
        "## 1 Introduction\n"
        f"Mismatched <img alt=\"mismatch\" src=\"{mismatched_uri}\"> image should ask users to export PNG/JPG again.\n\n"
        "## References\n"
        "[1] Synthetic reference.",
        encoding="utf-8",
    )
    content = extract_md_content(str(md), output_dir=str(work))
    paragraphs = [p for sec in content["sections"] for p in sec.get("paragraphs", [])]
    missing = content["_meta"].get("missing_images") or []
    joined = "\n".join(str(p) for p in paragraphs)
    assert_true(content["_meta"]["images_extracted"] == 0, f"mismatched data URI should not be copied: {content['_meta']}")
    assert_true(
        len(missing) == 1 and missing[0].get("reason") == "unreadable" and "JPEG" in str(missing[0].get("detail") or ""),
        f"mismatched data URI was not recorded as unreadable with detected format detail: {missing}",
    )
    assert_true("data:image" not in str(missing), f"raw data URI payload leaked into missing-image metadata: {missing}")
    assert_true("<img" not in joined and "data:image" not in joined, f"raw mismatched HTML image tag leaked into body content: {paragraphs}")

    result = run_generated_case("md_html_data_uri_mime_mismatch_generated", content, base_format())
    report = result["report"]
    codes = [item["code"] for item in report["issues"]]
    assert_true("CONTENT_IMAGE_UNREADABLE" in codes, f"QA did not report mismatched data URI as unreadable: {codes}")
    assert_true(report["passed"] is False, "mismatched data URI images should fail structural QA")
    action = str((report.get("repair_plan") or {}).get("next_action") or report.get("next_action") or "")
    assert_true("CONTENT_IMAGE_UNREADABLE" in action and "PNG" in action, f"mismatched data URI next action should tell users to export PNG/JPG: {action}")


@case
def md_table_cell_images_are_extracted_or_reported() -> None:
    work = new_workdir("md_table_cell_images")
    figures = work / "figures"
    figures.mkdir()
    (figures / "table panel.png").write_bytes(PNG_1X1)
    md = work / "table_cell_images.md"
    md.write_text(
        "# Table Cell Images\n\n"
        "| Item | Evidence |\n"
        "| --- | --- |\n"
        "| Existing | ![panel](figures/table%20panel.png) |\n"
        "| Missing | ![missing](figures/missing.png) |\n",
        encoding="utf-8",
    )
    content = extract_md_content(str(md), output_dir=str(work))
    paragraphs = [p for sec in content["sections"] for p in sec.get("paragraphs", [])]
    images = [p for p in _table_cell_media_items(paragraphs) if isinstance(p, dict) and p.get("role") == "image"]
    missing_markers = [p for p in paragraphs if isinstance(p, dict) and p.get("role") == "missing_image"]
    nested_missing = [p for p in _table_cell_media_items(paragraphs) if isinstance(p, dict) and p.get("role") == "missing_image"]
    missing = content["_meta"].get("missing_images") or []
    assert_true(any(isinstance(p, dict) and p.get("role") == "table" for p in paragraphs), f"Markdown table was not preserved: {paragraphs}")
    assert_true(content["_meta"]["images_extracted"] == 1, f"Markdown table-cell image was silently dropped: {content['_meta']}")
    assert_true(len(images) == 1, f"Markdown table-cell image was not attached to table cell content: {paragraphs}")
    assert_true(images[0].get("location") == "markdown_table_cell", f"table-cell image should keep its origin: {images}")
    assert_true(
        len(missing) == 1 and missing[0].get("reason") == "not_found" and missing[0].get("location") == "markdown_table_cell",
        f"missing Markdown table-cell image was not recorded with table-cell origin: {missing}",
    )
    assert_true(
        nested_missing and nested_missing[0].get("location") == "markdown_table_cell",
        f"missing Markdown table-cell image was not attached to table cell content: {paragraphs}",
    )
    assert_true(
        missing_markers and missing_markers[0].get("location") == "markdown_table_cell",
        f"missing Markdown table-cell image marker was not preserved: {paragraphs}",
    )

    result = run_generated_case("md_table_cell_images_generated", content, base_format())
    report = result["report"]
    codes = [item["code"] for item in report["issues"]]
    assert_true(result["manifest"]["counts"]["content_images_rendered"] == 1, f"existing table-cell image was not rendered once: {result['manifest']}")
    assert_true("CONTENT_IMAGE_MISSING" in codes, f"QA did not report missing Markdown table-cell image: {codes}")
    assert_true(report["passed"] is False, "missing Markdown table-cell image should fail structural QA")


@case
def md_table_cell_images_render_inside_generated_table() -> None:
    work = new_workdir("md_table_cell_image_render")
    figures = work / "figures"
    figures.mkdir()
    (figures / "table panel.png").write_bytes(PNG_1X1)
    md = work / "table_cell_image_render.md"
    md.write_text(
        "# Table Cell Image Render\n\n"
        "| Item | Evidence |\n"
        "| --- | --- |\n"
        "| Existing | ![panel](figures/table%20panel.png) |\n",
        encoding="utf-8",
    )
    content = extract_md_content(str(md), output_dir=str(work))
    result = run_generated_case("md_table_cell_image_render_generated", content, base_format())
    table_xmls = re.findall(r"<w:tbl\b.*?</w:tbl>", result["xml"], flags=re.S)
    table_drawings = sum(xml.count("<w:drawing>") for xml in table_xmls)
    total_drawings = result["xml"].count("<w:drawing>")
    assert_true(total_drawings == 1, f"expected one rendered image drawing, saw {total_drawings}")
    assert_true(table_drawings == 1, "Markdown table-cell image rendered outside the generated Word table")
    assert_true(result["manifest"]["counts"]["content_images_rendered"] == 1, f"table-cell image render count changed: {result['manifest']}")
    assert_true(result["report"]["passed"] is True, f"table-cell image render should pass QA: {result['report']}")


@case
def md_reference_style_images_are_extracted_or_reported() -> None:
    work = new_workdir("md_reference_style_images")
    figures = work / "figures"
    figures.mkdir()
    (figures / "ref panel.png").write_bytes(PNG_1X1)
    md = work / "reference_images.md"
    md.write_text(
        "# Reference Images\n\n"
        "Before ![diagram][fig-one] after.\n\n"
        "Missing ![missing][not-defined] should become a QA-visible marker.\n\n"
        "[fig-one]: figures/ref%20panel.png?raw=true#panel-a\n",
        encoding="utf-8",
    )
    content = extract_md_content(str(md), output_dir=str(work))
    paragraphs = [p for sec in content["sections"] for p in sec.get("paragraphs", [])]
    images = [p for p in paragraphs if isinstance(p, dict) and p.get("role") == "image"]
    missing = content["_meta"].get("missing_images") or []
    assert_true(content["_meta"]["images_extracted"] == 1, f"reference-style Markdown image was not copied: {content['_meta']}")
    assert_true(len(images) == 1, f"reference-style Markdown image was not preserved in content stream: {paragraphs}")
    assert_true(len(missing) == 1 and missing[0].get("reason") == "reference_not_found", f"undefined image reference was not reported clearly: {missing}")
    assert_true(
        all("[fig-one]:" not in str(p) for p in paragraphs),
        f"Markdown image reference definition leaked into body content: {paragraphs}",
    )


@case
def md_reference_style_image_titles_do_not_leak_into_body() -> None:
    work = new_workdir("md_reference_image_titles")
    figures = work / "figures"
    figures.mkdir()
    (figures / "ref panel.png").write_bytes(PNG_1X1)
    (figures / "angle panel.png").write_bytes(PNG_1X1)
    md = work / "reference_image_titles.md"
    md.write_text(
        "# Reference Image Titles\n\n"
        "Before ![diagram][fig-one] after.\n\n"
        "Wrapped ![angle][fig-two] after.\n\n"
        "[fig-one]: figures/ref%20panel.png\n"
        "  \"Reference caption title\"\n\n"
        "[fig-two]: <figures/angle panel.png>\n"
        "  'Wrapped reference title'\n",
        encoding="utf-8",
    )
    content = extract_md_content(str(md), output_dir=str(work))
    paragraphs = [p for sec in content["sections"] for p in sec.get("paragraphs", [])]
    images = [p for p in paragraphs if isinstance(p, dict) and p.get("role") == "image"]
    missing = content["_meta"].get("missing_images") or []
    joined = "\n".join(str(p) for p in paragraphs)
    assert_true(content["_meta"]["images_extracted"] == 2, f"reference image titles prevented image copying: {content['_meta']}")
    assert_true(len(images) == 2, f"reference image title paths were not preserved in content stream: {paragraphs}")
    assert_true(not missing, f"existing reference-title images were reported missing: {missing}")
    assert_true("[fig-one]:" not in joined and "[fig-two]:" not in joined, f"reference definitions leaked into body content: {paragraphs}")
    assert_true(
        "Reference caption title" not in joined and "Wrapped reference title" not in joined,
        f"Markdown reference image titles leaked into body content: {paragraphs}",
    )


@case
def md_shortcut_reference_images_are_extracted_or_reported() -> None:
    work = new_workdir("md_shortcut_reference_images")
    figures = work / "figures"
    figures.mkdir()
    (figures / "shortcut panel.png").write_bytes(PNG_1X1)
    md = work / "shortcut_reference_images.md"
    md.write_text(
        "# Shortcut Reference Images\n\n"
        "Before ![Shortcut Panel] after.\n\n"
        "Missing ![Missing Shortcut] should become a QA-visible marker.\n\n"
        "[Shortcut Panel]: figures/shortcut%20panel.png?raw=true#panel-a\n",
        encoding="utf-8",
    )
    content = extract_md_content(str(md), output_dir=str(work))
    paragraphs = [p for sec in content["sections"] for p in sec.get("paragraphs", [])]
    images = [p for p in paragraphs if isinstance(p, dict) and p.get("role") == "image"]
    missing = content["_meta"].get("missing_images") or []
    assert_true(content["_meta"]["images_extracted"] == 1, f"shortcut reference Markdown image was not copied: {content['_meta']}")
    assert_true(len(images) == 1, f"shortcut reference Markdown image was not preserved in content stream: {paragraphs}")
    assert_true(len(missing) == 1 and missing[0].get("reason") == "reference_not_found", f"undefined shortcut image reference was not reported clearly: {missing}")
    assert_true(
        all("[Shortcut Panel]:" not in str(p) for p in paragraphs),
        f"Markdown shortcut image reference definition leaked into body content: {paragraphs}",
    )


@case
def md_reference_definition_like_code_lines_stay_in_code_blocks() -> None:
    work = new_workdir("md_reference_code")
    md = work / "reference_code.md"
    md.write_text(
        "# Reference Code\n\n"
        "```markdown\n"
        "[fig-one]: figures/example.png\n"
        "![diagram][fig-one]\n"
        "```\n",
        encoding="utf-8",
    )
    content = extract_md_content(str(md), output_dir=str(work))
    paragraphs = [p for sec in content["sections"] for p in sec.get("paragraphs", [])]
    code = next((p for p in paragraphs if isinstance(p, dict) and p.get("role") == "code"), {})
    assert_true("[fig-one]: figures/example.png" in str(code.get("code") or ""), f"reference definition-like code line was stripped: {paragraphs}")
    assert_true(content["_meta"].get("images_extracted") == 0, f"code-block image reference should not be extracted: {content['_meta']}")


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

    result = run_generated_case("md_missing_images_generated", content, base_format())
    report = result["report"]
    codes = [item["code"] for item in report["issues"]]
    assert_true("CONTENT_IMAGE_MISSING" in codes, "QA did not report missing Markdown images")
    assert_true("CONTENT_IMAGE_REMOTE_UNSUPPORTED" in codes, "QA did not give remote Markdown images their own beginner-facing code")
    assert_true(report["passed"] is False, "missing images should fail QA")
    assert_true(
        "CONTENT_IMAGE_REMOTE_UNSUPPORTED" in str(report.get("next_action") or ""),
        f"remote image next_action did not name the actionable code: {report.get('next_action')}",
    )
    remote_step = next(
        (step for step in (report.get("repair_plan") or {}).get("steps", []) if step.get("code") == "CONTENT_IMAGE_REMOTE_UNSUPPORTED"),
        {},
    )
    assert_true("下载" in str(remote_step.get("user_action") or ""), f"remote image guide should tell users to download/localize the image: {remote_step}")


@case
def md_unreadable_local_images_are_reported_to_qa() -> None:
    work = new_workdir("md_unreadable_images")
    (work / "broken.png").write_bytes(b"not a real png")
    md = work / "unreadable_image.md"
    md.write_text(
        "# Unreadable Image\n\nBroken ![broken](broken.png) should tell users to replace the image file.",
        encoding="utf-8",
    )
    content = extract_md_content(str(md), output_dir=str(work))
    missing = content["_meta"].get("missing_images") or []
    paragraphs = [p for sec in content["sections"] for p in sec.get("paragraphs", [])]
    assert_true(content["_meta"].get("images_extracted") == 0, f"unreadable image should not be counted as extracted: {content['_meta']}")
    assert_true(
        any(isinstance(item, dict) and item.get("reason") == "unreadable" for item in missing),
        f"unreadable image was not recorded with a specific reason: {missing}",
    )
    assert_true(
        any(isinstance(p, dict) and p.get("role") == "missing_image" and p.get("reason") == "unreadable" for p in paragraphs),
        f"unreadable image marker not preserved in content stream: {paragraphs}",
    )

    result = run_generated_case("md_unreadable_images_generated", content, base_format())
    report = result["report"]
    codes = [item["code"] for item in report["issues"]]
    assert_true("CONTENT_IMAGE_UNREADABLE" in codes, f"QA did not report unreadable Markdown images: {codes}")
    assert_true("CONTENT_IMAGE_MISSING" not in codes, f"unreadable images should not collapse into generic missing-image guidance: {codes}")
    assert_true(
        "CONTENT_IMAGE_UNREADABLE" in str(report.get("next_action") or ""),
        f"unreadable image next_action did not name the actionable code: {report.get('next_action')}",
    )
    unreadable_step = next(
        (step for step in (report.get("repair_plan") or {}).get("steps", []) if step.get("code") == "CONTENT_IMAGE_UNREADABLE"),
        {},
    )
    action = str(unreadable_step.get("user_action") or "")
    assert_true("重新导出" in action and "PNG" in action, f"unreadable image guide should tell users to re-export a normal image: {unreadable_step}")


@case
def md_unsupported_local_image_formats_are_reported_to_qa() -> None:
    work = new_workdir("md_unsupported_image_formats")
    (work / "animated.gif").write_bytes(GIF_1X1)
    md = work / "unsupported_image_format.md"
    md.write_text(
        "# Unsupported Image Format\n\nGIF ![animated](animated.gif) should ask users to export PNG/JPG first.",
        encoding="utf-8",
    )
    content = extract_md_content(str(md), output_dir=str(work))
    missing = content["_meta"].get("missing_images") or []
    paragraphs = [p for sec in content["sections"] for p in sec.get("paragraphs", [])]
    assert_true(content["_meta"].get("images_extracted") == 0, f"unsupported GIF should not be counted as extracted: {content['_meta']}")
    assert_true(
        any(isinstance(item, dict) and item.get("reason") == "unreadable" and ".gif" in str(item.get("detail") or "") for item in missing),
        f"unsupported GIF was not recorded as an unreadable user-file issue: {missing}",
    )
    assert_true(
        any(isinstance(p, dict) and p.get("role") == "missing_image" and p.get("reason") == "unreadable" for p in paragraphs),
        f"unsupported GIF marker not preserved in content stream: {paragraphs}",
    )

    result = run_generated_case("md_unsupported_image_formats_generated", content, base_format())
    report = result["report"]
    codes = [item["code"] for item in report["issues"]]
    assert_true("CONTENT_IMAGE_UNREADABLE" in codes, f"QA did not report unsupported local image formats: {codes}")
    assert_true("CONTENT_IMAGE_MISSING" not in codes, f"unsupported local formats should not collapse into generic missing-image guidance: {codes}")
    action = str((report.get("repair_plan") or {}).get("next_action") or report.get("next_action") or "")
    assert_true("CONTENT_IMAGE_UNREADABLE" in action and "PNG" in action, f"unsupported-format next action should tell users to export PNG/JPG: {action}")


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


