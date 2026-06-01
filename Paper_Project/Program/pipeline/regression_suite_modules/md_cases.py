"""Markdown parser regression cases."""
from __future__ import annotations

from md_parser import extract_content as extract_md_content

from regression_suite_modules.generated_docx import run_generated_case
from regression_suite_modules.harness import (
    PNG_1X1,
    assert_true,
    base_format,
    case,
    new_workdir,
)

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


