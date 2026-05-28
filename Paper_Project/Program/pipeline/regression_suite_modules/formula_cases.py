"""Formula and OMML regression cases."""
from __future__ import annotations

from lxml import etree

from content_parser_modules.formula_extractor import _strip_trailing_formula_labels_from_xml
from latex_omath import latex_to_omath
from qa_conformance import check_conformance

from regression_suite_modules.generated_docx import omath_count, omath_para_count, run_generated_case
from regression_suite_modules.harness import assert_true, base_content, case

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
def latex_delimited_text_formula_with_number_renders_native() -> None:
    content = base_content([
        {"role": "formula", "source": "text", "text": r"$$E_{total}=\sum_{t=1}^{24}P(t)\Delta t$$ (1.1)"}
    ])
    result = run_generated_case("latex_delimited_numbered_text_formula", content)
    assert_true(result["manifest"]["counts"]["display_formulas_rendered"] == 1, "numbered LaTeX text formula was not rendered")
    assert_true(omath_para_count(result["xml"]) == 1, "numbered LaTeX text formula did not create display OMML")
    assert_true("[LaTeX error" not in result["xml"] and "$$" not in result["xml"], "LaTeX text leaked into final XML")
    assert_true(not result["report"]["issues"], f"unexpected QA issues: {result['report']['issues']}")


@case
def latex_delimited_appendix_formula_with_letter_number_renders_native() -> None:
    content = base_content([
        {
            "role": "formula",
            "source": "text",
            "text": r"$$R_{green}=\frac{E_{renew}}{E_{total}}\times100\%$$" + " \uff08A.1\uff09",
        }
    ])
    result = run_generated_case("latex_delimited_appendix_formula", content)
    assert_true(result["manifest"]["counts"]["display_formulas_rendered"] == 1, "appendix-numbered LaTeX text formula was not rendered")
    assert_true(omath_para_count(result["xml"]) == 1, "appendix-numbered formula did not create display OMML")
    assert_true("[LaTeX error" not in result["xml"] and "$$" not in result["xml"], "appendix formula leaked LaTeX/error text")
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
def latex_omath_sqrt_optional_degree_renders_radical_degree() -> None:
    xml = latex_to_omath(r"\sqrt[3]{x}", display=True)
    assert_true("[LaTeX error" not in xml, "nth-root optional degree produced a LaTeX error")
    root = etree.fromstring(xml.encode("utf-8"))
    ns = {"m": "http://schemas.openxmlformats.org/officeDocument/2006/math"}
    assert_true(root.find(".//m:rad/m:deg", ns) is not None, "nth-root did not create an OMML radical degree")
    texts = "".join(t.text or "" for t in root.findall(".//m:t", ns))
    assert_true("3" in texts and "x" in texts, f"nth-root text content was lost: {texts}")


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
    assert_true(("abc+def", "p") in styled_runs or (("abc", "p") in styled_runs and ("def", "p") in styled_runs), f"mixed-style grouped runs lost \\mathrm style: {styled_runs}")


@case
def latex_omath_keeps_literals_operators_and_brackets_upright() -> None:
    xml = latex_to_omath(r"P(t)=\max(0,PRE(t)-P_{total}(t))+\frac{x_1}{2}+\{z\}", display=True)
    root = etree.fromstring(xml.encode("utf-8"))
    ns = {"m": "http://schemas.openxmlformats.org/officeDocument/2006/math"}
    upright_chars = set("0123456789()[]{}=+-*/×÷<>≤≥≈≠,:;.%")
    bad_runs = []
    variable_runs = []
    for run in root.findall(".//m:r", ns):
        text = "".join(t.text or "" for t in run.findall("./m:t", ns))
        sty = run.find("./m:rPr/m:sty", ns)
        style = sty.get(f"{{{ns['m']}}}val") if sty is not None else None
        if text in {"P", "t", "x"} and style is None:
            variable_runs.append(text)
        if any(ch in upright_chars for ch in text) and style != "p":
            bad_runs.append((text, style))
    assert_true(not bad_runs, f"formula literals/operators/brackets should be upright, got {bad_runs}")
    assert_true({"P", "t", "x"}.issubset(set(variable_runs)), f"variables should keep default math style: {variable_runs}")



