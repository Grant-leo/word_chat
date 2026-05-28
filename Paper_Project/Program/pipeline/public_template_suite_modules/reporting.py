"""Markdown report writers for the public template compatibility suite."""
from __future__ import annotations

from typing import Any, Dict, List

from public_template_suite_modules.paths import TEST_ROOT
from public_template_suite_modules.scenarios import SCENARIOS


def write_readme(results: List[Dict[str, Any]]) -> None:
    lines = [
        "# Public Template Test Set",
        "",
        "Downloaded public templates are stored under `files/` and ignored by Git. Test run artifacts are stored under `runs/` and ignored by Git.",
        "",
        "The suite uses synthetic non-private content only.",
        "",
        "## Scenarios",
        "",
    ]
    for scenario in SCENARIOS:
        lines.append(f"- `{scenario['id']}`: {scenario['description']}")
    lines.extend(["", "## Results", ""])
    for item in results:
        status = "PASS" if item.get("qa_passed") else "FAIL"
        fmt = item.get("format") or {}
        lines.append(
            f"- **{status}** `{item.get('id')}`: {item.get('source')} | "
            f"paragraphs={fmt.get('paragraphs', 'n/a')} tables={fmt.get('tables', 'n/a')} "
            f"cover={fmt.get('cover_elements', 'n/a')} qa_errors={item.get('qa_errors', 'n/a')} "
            f"qa_warnings={item.get('qa_warnings', 'n/a')} "
            f"conformance_errors={item.get('conformance_errors', 'n/a')} "
            f"visual_errors={item.get('visual_errors', 'n/a')}"
        )
        lines.append(f"  - Source: {item.get('url')}")
        if item.get("error"):
            lines.append(f"  - Error: `{item.get('error')}`")
        for scenario in item.get("scenarios") or []:
            scenario_status = "PASS" if scenario.get("qa_passed") else "FAIL"
            lines.append(
                f"  - {scenario_status} `{scenario.get('id')}`: "
                f"errors={scenario.get('qa_errors', 'n/a')} warnings={scenario.get('qa_warnings', 'n/a')} "
                f"issues={scenario.get('qa_issues', 'n/a')} "
                f"conformance={scenario.get('conformance_issues', 'n/a')} "
                f"visual={scenario.get('visual_issues', 'n/a')}"
            )
            if scenario.get("error"):
                lines.append(f"    - Error: `{scenario.get('error')}`")
    lines.extend(
        [
            "",
            "## Rebuild",
            "",
            "```bash",
            "python Paper_Project/Program/pipeline/public_template_suite.py --download",
            "```",
            "",
            "Run a single scenario when narrowing a regression:",
            "",
            "```bash",
            "python Paper_Project/Program/pipeline/public_template_suite.py --template buaa_word_template --scenario rich_math",
            "```",
            "",
            "Filtered runs write `runs/last_selection_summary.json` and leave the canonical manifest/README untouched.",
            "",
            "Run visual QA when Word COM and Poppler are available:",
            "",
            "```bash",
            "python Paper_Project/Program/pipeline/public_template_suite.py --visual",
            "```",
            "",
            "Golden-baseline hashes can be updated during visual runs:",
            "",
            "```bash",
            "python Paper_Project/Program/pipeline/public_template_suite.py --visual --update-golden",
            "```",
            "",
            "## Add Manual Templates",
            "",
            "For templates that require login or manual download, place the DOCX under `files/` and add an item to `manifest.json`:",
            "",
            "```json",
            "{",
            "  \"id\": \"school_template_local\",",
            "  \"name\": \"school_template.docx\",",
            "  \"source\": \"manual download\",",
            "  \"file\": \"TestData/PublicTemplates/files/school_template.docx\",",
            "  \"license_hint\": \"local test only; do not redistribute\"",
            "}",
            "```",
            "",
            "`files/` and `runs/` are ignored by Git, so downloaded templates and generated DOCX/PDF/PNG artifacts stay local.",
            "",
        ]
    )
    (TEST_ROOT / "README.md").write_text("\n".join(lines), encoding="utf-8")

