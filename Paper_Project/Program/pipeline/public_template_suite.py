"""
public_template_suite.py - local public-template compatibility checks.

The corpus lives under TestData/PublicTemplates:
- manifest.json records public source URLs and hashes.
- files/ stores downloaded DOCX templates and is ignored by Git.
- runs/ stores generated outputs and is ignored by Git.

The suite uses only synthetic non-private content.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import struct
import subprocess
import sys
import urllib.request
import zlib
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional


PIPELINE_DIR = Path(__file__).resolve().parent
ROOT = PIPELINE_DIR.parents[2]
TEST_ROOT = ROOT / "TestData" / "PublicTemplates"
FILES_DIR = TEST_ROOT / "files"
RUNS_DIR = TEST_ROOT / "runs"
MANIFEST_PATH = TEST_ROOT / "manifest.json"
SELECTION_SUMMARY_PATH = RUNS_DIR / "last_selection_summary.json"

if str(PIPELINE_DIR) not in sys.path:
    sys.path.insert(0, str(PIPELINE_DIR))

from format_extractor import extract as extract_format
from md_parser import extract_content as extract_md_content
from qa_conformance import check_and_write as conformance_check_and_write
from qa_conformance import write_requirements as write_template_requirements
from qa_checker import check_output, write_reports as write_qa_reports
from script_generator import generate
from template_profiler import write_profile

try:
    from qa_visual import check_and_write as visual_check_and_write
except Exception:
    visual_check_and_write = None


DEFAULT_TEMPLATES: List[Dict[str, Any]] = [
    {
        "id": "buaa_word_template",
        "name": "BUAAThesis Template.docx",
        "source": "GitHub CheckBoxStudio/BUAAThesis",
        "url": "https://github.com/CheckBoxStudio/BUAAThesis/blob/master/Template.docx",
        "download_url": "https://raw.githubusercontent.com/CheckBoxStudio/BUAAThesis/master/Template.docx",
        "license_hint": "repository includes LICENSE file",
    },
    {
        "id": "scu_undergraduate_template",
        "name": "SCU_undergraduate_thesis_template.docx",
        "source": "GitHub SunnyHaze/scu-thesis-template",
        "url": "https://github.com/SunnyHaze/scu-thesis-template/blob/main/%E5%9B%9B%E5%B7%9D%E5%A4%A7%E5%AD%A6%E6%9C%AC%E7%A7%91%E6%AF%95%E4%B8%9A%E8%AE%BA%E6%96%87%E6%A8%A1%E6%9D%BF.docx",
        "download_url": "https://raw.githubusercontent.com/SunnyHaze/scu-thesis-template/main/%E5%9B%9B%E5%B7%9D%E5%A4%A7%E5%AD%A6%E6%9C%AC%E7%A7%91%E6%AF%95%E4%B8%9A%E8%AE%BA%E6%96%87%E6%A8%A1%E6%9D%BF.docx",
        "license_hint": "public GitHub repository; verify license before redistribution",
    },
    {
        "id": "scu_software_engineering_template",
        "name": "SCU_software_engineering_template.docx",
        "source": "GitHub SunnyHaze/scu-thesis-template",
        "url": "https://github.com/SunnyHaze/scu-thesis-template/blob/main/%E5%9B%9B%E5%B7%9D%E5%A4%A7%E5%AD%A6%E6%9C%AC%E7%A7%91%E6%AF%95%E4%B8%9A%E8%AE%BE%E8%AE%A1%E8%AE%BA%E6%96%87%E6%A8%A1%E6%9D%BF(%E8%BD%AF%E4%BB%B6%E5%AD%A6%E9%99%A2).docx",
        "download_url": "https://raw.githubusercontent.com/SunnyHaze/scu-thesis-template/main/%E5%9B%9B%E5%B7%9D%E5%A4%A7%E5%AD%A6%E6%9C%AC%E7%A7%91%E6%AF%95%E4%B8%9A%E8%AE%BE%E8%AE%A1%E8%AE%BA%E6%96%87%E6%A8%A1%E6%9D%BF(%E8%BD%AF%E4%BB%B6%E5%AD%A6%E9%99%A2).docx",
        "license_hint": "public GitHub repository; verify license before redistribution",
    },
    {
        "id": "opavon_thesis_cover_template",
        "name": "opavon_thesis_cover_template.docx",
        "source": "GitHub opavon/ThesisTemplate",
        "url": "https://github.com/opavon/ThesisTemplate/blob/main/Thesis_and_Cover_Template.docx",
        "download_url": "https://raw.githubusercontent.com/opavon/ThesisTemplate/main/Thesis_and_Cover_Template.docx",
        "license_hint": "public GitHub repository; verify license before redistribution",
    },
    {
        "id": "upm_official_word_thesis_template",
        "name": "UPM_official_word_thesis_template.docx",
        "source": "GitHub SinaAbdipoor/2025-upm-latex-thesis-template",
        "url": "https://github.com/SinaAbdipoor/2025-upm-latex-thesis-template/blob/main/Source/UPM%20Official%20Word%20Thesis%20Template/20240812102832Thesis_A4_Format_(Template)_(17.05.2024).docx",
        "download_url": "https://raw.githubusercontent.com/SinaAbdipoor/2025-upm-latex-thesis-template/main/Source/UPM%20Official%20Word%20Thesis%20Template/20240812102832Thesis_A4_Format_(Template)_(17.05.2024).docx",
        "license_hint": "public GitHub repository; verify license before redistribution",
    },
]

CN_TITLE = "\u5408\u6210\u516c\u5f00\u6d4b\u8bd5\u8bba\u6587\u6807\u9898"
CN_ABSTRACT = "\u672c\u6587\u7528\u4e8e\u516c\u5f00\u6a21\u677f\u517c\u5bb9\u6027\u56de\u5f52\u6d4b\u8bd5\uff0c\u4e0d\u5305\u542b\u4efb\u4f55\u79c1\u6709\u6570\u636e\u3002"
CN_KEYWORDS = "\u6a21\u677f\u6d4b\u8bd5\uff1b\u81ea\u52a8\u6392\u7248\uff1b\u8d28\u91cf\u68c0\u67e5"


ScenarioBuilder = Callable[[Path, Path], Dict[str, Any]]


def _png_chunk(kind: bytes, payload: bytes) -> bytes:
    return struct.pack(">I", len(payload)) + kind + payload + struct.pack(">I", zlib.crc32(kind + payload) & 0xFFFFFFFF)


def make_png_bytes(width: int, height: int, color_a: tuple[int, int, int], color_b: tuple[int, int, int]) -> bytes:
    rows = []
    for y in range(height):
        row = bytearray([0])
        for x in range(width):
            color = color_a if (x + y) % 2 == 0 else color_b
            row.extend(color)
        rows.append(bytes(row))
    raw = b"".join(rows)
    return b"".join(
        [
            b"\x89PNG\r\n\x1a\n",
            _png_chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)),
            _png_chunk(b"IDAT", zlib.compress(raw)),
            _png_chunk(b"IEND", b""),
        ]
    )


def write_sample_images(fig_dir: Path) -> Dict[str, str]:
    fig_dir.mkdir(parents=True, exist_ok=True)
    images = {
        "wide.png": make_png_bytes(320, 120, (33, 111, 181), (245, 247, 250)),
        "tall.png": make_png_bytes(120, 260, (45, 140, 98), (248, 249, 250)),
        "small.png": make_png_bytes(48, 48, (180, 68, 68), (255, 255, 255)),
    }
    for name, data in images.items():
        (fig_dir / name).write_bytes(data)
    return {name: name for name in images}


def content_base(fig_dir: Path, source: str, paragraphs: int, tables: int, images: int) -> Dict[str, Any]:
    return {
        "_meta": {
            "source": source,
            "sha256": "synthetic-public",
            "paragraphs": paragraphs,
            "tables_count": tables,
            "images_extracted": images,
            "images_dir": str(fig_dir),
            "missing_images": [],
            "image_extract_failures": [],
        },
        "title_info": {"title_cn": CN_TITLE, "title_en": "Synthetic Public Template Regression"},
        "cover_info": {
            "paper_title": CN_TITLE,
            "student_name": "\u6d4b\u8bd5\u5b66\u751f",
            "student_id": "000000",
            "advisor": "\u6d4b\u8bd5\u5bfc\u5e08",
            "college": "\u6d4b\u8bd5\u5b66\u9662",
            "class_name": "\u6d4b\u8bd5\u4e13\u4e1a",
        },
        "sections": [],
        "references": ["[1] Synthetic public regression reference."],
    }


def scenario_frontmatter_baseline(fig_dir: Path, work_dir: Path) -> Dict[str, Any]:
    write_sample_images(fig_dir)
    content = content_base(fig_dir, "synthetic_frontmatter.json", paragraphs=9, tables=1, images=1)
    content["sections"] = [
        {"heading": "\u6458\u8981", "level": 1, "role": "cn_abstract", "paragraphs": [CN_ABSTRACT], "images": []},
        {"heading": "\u5173\u952e\u8bcd", "level": 1, "role": "cn_keywords", "paragraphs": [CN_KEYWORDS], "images": []},
        {"heading": "Abstract", "level": 1, "role": "en_abstract", "paragraphs": ["This synthetic paper is used for public template regression testing."], "images": []},
        {"heading": "Key words", "level": 1, "role": "en_keywords", "paragraphs": ["template test; document automation; QA"], "images": []},
        {
            "heading": "1 Introduction",
            "level": 1,
            "role": "body",
            "images": ["wide.png"],
            "paragraphs": [
                "This paragraph checks body rendering, indentation, line spacing, and font fallback.",
                "\u8fd9\u4e2a\u6bb5\u843d\u68c0\u67e5\u4e2d\u82f1\u6587\u6df7\u6392\u3001\u9996\u884c\u7f29\u8fdb\u548c\u6b63\u6587\u6837\u5f0f\u3002",
                {"role": "formula", "latex": "a=b+c", "text": "a=b+c", "numbered": False},
                {"role": "figure", "image": "wide.png", "caption": "Figure 1 synthetic wide image"},
                {"role": "table", "table_rows": [["Metric", "Value"], ["Images", "1"], ["Tables", "1"]]},
                "The paragraph after the table checks spacing recovery.",
            ],
        },
    ]
    return content


def scenario_rich_math(fig_dir: Path, work_dir: Path) -> Dict[str, Any]:
    write_sample_images(fig_dir)
    inline_math = {"type": "inline", "latex": "E=mc^2", "text": "E=mc^2"}
    content = content_base(fig_dir, "synthetic_rich_math.json", paragraphs=8, tables=0, images=0)
    content["sections"] = [
        {
            "heading": "1 Formula Coverage",
            "level": 1,
            "role": "body",
            "images": [],
            "paragraphs": [
                {
                    "role": "rich_text",
                    "text": "Inline formula E=mc^2 should remain editable.",
                    "runs": [
                        {"type": "text", "text": "Inline formula "},
                        {"type": "math", "text": "E=mc^2", "math": [inline_math]},
                        {"type": "text", "text": " should remain editable."},
                    ],
                    "math": [inline_math],
                },
                {"role": "formula", "latex": "\\frac{a+b}{c+d}", "text": "(a+b)/(c+d)", "numbered": True},
                {"role": "formula", "latex": "\\sqrt{x^2+y^2}", "text": "sqrt(x^2+y^2)", "numbered": False},
                {"role": "formula", "latex": "\\sum_{i=1}^{n} i = \\frac{n(n+1)}{2}", "text": "sum i = n(n+1)/2", "numbered": True},
                "The following paragraph verifies that body text after several formulas returns to normal spacing.",
            ],
        },
        {
            "heading": "2 Formula Context",
            "level": 1,
            "role": "body",
            "images": [],
            "paragraphs": [
                "Formula paragraphs often appear between explanatory paragraphs; this section checks ordering.",
                {"role": "formula", "latex": "y=\\alpha x+\\beta", "text": "y=alpha*x+beta", "numbered": False},
            ],
        },
    ]
    return content


def scenario_media_tables_code(fig_dir: Path, work_dir: Path) -> Dict[str, Any]:
    write_sample_images(fig_dir)
    content = content_base(fig_dir, "synthetic_media_tables_code.json", paragraphs=11, tables=2, images=2)
    content["sections"] = [
        {
            "heading": "1 Media And Tables",
            "level": 1,
            "role": "body",
            "images": ["wide.png", "tall.png"],
            "paragraphs": [
                "The first image is wide and should be scaled within page margins.",
                {"role": "figure", "image": "wide.png", "caption": "Figure 1 wide synthetic figure"},
                "The second image is tall and checks aspect-ratio preservation.",
                {"role": "figure", "image": "tall.png", "caption": "Figure 2 tall synthetic figure"},
                {"role": "table", "table_rows": [["Item", "Description", "Result"], ["A", "Short value", "Pass"], ["B", "Longer wrapped cell content used to check table width and line wrapping.", "Review"]]},
                {"role": "table", "table_rows": [["Index", "Metric"], ["1", "Alpha"], ["2", "Beta"], ["3", "Gamma"]]},
            ],
        },
        {
            "heading": "2 Code Block",
            "level": 1,
            "role": "body",
            "images": [],
            "paragraphs": [
                {"role": "code", "language": "python", "code": "def add(a, b):\n    return a + b\n\nprint(add(1, 2))"},
                "The paragraph after code checks that monospace styling does not leak into normal body text.",
            ],
        },
        {
            "heading": "\u81f4\u8c22",
            "level": 1,
            "role": "acknowledgement",
            "images": [],
            "paragraphs": ["This acknowledgement paragraph checks back-matter placement."],
        },
    ]
    content["references"] = ["[1] Public synthetic table reference.", "[2] Public synthetic media reference."]
    return content


def scenario_long_multilevel(fig_dir: Path, work_dir: Path) -> Dict[str, Any]:
    write_sample_images(fig_dir)
    content = content_base(fig_dir, "synthetic_long_multilevel.json", paragraphs=36, tables=1, images=1)
    sections: List[Dict[str, Any]] = []
    for chapter in range(1, 4):
        sections.append(
            {
                "heading": f"{chapter} Chapter {chapter}",
                "level": 1,
                "role": "body",
                "images": [],
                "paragraphs": [
                    "This chapter opening paragraph checks heading level one spacing and numbering recovery.",
                    "A longer synthetic paragraph repeats enough content to stress pagination while remaining non-private. " * 3,
                ],
            }
        )
        for sub in range(1, 3):
            sections.append(
                {
                    "heading": f"{chapter}.{sub} Section {chapter}.{sub}",
                    "level": 2,
                    "role": "body",
                    "images": ["small.png"] if chapter == 2 and sub == 1 else [],
                    "paragraphs": [
                        "This second-level section checks left alignment, spacing before and after headings, and body continuation.",
                        {"role": "figure", "image": "small.png", "caption": "Figure synthetic small marker"} if chapter == 2 and sub == 1 else "Plain paragraph without a figure.",
                    ],
                }
            )
            sections.append(
                {
                    "heading": f"{chapter}.{sub}.1 Detail",
                    "level": 3,
                    "role": "body",
                    "images": [],
                    "paragraphs": [
                        "Third-level heading content checks whether the generator preserves hierarchy without over-promoting headings.",
                        {"role": "table", "table_rows": [["Level", "Observed"], ["H1", str(chapter)], ["H2", str(sub)]]} if chapter == 3 and sub == 2 else "No table in this subsection.",
                    ],
                }
            )
    sections.append(
        {
            "heading": "\u9644\u5f55 A",
            "level": 1,
            "role": "appendix",
            "images": [],
            "paragraphs": ["Appendix content checks back-matter heading recognition and spacing."],
        }
    )
    content["sections"] = sections
    content["references"] = ["[1] Synthetic long-document reference.", "[2] Synthetic appendix reference."]
    return content


def scenario_markdown_content(fig_dir: Path, work_dir: Path) -> Dict[str, Any]:
    asset_dir = work_dir / "md_assets"
    asset_dir.mkdir(parents=True, exist_ok=True)
    (asset_dir / "chart.png").write_bytes(make_png_bytes(240, 140, (82, 127, 179), (255, 255, 255)))
    md_path = work_dir / "synthetic_source.md"
    md_text = "\n".join(
        [
            "# Synthetic Markdown Thesis",
            "",
            "## Abstract",
            "",
            "Markdown input paragraph with inline formula $x^2+y^2=z^2$ and **bold** markers.",
            "",
            "## 1 Markdown Body",
            "",
            "This paragraph checks Markdown parser extraction and body rendering.",
            "",
            "$$",
            "\\frac{1}{n}\\sum_{i=1}^{n}x_i",
            "$$",
            "",
            "![Synthetic chart](md_assets/chart.png)",
            "",
            "| Metric | Value |",
            "| --- | --- |",
            "| Rows | 2 |",
            "| Formulas | yes |",
            "",
            "```python",
            "value = sum([1, 2, 3])",
            "print(value)",
            "```",
            "",
            "## References",
            "",
            "[1] Synthetic Markdown reference.",
        ]
    )
    md_path.write_text(md_text, encoding="utf-8")
    content = extract_md_content(str(md_path), output_dir=str(work_dir))
    content.setdefault("cover_info", {})
    content["cover_info"].update(
        {
            "paper_title": "Synthetic Markdown Thesis",
            "student_name": "Synthetic Student",
            "student_id": "MD0001",
            "advisor": "Synthetic Advisor",
            "college": "Synthetic College",
        }
    )
    return content


SCENARIOS: List[Dict[str, Any]] = [
    {"id": "frontmatter_baseline", "description": "cover, bilingual abstract, keywords, formula, figure, table", "builder": scenario_frontmatter_baseline},
    {"id": "rich_math", "description": "inline and display formulas intended for native equation rendering", "builder": scenario_rich_math},
    {"id": "media_tables_code", "description": "wide/tall figures, captions, wrapped tables, code, acknowledgement", "builder": scenario_media_tables_code},
    {"id": "long_multilevel", "description": "multi-level headings, pagination stress, appendix, references", "builder": scenario_long_multilevel},
    {"id": "markdown_content", "description": "Markdown parser path with image, table, formula, code", "builder": scenario_markdown_content},
]


def safe_download_url(value: Any) -> bool:
    if not isinstance(value, str) or not value.startswith(("http://", "https://")):
        return False
    if any(token in value for token in ("????", "\ufffd")):
        return False
    try:
        value.encode("ascii")
    except UnicodeEncodeError:
        return False
    return True


def read_manifest() -> Dict[str, Any]:
    default_by_id = {str(item.get("id")): item for item in DEFAULT_TEMPLATES}
    if MANIFEST_PATH.exists():
        with open(MANIFEST_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        if data.get("templates"):
            merged = []
            for item in data.get("templates") or []:
                item_id = str(item.get("id") or "")
                base = dict(default_by_id.get(item_id, {}))
                clean = {}
                for key, value in item.items():
                    if value in (None, ""):
                        continue
                    if key == "download_url" and base.get("download_url") and not safe_download_url(value):
                        continue
                    clean[key] = value
                base.update(clean)
                merged.append(base)
            seen = {str(item.get("id")) for item in merged}
            for item in DEFAULT_TEMPLATES:
                if str(item.get("id")) not in seen:
                    merged.append(dict(item))
            data["templates"] = merged
            return data
    return {"schema_version": 1, "templates": [dict(item) for item in DEFAULT_TEMPLATES]}


def write_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(value, f, ensure_ascii=False, indent=2)


def sha256_file(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def resolve_existing_file(item: Dict[str, Any]) -> Optional[Path]:
    raw = item.get("file")
    if not raw:
        return None
    path = Path(str(raw))
    candidates = [path]
    if not path.is_absolute():
        candidates.extend([ROOT / path, TEST_ROOT / path, FILES_DIR / path.name])
    for candidate in candidates:
        if candidate.exists() and candidate.is_file():
            return candidate
    return None


def download_template(item: Dict[str, Any], force: bool = False) -> Path:
    existing = resolve_existing_file(item)
    if existing and existing.exists() and not force:
        return existing

    FILES_DIR.mkdir(parents=True, exist_ok=True)
    if item.get("name"):
        name = str(item.get("name"))
    elif item.get("file"):
        name = Path(str(item.get("file"))).name
    else:
        name = f"{item.get('id', 'template')}.docx"
    path = FILES_DIR / name
    if path.exists() and path.stat().st_size > 2000 and not force:
        return path
    url = item.get("download_url") or item.get("url")
    if not safe_download_url(url):
        raise RuntimeError(f"missing or unsafe ASCII download_url for {item.get('id')}")
    req = urllib.request.Request(str(url), headers={"User-Agent": "word-chat-template-test/1.0"})
    with urllib.request.urlopen(req, timeout=60) as resp:
        data = resp.read()
    if len(data) < 2000 or not data.startswith(b"PK"):
        raise RuntimeError(f"downloaded file is not a DOCX zip: {len(data)} bytes")
    path.write_bytes(data)
    return path


def run_build(build_py: Path, cwd: Path) -> Dict[str, Any]:
    subprocess.run([sys.executable, "-m", "py_compile", str(build_py)], check=True, timeout=120)
    completed = subprocess.run(
        [sys.executable, str(build_py)],
        cwd=str(cwd),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=240,
    )
    return {
        "returncode": completed.returncode,
        "stdout_tail": (completed.stdout or "")[-2000:],
        "stderr_tail": (completed.stderr or "")[-2000:],
    }


def run_scenario(
    template_id: str,
    fmt: Dict[str, Any],
    scenario: Dict[str, Any],
    template_run_dir: Path,
    visual: bool = False,
    golden_dir: Optional[Path] = None,
    update_golden: bool = False,
) -> Dict[str, Any]:
    scenario_id = str(scenario["id"])
    scenario_dir = template_run_dir / scenario_id
    if scenario_dir.exists():
        shutil.rmtree(scenario_dir)
    scenario_dir.mkdir(parents=True, exist_ok=True)

    result: Dict[str, Any] = {
        "id": scenario_id,
        "description": scenario.get("description"),
        "qa_passed": False,
        "run_dir": str(scenario_dir.relative_to(ROOT)).replace("\\", "/"),
    }

    try:
        write_json(scenario_dir / "format.json", fmt)
        builder: ScenarioBuilder = scenario["builder"]
        content = builder(scenario_dir / "figures_src", scenario_dir)
        write_json(scenario_dir / "content.json", content)
        write_json(scenario_dir / "workflow_mode.json", {"mode": "developer", "template_id": template_id, "scenario": scenario_id})
        write_template_requirements(fmt, content, str(scenario_dir))
        generate(str(scenario_dir / "format.json"), str(scenario_dir / "content.json"), str(scenario_dir), "synthetic_output.docx")

        build_result = run_build(scenario_dir / "build_generated.py", scenario_dir)
        result["build_returncode"] = build_result["returncode"]
        if build_result["returncode"] != 0:
            result["error"] = (build_result["stderr_tail"] or build_result["stdout_tail"] or "build failed")[:2000]
            return result

        report = check_output(str(scenario_dir), mode="developer", output_docx_name="synthetic_output.docx")
        write_qa_reports(report, str(scenario_dir))
        issues = report.get("issues") or []
        result.update(
            {
                "qa_passed": bool(report.get("passed")),
                "qa_errors": sum(1 for issue in issues if issue.get("severity") == "error"),
                "qa_warnings": sum(1 for issue in issues if issue.get("severity") == "warning"),
                "qa_issues": len(issues),
                "content": {
                    "sections": len(content.get("sections") or []),
                    "paragraphs": content.get("_meta", {}).get("paragraphs"),
                    "tables": content.get("_meta", {}).get("tables_count"),
                    "images": content.get("_meta", {}).get("images_extracted"),
                    "references": len(content.get("references") or []),
                },
            }
        )
        conformance_report = conformance_check_and_write(str(scenario_dir), mode="developer", output_docx_name="synthetic_output.docx")
        conformance_issues = conformance_report.get("issues") or []
        result.update(
            {
                "conformance_passed": bool(conformance_report.get("passed")),
                "conformance_errors": sum(1 for issue in conformance_issues if issue.get("severity") == "error"),
                "conformance_warnings": sum(1 for issue in conformance_issues if issue.get("severity") == "warning"),
                "conformance_issues": len(conformance_issues),
                "conformance_counts": conformance_report.get("counts") or {},
            }
        )
        if not conformance_report.get("passed"):
            result["qa_passed"] = False
        if visual:
            if visual_check_and_write is None:
                result.update({"visual_passed": False, "visual_errors": 1, "visual_warnings": 0, "visual_issues": 1, "visual_error": "qa_visual is not available"})
                result["qa_passed"] = False
            else:
                visual_report = visual_check_and_write(
                    str(scenario_dir),
                    output_docx_name="synthetic_output.docx",
                    project_root=str(ROOT),
                    golden_dir=str(golden_dir) if golden_dir else None,
                    update_golden=update_golden,
                )
                visual_issues = visual_report.get("issues") or []
                result.update(
                    {
                        "visual_passed": bool(visual_report.get("passed")),
                        "visual_errors": sum(1 for issue in visual_issues if issue.get("severity") == "error"),
                        "visual_warnings": sum(1 for issue in visual_issues if issue.get("severity") == "warning"),
                        "visual_issues": len(visual_issues),
                        "visual_counts": visual_report.get("counts") or {},
                    }
                )
                if not visual_report.get("passed"):
                    result["qa_passed"] = False
    except Exception as exc:
        result["error"] = str(exc)
    return result


def run_template(
    item: Dict[str, Any],
    force_download: bool = False,
    scenario_filter: Optional[str] = None,
    visual: bool = False,
    golden_dir: Optional[Path] = None,
    update_golden: bool = False,
) -> Dict[str, Any]:
    docx_path = download_template(item, force=force_download)
    run_id = str(item.get("id") or docx_path.stem)
    template_run_dir = RUNS_DIR / run_id
    if template_run_dir.exists():
        shutil.rmtree(template_run_dir)
    template_run_dir.mkdir(parents=True, exist_ok=True)

    result = {k: item.get(k) for k in ("id", "name", "source", "url", "download_url", "license_hint")}
    result.update(
        {
            "file": str(docx_path.relative_to(ROOT)).replace("\\", "/") if docx_path.is_relative_to(ROOT) else str(docx_path),
            "bytes": docx_path.stat().st_size,
            "sha256": sha256_file(docx_path),
            "run_dir": str(template_run_dir.relative_to(ROOT)).replace("\\", "/"),
        }
    )

    fmt, fmt_md = extract_format(str(docx_path))
    write_json(template_run_dir / "format.json", fmt)
    (template_run_dir / "format_report.md").write_text(fmt_md, encoding="utf-8")
    profile = write_profile(fmt, str(template_run_dir), project_root=str(ROOT))
    result.update(
        {
            "format": {
                "paragraphs": len(fmt.get("paragraphs") or []),
                "tables": len(fmt.get("tables") or []),
                "sections": len(fmt.get("sections") or []),
                "cover_elements": len(fmt.get("cover") or []),
                "style_profiles": len(fmt.get("style_profiles") or {}),
            },
            "profile_risks": [k for k, v in (profile.get("risk_flags") or {}).items() if v],
            "scenarios": [],
        }
    )

    selected = [scenario for scenario in SCENARIOS if scenario_filter in (None, "", scenario["id"])]
    if not selected:
        raise RuntimeError(f"unknown scenario: {scenario_filter}")

    for scenario in selected:
        scenario_result = run_scenario(
            run_id,
            fmt,
            scenario,
            template_run_dir,
            visual=visual,
            golden_dir=golden_dir,
            update_golden=update_golden,
        )
        result["scenarios"].append(scenario_result)
        status = "PASS" if scenario_result.get("qa_passed") else "FAIL"
        visual_text = ""
        if visual:
            visual_text = f" visual={scenario_result.get('visual_issues', 'n/a')}"
        print(
            f"  {status} {run_id}/{scenario_result.get('id')} "
            f"issues={scenario_result.get('qa_issues', 'n/a')} "
            f"conformance={scenario_result.get('conformance_issues', 'n/a')}{visual_text}"
        )

    scenarios = result["scenarios"]
    result.update(
        {
            "qa_passed": bool(scenarios) and all(bool(s.get("qa_passed")) for s in scenarios),
            "qa_errors": sum(int(s.get("qa_errors") or 0) for s in scenarios),
            "qa_warnings": sum(int(s.get("qa_warnings") or 0) for s in scenarios),
            "qa_issues": sum(int(s.get("qa_issues") or 0) for s in scenarios),
            "conformance_errors": sum(int(s.get("conformance_errors") or 0) for s in scenarios),
            "conformance_warnings": sum(int(s.get("conformance_warnings") or 0) for s in scenarios),
            "conformance_issues": sum(int(s.get("conformance_issues") or 0) for s in scenarios),
            "visual_errors": sum(int(s.get("visual_errors") or 0) for s in scenarios),
            "visual_warnings": sum(int(s.get("visual_warnings") or 0) for s in scenarios),
            "visual_issues": sum(int(s.get("visual_issues") or 0) for s in scenarios),
        }
    )
    return result


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


def main() -> None:
    parser = argparse.ArgumentParser(description="Run public template compatibility checks.")
    parser.add_argument("--download", action="store_true", help="Download missing public templates before running.")
    parser.add_argument("--force-download", action="store_true", help="Re-download all templates.")
    parser.add_argument("--template", help="Run only one template id.")
    parser.add_argument("--scenario", help="Run only one scenario id.")
    parser.add_argument("--visual", action="store_true", help="Also run PDF/PNG visual QA for each generated DOCX.")
    parser.add_argument("--golden-dir", default=str(ROOT / "TestData" / "GoldenBaselines"), help="Directory for visual golden baseline JSON records.")
    parser.add_argument("--update-golden", action="store_true", help="Create or refresh visual golden baseline JSON records during visual QA.")
    args = parser.parse_args()

    TEST_ROOT.mkdir(parents=True, exist_ok=True)
    golden_dir = Path(args.golden_dir) if args.golden_dir else None
    manifest = read_manifest()
    templates = manifest.get("templates") or DEFAULT_TEMPLATES
    filtered_run = bool(args.template or args.scenario)
    if args.template:
        templates = [item for item in templates if str(item.get("id")) == args.template]
        if not templates:
            raise SystemExit(f"unknown template: {args.template}")

    results: List[Dict[str, Any]] = []
    failed = 0
    for item in templates:
        try:
            if args.download or args.force_download:
                result = run_template(
                    item,
                    force_download=args.force_download,
                    scenario_filter=args.scenario,
                    visual=args.visual,
                    golden_dir=golden_dir,
                    update_golden=args.update_golden,
                )
            else:
                expected = resolve_existing_file(item) or FILES_DIR / (item.get("name") or f"{item.get('id')}.docx")
                if not expected.exists():
                    raise RuntimeError(f"template file missing; rerun with --download: {expected}")
                result = run_template(
                    item,
                    force_download=False,
                    scenario_filter=args.scenario,
                    visual=args.visual,
                    golden_dir=golden_dir,
                    update_golden=args.update_golden,
                )
        except Exception as exc:
            result = {k: item.get(k) for k in ("id", "name", "source", "url", "download_url", "license_hint")}
            result.update({"qa_passed": False, "error": str(exc), "scenarios": []})
        if not result.get("qa_passed"):
            failed += 1
        status = "PASS" if result.get("qa_passed") else "FAIL"
        print(
            f"{status} {result.get('id')} scenarios={len(result.get('scenarios') or [])} "
            f"errors={result.get('qa_errors', 'n/a')} warnings={result.get('qa_warnings', 'n/a')}"
        )
        results.append(result)

    summary = {
        "schema_version": 1,
        "purpose": "Public non-private Word template regression corpus for local compatibility testing.",
        "files_ignored_by_git": True,
        "visual_qa_enabled": bool(args.visual),
        "golden_dir": str(golden_dir.relative_to(ROOT)).replace("\\", "/") if golden_dir and golden_dir.is_relative_to(ROOT) else (str(golden_dir) if golden_dir else None),
        "update_golden": bool(args.update_golden),
        "scenarios": [{k: scenario[k] for k in ("id", "description")} for scenario in SCENARIOS],
        "templates": results,
    }
    if filtered_run:
        write_json(SELECTION_SUMMARY_PATH, summary)
        print(f"Selection summary written to {SELECTION_SUMMARY_PATH.relative_to(ROOT)}")
    else:
        write_json(MANIFEST_PATH, summary)
        write_readme(results)
    print(f"RESULT passed={len(results) - failed} failed={failed} scenario_count={len(SCENARIOS)}")
    raise SystemExit(1 if failed else 0)


if __name__ == "__main__":
    main()
