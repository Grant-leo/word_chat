"""Synthetic non-private scenarios for public template compatibility checks."""
from __future__ import annotations

import struct
import zlib
from pathlib import Path
from typing import Any, Callable, Dict, List

from md_parser import extract_content as extract_md_content


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
