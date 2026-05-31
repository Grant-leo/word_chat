"""Stable Markdown parser entry point for format.json and content.json."""
from __future__ import annotations

import json

try:
    from md_parser_modules.content_extractor import extract_content
    from md_parser_modules.content_helpers import (
        _RE_BACKMATTER_HEADING,
        _RE_REF_HEADING,
        _detect_title,
        _is_format_section_heading,
        _is_markdown_table_separator,
        _latex_escape_text,
        _latex_from_formula_text,
        _looks_like_formula_text,
        _parse_markdown_table,
        _parse_paragraph_items,
        _parse_text_paragraph,
        _process_inline_math,
        _skip_format_section,
        _split_image_tokens_from_text,
        _split_markdown_table_row,
        _strip_md_formatting,
    )
    from md_parser_modules.format_extractor import (
        DEFAULT_FORMAT_TEXT,
        DEFAULT_PAGE,
        _build_format_dict,
        _find_format_section,
        _parse_page_geometry,
        _parse_yaml_frontmatter,
        extract_format,
    )
except ImportError:  # pragma: no cover - package-style imports
    from .md_parser_modules.content_extractor import extract_content
    from .md_parser_modules.content_helpers import (
        _RE_BACKMATTER_HEADING,
        _RE_REF_HEADING,
        _detect_title,
        _is_format_section_heading,
        _is_markdown_table_separator,
        _latex_escape_text,
        _latex_from_formula_text,
        _looks_like_formula_text,
        _parse_markdown_table,
        _parse_paragraph_items,
        _parse_text_paragraph,
        _process_inline_math,
        _skip_format_section,
        _split_image_tokens_from_text,
        _split_markdown_table_row,
        _strip_md_formatting,
    )
    from .md_parser_modules.format_extractor import (
        DEFAULT_FORMAT_TEXT,
        DEFAULT_PAGE,
        _build_format_dict,
        _find_format_section,
        _parse_page_geometry,
        _parse_yaml_frontmatter,
        extract_format,
    )

__all__ = ["extract_content", "extract_format"]

if __name__ == '__main__':
    import argparse
    from pathlib import Path

    parser = argparse.ArgumentParser(description="Markdown 格式/内容提取工具（建议优先使用 run_pipeline.py）。")
    parser.add_argument("md_path", nargs="?", default="Inputs/test.md", help="Markdown 文件路径。")
    parser.add_argument(
        "--output-dir",
        default=str(Path("Outputs") / "_md_parser_cli"),
        help="输出目录，默认 Outputs/_md_parser_cli，避免写入 Inputs/。",
    )
    args = parser.parse_args()
    path = args.md_path
    out_dir = Path(args.output_dir).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    stem = Path(path).stem

    # Test format extraction
    fmt, md_text = extract_format(path, output_dir=str(out_dir))
    format_json = out_dir / f'{stem}_format.json'
    with open(format_json, 'w', encoding='utf-8') as f:
        json.dump(fmt, f, ensure_ascii=False, indent=2)
    print(f'格式 JSON -> {format_json}')

    # Test content extraction
    cnt = extract_content(path, output_dir=str(out_dir))
    content_json = out_dir / f'{stem}_content.json'
    with open(content_json, 'w', encoding='utf-8') as f:
        json.dump(cnt, f, ensure_ascii=False, indent=2)
    print(f'内容 JSON -> {content_json}')
