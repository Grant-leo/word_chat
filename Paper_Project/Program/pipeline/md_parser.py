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
    import sys
    path = sys.argv[1] if len(sys.argv) > 1 else 'Inputs/test.md'

    # Test format extraction
    fmt, md_text = extract_format(path)
    with open(path.replace('.md', '_format.json'), 'w', encoding='utf-8') as f:
        json.dump(fmt, f, ensure_ascii=False, indent=2)
    print(f'Format JSON → {path.replace(".md", "_format.json")}')

    # Test content extraction
    cnt = extract_content(path)
    with open(path.replace('.md', '_content.json'), 'w', encoding='utf-8') as f:
        json.dump(cnt, f, ensure_ascii=False, indent=2)
    print(f'Content JSON → {path.replace(".md", "_content.json")}')
