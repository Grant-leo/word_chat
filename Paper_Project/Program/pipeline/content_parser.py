"""Stable DOCX content extraction entry point."""
from __future__ import annotations

import json
import os

try:
    from content_parser_modules.extractor import _content_placeholder_samples, _count_structured_body_tables, extract
except ImportError:  # pragma: no cover - package-style imports
    from .content_parser_modules.extractor import _content_placeholder_samples, _count_structured_body_tables, extract

__all__ = ["_content_placeholder_samples", "_count_structured_body_tables", "extract"]

if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser(description="DOCX 正文内容提取工具（建议优先使用 run_pipeline.py）。")
    parser.add_argument("docx_path", help="要提取的正文 DOCX 路径。")
    parser.add_argument(
        "--output-dir",
        default=os.path.join("Outputs", "_content_parser_cli"),
        help="输出目录，默认 Outputs/_content_parser_cli，避免写入 Inputs/。",
    )
    args = parser.parse_args()

    out_dir = os.path.abspath(args.output_dir)
    os.makedirs(out_dir, exist_ok=True)
    content = extract(args.docx_path, output_dir=out_dir)
    stem = os.path.splitext(os.path.basename(args.docx_path))[0]
    json_path = os.path.join(out_dir, f"{stem}_content.json")
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(content, f, ensure_ascii=False, indent=2)
    print(f'内容 JSON -> {json_path}')
    print(f'章节: {len(content["sections"])}  参考文献: {len(content["references"])}  图片: {content["_meta"]["images_extracted"]}')
