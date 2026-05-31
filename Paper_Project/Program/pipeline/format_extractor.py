"""Stable template format extraction entry point."""
from __future__ import annotations

import json
from pathlib import Path

try:
    from format_extractor_modules.extractor import extract
except ImportError:  # pragma: no cover - package-style imports
    from .format_extractor_modules.extractor import extract

__all__ = ["extract"]


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="模板格式提取工具（建议优先使用 run_pipeline.py）。")
    parser.add_argument("template_path", nargs="?", default="Templates/模板.docx", help="模板 DOCX/PDF 路径。")
    parser.add_argument(
        "--output-dir",
        default=str(Path("Outputs") / "_format_extractor_cli"),
        help="输出目录，默认 Outputs/_format_extractor_cli，避免写入 Templates/。",
    )
    args = parser.parse_args()

    path = args.template_path
    out_dir = Path(args.output_dir).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    fmt, md = extract(path, output_dir=str(out_dir))
    source = Path(path)
    json_path = str(out_dir / f"{source.stem}_format.json")
    with open(json_path, "w", encoding="utf-8") as handle:
        json.dump(fmt, handle, ensure_ascii=False, indent=2)
    md_path = str(out_dir / f"{source.stem}_格式提取.md")
    with open(md_path, "w", encoding="utf-8") as handle:
        handle.write(md)
    print(f"格式 JSON -> {json_path}")
    print(f"格式报告 -> {md_path}")
    print(f'段落: {len(fmt["paragraphs"])}  表格: {len(fmt["tables"])}  节: {len(fmt["sections"])}')
