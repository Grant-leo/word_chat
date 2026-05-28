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
    import sys

    path = sys.argv[1] if len(sys.argv) > 1 else "Templates/模板.docx"
    fmt, md = extract(path)
    source = Path(path)
    json_path = str(source.with_name(f"{source.stem}_format.json"))
    with open(json_path, "w", encoding="utf-8") as handle:
        json.dump(fmt, handle, ensure_ascii=False, indent=2)
    md_path = str(source.with_name(f"{source.stem}_格式提取.md"))
    with open(md_path, "w", encoding="utf-8") as handle:
        handle.write(md)
    print(f"Format JSON -> {json_path}")
    print(f"Format MD   -> {md_path}")
    print(f'Paragraphs: {len(fmt["paragraphs"])}  Tables: {len(fmt["tables"])}  Sections: {len(fmt["sections"])}')
