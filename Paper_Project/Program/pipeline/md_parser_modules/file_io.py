"""Markdown file reading helpers."""
from __future__ import annotations

from pathlib import Path


def read_markdown_text(md_path: str) -> str:
    data = Path(md_path).read_bytes()
    last_error: UnicodeDecodeError | None = None
    for encoding in ("utf-8-sig", "gb18030"):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError as exc:
            last_error = exc
    if last_error is not None:
        raise last_error
    return ""
