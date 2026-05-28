"""Format handoff checks for structural QA."""
from __future__ import annotations

import os
from typing import Any, Callable, Dict

try:
    from qa_checker_modules.metrics import _load_json
except ImportError:  # pragma: no cover - package-style imports
    from .metrics import _load_json

AddIssue = Callable[..., None]
def run_format_checks(paths: Dict[str, str], counts: Dict[str, Any], add: AddIssue) -> Dict[str, Any]:
    fmt: Dict[str, Any] = {}
    if os.path.exists(paths["format"]):
        try:
            fmt = _load_json(paths["format"])
            counts["format_paragraphs"] = len(fmt.get("paragraphs") or [])
            counts["format_tables"] = len(fmt.get("tables") or [])
            counts["format_sections"] = len(fmt.get("sections") or [])
            counts["cover_elements"] = len(fmt.get("cover") or [])
            counts["style_profiles"] = len(fmt.get("style_profiles") or {})
            source = str((fmt.get("_meta") or {}).get("source") or "")
            is_md_format = source.lower().endswith(".md")
            if not fmt.get("sections"):
                add("FORMAT_EMPTY", "error", "格式提取结果没有 section。")
            if not fmt.get("paragraphs"):
                add("FORMAT_EMPTY", "warning", "格式提取结果没有 paragraph，可能大量使用默认格式。")
            expected = {"body", "h1", "h2", "h3"}
            missing = sorted(expected - set((fmt.get("style_profiles") or {}).keys()))
            if missing and not is_md_format:
                add("STYLE_PROFILE_MISSING", "warning", "关键样式 profile 不完整。", ", ".join(missing))
            if not fmt.get("cover") and not is_md_format:
                add("COVER_NOT_EXTRACTED", "warning", "没有提取到封面结构；纯 MD 或无封面模板时可以忽略。")
        except Exception as exc:
            add("MISSING_FORMAT_JSON", "error", "format.json 无法读取。", str(exc))
    return fmt

