"""Body element dispatch for content_parser.py.

This module owns the central DOCX body traversal.  The public parser keeps
front-matter, output directory, and metadata orchestration; body_dispatcher
decides how paragraph/table OOXML becomes sections, references, images, tables,
code blocks, captions, and formulas.
"""
from __future__ import annotations

from dataclasses import dataclass
import re
from typing import Any, Dict, List

try:
    from formula_semantics import is_formula_problem_text
except ImportError:  # pragma: no cover - package-style imports
    from ..formula_semantics import is_formula_problem_text

try:
    from content_parser_modules.caption_flow import (
        is_figure_caption,
        is_table_caption,
        normalize_caption_spacing,
    )
    from content_parser_modules.formula_extractor import (
        _formula_item_from_text,
        _formula_problem_item_from_text,
        _looks_like_formula_text,
        _omml_text_looks_like_body,
        _rich_text_item_from_inline_formula_spans,
        set_clean_text_artifacts_func,
    )
    from content_parser_modules.heading_detector import (
        classify_section_role,
        detect_heading_level,
        is_backmatter_heading,
        normalize_heading_spacing,
    )
    from content_parser_modules.image_extractor import image_items_from_ooxml
    from content_parser_modules.paragraph_stream import (
        append_stream_run_group,
        paragraph_stream_items,
    )
    from content_parser_modules.placeholders import is_unfilled_placeholder_text
    from content_parser_modules.placeholders import is_template_instruction_text
    from content_parser_modules.reference_collector import ReferenceCollector
    from content_parser_modules.section_builder import make_body_section
    from content_parser_modules.source_toc import (
        is_source_toc_title,
        source_toc_skip_count_after_title,
    )
    from content_parser_modules.table_extractor import (
        code_text_from_table_rows,
        extract_table_rows_from_ooxml,
        looks_like_code_line,
        table_rows_look_like_code,
    )
    from content_parser_modules.text_cleaner import clean_code_text, clean_text_artifacts
except ImportError:  # pragma: no cover - package-style imports
    from .caption_flow import (
        is_figure_caption,
        is_table_caption,
        normalize_caption_spacing,
    )
    from .formula_extractor import (
        _formula_item_from_text,
        _formula_problem_item_from_text,
        _looks_like_formula_text,
        _omml_text_looks_like_body,
        _rich_text_item_from_inline_formula_spans,
        set_clean_text_artifacts_func,
    )
    from .heading_detector import (
        classify_section_role,
        detect_heading_level,
        is_backmatter_heading,
        normalize_heading_spacing,
    )
    from .image_extractor import image_items_from_ooxml
    from .paragraph_stream import append_stream_run_group, paragraph_stream_items
    from .placeholders import is_unfilled_placeholder_text
    from .placeholders import is_template_instruction_text
    from .reference_collector import ReferenceCollector
    from .section_builder import make_body_section
    from .source_toc import is_source_toc_title, source_toc_skip_count_after_title
    from .table_extractor import (
        code_text_from_table_rows,
        extract_table_rows_from_ooxml,
        looks_like_code_line,
        table_rows_look_like_code,
    )
    from .text_cleaner import clean_code_text, clean_text_artifacts


set_clean_text_artifacts_func(clean_text_artifacts)


@dataclass
class BodyDispatchResult:
    sections: List[Dict[str, Any]]
    references: List[Any]
    source_toc_skipped: int = 0
    placeholders_removed: int = 0


def append_text_or_code(section: Dict[str, Any], text: str, in_appendix: bool = False) -> None:
    """Append semantic blocks while preserving captions, code and inline math."""
    if not text:
        return
    text = clean_text_artifacts(text)
    if not text:
        return
    if is_unfilled_placeholder_text(text) or is_template_instruction_text(text):
        return
    if is_figure_caption(text):
        section["paragraphs"].append({"role": "figure_caption", "text": normalize_caption_spacing(text)})
    elif is_table_caption(text):
        section["paragraphs"].append({"role": "table_caption", "text": normalize_caption_spacing(text)})
    elif re.match(r"^\s*\$\$.+\$\$\s*(?:[\(\uff08]\s*\d+(?:\s*[-.]\s*\d+)?\s*[\)\uff09])?\s*$", text, re.S):
        section["paragraphs"].append(_formula_item_from_text(text))
    else:
        rich_item = _rich_text_item_from_inline_formula_spans(text)
        rich_has_text = bool(
            rich_item
            and any(
                r.get("type") == "text"
                and str(r.get("text") or "").strip(" \t\r\n，,。.;；:：()（）")
                for r in rich_item.get("runs") or []
            )
        )
        if rich_item and rich_has_text:
            section["paragraphs"].append(rich_item)
            return
        if is_formula_problem_text(text):
            section["paragraphs"].append(_formula_problem_item_from_text(text))
        elif _looks_like_formula_text(text):
            section["paragraphs"].append(_formula_item_from_text(text))
        elif in_appendix and (looks_like_code_line(text) or "\n" in text):
            section["paragraphs"].append({"role": "code", "code": clean_code_text(text)})
        elif rich_item:
            section["paragraphs"].append(rich_item)
        elif _omml_text_looks_like_body(text):
            section["paragraphs"].append(text)
        else:
            section["paragraphs"].append(text)


def _section_from_heading(text: str, level: int) -> Dict[str, Any]:
    clean_heading = text.split("（")[0].strip()
    if is_unfilled_placeholder_text(clean_heading) or is_template_instruction_text(text):
        return {}
    if not clean_heading:
        return {}
    m = re.match(r"(?i)^(Abstract\s*:?|Key\s*words?\s*:?|摘要\s*[：:]|关键词\s*[：:])\s*", text)
    if m:
        heading_part = m.group(1).strip()
        body_part = text[m.end() :].strip()
        body_part = re.sub(r"^[（(][^）)]*[）)]\s*", "", body_part)
    else:
        heading_part = clean_heading
        body_part = ""
    heading_part = normalize_heading_spacing(heading_part)
    role = classify_section_role(heading_part, level)
    section: Dict[str, Any] = {
        "heading": heading_part,
        "level": level,
        "role": role,
        "paragraphs": [],
        "images": [],
    }
    if role == "en_abstract":
        section["page_break_before"] = True
    if body_part:
        append_text_or_code(section, body_part, in_appendix=False)
    return section


def _append_paragraph_stream(section: Dict[str, Any], paragraph: Any, image_registry: Any) -> None:
    """Append text/image/math items from one Word paragraph in OOXML order."""
    text = paragraph.text.strip()
    stream_items = paragraph_stream_items(paragraph, image_registry)
    if not stream_items and text:
        stream_items = [{"role": "text", "text": text}]
    rich_runs: List[Dict[str, Any]] = []
    in_appendix = bool(re.search(r"(\u9644\s*\u5f55|\u914d\u7f6e|\u547d\u4ee4|\u4ee3\u7801)", section.get("heading", "")))

    def flush_rich_runs() -> None:
        nonlocal rich_runs
        append_stream_run_group(
            section,
            rich_runs,
            append_text_or_code_func=append_text_or_code,
            in_appendix=in_appendix,
        )
        rich_runs = []

    for item in stream_items:
        if item.get("role") == "image":
            flush_rich_runs()
            section["images"].append(item.get("image"))
            section["paragraphs"].append(item)
        elif item.get("role") == "math_inline":
            rich_runs.append({"type": "math", "text": item.get("text") or "", "math": item.get("math") or []})
        elif item.get("role") == "formula":
            flush_rich_runs()
            section["paragraphs"].append(item)
        elif item.get("role") == "text":
            txt = item.get("text") or ""
            if txt:
                rich_runs.append({"type": "text", "text": txt})
    flush_rich_runs()


def parse_body_sections(doc: Any, text_start: int, image_registry: Any) -> BodyDispatchResult:
    """Parse DOCX body children after front matter into content sections."""
    current_section = make_body_section()
    sections = [current_section]
    ref_collector = ReferenceCollector(
        clean_text_func=clean_text_artifacts,
        is_backmatter_heading_func=is_backmatter_heading,
        normalize_heading_spacing_func=normalize_heading_spacing,
        classify_section_role_func=classify_section_role,
        table_rows_look_like_code_func=table_rows_look_like_code,
        code_text_from_table_rows_func=code_text_from_table_rows,
        clean_code_func=clean_code_text,
    )

    body_children = list(doc.element.body)
    p_idx = 0
    source_toc_skip_remaining = 0
    source_toc_skipped = 0
    placeholders_removed = 0

    for child in body_children:
        tag = child.tag.split("}")[-1]

        if tag == "p":
            if p_idx < text_start:
                p_idx += 1
                continue
            paragraph = doc.paragraphs[p_idx]
            p_idx += 1
            text = paragraph.text.strip()
            level = detect_heading_level(paragraph)

            if is_unfilled_placeholder_text(text) or is_template_instruction_text(text):
                placeholders_removed += 1
                continue

            if source_toc_skip_remaining > 0:
                source_toc_skip_remaining -= 1
                source_toc_skipped += 1
                continue

            if is_source_toc_title(text):
                skip_after_title = source_toc_skip_count_after_title(doc.paragraphs, p_idx - 1)
                if skip_after_title:
                    source_toc_skip_remaining = skip_after_title
                    source_toc_skipped += 1
                    continue

            if ref_collector.start_if_heading(text):
                continue

            backmatter_section = ref_collector.exit_to_backmatter_section(text, level)
            if backmatter_section is not None:
                current_section = backmatter_section
                sections.append(current_section)
                continue

            if ref_collector.consume_text(text):
                continue
            if level > 0:
                section = _section_from_heading(text, level)
                if section:
                    current_section = section
                    sections.append(current_section)
            else:
                _append_paragraph_stream(current_section, paragraph, image_registry)

        elif tag == "tbl":
            if p_idx < text_start:
                continue
            rows = extract_table_rows_from_ooxml(child, clean_text_func=clean_text_artifacts)
            table_images = image_items_from_ooxml(child, doc.part.rels, image_registry, location="table_cell")
            if rows:
                if ref_collector.consume_table_rows(rows):
                    pass
                elif table_rows_look_like_code(rows):
                    current_section["paragraphs"].append(
                        {
                            "role": "code",
                            "code": code_text_from_table_rows(rows, clean_code_func=clean_code_text),
                            "table_rows": rows,
                        }
                    )
                else:
                    current_section["paragraphs"].append({"role": "table", "table_rows": rows})
            if table_images and not ref_collector.active:
                for image_item in table_images:
                    current_section["images"].append(image_item.get("image"))
                    current_section["paragraphs"].append(image_item)

    return BodyDispatchResult(
        sections=sections,
        references=ref_collector.finish(),
        source_toc_skipped=source_toc_skipped,
        placeholders_removed=placeholders_removed,
    )
