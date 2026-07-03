"""Compatibility exports for structural QA metric helpers."""
from __future__ import annotations

try:
    from qa_checker_modules.content_metrics import (
        _content_text_chars,
        _count_content_formulas,
        _count_content_images,
        _count_content_note_refs,
        _count_content_tables,
        _iter_content_image_refs,
        _iter_paragraph_items,
    )
    from qa_checker_modules.content_samples import (
        _content_toc_pollution_samples,
        _formula_number_conflict_samples,
        _fragmented_formula_samples,
        _low_res_image_fragment_samples,
        _placeholder_samples_from_texts,
    )
    from qa_checker_modules.docx_metrics import (
        _duplicate_front_matter_headings,
        _missing_heading_samples,
        _read_docx_xml,
        _xml_paragraph_texts,
        _xml_plain_text,
    )
    from qa_checker_modules.json_io import _load_json, _load_manifest_counts
except ImportError:  # pragma: no cover - package-style imports
    from .content_metrics import (
        _content_text_chars,
        _count_content_formulas,
        _count_content_images,
        _count_content_note_refs,
        _count_content_tables,
        _iter_content_image_refs,
        _iter_paragraph_items,
    )
    from .content_samples import (
        _content_toc_pollution_samples,
        _formula_number_conflict_samples,
        _fragmented_formula_samples,
        _low_res_image_fragment_samples,
        _placeholder_samples_from_texts,
    )
    from .docx_metrics import (
        _duplicate_front_matter_headings,
        _missing_heading_samples,
        _read_docx_xml,
        _xml_paragraph_texts,
        _xml_plain_text,
    )
    from .json_io import _load_json, _load_manifest_counts

__all__ = [
    "_content_text_chars",
    "_content_toc_pollution_samples",
    "_count_content_formulas",
    "_count_content_images",
    "_count_content_note_refs",
    "_count_content_tables",
    "_duplicate_front_matter_headings",
    "_formula_number_conflict_samples",
    "_fragmented_formula_samples",
    "_iter_content_image_refs",
    "_iter_paragraph_items",
    "_load_json",
    "_load_manifest_counts",
    "_low_res_image_fragment_samples",
    "_missing_heading_samples",
    "_placeholder_samples_from_texts",
    "_read_docx_xml",
    "_xml_paragraph_texts",
    "_xml_plain_text",
]

