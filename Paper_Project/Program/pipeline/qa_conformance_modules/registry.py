"""Shared constants for strict DOCX conformance QA."""
from __future__ import annotations


VALID_MODES = {"user", "developer"}

STYLE_ROLES = [
    "body", "h1", "h2", "h3",
    "cn_title", "en_title", "cn_abstract_heading", "cn_abstract_body",
    "en_abstract_heading", "en_abstract_body", "figure_caption",
    "table_caption", "table_body", "table_header", "formula",
    "reference", "reference_english",
]

TEXT_ONLY_ROLES = {"body", "h1", "h2", "h3", "figure_caption", "table_caption", "reference"}
BACKMATTER_ROLES = {"references", "acknowledgement", "appendix"}
FRONTMATTER_ROLES = {"cn_abstract", "cn_keywords", "en_abstract", "en_keywords"}

