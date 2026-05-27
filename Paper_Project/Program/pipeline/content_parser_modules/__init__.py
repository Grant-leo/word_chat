"""Submodules used by content_parser.py.

The public entry point remains ``content_parser.extract``.  These modules hold
cohesive parsing rules so the top-level parser can stay as an orchestration
layer instead of accumulating every DOCX heuristic in one file.

Current modules cover placeholders, style helpers, text cleanup, front matter,
caption flow, paragraph streams, source TOC filtering, image extraction, table
extraction, formula extraction/repair, heading detection, reference collection,
body dispatch, and section construction/post-processing.
"""
