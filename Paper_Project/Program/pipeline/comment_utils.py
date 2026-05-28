"""Compatibility entry for generated Word comment support.

Generated scripts import ``CommentCollector`` from this module. The OOXML
implementation lives in ``comment_utils_modules.collector`` so the top-level
pipeline helper stays small and easy to audit.
"""
from __future__ import annotations

from comment_utils_modules.collector import COMMENT_FONT, W, CommentCollector, _xml_escape

__all__ = ["COMMENT_FONT", "W", "CommentCollector", "_xml_escape"]
