"""
sync_tracker.py — OOXML bookmark injection for SyncTeX-style source tracking.

Usage in build scripts:
    from pipeline.sync_tracker import SyncTracker
    st = SyncTracker(doc)
    body = st.track(body)
    heading1 = st.track(heading1)
    heading2 = st.track(heading2)
    heading3 = st.track(heading3)
    insert_figure = st.track_multi(insert_figure, count=2)

Each call to a tracked function automatically inserts a hidden bookmark
_src_L{line} into the created paragraph, linking the DOCX element back
to the source line that generated it.
"""
import sys
import os
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

WML = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


class SyncTracker:
    def __init__(self, doc, build_script_path=""):
        self.doc = doc
        self.build_script = build_script_path
        self._bookmark_id = 0
        self._enabled = os.environ.get("DOCX_SYNC_DISABLE", "0") != "1"

    def _add_bookmark(self, paragraph, source_line):
        if not self._enabled or paragraph is None:
            return
        bm_id = self._bookmark_id
        self._bookmark_id += 1
        name = f"_src_L{source_line}"

        bk_start = OxmlElement("w:bookmarkStart")
        bk_start.set(qn("w:id"), str(bm_id))
        bk_start.set(qn("w:name"), name)
        paragraph._element.insert(0, bk_start)

        bk_end = OxmlElement("w:bookmarkEnd")
        bk_end.set(qn("w:id"), str(bm_id))
        paragraph._element.append(bk_end)

    def track(self, func):
        """Return a wrapper that injects _src_L{line} bookmarks into
        paragraphs created by func, capturing the caller's source line."""
        def wrapper(*args, **kwargs):
            source_line = sys._getframe(1).f_lineno
            result = func(*args, **kwargs)
            if result is not None:
                self._add_bookmark(result, source_line)
            return result
        return wrapper

    def track_multi(self, func, count=2):
        """Like track() but maps all returned paragraphs (up to count)
        to the same source line. Useful for insert_figure etc."""
        def wrapper(*args, **kwargs):
            source_line = sys._getframe(1).f_lineno
            result = func(*args, **kwargs)
            if result is not None and isinstance(result, (list, tuple)):
                for p in result[:count]:
                    self._add_bookmark(p, source_line)
            return result
        return wrapper
