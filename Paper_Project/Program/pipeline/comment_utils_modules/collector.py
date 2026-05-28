"""OOXML-level Word comment injection."""
from __future__ import annotations

import datetime
import shutil
import zipfile

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
COMMENT_FONT = "\u5fae\u8f6f\u96c5\u9ed1"


class CommentCollector:
    """Collect comments during document building and inject them post-save."""

    def __init__(self):
        self._comments = []
        self._next_id = 0

    def add(self, paragraph, text, author="AI"):
        """Add a comment to a paragraph and return the comment id."""
        cid = self._next_id
        self._next_id += 1

        date_str = datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ")
        initials = "".join(w[0].upper() for w in author.split() if w) or "AI"
        self._comments.append(
            {
                "id": cid,
                "author": author,
                "date": date_str,
                "initials": initials,
                "text": text,
            }
        )

        p_el = paragraph._element
        crs = OxmlElement("w:commentRangeStart")
        crs.set(qn("w:id"), str(cid))
        p_el.insert(0, crs)

        cre = OxmlElement("w:commentRangeEnd")
        cre.set(qn("w:id"), str(cid))
        p_el.append(cre)

        ref_r = OxmlElement("w:r")
        ref_rPr = OxmlElement("w:rPr")
        rs = OxmlElement("w:rStyle")
        rs.set(qn("w:val"), "CommentReference")
        ref_rPr.append(rs)
        ref_r.append(ref_rPr)
        cref = OxmlElement("w:commentReference")
        cref.set(qn("w:id"), str(cid))
        ref_r.append(cref)
        p_el.append(ref_r)

        return cid

    def save(self, docx_path):
        """Post-process a saved .docx to inject word/comments.xml."""
        if not self._comments:
            return

        parts = [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            f'<w:comments xmlns:w="{W}"'
            ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
            ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        ]
        for comment in self._comments:
            parts.append(
                f'<w:comment w:id="{comment["id"]}" w:author="{_xml_escape(comment["author"])}" '
                f'w:date="{comment["date"]}" w:initials="{_xml_escape(comment["initials"])}">'
            )
            parts.append(
                f'<w:p><w:r><w:rPr><w:rFonts w:ascii="Microsoft YaHei" w:eastAsia="{COMMENT_FONT}"/>'
                f'<w:sz w:val="18"/></w:rPr><w:t xml:space="preserve">{_xml_escape(comment["text"])}</w:t>'
                f"</w:r></w:p>"
            )
            parts.append("</w:comment>")
        parts.append("</w:comments>")
        self._inject(docx_path, "\n".join(parts))

    def _inject(self, docx_path, comments_xml):
        """Open a DOCX zip, inject comments.xml, and update package metadata."""
        tmp = docx_path + ".tmp"
        with zipfile.ZipFile(docx_path, "r") as zin:
            with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    if item.filename == "word/comments.xml":
                        continue
                    data = zin.read(item.filename)

                    if item.filename == "[Content_Types].xml":
                        text = data.decode("utf-8")
                        if "comments+xml" not in text:
                            text = text.replace(
                                "</Types>",
                                '<Override PartName="/word/comments.xml" '
                                'ContentType="application/vnd.openxmlformats-officedocument.'
                                'wordprocessingml.comments+xml"/></Types>',
                            )
                        data = text.encode("utf-8")

                    elif item.filename == "word/_rels/document.xml.rels":
                        text = data.decode("utf-8")
                        if "comments.xml" not in text:
                            text = text.replace(
                                "</Relationships>",
                                '<Relationship Id="rIdComments" '
                                'Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                                'relationships/comments" Target="comments.xml"/>'
                                "</Relationships>",
                            )
                        data = text.encode("utf-8")

                    zout.writestr(item, data)

                zout.writestr("word/comments.xml", comments_xml.encode("utf-8"))

        shutil.move(tmp, docx_path)


def _xml_escape(s):
    """Escape XML special characters."""
    return (
        str(s)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )
