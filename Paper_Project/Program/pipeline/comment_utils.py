"""
comment_utils.py — Add Word comments to python-docx documents.

Operates at the OOXML level: injects w:commentRangeStart/End markers,
w:commentReference runs into the paragraph XML, and post-processes the
saved .docx zip to add word/comments.xml.

Usage in build_generated.py:
    from comment_utils import CommentCollector

    cc = CommentCollector()
    p = body("Some text")
    cc.add(p, "This needs review", author="导师")

    # ... build rest of doc ...

    doc.save('output.docx')
    cc.save('output.docx')  # injects comments into the saved docx
"""

from lxml import etree
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import datetime
import zipfile
import os
import shutil


W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'


class CommentCollector:
    """Collects comments during document building, injects them post-save."""

    def __init__(self):
        self._comments = []  # list of (id, author, text, date)
        self._next_id = 0

    def add(self, paragraph, text, author="AI"):
        """Add a comment to a paragraph. Returns comment ID.

        Must be called BEFORE doc.save() — injects OOXML markers into the paragraph.
        """
        cid = self._next_id
        self._next_id += 1

        # Build comment metadata
        date_str = datetime.datetime.now().strftime('%Y-%m-%dT%H:%M:%SZ')
        initials = ''.join(w[0].upper() for w in author.split() if w) or 'AI'
        self._comments.append({
            'id': cid, 'author': author, 'date': date_str,
            'initials': initials, 'text': text,
        })

        # Inject OOXML markers into the paragraph
        p_el = paragraph._element

        # Comment Range Start
        crs = OxmlElement('w:commentRangeStart')
        crs.set(qn('w:id'), str(cid))
        p_el.insert(0, crs)

        # Comment Range End
        cre = OxmlElement('w:commentRangeEnd')
        cre.set(qn('w:id'), str(cid))
        p_el.append(cre)

        # Comment Reference run (marker showing where comment attaches)
        ref_r = OxmlElement('w:r')
        ref_rPr = OxmlElement('w:rPr')
        rs = OxmlElement('w:rStyle')
        rs.set(qn('w:val'), 'CommentReference')
        ref_rPr.append(rs)
        ref_r.append(ref_rPr)
        cref = OxmlElement('w:commentReference')
        cref.set(qn('w:id'), str(cid))
        ref_r.append(cref)
        p_el.append(ref_r)

        return cid

    def save(self, docx_path):
        """Post-process a saved .docx to inject word/comments.xml.

        Must be called AFTER doc.save() with the same path.
        """
        if not self._comments:
            return

        # Build comments.xml
        parts = [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            f'<w:comments xmlns:w="{W}"'
            ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
            ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        ]
        for c in self._comments:
            parts.append(f'<w:comment w:id="{c["id"]}" w:author="{_xml_escape(c["author"])}" '
                         f'w:date="{c["date"]}" w:initials="{c["initials"]}">')
            parts.append(f'<w:p><w:r><w:rPr><w:rFonts w:ascii="Microsoft YaHei" w:eastAsia="微软雅黑"/>'
                         f'<w:sz w:val="18"/></w:rPr><w:t xml:space="preserve">{_xml_escape(c["text"])}</w:t>'
                         f'</w:r></w:p>')
            parts.append('</w:comment>')
        parts.append('</w:comments>')
        comments_xml = '\n'.join(parts)

        # Inject into the saved docx zip
        self._inject(docx_path, comments_xml)

    def _inject(self, docx_path, comments_xml):
        """Open docx as zip, inject comments.xml, update rels and content types."""
        tmp = docx_path + '.tmp'
        with zipfile.ZipFile(docx_path, 'r') as zin:
            with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    if item.filename == 'word/comments.xml':
                        continue
                    data = zin.read(item.filename)

                    if item.filename == '[Content_Types].xml':
                        text = data.decode('utf-8')
                        if 'comments+xml' not in text:
                            text = text.replace(
                                '</Types>',
                                '<Override PartName="/word/comments.xml" '
                                'ContentType="application/vnd.openxmlformats-officedocument.'
                                'wordprocessingml.comments+xml"/></Types>')
                        data = text.encode('utf-8')

                    elif item.filename == 'word/_rels/document.xml.rels':
                        text = data.decode('utf-8')
                        if 'comments.xml' not in text:
                            text = text.replace(
                                '</Relationships>',
                                '<Relationship Id="rIdComments" '
                                'Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                                'relationships/comments" Target="comments.xml"/>'
                                '</Relationships>')
                        data = text.encode('utf-8')

                    zout.writestr(item, data)

                # Add comments.xml
                zout.writestr('word/comments.xml', comments_xml.encode('utf-8'))

        # Atomic replace
        shutil.move(tmp, docx_path)


def _xml_escape(s):
    """Escape XML special characters."""
    return (s.replace('&', '&amp;')
            .replace('<', '&lt;')
            .replace('>', '&gt;')
            .replace('"', '&quot;')
            .replace("'", '&apos;'))
