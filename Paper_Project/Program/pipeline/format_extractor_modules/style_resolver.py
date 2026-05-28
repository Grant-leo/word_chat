"""Style inheritance resolver for python-docx templates."""
from __future__ import annotations

from docx.oxml.ns import qn

from .ooxml import ALIGN_MAP, pt, tag, twips_to_cm, val


class StyleResolver:
    """Resolve formatting by walking the style inheritance tree in styles.xml."""

    def __init__(self, doc):
        self.styles = {}
        self._load_styles(doc)

    def _load_styles(self, doc):
        for style in doc.styles:
            sid = style.style_id
            entry = {}
            font = getattr(style, "font", None)
            if font is not None:
                if font.name:
                    entry["font"] = font.name
                if font.size:
                    entry["size"] = font.size.pt
                if font.bold is not None:
                    entry["bold"] = font.bold
                if font.italic is not None:
                    entry["italic"] = font.italic
            try:
                pf = style.paragraph_format
                if pf.line_spacing:
                    entry["ls"] = pf.line_spacing
                if pf.alignment is not None:
                    entry["align"] = ALIGN_MAP.get(pf.alignment, "DEFAULT")
                if pf.first_line_indent:
                    entry["indent"] = pf.first_line_indent.cm
                if pf.left_indent:
                    entry["left_indent_cm"] = pf.left_indent.cm
                if pf.right_indent:
                    entry["right_indent_cm"] = pf.right_indent.cm
            except AttributeError:
                pass
            try:
                pPr = style._element.find(qn("w:pPr"))
                ind = pPr.find(qn("w:ind")) if pPr is not None else None
                if ind is not None:
                    left = ind.get(qn("w:left"))
                    right = ind.get(qn("w:right"))
                    first = ind.get(qn("w:firstLine"))
                    hanging = ind.get(qn("w:hanging"))
                    if left is not None:
                        entry["left_indent_cm"] = twips_to_cm(left)
                    if right is not None:
                        entry["right_indent_cm"] = twips_to_cm(right)
                    if first is not None:
                        entry["indent"] = twips_to_cm(first) or 0
                    if hanging is not None:
                        entry["hanging_indent_cm"] = twips_to_cm(hanging)
            except Exception:
                pass
            try:
                entry["base"] = style.base_style.style_id if style.base_style else None
            except (AttributeError, ValueError):
                entry["base"] = None
            self.styles[sid] = entry

    def resolve(self, p_elem, r_elem):
        """Return fully resolved formatting for a run."""
        result = {
            "font": None,
            "size": None,
            "bold": False,
            "italic": False,
            "ls": None,
            "align": "DEFAULT",
            "indent": None,
            "left_indent_cm": None,
            "right_indent_cm": None,
            "hanging_indent_cm": None,
        }

        rPr = r_elem.find(qn("w:rPr"))
        if rPr is not None:
            for child in rPr:
                child_tag = tag(child)
                if child_tag == "rFonts":
                    result["font"] = val(child, "w:ascii") or val(child, "w:hAnsi")
                elif child_tag == "sz":
                    result["size"] = pt(val(child))
                elif child_tag == "b":
                    result["bold"] = True
                elif child_tag == "i":
                    result["italic"] = True
                elif child_tag == "color":
                    result["color"] = val(child)

        pPr = p_elem.find(qn("w:pPr"))
        style_id = None
        if pPr is not None:
            for child in pPr:
                child_tag = tag(child)
                if child_tag == "pStyle":
                    style_id = val(child)
                elif child_tag == "jc":
                    value = val(child)
                    align_map = {"left": "LEFT", "center": "CENTER", "right": "RIGHT", "both": "JUSTIFY"}
                    result["align"] = align_map.get(value, "DEFAULT")
                elif child_tag == "spacing":
                    line = child.get(qn("w:line"))
                    if line:
                        result["ls"] = int(line) / 240.0
                elif child_tag == "ind":
                    first = child.get(qn("w:firstLine"))
                    if first:
                        result["indent"] = round(int(first) / 567.0, 1)
                    left = child.get(qn("w:left"))
                    right = child.get(qn("w:right"))
                    hanging = child.get(qn("w:hanging"))
                    if left is not None:
                        result["left_indent_cm"] = twips_to_cm(left)
                    if right is not None:
                        result["right_indent_cm"] = twips_to_cm(right)
                    if hanging is not None:
                        result["hanging_indent_cm"] = twips_to_cm(hanging)

        self._resolve_style(style_id, result)
        return result

    def _resolve_style(self, sid, result):
        """Walk the style inheritance chain to fill missing values."""
        visited = set()
        while sid and sid not in visited:
            visited.add(sid)
            style = self.styles.get(sid, {})
            if result["font"] is None:
                result["font"] = style.get("font")
            if result["size"] is None:
                result["size"] = style.get("size")
            if result["align"] == "DEFAULT":
                result["align"] = style.get("align", "DEFAULT")
            if result["ls"] is None:
                result["ls"] = style.get("ls")
            if result.get("indent") is None:
                result["indent"] = style.get("indent")
            if result.get("left_indent_cm") is None:
                result["left_indent_cm"] = style.get("left_indent_cm")
            if result.get("right_indent_cm") is None:
                result["right_indent_cm"] = style.get("right_indent_cm")
            if result.get("hanging_indent_cm") is None:
                result["hanging_indent_cm"] = style.get("hanging_indent_cm")
            sid = style.get("base")
        if result["font"] is None:
            result["font"] = "Times New Roman"
        if result["size"] is None:
            result["size"] = 12.0
        if result["ls"] is None:
            result["ls"] = 1.15
        if result["indent"] is None:
            result["indent"] = 0
        if result["align"] == "DEFAULT":
            result["align"] = "LEFT"

