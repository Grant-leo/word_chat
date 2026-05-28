"""Cover, declaration, image, and cover-table extraction helpers."""
from __future__ import annotations

import os
import re
from io import BytesIO

try:
    from PIL import Image
except Exception:
    Image = None


def crop_blob_by_src_rect(blob, src_rect, ext):
    """Crop a DOCX image blob according to DrawingML a:srcRect."""
    if not src_rect or Image is None:
        return blob, ext
    try:
        vals = {key: int(src_rect.get(key, 0) or 0) for key in ("l", "t", "r", "b")}
    except Exception:
        return blob, ext
    if not any(vals.values()):
        return blob, ext
    try:
        img = Image.open(BytesIO(blob))
        width, height = img.size
        left = max(0, int(width * vals["l"] / 100000.0))
        top = max(0, int(height * vals["t"] / 100000.0))
        right = min(width, int(width * (1 - vals["r"] / 100000.0)))
        bottom = min(height, int(height * (1 - vals["b"] / 100000.0)))
        if right <= left or bottom <= top:
            return blob, ext
        out = BytesIO()
        img.crop((left, top, right, bottom)).save(out, format="PNG")
        return out.getvalue(), "png"
    except Exception:
        return blob, ext


def cover_table_role(rows_data):
    """Infer cover table role from structure instead of school-specific text."""

    def cell_text(cell):
        return "".join(run.get("t", "") for paragraph in cell.get("p", []) for run in paragraph.get("r", []))

    first_col = []
    for row in rows_data or []:
        if row:
            first_col.append(re.sub(r"[\s：:]+", "", cell_text(row[0])))
    if first_col and len(rows_data or []) <= 2 and all(text.endswith("编码") for text in first_col if text):
        return "cover_code_table"
    if len(rows_data or []) >= 3 and sum(1 for text in first_col if text.endswith(("题目", "姓名", "学号", "学院", "班级", "老师", "教师"))) >= 2:
        return "cover_info_table"
    return "cover_table"


def ooxml_attrs(el, ns):
    """Return OOXML attributes without namespace prefixes for JSON storage."""
    if el is None:
        return {}
    out = {}
    for key, value in el.attrib.items():
        clean_key = key.split("}")[-1] if "}" in key else key
        out[clean_key] = value
    return out


def extract_margin_box(parent, ns, child_name):
    box = parent.find(f"{{{ns}}}{child_name}") if parent is not None else None
    if box is None:
        return {}
    result = {}
    for side in ("top", "left", "bottom", "right", "start", "end"):
        el = box.find(f"{{{ns}}}{side}")
        if el is not None:
            result[side] = ooxml_attrs(el, ns)
    return result


def extract_tbl_props(tbl_elem):
    """Extract table-level layout properties so cover tables can be replayed."""
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    props = {
        "tblW": {},
        "tblInd": {},
        "jc": None,
        "tblLayout": None,
        "cellMar": {},
        "grid_cols": [],
        "rows": [],
    }
    tblPr = tbl_elem.find(f"{{{W}}}tblPr")
    if tblPr is not None:
        for name in ("tblW", "tblInd", "tblLayout"):
            el = tblPr.find(f"{{{W}}}{name}")
            if el is not None:
                props[name] = ooxml_attrs(el, W)
        jc = tblPr.find(f"{{{W}}}jc")
        if jc is not None:
            props["jc"] = jc.get(f"{{{W}}}val")
        props["cellMar"] = extract_margin_box(tblPr, W, "tblCellMar")
    grid = tbl_elem.find(f"{{{W}}}tblGrid")
    if grid is not None:
        for grid_col in grid.findall(f"{{{W}}}gridCol"):
            width = grid_col.get(f"{{{W}}}w")
            if width:
                props["grid_cols"].append(width)
    for tr in tbl_elem.findall(f"{{{W}}}tr"):
        row_props = {"height": {}, "cantSplit": False}
        trPr = tr.find(f"{{{W}}}trPr")
        if trPr is not None:
            height = trPr.find(f"{{{W}}}trHeight")
            if height is not None:
                row_props["height"] = ooxml_attrs(height, W)
            row_props["cantSplit"] = trPr.find(f"{{{W}}}cantSplit") is not None
        props["rows"].append(row_props)
    return props


def extract_tc_props(tcPr):
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    info = {"tcW": {}, "tcMar": {}, "vAlign": None, "gridSpan": None, "vMerge": None}
    if tcPr is None:
        return info
    tcW = tcPr.find(f"{{{W}}}tcW")
    if tcW is not None:
        info["tcW"] = ooxml_attrs(tcW, W)
    info["tcMar"] = extract_margin_box(tcPr, W, "tcMar")
    v_align = tcPr.find(f"{{{W}}}vAlign")
    if v_align is not None:
        info["vAlign"] = v_align.get(f"{{{W}}}val")
    grid_span = tcPr.find(f"{{{W}}}gridSpan")
    if grid_span is not None:
        info["gridSpan"] = grid_span.get(f"{{{W}}}val")
    v_merge = tcPr.find(f"{{{W}}}vMerge")
    if v_merge is not None:
        info["vMerge"] = v_merge.get(f"{{{W}}}val") or "continue"
    return info


def extract_cover(doc, assets_dir=None):
    """Walk template body from start and extract cover/declaration elements."""
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    A = "http://schemas.openxmlformats.org/drawingml/2006/main"
    R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"

    if assets_dir:
        os.makedirs(assets_dir, exist_ok=True)
    image_counter = 0

    def save_image_by_rid(rid, src_rect=None):
        nonlocal image_counter
        if not rid or rid not in doc.part.rels:
            return None
        rel = doc.part.rels[rid]
        if "image" not in rel.reltype:
            return None
        ext = rel.target_ref.rsplit(".", 1)[-1].lower() if "." in rel.target_ref else "png"
        if ext not in ("png", "jpg", "jpeg", "gif", "bmp", "tif", "tiff", "emf", "wmf"):
            ext = "png"
        blob, ext = crop_blob_by_src_rect(rel.target_part.blob, src_rect, ext)
        image_counter += 1
        filename = f"cover_img_{image_counter:03d}.{ext}"
        if assets_dir:
            with open(os.path.join(assets_dir, filename), "wb") as handle:
                handle.write(blob)
        return filename

    def image_payload(elem):
        extent = None
        src_rect = {}
        rid = None
        drawing_nodes = list(elem.iter(f"{{{WP}}}inline")) + list(elem.iter(f"{{{WP}}}anchor"))
        for drawing in drawing_nodes:
            ext = drawing.find(f"{{{WP}}}extent")
            if ext is not None:
                extent = {"cx": ext.get("cx", "0"), "cy": ext.get("cy", "0")}
            for src in drawing.iter(f"{{{A}}}srcRect"):
                src_rect = {"l": src.get("l", "0"), "t": src.get("t", "0"), "r": src.get("r", "0"), "b": src.get("b", "0")}
            for blip in drawing.iter(f"{{{A}}}blip"):
                rid = blip.get(f"{{{R}}}embed")
                if rid:
                    break
            if rid:
                break
        asset = save_image_by_rid(rid, src_rect) if rid else None
        if not asset:
            return None
        return {"asset": asset, "extent": extent, "srcRect": src_rect, "rEmbed": rid}

    stop_kw = ["摘 要", "摘要", "目  录", "目录", "目 录", "ABSTRACT", "第1章", "1.1 ", "1.1."]
    skip_kw = ["页边距要求", "碳素笔", "完成后删除", "封面要求", "1.论文题目", "毕业论文（设计）题目为", "按答辩时间", "提交论文", "禁止使用"]

    def front_stop_match(text):
        compact = re.sub(r"[\s\u3000:：]+", "", str(text or "")[:40]).upper()
        if compact in {"摘要", "中文摘要", "目录", "目次", "ABSTRACT", "CONTENTS"}:
            return True
        return bool(
            compact.startswith(("摘要", "中文摘要", "目录", "目次", "ABSTRACT", "CONTENTS"))
            or re.match(r"^第[一二三四五六七八九十百千万\d]+章", compact)
            or re.match(r"^\d+(?:\.\d+)+", compact)
        )

    elements = []
    for child in doc.element.body:
        child_tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if child_tag == "sectPr":
            break

        if child_tag == "tbl":
            tbl_props = extract_tbl_props(child)
            rows_data = []
            for row in child.findall(f"{{{W}}}tr"):
                cells_data = []
                for tc in row.findall(f"{{{W}}}tc"):
                    tcPr = tc.find(f"{{{W}}}tcPr")
                    tcW = tcPr.find(f"{{{W}}}tcW") if tcPr is not None else None
                    cell_w = int(tcW.get(f"{{{W}}}w", "0")) if tcW is not None else 0
                    cell_borders = {}
                    tc_borders = tcPr.find(f"{{{W}}}tcBorders") if tcPr is not None else None
                    if tc_borders is not None:
                        for border in tc_borders:
                            border_tag = border.tag.split("}")[-1]
                            border_val = border.get(f"{{{W}}}val", "nil")
                            if border_val not in (None, "nil", "none"):
                                cell_borders[border_tag] = {
                                    "val": border_val,
                                    "sz": border.get(f"{{{W}}}sz", "0"),
                                    "color": border.get(f"{{{W}}}color", "000000"),
                                }
                    cell_paras = []
                    for paragraph in tc.findall(f"{{{W}}}p"):
                        pPr = paragraph.find(f"{{{W}}}pPr")
                        jc = pPr.find(f"{{{W}}}jc") if pPr is not None else None
                        palign = jc.get(f"{{{W}}}val") if jc is not None else None
                        spacing = pPr.find(f"{{{W}}}spacing") if pPr is not None else None
                        p_line = spacing.get(f"{{{W}}}line") if spacing is not None else None
                        p_rule = spacing.get(f"{{{W}}}lineRule") if spacing is not None else None
                        p_before = spacing.get(f"{{{W}}}before") if spacing is not None else None
                        p_after = spacing.get(f"{{{W}}}after") if spacing is not None else None
                        p_ind = pPr.find(f"{{{W}}}ind") if pPr is not None else None
                        p_first = p_ind.get(f"{{{W}}}firstLine") if p_ind is not None else None
                        runs = []
                        for run in paragraph.findall(f"{{{W}}}r"):
                            rPr = run.find(f"{{{W}}}rPr")
                            fn_ascii, fn_ea, font_size, font_bold = "", "", 0, False
                            if rPr is not None:
                                run_fonts = rPr.find(f"{{{W}}}rFonts")
                                if run_fonts is not None:
                                    fn_ascii = run_fonts.get(f"{{{W}}}ascii", "") or ""
                                    fn_ea = run_fonts.get(f"{{{W}}}eastAsia", "") or ""
                                size = rPr.find(f"{{{W}}}sz")
                                if size is not None:
                                    font_size = int(size.get(f"{{{W}}}val", "0")) // 2
                                bold_el = rPr.find(f"{{{W}}}b")
                                font_bold = bold_el is not None and bold_el.get(f"{{{W}}}val", "1") not in ("0", "false", "False")
                            text = "".join(t.text or "" for t in run.findall(f"{{{W}}}t"))
                            payload = image_payload(run)
                            if text:
                                runs.append({"t": text, "fn": fn_ascii, "fe": fn_ea, "sz": font_size, "b": font_bold})
                            if payload:
                                runs.append({"t": "", "fn": fn_ascii, "fe": fn_ea, "sz": font_size, "b": font_bold, **payload})
                            if not text and not payload:
                                runs.append({"t": text, "fn": fn_ascii, "fe": fn_ea, "sz": font_size, "b": font_bold})
                        cell_paras.append(
                            {
                                "al": palign,
                                "ls_val": p_line,
                                "ls_rule": p_rule,
                                "sp_before": p_before,
                                "sp_after": p_after,
                                "fl_indent": p_first,
                                "r": runs,
                            }
                        )
                    cells_data.append({"w": cell_w, "tcPr": extract_tc_props(tcPr), "borders": cell_borders, "p": cell_paras})
                rows_data.append(cells_data)
            elements.append({"type": "table", "role": cover_table_role(rows_data), "tblPr": tbl_props, "rows": rows_data})
            continue

        if child_tag != "p":
            continue

        pPr = child.find(f"{{{W}}}pPr")
        has_sectPr = pPr.find(f"{{{W}}}sectPr") is not None if pPr is not None else False
        jc = pPr.find(f"{{{W}}}jc") if pPr is not None else None
        palign = jc.get(f"{{{W}}}val") if jc is not None else None

        spacing = pPr.find(f"{{{W}}}spacing") if pPr is not None else None
        line_val = spacing.get(f"{{{W}}}line") if spacing is not None else None
        line_rule = spacing.get(f"{{{W}}}lineRule") if spacing is not None else None
        before_val = spacing.get(f"{{{W}}}before") if spacing is not None else None
        after_val = spacing.get(f"{{{W}}}after") if spacing is not None else None

        indent = pPr.find(f"{{{W}}}ind") if pPr is not None else None
        first_line = indent.get(f"{{{W}}}firstLine") if indent is not None else None

        runs = []
        for run in child.findall(f"{{{W}}}r"):
            rPr = run.find(f"{{{W}}}rPr")
            fn_ascii, fn_ea, font_size, font_bold = "", "", 0, False
            if rPr is not None:
                run_fonts = rPr.find(f"{{{W}}}rFonts")
                if run_fonts is not None:
                    fn_ascii = run_fonts.get(f"{{{W}}}ascii", "") or ""
                    fn_ea = run_fonts.get(f"{{{W}}}eastAsia", "") or ""
                size = rPr.find(f"{{{W}}}sz")
                if size is not None:
                    font_size = int(size.get(f"{{{W}}}val", "0")) // 2
                bold_el = rPr.find(f"{{{W}}}b")
                font_bold = bold_el is not None and bold_el.get(f"{{{W}}}val", "1") not in ("0", "false", "False")
            text = "".join(t.text or "" for t in run.findall(f"{{{W}}}t"))
            runs.append({"t": text, "fn": fn_ascii, "fe": fn_ea, "sz": font_size, "b": font_bold})

        full_text = "".join(run["t"] for run in runs).strip()
        fmt_font = any(font in full_text for font in ["黑体", "宋体", "楷体", "华文", "方正", "Times New Roman"])
        fmt_size = any(key in full_text for key in ["二号", "三号", "四号", "小四", "五号", "小五", "号加粗", "号居中", "pt", "号字"])
        fmt_align = any(key in full_text for key in ["居中", "加粗", "缩进", "对齐", "行距", "段前", "段后", "固定值", "倍行距", "1.5倍", "双倍"])
        fmt_paren = "（" in full_text and "）" in full_text
        is_fmt_note = (fmt_paren and (fmt_font or fmt_size)) or (fmt_font and fmt_size and fmt_align)
        is_fmt_note = is_fmt_note or (fmt_paren and any(key in full_text for key in ["空一行", "空两行", "空行", "空  行"]))
        is_fmt_header = any(key in full_text[:60] for key in ["页眉页脚", "页眉", "页码", "字体要求", "字号要求", "格式要求", "排版要求"])

        if full_text and (any(key in full_text[:20] for key in stop_kw) or front_stop_match(full_text)):
            if not is_fmt_note:
                break

        if full_text:
            if is_fmt_note or is_fmt_header:
                continue
            if any(key in full_text[:30] for key in skip_kw):
                continue
            if len(full_text) > 40 and "删除" in full_text[:80]:
                continue

        img_payload = image_payload(child)
        base_payload = {
            "al": palign,
            "ls_val": line_val,
            "ls_rule": line_rule,
            "sp_before": before_val,
            "sp_after": after_val,
            "fl_indent": first_line,
            "r": runs,
            "section_break_after": has_sectPr,
        }
        if img_payload:
            elements.append({"type": "image", **base_payload, **img_payload})
        elif not full_text:
            elements.append({"type": "empty", **base_payload})
        else:
            elements.append({"type": "para", **base_payload})

    return elements
