"""Privacy-safe DOCX source-structure audit helpers."""
from __future__ import annotations

import os
import re
import zipfile
from typing import Any, Dict, Iterable, List
from xml.etree import ElementTree as ET


W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
SUPPORTED_WORD_IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff"}
_TRANSPARENT_CONTENT_CONTAINERS = {"customXml", "smartTag"}
_ACCEPTED_REVISION_CONTAINERS = {"ins", "moveTo"}
_DELETED_REVISION_CONTAINERS = {"del", "moveFrom"}
ISSUE_MESSAGES = {
    "SOURCE_TEXTBOX_UNSUPPORTED": "Source DOCX contains textbox content; visible text is downgraded into the body stream but original floating position needs review.",
    "SOURCE_FOOTNOTE_UNSUPPORTED": "Source DOCX contains footnotes; anchors and visible note text are extracted, but final numbering needs QA/visual review.",
    "SOURCE_ENDNOTE_UNSUPPORTED": "Source DOCX contains endnotes; anchors and visible note text are extracted, but final numbering needs QA/visual review.",
    "TRACKED_CHANGES_PRESENT": "Source DOCX contains tracked changes; accept or reject revisions before running.",
    "COMMENTS_PRESENT": "Source DOCX contains comments; comments are not part of the rendered paper body.",
    "CONTENT_CONTROL_UNSUPPORTED": "Source DOCX contains content controls; fields may need manual confirmation.",
    "SOURCE_EMBEDDED_OBJECT_UNSUPPORTED": "Source DOCX contains embedded/OLE objects that cannot be rendered safely.",
    "SOURCE_LANDSCAPE_SECTION_UNSUPPORTED": "Source DOCX contains landscape sections; final pagination needs visual/manual review.",
    "CONTENT_IMAGE_FORMAT_UNSUPPORTED": "Source DOCX contains images in a format outside the stable PNG/JPG path.",
    "COMPLEX_TABLE_UNSUPPORTED": "Source DOCX contains deeply nested, overwide, or irregular table structures that need manual/visual review.",
    "TABLE_MERGE_UNSUPPORTED": "Source DOCX contains merged table cells; basic gridSpan/hMerge/vMerge, common table layout details, and explicit borders are preserved, but complex table layout still needs visual review.",
}


def _local_name(elem: ET.Element) -> str:
    return str(elem.tag).rsplit("}", 1)[-1]


def _sdt_content_children(elem: ET.Element) -> List[ET.Element]:
    content = elem.find(W_NS + "sdtContent") if elem is not None else None
    return list(content) if content is not None else []


def _iter_final_view_children(elem: ET.Element) -> Iterable[ET.Element]:
    for child in list(elem):
        local_name = _local_name(child)
        if local_name == "sdt":
            for nested in _sdt_content_children(child):
                yield from _iter_final_view_children(nested)
        elif local_name in _TRANSPARENT_CONTENT_CONTAINERS or local_name in _ACCEPTED_REVISION_CONTAINERS:
            yield from _iter_final_view_children(child)
        elif local_name in _DELETED_REVISION_CONTAINERS:
            continue
        else:
            yield child


def _iter_visible_table_elements(container: ET.Element) -> List[ET.Element]:
    tables: List[ET.Element] = []

    def walk(elem: ET.Element) -> None:
        if elem.tag == W_NS + "tbl":
            tables.append(elem)
        for child in _iter_final_view_children(elem):
            walk(child)

    walk(container)
    return tables


def _iter_table_row_elements(container: ET.Element) -> List[ET.Element]:
    rows: List[ET.Element] = []
    for child in _iter_final_view_children(container):
        local_name = _local_name(child)
        if local_name == "tr":
            rows.append(child)
    return rows


def _iter_table_cell_elements(row: ET.Element) -> List[ET.Element]:
    cells: List[ET.Element] = []
    for child in _iter_final_view_children(row):
        local_name = _local_name(child)
        if local_name == "tc":
            cells.append(child)
    return cells


def _safe_issue(code: str, severity: str, detail: str = "") -> Dict[str, Any]:
    return {
        "code": code,
        "severity": severity,
        "message": ISSUE_MESSAGES.get(code, code),
        "detail": detail,
    }


def _read_text(zf: zipfile.ZipFile, name: str) -> str:
    try:
        return zf.read(name).decode("utf-8", errors="ignore")
    except Exception:
        return ""


def _document_parts(names: Iterable[str]) -> List[str]:
    parts = [
        name for name in names
        if name.startswith("word/")
        and name.endswith(".xml")
        and (
            name == "word/document.xml"
            or name.startswith("word/header")
            or name.startswith("word/footer")
            or name.startswith("word/footnotes")
            or name.startswith("word/endnotes")
            or name.startswith("word/comments")
        )
    ]
    return sorted(parts)


def _count_real_notes(xml_text: str, note_tag: str) -> int:
    if not xml_text:
        return 0
    try:
        root = ET.fromstring(xml_text.encode("utf-8"))
    except Exception:
        return 1 if f"<w:{note_tag}" in xml_text else 0
    count = 0
    for node in root.iter(W_NS + note_tag):
        note_type = node.attrib.get(W_NS + "type", "")
        if note_type not in {"separator", "continuationSeparator", "continuationNotice"}:
            count += 1
    return count


def _nested_table_stats(xml_text: str) -> Dict[str, int]:
    empty = {"nested_table_count": 0, "nested_table_max_depth": 0}
    if not xml_text:
        return empty
    try:
        root = ET.fromstring(xml_text.encode("utf-8"))
    except Exception:
        count = 1 if re.search(r"<w:tc\b[\s\S]*<w:tbl\b", xml_text) else 0
        return {"nested_table_count": count, "nested_table_max_depth": 1 if count else 0}

    def walk(elem: ET.Element, table_depth: int = 0) -> Dict[str, int]:
        nested_count = 0
        max_depth = 0
        for child in _iter_final_view_children(elem):
            if child.tag == W_NS + "tbl":
                if table_depth >= 1:
                    nested_count += 1
                    max_depth = max(max_depth, table_depth)
                child_stats = walk(child, table_depth + 1)
            else:
                child_stats = walk(child, table_depth)
            nested_count += child_stats["nested_table_count"]
            max_depth = max(max_depth, child_stats["nested_table_max_depth"])
        return {"nested_table_count": nested_count, "nested_table_max_depth": max_depth}

    return walk(root, 0)


def _max_table_columns(xml_text: str) -> int:
    if not xml_text:
        return 0
    try:
        root = ET.fromstring(xml_text.encode("utf-8"))
    except Exception:
        return 0
    max_cols = 0
    for table in _iter_visible_table_elements(root):
        max_cols = max(max_cols, _table_max_columns(table))
    return max_cols


def _wide_table_count(xml_text: str, threshold: int = 8) -> int:
    if not xml_text:
        return 0
    try:
        root = ET.fromstring(xml_text.encode("utf-8"))
    except Exception:
        return 0
    wide = 0
    for table in _iter_visible_table_elements(root):
        if _table_max_columns(table) > threshold:
            wide += 1
    return wide


def _table_structure_stats(xml_text: str) -> Dict[str, int]:
    empty = {"table_count": 0, "grid_span_count": 0, "hmerge_count": 0, "vmerge_count": 0}
    if not xml_text:
        return empty
    try:
        root = ET.fromstring(xml_text.encode("utf-8"))
    except Exception:
        return empty
    stats = dict(empty)
    tables = _iter_visible_table_elements(root)
    stats["table_count"] = len(tables)
    for table in tables:
        for row in _iter_table_row_elements(table):
            for cell in _iter_table_cell_elements(row):
                if _table_cell_grid_span(cell) > 1:
                    stats["grid_span_count"] += 1
                if _table_cell_hmerge_kind(cell):
                    stats["hmerge_count"] += 1
                if _table_cell_vmerge_kind(cell):
                    stats["vmerge_count"] += 1
    return stats


def _table_max_columns(table: ET.Element) -> int:
    table_max = 0
    for row in _iter_table_row_elements(table):
        cols = _row_grid_before(row)
        active_hmerge = False
        gridspan_backed_remaining = 0
        for cell in _iter_table_cell_elements(row):
            span = _table_cell_grid_span(cell)
            hmerge_kind = _table_cell_hmerge_kind(cell)
            if hmerge_kind == "continue" and active_hmerge and gridspan_backed_remaining > 0:
                gridspan_backed_remaining = max(0, gridspan_backed_remaining - span)
                continue
            cols += span
            if hmerge_kind == "restart":
                active_hmerge = True
                gridspan_backed_remaining = max(0, span - 1)
            elif hmerge_kind == "continue" and active_hmerge:
                gridspan_backed_remaining = 0
            else:
                active_hmerge = False
                gridspan_backed_remaining = 0
        cols += _row_grid_after(row)
        table_max = max(table_max, cols)
    return table_max


def _table_cell_grid_span(cell: ET.Element) -> int:
    tc_pr = cell.find(W_NS + "tcPr")
    if tc_pr is None:
        return 1
    grid_span = tc_pr.find(W_NS + "gridSpan")
    if grid_span is None:
        return 1
    try:
        return max(1, int(grid_span.attrib.get(W_NS + "val", "1")))
    except Exception:
        return 1


def _table_cell_vmerge_kind(cell: ET.Element) -> str:
    tc_pr = cell.find(W_NS + "tcPr")
    if tc_pr is None:
        return ""
    vmerge = tc_pr.find(W_NS + "vMerge")
    if vmerge is None:
        return ""
    return str(vmerge.attrib.get(W_NS + "val") or "continue").strip() or "continue"


def _table_cell_hmerge_kind(cell: ET.Element) -> str:
    tc_pr = cell.find(W_NS + "tcPr")
    if tc_pr is None:
        return ""
    hmerge = tc_pr.find(W_NS + "hMerge")
    if hmerge is None:
        return ""
    return str(hmerge.attrib.get(W_NS + "val") or "continue").strip() or "continue"


def _row_grid_before(row: ET.Element) -> int:
    tr_pr = row.find(W_NS + "trPr")
    if tr_pr is None:
        return 0
    grid_before = tr_pr.find(W_NS + "gridBefore")
    if grid_before is None:
        return 0
    try:
        return max(0, int(grid_before.attrib.get(W_NS + "val", "0")))
    except Exception:
        return 0


def _row_grid_after(row: ET.Element) -> int:
    tr_pr = row.find(W_NS + "trPr")
    if tr_pr is None:
        return 0
    grid_after = tr_pr.find(W_NS + "gridAfter")
    if grid_after is None:
        return 0
    try:
        return max(0, int(grid_after.attrib.get(W_NS + "val", "0")))
    except Exception:
        return 0


def _row_hmerge_spans(row: ET.Element) -> Dict[int, int]:
    spans: Dict[int, int] = {}
    active_start: int | None = None
    active_width = 0
    gridspan_backed_remaining = 0
    col_idx = _row_grid_before(row)
    for cell in _iter_table_cell_elements(row):
        span = _table_cell_grid_span(cell)
        hmerge_kind = _table_cell_hmerge_kind(cell)
        advance = span
        if hmerge_kind == "restart":
            active_start = col_idx
            active_width = span
            gridspan_backed_remaining = max(0, span - 1)
            spans[active_start] = active_width
        elif hmerge_kind == "continue" and active_start is not None:
            if gridspan_backed_remaining > 0:
                gridspan_backed_remaining = max(0, gridspan_backed_remaining - span)
                advance = 0
            else:
                active_width += span
                spans[active_start] = active_width
        else:
            active_start = None
            active_width = 0
            gridspan_backed_remaining = 0
        col_idx += advance
    return spans


def _row_irregular_hmerge_count(row: ET.Element) -> int:
    count = 0
    active_hmerge = False
    gridspan_backed_remaining = 0
    flagged_group = False
    for cell in _iter_table_cell_elements(row):
        span = _table_cell_grid_span(cell)
        hmerge_kind = _table_cell_hmerge_kind(cell)
        if hmerge_kind == "restart":
            active_hmerge = True
            gridspan_backed_remaining = max(0, span - 1)
            flagged_group = False
        elif hmerge_kind == "continue" and active_hmerge and gridspan_backed_remaining > 0:
            if not flagged_group:
                count += 1
                flagged_group = True
            gridspan_backed_remaining = max(0, gridspan_backed_remaining - span)
        elif hmerge_kind == "continue" and active_hmerge:
            gridspan_backed_remaining = 0
        else:
            active_hmerge = False
            gridspan_backed_remaining = 0
            flagged_group = False
    return count


def _table_geometry_stats(xml_text: str) -> Dict[str, int]:
    empty = {
        "irregular_table_count": 0,
        "irregular_vmerge_count": 0,
        "irregular_grid_span_count": 0,
        "irregular_hmerge_count": 0,
    }
    if not xml_text:
        return empty
    try:
        root = ET.fromstring(xml_text.encode("utf-8"))
    except Exception:
        return empty

    stats = dict(empty)
    for table in _iter_visible_table_elements(root):
        grid = table.find(W_NS + "tblGrid")
        grid_cols = len(grid.findall(W_NS + "gridCol")) if grid is not None else 0
        active_vmerges: Dict[int, Dict[str, int]] = {}
        table_irregular = False
        for row in _iter_table_row_elements(table):
            irregular_hmerges = _row_irregular_hmerge_count(row)
            if irregular_hmerges:
                stats["irregular_hmerge_count"] += irregular_hmerges
                table_irregular = True
            row_hmerge_spans = _row_hmerge_spans(row)
            col_idx = _row_grid_before(row)
            next_active: Dict[int, Dict[str, int]] = {}
            active_hmerge = False
            gridspan_backed_remaining = 0
            for cell in _iter_table_cell_elements(row):
                span = _table_cell_grid_span(cell)
                hmerge_kind = _table_cell_hmerge_kind(cell)
                vmerge_kind = _table_cell_vmerge_kind(cell)
                if hmerge_kind == "continue" and active_hmerge and gridspan_backed_remaining > 0 and not vmerge_kind:
                    gridspan_backed_remaining = max(0, gridspan_backed_remaining - span)
                    continue
                if grid_cols and col_idx + span > grid_cols:
                    stats["irregular_grid_span_count"] += 1
                    table_irregular = True
                if vmerge_kind == "restart":
                    hmerge_width = row_hmerge_spans.get(col_idx, span)
                    active = {"span": span, "hmerge_width": hmerge_width}
                    for offset in range(span):
                        next_active[col_idx + offset] = active
                elif vmerge_kind == "continue":
                    expected = active_vmerges.get(col_idx)
                    cell_irregular = False
                    if expected is None:
                        cell_irregular = True
                        expected_span = span
                        expected_hmerge_width = span
                    else:
                        expected_span = expected.get("span") or span
                        expected_hmerge_width = expected.get("hmerge_width") or expected_span
                    if expected is not None and expected_span != span:
                        cell_irregular = True
                    current_hmerge_width = row_hmerge_spans.get(col_idx, span)
                    if expected is not None and expected_hmerge_width > expected_span and current_hmerge_width != expected_hmerge_width:
                        cell_irregular = True
                    if cell_irregular:
                        stats["irregular_vmerge_count"] += 1
                        table_irregular = True
                    active = {"span": expected_span, "hmerge_width": expected_hmerge_width}
                    for offset in range(span):
                        next_active[col_idx + offset] = active
                if hmerge_kind == "restart":
                    active_hmerge = True
                    gridspan_backed_remaining = max(0, span - 1)
                elif hmerge_kind == "continue" and active_hmerge:
                    gridspan_backed_remaining = 0
                else:
                    active_hmerge = False
                    gridspan_backed_remaining = 0
                col_idx += span
            active_vmerges = next_active
        if table_irregular:
            stats["irregular_table_count"] += 1
    return stats


def _is_landscape_section(sect_pr: ET.Element) -> bool:
    pg_sz = sect_pr.find(W_NS + "pgSz")
    if pg_sz is None:
        return False
    orient = str(pg_sz.attrib.get(W_NS + "orient") or "").strip().lower()
    return orient == "landscape"


def _paragraph_section_properties(elem: ET.Element) -> ET.Element | None:
    if elem.tag != W_NS + "p":
        return None
    p_pr = elem.find(W_NS + "pPr")
    if p_pr is None:
        return None
    return p_pr.find(W_NS + "sectPr")


def _wide_tables_in_elements(elements: Iterable[ET.Element], threshold: int = 8) -> int:
    wide = 0
    seen: set[int] = set()
    for elem in elements:
        for table in _iter_visible_table_elements(elem):
            table_id = id(table)
            if table_id in seen:
                continue
            seen.add(table_id)
            if _table_max_columns(table) > threshold:
                wide += 1
    return wide


def _landscape_wide_table_risk_count(xml_text: str, threshold: int = 8) -> int:
    if not xml_text:
        return 0
    try:
        root = ET.fromstring(xml_text.encode("utf-8"))
    except Exception:
        return 0
    body = root.find(W_NS + "body")
    if body is None:
        return 0
    risk_count = 0
    section_elements: List[ET.Element] = []
    for child in list(body):
        if child.tag == W_NS + "sectPr":
            if _is_landscape_section(child):
                risk_count += _wide_tables_in_elements(section_elements, threshold=threshold)
            section_elements = []
            continue
        section_elements.append(child)
        sect_pr = _paragraph_section_properties(child)
        if sect_pr is not None:
            if _is_landscape_section(sect_pr):
                risk_count += _wide_tables_in_elements(section_elements, threshold=threshold)
            section_elements = []
    return risk_count


def _media_formats(names: Iterable[str]) -> Dict[str, int]:
    formats: Dict[str, int] = {}
    for name in names:
        if not name.startswith("word/media/"):
            continue
        ext = os.path.splitext(name)[1].lower() or "<no_ext>"
        formats[ext] = formats.get(ext, 0) + 1
    return formats


def audit_docx_source(docx_path: str) -> Dict[str, Any]:
    """Audit a DOCX for structures that can silently degrade extraction.

    The result intentionally avoids paragraph text and absolute paths. It records
    only structural counts, format families, and issue codes.
    """
    result: Dict[str, Any] = {
        "schema_version": 1,
        "file_type": "docx",
        "counts": {},
        "issues": [],
    }
    try:
        with zipfile.ZipFile(docx_path) as zf:
            names = zf.namelist()
            xml_parts = {name: _read_text(zf, name) for name in _document_parts(names)}
            document_xml = xml_parts.get("word/document.xml", "")
            all_xml = "\n".join(xml_parts.values())

            nested_stats = _nested_table_stats(document_xml)
            geometry_stats = _table_geometry_stats(document_xml)
            table_stats = _table_structure_stats(document_xml)
            wide_table_count = _wide_table_count(document_xml)
            landscape_section_count = len(re.findall(r'w:orient=["\']landscape["\']', all_xml))
            counts: Dict[str, Any] = {
                "xml_parts": len(xml_parts),
                "textbox_count": len(re.findall(r"<w:txbxContent\b|<wps:txbx\b|<v:textbox\b", all_xml)),
                "tracked_change_count": len(re.findall(r"<w:(?:ins|del|moveFrom|moveTo)\b", all_xml)),
                "comment_count": _count_real_notes(xml_parts.get("word/comments.xml", ""), "comment"),
                "content_control_count": len(re.findall(r"<w:sdt\b", all_xml)),
                "embedded_object_count": len([n for n in names if n.startswith("word/embeddings/")])
                + len(re.findall(r"<w:object\b|oleObject", all_xml)),
                "landscape_section_count": landscape_section_count,
                "footnote_count": _count_real_notes(xml_parts.get("word/footnotes.xml", ""), "footnote"),
                "endnote_count": _count_real_notes(xml_parts.get("word/endnotes.xml", ""), "endnote"),
                "table_count": table_stats["table_count"],
                "grid_span_count": table_stats["grid_span_count"],
                "hmerge_count": table_stats["hmerge_count"],
                "vmerge_count": table_stats["vmerge_count"],
                "nested_table_count": nested_stats["nested_table_count"],
                "nested_table_max_depth": nested_stats["nested_table_max_depth"],
                "max_table_columns": _max_table_columns(document_xml),
                "wide_table_count": wide_table_count,
                "landscape_wide_table_risk_count": _landscape_wide_table_risk_count(document_xml),
                "irregular_table_count": geometry_stats["irregular_table_count"],
                "irregular_vmerge_count": geometry_stats["irregular_vmerge_count"],
                "irregular_grid_span_count": geometry_stats["irregular_grid_span_count"],
                "irregular_hmerge_count": geometry_stats["irregular_hmerge_count"],
            }
            counts["merged_cell_count"] = counts["grid_span_count"] + counts["hmerge_count"] + counts["vmerge_count"]
            formats = _media_formats(names)
            counts["image_format_counts"] = formats
            result["counts"] = counts

            issues: List[Dict[str, Any]] = []
            if counts["textbox_count"]:
                issues.append(_safe_issue("SOURCE_TEXTBOX_UNSUPPORTED", "warning", f"textboxes={counts['textbox_count']}"))
            if counts["footnote_count"]:
                issues.append(_safe_issue("SOURCE_FOOTNOTE_UNSUPPORTED", "warning", f"footnotes={counts['footnote_count']}"))
            if counts["endnote_count"]:
                issues.append(_safe_issue("SOURCE_ENDNOTE_UNSUPPORTED", "warning", f"endnotes={counts['endnote_count']}"))
            if counts["tracked_change_count"]:
                issues.append(_safe_issue("TRACKED_CHANGES_PRESENT", "error", f"tracked_changes={counts['tracked_change_count']}"))
            if counts["comment_count"]:
                issues.append(_safe_issue("COMMENTS_PRESENT", "warning", f"comments={counts['comment_count']}"))
            if counts["content_control_count"]:
                issues.append(_safe_issue("CONTENT_CONTROL_UNSUPPORTED", "warning", f"content_controls={counts['content_control_count']}"))
            if counts["embedded_object_count"]:
                issues.append(_safe_issue("SOURCE_EMBEDDED_OBJECT_UNSUPPORTED", "error", f"embedded_objects={counts['embedded_object_count']}"))
            if counts["landscape_section_count"]:
                issues.append(_safe_issue("SOURCE_LANDSCAPE_SECTION_UNSUPPORTED", "warning", f"landscape_sections={counts['landscape_section_count']}"))
            unsupported_formats = {
                ext: count for ext, count in formats.items()
                if ext not in SUPPORTED_WORD_IMAGE_EXTS
            }
            if unsupported_formats:
                detail = ", ".join(f"{ext}:{count}" for ext, count in sorted(unsupported_formats.items()))
                issues.append(_safe_issue("CONTENT_IMAGE_FORMAT_UNSUPPORTED", "error", detail))
            if counts["merged_cell_count"]:
                detail = f"merged_cells={counts['merged_cell_count']}"
                if counts.get("hmerge_count"):
                    detail += f"; hmerge={counts['hmerge_count']}"
                issues.append(_safe_issue("TABLE_MERGE_UNSUPPORTED", "warning", detail))
            if counts["nested_table_max_depth"] > 4 or counts["max_table_columns"] > 8 or counts["irregular_table_count"]:
                detail = (
                    f"nested_tables={counts['nested_table_count']}; "
                    f"nested_depth={counts['nested_table_max_depth']}; "
                    f"max_columns={counts['max_table_columns']}; "
                    f"wide_tables={counts['wide_table_count']}; "
                    f"landscape_wide_tables={counts['landscape_wide_table_risk_count']}; "
                    f"irregular_tables={counts['irregular_table_count']}; "
                    f"irregular_vmerges={counts['irregular_vmerge_count']}; "
                    f"irregular_grid_spans={counts['irregular_grid_span_count']}; "
                    f"irregular_hmerges={counts['irregular_hmerge_count']}"
                )
                issues.append(_safe_issue("COMPLEX_TABLE_UNSUPPORTED", "warning", detail))
            result["issues"] = issues
    except zipfile.BadZipFile:
        result["file_type"] = "invalid_docx"
        result["issues"] = [_safe_issue("SOURCE_FORMAT_UNSUPPORTED", "error", "bad docx zip container")]
    except Exception as exc:
        result["issues"] = [_safe_issue("SOURCE_FORMAT_UNSUPPORTED", "error", type(exc).__name__)]
    return result
