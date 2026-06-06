"""DOCX table extraction helpers for content parsing."""
from __future__ import annotations

import re
from typing import Any, Callable, Dict, List, Optional

try:
    from content_parser_modules.paragraph_stream import math_entry_from_ooxml
except ImportError:  # pragma: no cover - package-style imports
    from .paragraph_stream import math_entry_from_ooxml


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_MARGIN_SIDE_MAP = {
    "top": "top",
    "bottom": "bottom",
    "left": "left",
    "right": "right",
    "start": "left",
    "end": "right",
}
_BORDER_SIDES = {"top", "left", "bottom", "right", "insideH", "insideV", "tl2br", "tr2bl", "start", "end"}
_BORDER_ATTRS = ("val", "sz", "color", "space")


def paragraph_plain_text_from_ooxml(p_elem: Any) -> str:
    pieces: List[str] = []
    for run in p_elem.findall(f"{{{W_NS}}}r"):
        part = "".join(t.text or "" for t in run.findall(f"{{{W_NS}}}t"))
        if run.find(f"{{{W_NS}}}br") is not None and not part:
            part = "\n"
        pieces.append(part)
    return "".join(pieces)


def _int_attr(elem: Any, name: str, default: int = 1) -> int:
    try:
        value = elem.get(f"{{{W_NS}}}{name}") if elem is not None else None
        return max(default, int(value)) if value is not None else default
    except Exception:
        return default


def _cell_grid_span(tc_elem: Any) -> int:
    tc_pr = tc_elem.find(f"{{{W_NS}}}tcPr")
    grid_span = tc_pr.find(f"{{{W_NS}}}gridSpan") if tc_pr is not None else None
    return _int_attr(grid_span, "val", 1)


def _cell_width_twips(tc_elem: Any) -> int:
    tc_pr = tc_elem.find(f"{{{W_NS}}}tcPr")
    tc_w = tc_pr.find(f"{{{W_NS}}}tcW") if tc_pr is not None else None
    return _int_attr(tc_w, "w", 0)


def _cell_vmerge_kind(tc_elem: Any) -> str:
    tc_pr = tc_elem.find(f"{{{W_NS}}}tcPr")
    vmerge = tc_pr.find(f"{{{W_NS}}}vMerge") if tc_pr is not None else None
    if vmerge is None:
        return ""
    return (vmerge.get(f"{{{W_NS}}}val") or "continue").lower()


def _row_height_twips(tr_elem: Any) -> Dict[str, Any]:
    tr_pr = tr_elem.find(f"{{{W_NS}}}trPr")
    tr_height = tr_pr.find(f"{{{W_NS}}}trHeight") if tr_pr is not None else None
    val = _int_attr(tr_height, "val", 0)
    if val <= 0:
        return {}
    out: Dict[str, Any] = {"val": val}
    rule = (tr_height.get(f"{{{W_NS}}}hRule") or "").strip() if tr_height is not None else ""
    if rule:
        out["rule"] = rule
    return out


def _row_repeats_header(tr_elem: Any) -> bool:
    tr_pr = tr_elem.find(f"{{{W_NS}}}trPr")
    header = tr_pr.find(f"{{{W_NS}}}tblHeader") if tr_pr is not None else None
    if header is None:
        return False
    val = (header.get(f"{{{W_NS}}}val") or "true").strip().lower()
    return val not in {"0", "false", "off", "no"}


def _margins_twips(parent_elem: Any, container_name: str) -> Dict[str, int]:
    container = parent_elem.find(f"{{{W_NS}}}{container_name}") if parent_elem is not None else None
    margins: Dict[str, int] = {}
    if container is None:
        return margins
    for child in list(container):
        side_name = str(child.tag).rsplit("}", 1)[-1]
        side = _MARGIN_SIDE_MAP.get(side_name)
        if not side:
            continue
        width = _int_attr(child, "w", 0)
        if width > 0:
            margins[side] = width
    return margins


def _border_spec(border_elem: Any) -> Dict[str, str]:
    spec: Dict[str, str] = {}
    if border_elem is None:
        return spec
    for attr in _BORDER_ATTRS:
        value = border_elem.get(f"{{{W_NS}}}{attr}")
        if value is not None:
            spec[attr] = str(value)
    return spec


def _borders_from_elem(parent_elem: Any, container_name: str) -> Dict[str, Dict[str, str]]:
    container = parent_elem.find(f"{{{W_NS}}}{container_name}") if parent_elem is not None else None
    borders: Dict[str, Dict[str, str]] = {}
    if container is None:
        return borders
    for child in list(container):
        side = str(child.tag).rsplit("}", 1)[-1]
        if side not in _BORDER_SIDES:
            continue
        spec = _border_spec(child)
        if spec:
            borders[side] = spec
    return borders


def _cell_v_align(tc_elem: Any) -> str:
    tc_pr = tc_elem.find(f"{{{W_NS}}}tcPr")
    v_align = tc_pr.find(f"{{{W_NS}}}vAlign") if tc_pr is not None else None
    return (v_align.get(f"{{{W_NS}}}val") or "").strip() if v_align is not None else ""


def _cell_text_from_ooxml(tc_elem: Any, clean_text_func: Optional[Callable[..., str]] = None) -> str:
    paras: List[str] = []
    for p in tc_elem.findall(f"{{{W_NS}}}p"):
        raw = paragraph_plain_text_from_ooxml(p)
        if clean_text_func is not None:
            txt = clean_text_func(raw, preserve_newlines=True).rstrip()
        else:
            txt = raw.rstrip()
        if txt:
            paras.append(txt)
    return "\n".join(paras).strip()


def _table_grid_widths_twips(tbl_elem: Any) -> List[int]:
    grid = tbl_elem.find(f"{{{W_NS}}}tblGrid")
    widths: List[int] = []
    if grid is not None:
        for grid_col in grid.findall(f"{{{W_NS}}}gridCol"):
            width = _int_attr(grid_col, "w", 0)
            widths.append(width)
    return widths


def _row_grid_before(tr_elem: Any) -> int:
    tr_pr = tr_elem.find(f"{{{W_NS}}}trPr")
    grid_before = tr_pr.find(f"{{{W_NS}}}gridBefore") if tr_pr is not None else None
    return _int_attr(grid_before, "val", 0)


def _fallback_row_widths_twips(tbl_elem: Any, ncols: int) -> List[int]:
    best: List[int] = []
    best_positive = 0
    for tr in tbl_elem.findall(f"{{{W_NS}}}tr"):
        widths: List[int] = [0] * _row_grid_before(tr)
        for tc in tr.findall(f"{{{W_NS}}}tc"):
            colspan = _cell_grid_span(tc)
            width = _cell_width_twips(tc)
            if width > 0 and colspan > 1:
                per_col = max(1, width // colspan)
                widths.extend([per_col] * colspan)
            else:
                widths.extend([width] + ([0] * (colspan - 1)))
        positive = sum(1 for value in widths if value > 0)
        if positive > best_positive:
            best = widths
            best_positive = positive
        if best_positive >= ncols:
            break
    return best[:ncols]


def _normalized_widths(widths: List[int], ncols: int) -> List[int]:
    if not widths or ncols <= 0:
        return []
    out = [int(value or 0) for value in widths[:ncols]]
    if len(out) < ncols:
        out.extend([0] * (ncols - len(out)))
    return out if any(value > 0 for value in out) else []


def _run_text_preserve_breaks(run_elem: Any) -> str:
    parts: List[str] = []
    for child in list(run_elem):
        local_name = str(child.tag).rsplit("}", 1)[-1]
        if local_name == "t":
            parts.append(child.text or "")
        elif local_name == "tab":
            parts.append("\t")
        elif local_name in {"br", "cr"}:
            parts.append("\n")
    return "".join(parts)


def _append_clean_text_part(
    text_parts: List[str],
    buf: List[str],
    clean_text_func: Optional[Callable[..., str]] = None,
) -> None:
    raw = "".join(buf)
    buf.clear()
    if clean_text_func is not None:
        txt = clean_text_func(raw, preserve_newlines=True).rstrip()
    else:
        txt = raw.rstrip()
    if txt:
        text_parts.append(txt)


def _clean_cell_text(raw: str, clean_text_func: Optional[Callable[..., str]] = None) -> str:
    if clean_text_func is not None:
        return clean_text_func(raw, preserve_newlines=True).rstrip()
    return raw.rstrip()


def _local_name(elem: Any) -> str:
    return str(elem.tag).rsplit("}", 1)[-1]


def _paragraph_has_structured_runs(p_elem: Any) -> bool:
    return any(_local_name(elem) in {"oMath", "oMathPara", "footnoteReference", "endnoteReference"} for elem in p_elem.iter())


def _append_rich_text_run(runs: List[Dict[str, Any]], text: str) -> None:
    if not text:
        return
    if runs and runs[-1].get("type") == "text":
        runs[-1]["text"] = str(runs[-1].get("text") or "") + text
    else:
        runs.append({"type": "text", "text": text})


def _extract_structured_paragraph_items(
    p_elem: Any,
    text_parts: List[str],
    clean_text_func: Optional[Callable[..., str]] = None,
    image_rels: Any = None,
    image_registry: Any = None,
    image_items_func: Optional[Callable[..., List[Dict[str, Any]]]] = None,
    image_run_items_func: Optional[Callable[..., List[Dict[str, Any]]]] = None,
    notes: Optional[Dict[str, Dict[str, str]]] = None,
) -> List[Dict[str, Any]]:
    items: List[Dict[str, Any]] = []
    tokens: List[Dict[str, Any]] = []
    buf: List[str] = []
    seen_rids = set()

    def flush_text() -> None:
        text = "".join(buf)
        buf.clear()
        if text:
            tokens.append({"type": "text", "text": text})

    def extract_run_images(run_elem: Any) -> List[Dict[str, Any]]:
        if image_rels is None or image_registry is None:
            return []
        if image_run_items_func is not None:
            return image_run_items_func(run_elem, image_rels, image_registry, seen_rids, location="table_cell") or []
        if image_items_func is not None:
            return image_items_func(run_elem, image_rels, image_registry, location="table_cell") or []
        return []

    def append_math(math_elem: Any, math_type: str = "inline") -> None:
        flush_text()
        entry = math_entry_from_ooxml(math_elem, math_type)
        tokens.append({"type": "math", "text": entry.get("text") or "", "math": [entry]})

    def append_note_ref(note_elem: Any, note_type: str) -> None:
        flush_text()
        source_id = str(note_elem.get(f"{{{W_NS}}}id") or "").strip()
        tokens.append(
            {
                "type": "note_ref",
                "note_type": note_type,
                "source_id": source_id,
                "text": ((notes or {}).get(note_type) or {}).get(source_id, ""),
            }
        )

    def consume_run(run_elem: Any) -> None:
        for part in list(run_elem):
            local_name = _local_name(part)
            if local_name in {"drawing", "pict"}:
                flush_text()
                images = [dict(item) for item in extract_run_images(run_elem) if isinstance(item, dict)]
                if images:
                    tokens.append({"type": "image", "items": images})
            elif local_name == "oMath":
                append_math(part, "inline")
            elif local_name in {"footnoteReference", "endnoteReference"}:
                append_note_ref(part, "footnote" if local_name == "footnoteReference" else "endnote")
            elif local_name == "t":
                if part.text:
                    buf.append(part.text)
            elif local_name == "tab":
                buf.append("\t")
            elif local_name in {"br", "cr"}:
                buf.append("\n")

    for child in list(p_elem):
        local_name = _local_name(child)
        if local_name == "r":
            consume_run(child)
        elif local_name == "hyperlink":
            for run in list(child):
                if _local_name(run) == "r":
                    consume_run(run)
        elif local_name in {"oMath", "oMathPara"}:
            append_math(child, "display" if local_name == "oMathPara" else "inline")
    flush_text()

    plain_raw = "".join(str(token.get("text") or "") for token in tokens if token.get("type") in {"text", "math"})
    plain_text = _clean_cell_text(plain_raw, clean_text_func=clean_text_func)
    replace_index = len(text_parts)
    if plain_text:
        text_parts.append(plain_text)

    rich_runs: List[Dict[str, Any]] = []
    math_items: List[Dict[str, Any]] = []
    note_items: List[Dict[str, Any]] = []
    for token in tokens:
        kind = token.get("type")
        if kind == "text":
            _append_rich_text_run(rich_runs, str(token.get("text") or ""))
        elif kind == "math":
            entries = token.get("math") or []
            math_items.extend(entries)
            rich_runs.append({"type": "math", "text": token.get("text") or "", "math": entries})
        elif kind == "note_ref":
            note_run = {
                "type": "note_ref",
                "note_type": token.get("note_type") or "footnote",
                "source_id": token.get("source_id") or "",
                "text": token.get("text") or "",
            }
            note_items.append(note_run)
            rich_runs.append(note_run)

    if (math_items or note_items) and rich_runs:
        rich_item: Dict[str, Any] = {
            "role": "rich_text",
            "location": "table_cell",
            "replace_paragraph_index": replace_index,
            "text": plain_text,
            "runs": rich_runs,
        }
        if math_items:
            rich_item["math"] = math_items
        if note_items:
            rich_item["notes"] = [
                {
                    "type": note.get("note_type") or "footnote",
                    "source_id": note.get("source_id") or "",
                    "text": note.get("text") or "",
                }
                for note in note_items
            ]
        items.append(rich_item)

    seen_textish = False
    for token in tokens:
        kind = token.get("type")
        if kind == "image":
            image_index = replace_index + (1 if seen_textish and plain_text else 0)
            for image_item in token.get("items") or []:
                image_item.setdefault("after_paragraph_index", image_index)
                items.append(image_item)
        elif kind in {"text", "math", "note_ref"} and str(token.get("text") or "").strip():
            seen_textish = True
    return items


def _extract_ordered_paragraph_text_and_media(
    p_elem: Any,
    text_parts: List[str],
    clean_text_func: Optional[Callable[..., str]] = None,
    image_rels: Any = None,
    image_registry: Any = None,
    image_items_func: Optional[Callable[..., List[Dict[str, Any]]]] = None,
    image_run_items_func: Optional[Callable[..., List[Dict[str, Any]]]] = None,
    notes: Optional[Dict[str, Dict[str, str]]] = None,
) -> List[Dict[str, Any]]:
    if _paragraph_has_structured_runs(p_elem):
        return _extract_structured_paragraph_items(
            p_elem,
            text_parts,
            clean_text_func=clean_text_func,
            image_rels=image_rels,
            image_registry=image_registry,
            image_items_func=image_items_func,
            image_run_items_func=image_run_items_func,
            notes=notes,
        )

    media_items: List[Dict[str, Any]] = []
    buf: List[str] = []
    seen_rids = set()
    start_text_count = len(text_parts)

    def extract_run_images(run_elem: Any) -> List[Dict[str, Any]]:
        if image_rels is None or image_registry is None:
            return []
        if image_run_items_func is not None:
            return image_run_items_func(run_elem, image_rels, image_registry, seen_rids, location="table_cell") or []
        if image_items_func is not None:
            return image_items_func(run_elem, image_rels, image_registry, location="table_cell") or []
        return []

    def add_media_from_run(run_elem: Any) -> None:
        _append_clean_text_part(text_parts, buf, clean_text_func=clean_text_func)
        for image_item in extract_run_images(run_elem):
            if isinstance(image_item, dict):
                image_item = dict(image_item)
                image_item.setdefault("after_paragraph_index", len(text_parts))
                media_items.append(image_item)

    def consume_run(run_elem: Any) -> None:
        for part in list(run_elem):
            local_name = str(part.tag).rsplit("}", 1)[-1]
            if local_name in {"drawing", "pict"}:
                add_media_from_run(run_elem)
            elif local_name == "t":
                if part.text:
                    buf.append(part.text)
            elif local_name == "tab":
                buf.append("\t")
            elif local_name in {"br", "cr"}:
                buf.append("\n")

    for child in list(p_elem):
        local_name = str(child.tag).rsplit("}", 1)[-1]
        if local_name == "r":
            consume_run(child)
        elif local_name == "hyperlink":
            for run in list(child):
                if str(run.tag).rsplit("}", 1)[-1] == "r":
                    consume_run(run)
    _append_clean_text_part(text_parts, buf, clean_text_func=clean_text_func)

    if not media_items and len(text_parts) == start_text_count:
        raw = paragraph_plain_text_from_ooxml(p_elem)
        if clean_text_func is not None:
            txt = clean_text_func(raw, preserve_newlines=True).rstrip()
        else:
            txt = raw.rstrip()
        if txt:
            text_parts.append(txt)
    return media_items


def _cell_text_and_nested_items_from_ooxml(
    tc_elem: Any,
    clean_text_func: Optional[Callable[..., str]] = None,
    nested_depth: int = 0,
    max_nested_depth: int = 2,
    image_rels: Any = None,
    image_registry: Any = None,
    image_items_func: Optional[Callable[..., List[Dict[str, Any]]]] = None,
    image_run_items_func: Optional[Callable[..., List[Dict[str, Any]]]] = None,
    notes: Optional[Dict[str, Dict[str, str]]] = None,
) -> tuple[str, List[Dict[str, Any]]]:
    text_parts: List[str] = []
    nested_items: List[Dict[str, Any]] = []
    for child in list(tc_elem):
        local_name = str(child.tag).rsplit("}", 1)[-1]
        if local_name == "p":
            nested_items.extend(
                _extract_ordered_paragraph_text_and_media(
                    child,
                    text_parts,
                    clean_text_func=clean_text_func,
                    image_rels=image_rels,
                    image_registry=image_registry,
                    image_items_func=image_items_func,
                    image_run_items_func=image_run_items_func,
                    notes=notes,
                )
            )
        elif local_name == "tbl" and nested_depth < max_nested_depth:
            nested_data = extract_table_from_ooxml(
                child,
                clean_text_func=clean_text_func,
                nested_depth=nested_depth + 1,
                max_nested_depth=max_nested_depth,
                image_rels=image_rels,
                image_registry=image_registry,
                image_items_func=image_items_func,
                image_run_items_func=image_run_items_func,
                notes=notes,
            )
            if nested_data.get("table_rows"):
                nested_item: Dict[str, Any] = {
                    "role": "table",
                    "location": "nested_table_cell",
                    "after_paragraph_index": len(text_parts),
                }
                nested_item.update(nested_data)
                nested_items.append(nested_item)
    return "\n".join(text_parts).strip(), nested_items


def extract_table_from_ooxml(
    tbl_elem: Any,
    clean_text_func: Optional[Callable[..., str]] = None,
    nested_depth: int = 0,
    max_nested_depth: int = 2,
    image_rels: Any = None,
    image_registry: Any = None,
    image_items_func: Optional[Callable[..., List[Dict[str, Any]]]] = None,
    image_run_items_func: Optional[Callable[..., List[Dict[str, Any]]]] = None,
    notes: Optional[Dict[str, Dict[str, str]]] = None,
) -> Dict[str, Any]:
    """Extract table text plus basic merged-cell geometry on a stable grid."""
    rows: List[List[str]] = []
    merges: List[Dict[str, int]] = []
    row_heights: List[Dict[str, Any]] = []
    header_flags: List[bool] = []
    cell_overrides: List[Dict[str, Any]] = []
    table_cell_items: List[Dict[str, Any]] = []
    active_vmerges: Dict[int, Dict[str, int]] = {}
    tbl_pr = tbl_elem.find(f"{{{W_NS}}}tblPr")
    table_cell_margins = _margins_twips(tbl_pr, "tblCellMar")
    table_borders = _borders_from_elem(tbl_pr, "tblBorders")

    for tr in tbl_elem.findall(f"{{{W_NS}}}tr"):
        row_idx = len(rows)
        cells: List[str] = [""] * _row_grid_before(tr)
        seen_vmerge_cols = set()
        continued_records = set()
        row_heights.append(_row_height_twips(tr))
        header_flags.append(_row_repeats_header(tr))
        for tc in tr.findall(f"{{{W_NS}}}tc"):
            col_idx = len(cells)
            colspan = _cell_grid_span(tc)
            vmerge_kind = _cell_vmerge_kind(tc)
            is_vmerge_continue = vmerge_kind == "continue"
            if is_vmerge_continue:
                text = ""
                nested_items = []
            else:
                text, nested_items = _cell_text_and_nested_items_from_ooxml(
                    tc,
                    clean_text_func=clean_text_func,
                    nested_depth=nested_depth,
                    max_nested_depth=max_nested_depth,
                    image_rels=image_rels,
                    image_registry=image_registry,
                    image_items_func=image_items_func,
                    image_run_items_func=image_run_items_func,
                    notes=notes,
                )
            tc_pr = tc.find(f"{{{W_NS}}}tcPr")
            cell_margins = _margins_twips(tc_pr, "tcMar")
            cell_borders = _borders_from_elem(tc_pr, "tcBorders")
            v_align = _cell_v_align(tc)
            if nested_items:
                table_cell_items.append({"row": row_idx, "col": col_idx, "items": nested_items})
            if v_align or cell_margins or cell_borders:
                override: Dict[str, Any] = {"row": row_idx, "col": col_idx}
                if v_align:
                    override["v_align"] = v_align
                if cell_margins:
                    override["margins_twips"] = cell_margins
                if cell_borders:
                    override["borders"] = cell_borders
                cell_overrides.append(override)

            cells.append(text)
            for _ in range(colspan - 1):
                cells.append("")

            if vmerge_kind == "restart":
                record = {"row": row_idx, "col": col_idx, "rowspan": 1, "colspan": colspan}
                merges.append(record)
                for offset in range(colspan):
                    active_vmerges[col_idx + offset] = record
                    seen_vmerge_cols.add(col_idx + offset)
            elif is_vmerge_continue:
                record = active_vmerges.get(col_idx)
                if record is not None:
                    record_id = id(record)
                    if record_id not in continued_records:
                        record["rowspan"] = int(record.get("rowspan") or 1) + 1
                        continued_records.add(record_id)
                    for offset in range(int(record.get("colspan") or colspan or 1)):
                        seen_vmerge_cols.add(col_idx + offset)
            elif colspan > 1:
                merges.append({"row": row_idx, "col": col_idx, "rowspan": 1, "colspan": colspan})

        rows.append(cells)
        for col in list(active_vmerges):
            if col not in seen_vmerge_cols:
                active_vmerges.pop(col, None)

    merges = [
        merge
        for merge in merges
        if int(merge.get("rowspan") or 1) > 1 or int(merge.get("colspan") or 1) > 1
    ]
    ncols = max((len(row) for row in rows), default=0)
    widths = _normalized_widths(_table_grid_widths_twips(tbl_elem), ncols)
    if not widths:
        widths = _normalized_widths(_fallback_row_widths_twips(tbl_elem, ncols), ncols)
    table_data: Dict[str, Any] = {"table_rows": rows, "table_merges": merges}
    if widths:
        table_data["table_col_widths_twips"] = widths
    if table_borders:
        table_data["table_borders"] = table_borders
    if any(row_heights):
        table_data["table_row_heights_twips"] = row_heights
    repeat_header_rows = 0
    for flag in header_flags:
        if not flag:
            break
        repeat_header_rows += 1
    table_data["table_repeat_header_rows"] = repeat_header_rows
    if table_cell_margins:
        table_data["table_cell_margins_twips"] = table_cell_margins
    if cell_overrides:
        table_data["table_cell_overrides"] = cell_overrides
    if table_cell_items:
        table_data["table_cell_items"] = table_cell_items
    return table_data


def extract_table_rows_from_ooxml(tbl_elem: Any, clean_text_func: Optional[Callable[..., str]] = None) -> List[List[str]]:
    """Preserve cell paragraph breaks so code/config tables do not collapse."""
    return extract_table_from_ooxml(tbl_elem, clean_text_func=clean_text_func).get("table_rows") or []


def looks_like_code_line(text: str) -> bool:
    """Heuristic for network/device configuration or command-line code."""
    t = (text or "").strip()
    if not t or len(t) > 220:
        return False
    if re.match(r"^[A-Za-z0-9_.-]+[>#]", t):
        return True
    if re.match(
        r"^(interface|vlan|ip route|ip address|router|switchport|acl|rule|nat|dhcp|dns|ospf|bgp|display|show|ping|tracert|undo|quit|return|sysname|description|gateway|firewall|security-policy)\b",
        t,
        re.I,
    ):
        return True
    if re.match(r"^[a-z][a-z0-9_-]+\s+[-A-Za-z0-9_/.:]+", t) and any(ch in t for ch in ["/", ".", "-", "_"]):
        return True
    return False


def table_rows_look_like_code(rows: List[List[str]]) -> bool:
    """Classify one-/two-column command tables as code, not academic tables."""
    flat: List[str] = []
    for row in rows or []:
        for cell in row or []:
            for line in str(cell or "").splitlines():
                if line.strip():
                    flat.append(line.strip())
    if not flat:
        return False
    ncols = max((len(row) for row in rows or []), default=0)
    hits = sum(1 for value in flat if looks_like_code_line(value))
    if ncols <= 1 and len(flat) >= 2 and hits >= 2:
        return True
    if ncols <= 2 and len(flat) >= 4 and hits >= max(2, len(flat) // 3):
        return True
    return False


def code_text_from_table_rows(rows: List[List[str]], clean_code_func: Optional[Callable[[str], str]] = None) -> str:
    lines: List[str] = []
    for row in rows or []:
        cells = [str(cell or "").rstrip() for cell in row]
        if len(cells) == 1:
            lines.append(cells[0])
        else:
            lines.append("    ".join(cells).rstrip())
    text = "\n".join(lines).rstrip()
    return clean_code_func(text) if clean_code_func is not None else text.strip()
