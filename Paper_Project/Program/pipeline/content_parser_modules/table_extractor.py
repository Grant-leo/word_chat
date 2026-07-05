"""DOCX table extraction helpers for content parsing."""
from __future__ import annotations

import re
from typing import Any, Callable, Dict, List, Optional

try:
    from content_parser_modules.paragraph_stream import math_entry_from_ooxml
    from content_parser_modules.formula_text_items import _rich_text_item_from_inline_formula_spans
except ImportError:  # pragma: no cover - package-style imports
    from .paragraph_stream import math_entry_from_ooxml
    from .formula_text_items import _rich_text_item_from_inline_formula_spans


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
_TRANSPARENT_CONTENT_CONTAINERS = {"customXml", "smartTag"}
_ACCEPTED_REVISION_CONTAINERS = {"ins", "moveTo"}
_DELETED_REVISION_CONTAINERS = {"del", "moveFrom"}


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


def _cell_hmerge_kind(tc_elem: Any) -> str:
    tc_pr = tc_elem.find(f"{{{W_NS}}}tcPr")
    hmerge = tc_pr.find(f"{{{W_NS}}}hMerge") if tc_pr is not None else None
    if hmerge is None:
        return ""
    return (hmerge.get(f"{{{W_NS}}}val") or "continue").lower()


def _merge_int(merge: Dict[str, int], key: str, default: int = 1) -> int:
    try:
        return max(default, int(merge.get(key) or default))
    except Exception:
        return default


def _coalesce_legacy_hmerge_vmerge_merges(merges: List[Dict[str, int]]) -> List[Dict[str, int]]:
    horizontal_by_start: Dict[tuple, tuple] = {}
    verticals: List[tuple] = []
    for idx, merge in enumerate(merges):
        row = _merge_int(merge, "row", 0)
        col = _merge_int(merge, "col", 0)
        rowspan = _merge_int(merge, "rowspan", 1)
        colspan = _merge_int(merge, "colspan", 1)
        if rowspan == 1 and colspan > 1:
            horizontal_by_start.setdefault((row, col), (idx, colspan))
        if rowspan > 1:
            verticals.append((idx, row, col, rowspan, colspan))

    consumed = set()
    for idx, row, col, rowspan, colspan in verticals:
        if idx in consumed:
            continue
        start_horizontal = horizontal_by_start.get((row, col))
        if not start_horizontal:
            continue
        _, width = start_horizontal
        if width < colspan:
            continue
        row_horizontal_indices = []
        for row_offset in range(rowspan):
            horizontal = horizontal_by_start.get((row + row_offset, col))
            if not horizontal or horizontal[1] != width:
                row_horizontal_indices = []
                break
            row_horizontal_indices.append(horizontal[0])
        if not row_horizontal_indices:
            continue
        merge = merges[idx]
        if width > colspan:
            merge["colspan"] = width
        consumed.update(row_horizontal_indices)
        for other_idx, other_row, other_col, other_rowspan, _ in verticals:
            if (
                other_idx != idx
                and other_row == row
                and other_rowspan == rowspan
                and col < other_col < col + width
            ):
                consumed.add(other_idx)

    return [merge for idx, merge in enumerate(merges) if idx not in consumed]


def _dedupe_exact_merges(merges: List[Dict[str, int]]) -> List[Dict[str, int]]:
    deduped: List[Dict[str, int]] = []
    seen = set()
    for merge in merges:
        key = (
            _merge_int(merge, "row", 0),
            _merge_int(merge, "col", 0),
            _merge_int(merge, "rowspan", 1),
            _merge_int(merge, "colspan", 1),
        )
        if key in seen:
            continue
        seen.add(key)
        deduped.append(merge)
    return deduped


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


def _row_grid_after(tr_elem: Any) -> int:
    tr_pr = tr_elem.find(f"{{{W_NS}}}trPr")
    grid_after = tr_pr.find(f"{{{W_NS}}}gridAfter") if tr_pr is not None else None
    return _int_attr(grid_after, "val", 0)


def _fallback_row_widths_twips(tbl_elem: Any, ncols: int) -> List[int]:
    best: List[int] = []
    best_positive = 0
    for tr in _iter_table_row_elements(tbl_elem):
        widths: List[int] = [0] * _row_grid_before(tr)
        for tc in _iter_table_cell_elements(tr):
            colspan = _cell_grid_span(tc)
            width = _cell_width_twips(tc)
            if width > 0 and colspan > 1:
                per_col = max(1, width // colspan)
                widths.extend([per_col] * colspan)
            else:
                widths.extend([width] + ([0] * (colspan - 1)))
        widths.extend([0] * _row_grid_after(tr))
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


def _table_widths_twips(tbl_elem: Any, ncols: int) -> List[int]:
    grid_widths = _normalized_widths(_table_grid_widths_twips(tbl_elem), ncols)
    fallback_widths = _normalized_widths(_fallback_row_widths_twips(tbl_elem, ncols), ncols)
    if not grid_widths:
        return fallback_widths
    if not any(width <= 0 for width in grid_widths):
        return grid_widths

    repaired = list(grid_widths)
    positive = [width for width in fallback_widths + grid_widths if width > 0]
    default_width = max(1, sum(positive) // len(positive)) if positive else 0
    for idx, width in enumerate(repaired):
        if width > 0:
            continue
        fallback = fallback_widths[idx] if idx < len(fallback_widths) else 0
        repaired[idx] = fallback if fallback > 0 else default_width
    return repaired if any(width > 0 for width in repaired) else []


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
) -> Optional[tuple[int, str]]:
    raw = "".join(buf)
    buf.clear()
    if clean_text_func is not None:
        txt = clean_text_func(raw, preserve_newlines=True).rstrip()
    else:
        txt = raw.rstrip()
    if txt:
        idx = len(text_parts)
        text_parts.append(txt)
        return idx, txt
    return None


def _clean_cell_text(raw: str, clean_text_func: Optional[Callable[..., str]] = None) -> str:
    if clean_text_func is not None:
        return clean_text_func(raw, preserve_newlines=True).rstrip()
    return raw.rstrip()


def _local_name(elem: Any) -> str:
    return str(elem.tag).rsplit("}", 1)[-1]


def _sdt_content_children(elem: Any) -> List[Any]:
    content = elem.find(f"{{{W_NS}}}sdtContent") if elem is not None else None
    return list(content) if content is not None else []


def _iter_table_row_elements(container_elem: Any) -> List[Any]:
    rows: List[Any] = []
    for child in list(container_elem):
        local_name = _local_name(child)
        if local_name == "tr":
            rows.append(child)
        elif local_name == "sdt":
            for nested in _sdt_content_children(child):
                rows.extend(_iter_table_row_elements(nested))
        elif local_name in _TRANSPARENT_CONTENT_CONTAINERS or local_name in _ACCEPTED_REVISION_CONTAINERS:
            rows.extend(_iter_table_row_elements(child))
        elif local_name in _DELETED_REVISION_CONTAINERS:
            continue
    return rows


def _iter_table_cell_elements(row_elem: Any) -> List[Any]:
    cells: List[Any] = []
    for child in list(row_elem):
        local_name = _local_name(child)
        if local_name == "tc":
            cells.append(child)
        elif local_name == "sdt":
            for nested in _sdt_content_children(child):
                cells.extend(_iter_table_cell_elements(nested))
        elif local_name in _TRANSPARENT_CONTENT_CONTAINERS or local_name in _ACCEPTED_REVISION_CONTAINERS:
            cells.extend(_iter_table_cell_elements(child))
        elif local_name in _DELETED_REVISION_CONTAINERS:
            continue
    return cells


def _flatten_over_depth_table_content_into(
    tbl_elem: Any,
    text_parts: List[str],
    nested_items: List[Dict[str, Any]],
    clean_text_func: Optional[Callable[..., str]] = None,
    image_rels: Any = None,
    image_registry: Any = None,
    image_items_func: Optional[Callable[..., List[Dict[str, Any]]]] = None,
    image_run_items_func: Optional[Callable[..., List[Dict[str, Any]]]] = None,
    notes: Optional[Dict[str, Dict[str, str]]] = None,
) -> None:
    """Fail open for nested tables deeper than the structured-table limit."""

    def consume_child(child: Any) -> None:
        local_name = _local_name(child)
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
        elif local_name == "tbl":
            _flatten_over_depth_table_content_into(
                child,
                text_parts,
                nested_items,
                clean_text_func=clean_text_func,
                image_rels=image_rels,
                image_registry=image_registry,
                image_items_func=image_items_func,
                image_run_items_func=image_run_items_func,
                notes=notes,
            )
        elif local_name == "sdt":
            for nested in _sdt_content_children(child):
                consume_child(nested)
        elif local_name in _TRANSPARENT_CONTENT_CONTAINERS or local_name in _ACCEPTED_REVISION_CONTAINERS:
            for nested in list(child):
                consume_child(nested)
        elif local_name in _DELETED_REVISION_CONTAINERS:
            return

    for row in _iter_table_row_elements(tbl_elem):
        for cell in _iter_table_cell_elements(row):
            for child in list(cell):
                consume_child(child)


def _row_hmerge_spans(row_elem: Any, grid_before: int = 0) -> Dict[int, int]:
    spans: Dict[int, int] = {}
    active_start: Optional[int] = None
    active_width = 0
    gridspan_backed_remaining = 0
    col_idx = grid_before
    for tc in _iter_table_cell_elements(row_elem):
        colspan = _cell_grid_span(tc)
        hmerge_kind = _cell_hmerge_kind(tc)
        advance = colspan
        if hmerge_kind == "restart":
            active_start = col_idx
            active_width = colspan
            gridspan_backed_remaining = max(0, colspan - 1)
            spans[active_start] = active_width
        elif hmerge_kind == "continue" and active_start is not None:
            if gridspan_backed_remaining > 0:
                gridspan_backed_remaining = max(0, gridspan_backed_remaining - colspan)
                advance = 0
            else:
                active_width += colspan
                spans[active_start] = active_width
        else:
            active_start = None
            active_width = 0
            gridspan_backed_remaining = 0
        col_idx += advance
    return spans


def _paragraph_has_structured_runs(p_elem: Any) -> bool:
    return any(_local_name(elem) in {"oMath", "oMathPara", "footnoteReference", "endnoteReference"} for elem in p_elem.iter())


def _append_rich_text_run(runs: List[Dict[str, Any]], text: str) -> None:
    if not text:
        return
    if runs and runs[-1].get("type") == "text":
        runs[-1]["text"] = str(runs[-1].get("text") or "") + text
    else:
        runs.append({"type": "text", "text": text})


def _append_text_with_inline_formula_runs(
    runs: List[Dict[str, Any]],
    math_items: List[Dict[str, Any]],
    text: str,
) -> None:
    rich_item = _rich_text_item_from_inline_formula_spans(text)
    if not isinstance(rich_item, dict) or not rich_item.get("runs"):
        _append_rich_text_run(runs, text)
        return
    for run in rich_item.get("runs") or []:
        if not isinstance(run, dict):
            continue
        kind = run.get("type") or ("math" if run.get("math") else "text")
        if kind == "math":
            entries = [entry for entry in (run.get("math") or []) if isinstance(entry, dict)]
            if entries:
                math_items.extend(entries)
                copied = dict(run)
                copied["math"] = entries
                runs.append(copied)
        else:
            _append_rich_text_run(runs, str(run.get("text") or ""))


def _rich_text_replacement_for_text(index: int, text: str) -> Optional[Dict[str, Any]]:
    rich_item = _rich_text_item_from_inline_formula_spans(text)
    if not isinstance(rich_item, dict):
        return None
    rich_item = dict(rich_item)
    rich_item["location"] = "table_cell"
    rich_item["replace_paragraph_index"] = index
    return rich_item


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
            elif local_name in _ACCEPTED_REVISION_CONTAINERS:
                consume_transparent_container(part)
            elif local_name in _DELETED_REVISION_CONTAINERS or local_name == "delText":
                continue

    def consume_hyperlink(hyperlink_elem: Any) -> None:
        for part in list(hyperlink_elem):
            local_name = _local_name(part)
            if local_name == "r":
                consume_run(part)
            elif local_name == "hyperlink":
                consume_hyperlink(part)
            elif local_name == "sdt":
                consume_inline_sdt(part)
            elif local_name == "fldSimple":
                consume_field(part)
            elif local_name in _TRANSPARENT_CONTENT_CONTAINERS:
                consume_transparent_container(part)
            elif local_name in _ACCEPTED_REVISION_CONTAINERS:
                consume_transparent_container(part)
            elif local_name in _DELETED_REVISION_CONTAINERS:
                continue
            elif local_name in {"oMath", "oMathPara"}:
                append_math(part, "display" if local_name == "oMathPara" else "inline")
            elif local_name == "p":
                for nested in list(part):
                    nested_name = _local_name(nested)
                    if nested_name == "r":
                        consume_run(nested)
                    elif nested_name == "hyperlink":
                        consume_hyperlink(nested)
                    elif nested_name == "sdt":
                        consume_inline_sdt(nested)
                    elif nested_name == "fldSimple":
                        consume_field(nested)
                    elif nested_name in _TRANSPARENT_CONTENT_CONTAINERS:
                        consume_transparent_container(nested)
                    elif nested_name in _ACCEPTED_REVISION_CONTAINERS:
                        consume_transparent_container(nested)
                    elif nested_name in _DELETED_REVISION_CONTAINERS:
                        continue
                    elif nested_name in {"oMath", "oMathPara"}:
                        append_math(nested, "display" if nested_name == "oMathPara" else "inline")

    def consume_inline_sdt(sdt_elem: Any) -> None:
        for part in _sdt_content_children(sdt_elem):
            local_name = _local_name(part)
            if local_name == "r":
                consume_run(part)
            elif local_name == "hyperlink":
                consume_hyperlink(part)
            elif local_name in {"oMath", "oMathPara"}:
                append_math(part, "display" if local_name == "oMathPara" else "inline")
            elif local_name == "sdt":
                consume_inline_sdt(part)
            elif local_name == "fldSimple":
                consume_field(part)
            elif local_name in _TRANSPARENT_CONTENT_CONTAINERS:
                consume_transparent_container(part)
            elif local_name in _ACCEPTED_REVISION_CONTAINERS:
                consume_transparent_container(part)
            elif local_name in _DELETED_REVISION_CONTAINERS:
                continue
            elif local_name == "p":
                for nested in list(part):
                    nested_name = _local_name(nested)
                    if nested_name == "r":
                        consume_run(nested)
                    elif nested_name == "hyperlink":
                        consume_hyperlink(nested)
                    elif nested_name in {"oMath", "oMathPara"}:
                        append_math(nested, "display" if nested_name == "oMathPara" else "inline")
                    elif nested_name == "sdt":
                        consume_inline_sdt(nested)
                    elif nested_name == "fldSimple":
                        consume_field(nested)
                    elif nested_name in _TRANSPARENT_CONTENT_CONTAINERS:
                        consume_transparent_container(nested)
                    elif nested_name in _ACCEPTED_REVISION_CONTAINERS:
                        consume_transparent_container(nested)
                    elif nested_name in _DELETED_REVISION_CONTAINERS:
                        continue

    def consume_field(field_elem: Any) -> None:
        for part in list(field_elem):
            local_name = _local_name(part)
            if local_name == "r":
                consume_run(part)
            elif local_name == "hyperlink":
                consume_hyperlink(part)
            elif local_name == "sdt":
                consume_inline_sdt(part)
            elif local_name == "fldSimple":
                consume_field(part)
            elif local_name in _TRANSPARENT_CONTENT_CONTAINERS:
                consume_transparent_container(part)
            elif local_name in _ACCEPTED_REVISION_CONTAINERS:
                consume_transparent_container(part)
            elif local_name in _DELETED_REVISION_CONTAINERS:
                continue
            elif local_name in {"oMath", "oMathPara"}:
                append_math(part, "display" if local_name == "oMathPara" else "inline")
            elif local_name == "p":
                for nested in list(part):
                    nested_name = _local_name(nested)
                    if nested_name == "r":
                        consume_run(nested)
                    elif nested_name == "hyperlink":
                        consume_hyperlink(nested)
                    elif nested_name == "sdt":
                        consume_inline_sdt(nested)
                    elif nested_name == "fldSimple":
                        consume_field(nested)
                    elif nested_name in _TRANSPARENT_CONTENT_CONTAINERS:
                        consume_transparent_container(nested)
                    elif nested_name in _ACCEPTED_REVISION_CONTAINERS:
                        consume_transparent_container(nested)
                    elif nested_name in _DELETED_REVISION_CONTAINERS:
                        continue
                    elif nested_name in {"oMath", "oMathPara"}:
                        append_math(nested, "display" if nested_name == "oMathPara" else "inline")

    def consume_transparent_container(container_elem: Any) -> None:
        for part in list(container_elem):
            local_name = _local_name(part)
            if local_name == "r":
                consume_run(part)
            elif local_name == "hyperlink":
                consume_hyperlink(part)
            elif local_name == "sdt":
                consume_inline_sdt(part)
            elif local_name == "fldSimple":
                consume_field(part)
            elif local_name in _TRANSPARENT_CONTENT_CONTAINERS:
                consume_transparent_container(part)
            elif local_name in _ACCEPTED_REVISION_CONTAINERS:
                consume_transparent_container(part)
            elif local_name in _DELETED_REVISION_CONTAINERS:
                continue
            elif local_name in {"oMath", "oMathPara"}:
                append_math(part, "display" if local_name == "oMathPara" else "inline")
            elif local_name == "p":
                for nested in list(part):
                    nested_name = _local_name(nested)
                    if nested_name == "r":
                        consume_run(nested)
                    elif nested_name == "hyperlink":
                        consume_hyperlink(nested)
                    elif nested_name == "sdt":
                        consume_inline_sdt(nested)
                    elif nested_name == "fldSimple":
                        consume_field(nested)
                    elif nested_name in _TRANSPARENT_CONTENT_CONTAINERS:
                        consume_transparent_container(nested)
                    elif nested_name in _ACCEPTED_REVISION_CONTAINERS:
                        consume_transparent_container(nested)
                    elif nested_name in _DELETED_REVISION_CONTAINERS:
                        continue
                    elif nested_name in {"oMath", "oMathPara"}:
                        append_math(nested, "display" if nested_name == "oMathPara" else "inline")

    for child in list(p_elem):
        local_name = _local_name(child)
        if local_name == "r":
            consume_run(child)
        elif local_name == "hyperlink":
            consume_hyperlink(child)
        elif local_name in {"oMath", "oMathPara"}:
            append_math(child, "display" if local_name == "oMathPara" else "inline")
        elif local_name == "sdt":
            consume_inline_sdt(child)
        elif local_name == "fldSimple":
            consume_field(child)
        elif local_name in _TRANSPARENT_CONTENT_CONTAINERS:
            consume_transparent_container(child)
        elif local_name in _ACCEPTED_REVISION_CONTAINERS:
            consume_transparent_container(child)
        elif local_name in _DELETED_REVISION_CONTAINERS:
            continue
    flush_text()

    def flush_segment(segment_tokens: List[Dict[str, Any]]) -> None:
        if not segment_tokens:
            return
        plain_raw = "".join(
            str(token.get("text") or "")
            for token in segment_tokens
            if token.get("type") in {"text", "math"}
        )
        plain_text = _clean_cell_text(plain_raw, clean_text_func=clean_text_func)
        replace_index = len(text_parts)
        if plain_text:
            text_parts.append(plain_text)

        rich_runs: List[Dict[str, Any]] = []
        math_items: List[Dict[str, Any]] = []
        note_items: List[Dict[str, Any]] = []
        for token in segment_tokens:
            kind = token.get("type")
            if kind == "text":
                _append_text_with_inline_formula_runs(rich_runs, math_items, str(token.get("text") or ""))
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

    segment: List[Dict[str, Any]] = []
    for token in tokens:
        kind = token.get("type")
        if kind == "image":
            flush_segment(segment)
            segment = []
            for image_item in token.get("items") or []:
                image_item.setdefault("after_paragraph_index", len(text_parts))
                items.append(image_item)
        else:
            segment.append(token)
    flush_segment(segment)
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
        appended_text = _append_clean_text_part(text_parts, buf, clean_text_func=clean_text_func)
        if appended_text is not None:
            rich_item = _rich_text_replacement_for_text(appended_text[0], appended_text[1])
            if rich_item:
                media_items.append(rich_item)
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
            elif local_name in _ACCEPTED_REVISION_CONTAINERS:
                consume_transparent_container(part)
            elif local_name in _DELETED_REVISION_CONTAINERS or local_name == "delText":
                continue

    def consume_hyperlink(hyperlink_elem: Any) -> None:
        for part in list(hyperlink_elem):
            local_name = _local_name(part)
            if local_name == "r":
                consume_run(part)
            elif local_name == "hyperlink":
                consume_hyperlink(part)
            elif local_name == "sdt":
                consume_inline_sdt(part)
            elif local_name == "fldSimple":
                consume_field(part)
            elif local_name in _TRANSPARENT_CONTENT_CONTAINERS:
                consume_transparent_container(part)
            elif local_name in _ACCEPTED_REVISION_CONTAINERS:
                consume_transparent_container(part)
            elif local_name in _DELETED_REVISION_CONTAINERS:
                continue
            elif local_name == "p":
                for nested in list(part):
                    nested_name = _local_name(nested)
                    if nested_name == "r":
                        consume_run(nested)
                    elif nested_name == "hyperlink":
                        consume_hyperlink(nested)
                    elif nested_name == "sdt":
                        consume_inline_sdt(nested)
                    elif nested_name == "fldSimple":
                        consume_field(nested)
                    elif nested_name in _TRANSPARENT_CONTENT_CONTAINERS:
                        consume_transparent_container(nested)
                    elif nested_name in _ACCEPTED_REVISION_CONTAINERS:
                        consume_transparent_container(nested)
                    elif nested_name in _DELETED_REVISION_CONTAINERS:
                        continue

    def consume_inline_sdt(sdt_elem: Any) -> None:
        for part in _sdt_content_children(sdt_elem):
            local_name = _local_name(part)
            if local_name == "r":
                consume_run(part)
            elif local_name == "hyperlink":
                consume_hyperlink(part)
            elif local_name == "sdt":
                consume_inline_sdt(part)
            elif local_name == "fldSimple":
                consume_field(part)
            elif local_name in _TRANSPARENT_CONTENT_CONTAINERS:
                consume_transparent_container(part)
            elif local_name in _ACCEPTED_REVISION_CONTAINERS:
                consume_transparent_container(part)
            elif local_name in _DELETED_REVISION_CONTAINERS:
                continue
            elif local_name == "p":
                for nested in list(part):
                    nested_name = _local_name(nested)
                    if nested_name == "r":
                        consume_run(nested)
                    elif nested_name == "hyperlink":
                        consume_hyperlink(nested)
                    elif nested_name == "sdt":
                        consume_inline_sdt(nested)
                    elif nested_name == "fldSimple":
                        consume_field(nested)
                    elif nested_name in _TRANSPARENT_CONTENT_CONTAINERS:
                        consume_transparent_container(nested)
                    elif nested_name in _ACCEPTED_REVISION_CONTAINERS:
                        consume_transparent_container(nested)
                    elif nested_name in _DELETED_REVISION_CONTAINERS:
                        continue

    def consume_field(field_elem: Any) -> None:
        for part in list(field_elem):
            local_name = _local_name(part)
            if local_name == "r":
                consume_run(part)
            elif local_name == "hyperlink":
                consume_hyperlink(part)
            elif local_name == "sdt":
                consume_inline_sdt(part)
            elif local_name == "fldSimple":
                consume_field(part)
            elif local_name in _TRANSPARENT_CONTENT_CONTAINERS:
                consume_transparent_container(part)
            elif local_name in _ACCEPTED_REVISION_CONTAINERS:
                consume_transparent_container(part)
            elif local_name in _DELETED_REVISION_CONTAINERS:
                continue
            elif local_name == "p":
                for nested in list(part):
                    nested_name = _local_name(nested)
                    if nested_name == "r":
                        consume_run(nested)
                    elif nested_name == "hyperlink":
                        consume_hyperlink(nested)
                    elif nested_name == "sdt":
                        consume_inline_sdt(nested)
                    elif nested_name == "fldSimple":
                        consume_field(nested)
                    elif nested_name in _TRANSPARENT_CONTENT_CONTAINERS:
                        consume_transparent_container(nested)
                    elif nested_name in _ACCEPTED_REVISION_CONTAINERS:
                        consume_transparent_container(nested)
                    elif nested_name in _DELETED_REVISION_CONTAINERS:
                        continue

    def consume_transparent_container(container_elem: Any) -> None:
        for part in list(container_elem):
            local_name = _local_name(part)
            if local_name == "r":
                consume_run(part)
            elif local_name == "hyperlink":
                consume_hyperlink(part)
            elif local_name == "sdt":
                consume_inline_sdt(part)
            elif local_name == "fldSimple":
                consume_field(part)
            elif local_name in _TRANSPARENT_CONTENT_CONTAINERS:
                consume_transparent_container(part)
            elif local_name in _ACCEPTED_REVISION_CONTAINERS:
                consume_transparent_container(part)
            elif local_name in _DELETED_REVISION_CONTAINERS:
                continue
            elif local_name == "p":
                for nested in list(part):
                    nested_name = _local_name(nested)
                    if nested_name == "r":
                        consume_run(nested)
                    elif nested_name == "hyperlink":
                        consume_hyperlink(nested)
                    elif nested_name == "sdt":
                        consume_inline_sdt(nested)
                    elif nested_name == "fldSimple":
                        consume_field(nested)
                    elif nested_name in _TRANSPARENT_CONTENT_CONTAINERS:
                        consume_transparent_container(nested)
                    elif nested_name in _ACCEPTED_REVISION_CONTAINERS:
                        consume_transparent_container(nested)
                    elif nested_name in _DELETED_REVISION_CONTAINERS:
                        continue

    for child in list(p_elem):
        local_name = str(child.tag).rsplit("}", 1)[-1]
        if local_name == "r":
            consume_run(child)
        elif local_name == "hyperlink":
            consume_hyperlink(child)
        elif local_name == "sdt":
            consume_inline_sdt(child)
        elif local_name == "fldSimple":
            consume_field(child)
        elif local_name in _TRANSPARENT_CONTENT_CONTAINERS:
            consume_transparent_container(child)
        elif local_name in _ACCEPTED_REVISION_CONTAINERS:
            consume_transparent_container(child)
        elif local_name in _DELETED_REVISION_CONTAINERS:
            continue
    appended_text = _append_clean_text_part(text_parts, buf, clean_text_func=clean_text_func)
    if appended_text is not None:
        rich_item = _rich_text_replacement_for_text(appended_text[0], appended_text[1])
        if rich_item:
            media_items.append(rich_item)

    if not media_items and len(text_parts) == start_text_count:
        raw = paragraph_plain_text_from_ooxml(p_elem)
        if clean_text_func is not None:
            txt = clean_text_func(raw, preserve_newlines=True).rstrip()
        else:
            txt = raw.rstrip()
        if txt:
            idx = len(text_parts)
            text_parts.append(txt)
            rich_item = _rich_text_replacement_for_text(idx, txt)
            if rich_item:
                media_items.append(rich_item)
    return media_items


def _cell_text_and_nested_items_from_ooxml(
    tc_elem: Any,
    clean_text_func: Optional[Callable[..., str]] = None,
    nested_depth: int = 0,
    max_nested_depth: int = 6,
    image_rels: Any = None,
    image_registry: Any = None,
    image_items_func: Optional[Callable[..., List[Dict[str, Any]]]] = None,
    image_run_items_func: Optional[Callable[..., List[Dict[str, Any]]]] = None,
    notes: Optional[Dict[str, Dict[str, str]]] = None,
) -> tuple[str, List[Dict[str, Any]]]:
    text_parts: List[str] = []
    nested_items: List[Dict[str, Any]] = []

    def consume_cell_child(child: Any) -> None:
        local_name = _local_name(child)
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
        elif local_name == "tbl":
            if nested_depth < max_nested_depth:
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
            else:
                _flatten_over_depth_table_content_into(
                    child,
                    text_parts,
                    nested_items,
                    clean_text_func=clean_text_func,
                    image_rels=image_rels,
                    image_registry=image_registry,
                    image_items_func=image_items_func,
                    image_run_items_func=image_run_items_func,
                    notes=notes,
                )
        elif local_name == "sdt":
            for sdt_child in _sdt_content_children(child):
                consume_cell_child(sdt_child)
        elif local_name in _TRANSPARENT_CONTENT_CONTAINERS:
            for nested_child in list(child):
                consume_cell_child(nested_child)
        elif local_name in _ACCEPTED_REVISION_CONTAINERS:
            for nested_child in list(child):
                consume_cell_child(nested_child)
        elif local_name in _DELETED_REVISION_CONTAINERS:
            return

    for child in list(tc_elem):
        consume_cell_child(child)
    return "\n".join(text_parts).strip(), nested_items


def extract_table_from_ooxml(
    tbl_elem: Any,
    clean_text_func: Optional[Callable[..., str]] = None,
    nested_depth: int = 0,
    max_nested_depth: int = 6,
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
    row_grid_before: List[int] = []
    row_grid_after: List[int] = []
    cell_overrides: List[Dict[str, Any]] = []
    table_cell_items: List[Dict[str, Any]] = []
    active_vmerges: Dict[int, Dict[str, int]] = {}
    tbl_pr = tbl_elem.find(f"{{{W_NS}}}tblPr")
    table_cell_margins = _margins_twips(tbl_pr, "tblCellMar")
    table_borders = _borders_from_elem(tbl_pr, "tblBorders")

    for tr in _iter_table_row_elements(tbl_elem):
        row_idx = len(rows)
        grid_before = _row_grid_before(tr)
        grid_after = _row_grid_after(tr)
        row_grid_before.append(grid_before)
        row_grid_after.append(grid_after)
        row_hmerge_spans = _row_hmerge_spans(tr, grid_before)
        cells: List[str] = [""] * grid_before
        seen_vmerge_cols = set()
        continued_records = set()
        current_hmerge_record: Optional[Dict[str, int]] = None
        hmerge_gridspan_remaining = 0
        row_heights.append(_row_height_twips(tr))
        header_flags.append(_row_repeats_header(tr))
        for tc in _iter_table_cell_elements(tr):
            col_idx = len(cells)
            colspan = _cell_grid_span(tc)
            vmerge_kind = _cell_vmerge_kind(tc)
            hmerge_kind = _cell_hmerge_kind(tc)
            vmerge_record = active_vmerges.get(col_idx) if vmerge_kind == "continue" else None
            vmerge_record_colspan = int((vmerge_record or {}).get("colspan") or 1)
            required_hmerge_width = int((vmerge_record or {}).get("hmerge_colspan") or 1)
            current_hmerge_width = row_hmerge_spans.get(col_idx, colspan)
            is_vmerge_continue = (
                vmerge_kind == "continue"
                and vmerge_record is not None
                and vmerge_record_colspan == colspan
                and (required_hmerge_width <= 1 or current_hmerge_width == required_hmerge_width)
                and id(vmerge_record) not in continued_records
                and all(active_vmerges.get(col_idx + offset) is vmerge_record for offset in range(colspan))
            )
            is_gridspan_duplicate_hmerge = (
                hmerge_kind == "continue"
                and current_hmerge_record is not None
                and hmerge_gridspan_remaining > 0
            )
            duplicate_probe_text: Optional[str] = None
            duplicate_probe_nested_items: Optional[List[Dict[str, Any]]] = None
            vmerge_probe_text: Optional[str] = None
            vmerge_probe_nested_items: Optional[List[Dict[str, Any]]] = None
            if is_gridspan_duplicate_hmerge:
                duplicate_probe_text, duplicate_probe_nested_items = _cell_text_and_nested_items_from_ooxml(
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
                if not duplicate_probe_text and not duplicate_probe_nested_items:
                    hmerge_gridspan_remaining = max(0, hmerge_gridspan_remaining - colspan)
                    continue
                is_vmerge_continue = False
            if is_vmerge_continue:
                vmerge_probe_text, vmerge_probe_nested_items = _cell_text_and_nested_items_from_ooxml(
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
                if vmerge_probe_text or vmerge_probe_nested_items:
                    is_vmerge_continue = False
            is_hmerge_continue = (
                hmerge_kind == "continue"
                and current_hmerge_record is not None
                and hmerge_gridspan_remaining <= 0
            )
            if is_vmerge_continue or is_hmerge_continue:
                text = ""
                nested_items = []
            elif duplicate_probe_text is not None:
                text = duplicate_probe_text
                nested_items = duplicate_probe_nested_items or []
            elif vmerge_probe_text is not None:
                text = vmerge_probe_text
                nested_items = vmerge_probe_nested_items or []
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
                hmerge_width = row_hmerge_spans.get(col_idx, colspan)
                if hmerge_width > colspan:
                    record["hmerge_colspan"] = hmerge_width
                merges.append(record)
                for offset in range(colspan):
                    active_vmerges[col_idx + offset] = record
                    seen_vmerge_cols.add(col_idx + offset)
            elif is_vmerge_continue:
                record_id = id(vmerge_record)
                vmerge_record["rowspan"] = int(vmerge_record.get("rowspan") or 1) + 1
                continued_records.add(record_id)
                for offset in range(vmerge_record_colspan):
                    seen_vmerge_cols.add(col_idx + offset)

            if hmerge_kind == "restart":
                current_hmerge_record = {"row": row_idx, "col": col_idx, "rowspan": 1, "colspan": colspan}
                hmerge_gridspan_remaining = max(0, colspan - 1)
                merges.append(current_hmerge_record)
            elif is_hmerge_continue:
                current_hmerge_record["colspan"] = int(current_hmerge_record.get("colspan") or 1) + colspan
                hmerge_gridspan_remaining = 0
            elif hmerge_kind == "continue":
                current_hmerge_record = None
                hmerge_gridspan_remaining = 0
            elif colspan > 1 and vmerge_kind != "restart" and not is_vmerge_continue:
                merges.append({"row": row_idx, "col": col_idx, "rowspan": 1, "colspan": colspan})
                current_hmerge_record = None
                hmerge_gridspan_remaining = 0
            else:
                current_hmerge_record = None
                hmerge_gridspan_remaining = 0

        if grid_after > 0:
            cells.extend([""] * grid_after)
        rows.append(cells)
        for col in list(active_vmerges):
            if col not in seen_vmerge_cols:
                active_vmerges.pop(col, None)

    merges = _coalesce_legacy_hmerge_vmerge_merges(merges)
    for merge in merges:
        merge.pop("hmerge_colspan", None)
    merges = _dedupe_exact_merges(merges)
    merges = [
        merge
        for merge in merges
        if int(merge.get("rowspan") or 1) > 1 or int(merge.get("colspan") or 1) > 1
    ]
    ncols = max((len(row) for row in rows), default=0)
    widths = _table_widths_twips(tbl_elem, ncols)
    table_data: Dict[str, Any] = {"table_rows": rows, "table_merges": merges}
    if any(row_grid_before):
        table_data["table_row_grid_before"] = row_grid_before
    if any(row_grid_after):
        table_data["table_row_grid_after"] = row_grid_after
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
