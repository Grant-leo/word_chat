"""DOCX table extraction helpers for content parsing."""
from __future__ import annotations

import re
from typing import Any, Callable, List, Optional


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def paragraph_plain_text_from_ooxml(p_elem: Any) -> str:
    pieces: List[str] = []
    for run in p_elem.findall(f"{{{W_NS}}}r"):
        part = "".join(t.text or "" for t in run.findall(f"{{{W_NS}}}t"))
        if run.find(f"{{{W_NS}}}br") is not None and not part:
            part = "\n"
        pieces.append(part)
    return "".join(pieces)


def extract_table_rows_from_ooxml(tbl_elem: Any, clean_text_func: Optional[Callable[..., str]] = None) -> List[List[str]]:
    """Preserve cell paragraph breaks so code/config tables do not collapse."""
    rows: List[List[str]] = []
    for tr in tbl_elem.findall(f"{{{W_NS}}}tr"):
        cells: List[str] = []
        for tc in tr.findall(f"{{{W_NS}}}tc"):
            paras: List[str] = []
            for p in tc.findall(f"{{{W_NS}}}p"):
                raw = paragraph_plain_text_from_ooxml(p)
                if clean_text_func is not None:
                    txt = clean_text_func(raw, preserve_newlines=True).rstrip()
                else:
                    txt = raw.rstrip()
                if txt:
                    paras.append(txt)
            cells.append("\n".join(paras).strip())
        rows.append(cells)
    return rows


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
