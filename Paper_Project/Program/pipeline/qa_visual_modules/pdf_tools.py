"""Poppler/PDF helper functions for visual QA."""
from __future__ import annotations

import os
import re
import shutil
from typing import Any, Dict, List

try:
    from qa_visual_modules.exporters import _run
except ImportError:  # pragma: no cover - package-style imports
    from .exporters import _run


def _tool_candidates(name: str) -> List[str]:
    candidates: List[str] = []

    def add(path: str | None) -> None:
        if not path:
            return
        norm = os.path.normcase(os.path.abspath(path))
        if norm not in {os.path.normcase(os.path.abspath(item)) for item in candidates}:
            candidates.append(path)

    add(shutil.which(name))
    finder = "where.exe" if os.name == "nt" else "which"
    args = [finder, name] if os.name == "nt" else [finder, "-a", name]
    try:
        result = _run(args, timeout=10)
        if result.returncode == 0:
            for line in (result.stdout or "").splitlines():
                add(line.strip())
    except Exception:
        pass
    return candidates


def _run_tool(name: str, args: List[str], timeout: int = 120):
    candidates = _tool_candidates(name)
    if not candidates:
        return None, None, []
    failures: List[str] = []
    last_result = None
    for exe in candidates:
        try:
            result = _run([exe] + args, timeout=timeout)
        except Exception as exc:
            failures.append((str(exc) or exc.__class__.__name__)[:500])
            continue
        if result.returncode == 0:
            return result, exe, failures
        last_result = result
        detail = (result.stderr or result.stdout or "").strip()
        if detail:
            failures.append(detail[:500])
    return last_result, candidates[-1], failures


def _pdfinfo(pdf_path: str) -> Dict[str, Any]:
    result, _exe, failures = _run_tool("pdfinfo", [pdf_path], timeout=60)
    if result is None:
        if failures:
            return {"available": True, "error": " | ".join(failures[-3:])}
        return {"available": False}
    if result.returncode != 0:
        detail = (result.stderr or result.stdout or "").strip()[:500]
        if failures:
            detail = " | ".join(failures[-3:])
        return {"available": True, "error": detail}
    info: Dict[str, Any] = {"available": True, "raw": result.stdout}
    m = re.search(r"^Pages:\s*(\d+)", result.stdout, re.M)
    if m:
        info["pages"] = int(m.group(1))
    m = re.search(r"^Page size:\s*([\d.]+)\s*x\s*([\d.]+)\s*pts", result.stdout, re.M)
    if m:
        info["page_width_pt"] = float(m.group(1))
        info["page_height_pt"] = float(m.group(2))
    return info


def _pdf_pages_text(pdf_path: str, visual_dir: str) -> List[str]:
    if not _tool_candidates("pdftotext"):
        return []
    txt_path = os.path.join(visual_dir, "rendered.txt")
    result, _exe, _failures = _run_tool("pdftotext", ["-layout", pdf_path, txt_path], timeout=120)
    if result is None:
        return []
    if result.returncode != 0 or not os.path.exists(txt_path):
        return []
    with open(txt_path, "r", encoding="utf-8", errors="ignore") as f:
        return f.read().split("\f")


def _find_page(pages: List[str], patterns: List[str], start: int = 0) -> int | None:
    for idx, text in enumerate(pages[start:], start):
        compact = re.sub(r"\s+", "", text or "")
        for pat in patterns:
            if re.search(pat, compact, re.I) or re.search(pat, text or "", re.I):
                return idx + 1
    return None


def _find_pages(pages: List[str], patterns: List[str], start: int = 0) -> List[int]:
    matches: List[int] = []
    for idx, text in enumerate(pages[start:], start):
        compact = re.sub(r"\s+", "", text or "")
        for pat in patterns:
            if re.search(pat, compact, re.I) or re.search(pat, text or "", re.I):
                matches.append(idx + 1)
                break
    return matches


def _is_front_matter_list_page(text: str) -> bool:
    raw = text or ""
    compact = re.sub(r"\s+", "", raw)
    lower = raw.lower()
    compact_lower = compact.lower()
    if re.search(r"\blist\s+of\s+(figures|tables)\b", lower):
        return True
    if "contents" in compact_lower:
        return True
    if any(token in compact for token in ("目录", "图清单", "表清单", "图目录", "表目录")):
        return True
    if re.search(r"\.{4,}\s*(?:[ivxlcdm]+|\d+)\s*$", lower, re.M) and re.search(r"\b(table|figure|fig\.?|chapter)\b", lower):
        return True
    return False


def _add_sample(samples: List[int], page: int | None, page_count: int, limit: int = 6) -> None:
    if page and 1 <= page <= page_count and page not in samples and len(samples) < limit:
        samples.append(page)


def _sample_pages(page_count: int, pages_text: List[str]) -> List[int]:
    if page_count <= 0:
        return []
    samples: List[int] = []
    toc_page = _find_page(pages_text, [r"目录", r"contents"])
    body_page = _find_page(pages_text, [r"第\d+章", r"chapter\s*1", r"1\.\s*[A-Za-z]"])
    for page in (1, toc_page, body_page):
        _add_sample(samples, page, page_count)

    risk_page_patterns = [
        [r"图\s*\d+", r"\bfig\.?\s*\d+", r"\bfigure\s*\d+", r"插图"],
        [r"表\s*\d+", r"\btable\s*\d+", r"三线表"],
        [r"公式", r"方程", r"\bequation\b", r"\beq\.?\s*\(?\d+"],
    ]
    for risk_index, patterns in enumerate(risk_page_patterns):
        for page in _find_pages(pages_text, patterns):
            if page - 1 < len(pages_text) and _is_front_matter_list_page(pages_text[page - 1]):
                continue
            if page not in samples:
                _add_sample(samples, page, page_count)
                if risk_index == 1:
                    _add_sample(samples, page + 1, page_count)
                break

    for page in (
        3 if page_count >= 3 else None,
        page_count // 2 if page_count >= 8 else None,
        page_count,
    ):
        _add_sample(samples, page, page_count)
    return sorted(samples)


def _render_samples(pdf_path: str, visual_dir: str, pages: List[int]) -> List[str]:
    if not _tool_candidates("pdftoppm"):
        return []
    sample_dir = os.path.join(visual_dir, "samples")
    os.makedirs(sample_dir, exist_ok=True)
    rendered: List[str] = []
    for page in pages:
        prefix = os.path.join(sample_dir, f"page_{page:03d}")
        result, _exe, _failures = _run_tool("pdftoppm", ["-png", "-f", str(page), "-l", str(page), "-r", "120", pdf_path, prefix], timeout=120)
        if result is None:
            continue
        if result.returncode == 0:
            matches = [os.path.join(sample_dir, f) for f in os.listdir(sample_dir) if f.startswith(f"page_{page:03d}") and f.endswith(".png")]
            rendered.extend(sorted(matches))
    return rendered


def _render_all_pages(pdf_path: str, visual_dir: str, page_count: int) -> List[str]:
    if not _tool_candidates("pdftoppm") or page_count <= 0:
        return []
    page_dir = os.path.join(visual_dir, "pages")
    if os.path.isdir(page_dir):
        shutil.rmtree(page_dir, ignore_errors=True)
    os.makedirs(page_dir, exist_ok=True)
    prefix = os.path.join(page_dir, "page")
    result, _exe, _failures = _run_tool("pdftoppm", ["-png", "-r", "110", pdf_path, prefix], timeout=max(120, page_count * 10))
    if result is None:
        return []
    if result.returncode != 0:
        return []
    return sorted(os.path.join(page_dir, f) for f in os.listdir(page_dir) if f.lower().endswith(".png"))
