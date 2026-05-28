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


def _pdfinfo(pdf_path: str) -> Dict[str, Any]:
    exe = shutil.which("pdfinfo")
    if not exe:
        return {"available": False}
    result = _run([exe, pdf_path], timeout=60)
    if result.returncode != 0:
        return {"available": True, "error": result.stderr.strip()[:500]}
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
    exe = shutil.which("pdftotext")
    if not exe:
        return []
    txt_path = os.path.join(visual_dir, "rendered.txt")
    result = _run([exe, "-layout", pdf_path, txt_path], timeout=120)
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


def _sample_pages(page_count: int, pages_text: List[str]) -> List[int]:
    if page_count <= 0:
        return []
    samples = {1}
    toc_page = _find_page(pages_text, [r"目录", r"contents"])
    body_page = _find_page(pages_text, [r"第\d+章", r"chapter\s*1", r"1\.\s*[A-Za-z]"])
    for page in (toc_page, body_page, 3 if page_count >= 3 else None, page_count // 2 if page_count >= 8 else None):
        if page and 1 <= page <= page_count:
            samples.add(page)
    return sorted(samples)[:6]


def _render_samples(pdf_path: str, visual_dir: str, pages: List[int]) -> List[str]:
    exe = shutil.which("pdftoppm")
    if not exe:
        return []
    sample_dir = os.path.join(visual_dir, "samples")
    os.makedirs(sample_dir, exist_ok=True)
    rendered: List[str] = []
    for page in pages:
        prefix = os.path.join(sample_dir, f"page_{page:03d}")
        result = _run([exe, "-png", "-f", str(page), "-l", str(page), "-r", "120", pdf_path, prefix], timeout=120)
        if result.returncode == 0:
            matches = [os.path.join(sample_dir, f) for f in os.listdir(sample_dir) if f.startswith(f"page_{page:03d}") and f.endswith(".png")]
            rendered.extend(sorted(matches))
    return rendered


def _render_all_pages(pdf_path: str, visual_dir: str, page_count: int) -> List[str]:
    exe = shutil.which("pdftoppm")
    if not exe or page_count <= 0:
        return []
    page_dir = os.path.join(visual_dir, "pages")
    if os.path.isdir(page_dir):
        shutil.rmtree(page_dir, ignore_errors=True)
    os.makedirs(page_dir, exist_ok=True)
    prefix = os.path.join(page_dir, "page")
    result = _run([exe, "-png", "-r", "110", pdf_path, prefix], timeout=max(120, page_count * 10))
    if result.returncode != 0:
        return []
    return sorted(os.path.join(page_dir, f) for f in os.listdir(page_dir) if f.lower().endswith(".png"))

