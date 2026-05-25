"""
qa_visual.py - optional PDF/render QA for generated DOCX outputs.

This module is deliberately best-effort: if Word COM or Poppler tools are not
available, it reports a clear issue instead of guessing visual quality.
"""
from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import tempfile
from datetime import datetime
from typing import Any, Dict, List

try:
    from privacy import sanitize_value
except Exception:  # pragma: no cover
    def sanitize_value(value: Any, project_root: str | None = None) -> Any:
        return value


def _issue(code: str, severity: str, message: str, detail: str = "") -> Dict[str, Any]:
    return {"code": code, "severity": severity, "message": message, "detail": detail}


def _load_json(path: str) -> Dict[str, Any]:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def _run(cmd: List[str], timeout: int = 120) -> subprocess.CompletedProcess:
    return subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=timeout)


def _export_pdf_with_word(docx_path: str, pdf_path: str) -> str:
    powershell = shutil.which("powershell") or shutil.which("pwsh")
    if not powershell:
        raise RuntimeError("PowerShell is not available; cannot drive Word COM.")

    script = f"""
$ErrorActionPreference = 'Stop'
$docx = {json.dumps(os.path.abspath(docx_path))}
$pdf = {json.dumps(os.path.abspath(pdf_path))}
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $null
try {{
  $doc = $word.Documents.Open($docx, $false, $true)
  $doc.ExportAsFixedFormat($pdf, 17)
  $doc.Close($false)
  $doc = $null
}} finally {{
  if ($doc -ne $null) {{
    try {{ $doc.Close($false) }} catch {{ }}
  }}
  if ($word -ne $null) {{
    try {{ $word.Quit() }} catch {{ }}
  }}
}}
"""
    with tempfile.TemporaryDirectory(prefix="wordchat_visual_ps_") as td:
        ps1 = os.path.join(td, "export.ps1")
        with open(ps1, "w", encoding="utf-8-sig") as f:
            f.write(script)
        result = _run([powershell, "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", ps1], timeout=180)
    if result.returncode != 0:
        raise RuntimeError((result.stderr or result.stdout or "Word COM export failed").strip()[:1000])
    if not os.path.exists(pdf_path):
        raise RuntimeError("Word COM finished but PDF was not created.")
    return pdf_path


def _export_pdf(docx_path: str, visual_dir: str) -> str:
    # Copy to ASCII temp path first; Word/WPS COM can be fragile with long CJK paths.
    work = tempfile.mkdtemp(prefix="wordchat_visual_docx_")
    try:
        safe_docx = os.path.join(work, "input.docx")
        safe_pdf = os.path.join(work, "output.pdf")
        shutil.copy2(docx_path, safe_docx)
        _export_pdf_with_word(safe_docx, safe_pdf)
        final_pdf = os.path.join(visual_dir, "rendered.pdf")
        shutil.copy2(safe_pdf, final_pdf)
        return final_pdf
    finally:
        shutil.rmtree(work, ignore_errors=True)


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


def check_visual(out_dir: str, output_docx_name: str = "最终论文.docx", project_root: str | None = None) -> Dict[str, Any]:
    out_dir = os.path.abspath(out_dir)
    docx_path = os.path.join(out_dir, output_docx_name)
    visual_dir = os.path.join(out_dir, "visual_qa")
    os.makedirs(visual_dir, exist_ok=True)

    issues: List[Dict[str, Any]] = []
    counts: Dict[str, Any] = {}
    artifacts: Dict[str, Any] = {}

    if not os.path.exists(docx_path):
        issues.append(_issue("MISSING_DOCX", "error", "Cannot run visual QA because final DOCX is missing.", docx_path))
    else:
        try:
            pdf_path = _export_pdf(docx_path, visual_dir)
            artifacts["pdf"] = pdf_path
            info = _pdfinfo(pdf_path)
            counts.update({k: v for k, v in info.items() if k in {"pages", "page_width_pt", "page_height_pt"}})
            if not info.get("available"):
                issues.append(_issue("PDFINFO_UNAVAILABLE", "error", "pdfinfo is not available; visual QA cannot verify page count or paper size."))
            elif info.get("error"):
                issues.append(_issue("PDFINFO_FAILED", "error", "pdfinfo failed.", str(info.get("error"))))
            if int(info.get("pages") or 0) <= 0:
                issues.append(_issue("PDF_PAGE_COUNT_INVALID", "error", "Rendered PDF has no pages."))

            text_tool_available = shutil.which("pdftotext") is not None
            pages_text = _pdf_pages_text(pdf_path, visual_dir)
            counts["text_pages"] = len([p for p in pages_text if p.strip()])
            if not text_tool_available:
                issues.append(_issue("PDFTOTEXT_UNAVAILABLE", "error", "pdftotext is not available; visual QA cannot verify rendered text."))
            elif not pages_text:
                issues.append(_issue("PDFTOTEXT_FAILED", "error", "pdftotext did not produce readable page text."))
            blank_pages = [idx + 1 for idx, page in enumerate(pages_text) if not page.strip()]
            if len(blank_pages) > max(2, int(info.get("pages") or 0) // 8):
                issues.append(_issue("MANY_BLANK_PAGES", "warning", "Rendered PDF has many text-empty pages.", ",".join(map(str, blank_pages[:12]))))
            if pages_text and not _find_page(pages_text, [r"目录", r"contents"]) and int(info.get("pages") or 0) >= 6:
                issues.append(_issue("TOC_TEXT_NOT_FOUND", "warning", "Rendered PDF text does not expose a TOC page."))

            samples = _sample_pages(int(info.get("pages") or 0), pages_text)
            rendered = _render_samples(pdf_path, visual_dir, samples)
            artifacts["samples"] = rendered
            counts["sample_pages"] = samples
            counts["sample_images"] = len(rendered)
            if samples and len(rendered) < len(samples):
                issues.append(_issue("SAMPLE_RENDER_FAILED", "error", "Could not render all PDF sample PNGs; install pdftoppm/Poppler."))
        except Exception as exc:
            issues.append(_issue("PDF_EXPORT_FAILED", "error", "DOCX could not be exported to PDF for visual QA.", str(exc)))

    passed = not any(i.get("severity") == "error" for i in issues)
    return {
        "schema_version": 1,
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "passed": passed,
        "output_dir_name": os.path.basename(out_dir),
        "counts": counts,
        "issues": issues,
        "artifacts": sanitize_value(artifacts, project_root),
    }


def report_to_markdown(report: Dict[str, Any]) -> str:
    lines = [
        "# Visual QA Report",
        "",
        f"- Result: {'passed' if report.get('passed') else 'failed'}",
        f"- Output: `{report.get('output_dir_name')}`",
        "",
        "## Counts",
        "",
    ]
    for key, value in sorted((report.get("counts") or {}).items()):
        lines.append(f"- `{key}`: {value}")
    lines.extend(["", "## Issues", ""])
    issues = report.get("issues") or []
    if not issues:
        lines.append("- No visual QA issues detected by automated checks.")
    else:
        for item in issues:
            lines.append(f"- **{item.get('severity')}** `{item.get('code')}`: {item.get('message')}")
            if item.get("detail"):
                lines.append(f"  Detail: `{item.get('detail')}`")
    lines.append("")
    return "\n".join(lines)


def write_reports(report: Dict[str, Any], out_dir: str) -> None:
    with open(os.path.join(out_dir, "visual_report.json"), "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)
    with open(os.path.join(out_dir, "visual_report.md"), "w", encoding="utf-8") as f:
        f.write(report_to_markdown(report))


def check_and_write(out_dir: str, output_docx_name: str = "最终论文.docx", project_root: str | None = None) -> Dict[str, Any]:
    report = check_visual(out_dir, output_docx_name=output_docx_name, project_root=project_root)
    write_reports(report, out_dir)
    return report


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Run optional visual QA on a generated DOCX output directory.")
    parser.add_argument("out_dir")
    parser.add_argument("--docx", default="最终论文.docx")
    args = parser.parse_args()
    result = check_and_write(args.out_dir, output_docx_name=args.docx)
    print(report_to_markdown(result))
    raise SystemExit(0 if result.get("passed") else 1)
