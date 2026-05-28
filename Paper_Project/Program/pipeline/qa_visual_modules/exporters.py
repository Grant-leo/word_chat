"""DOCX-to-PDF export helpers for visual QA."""
from __future__ import annotations

import json
import os
import shutil
import subprocess
import tempfile
from typing import List


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


def _export_pdf_with_wps(docx_path: str, pdf_path: str) -> str:
    powershell = shutil.which("powershell") or shutil.which("pwsh")
    if not powershell:
        raise RuntimeError("PowerShell is not available; cannot drive WPS COM.")

    script = f"""
$ErrorActionPreference = 'Stop'
$docx = {json.dumps(os.path.abspath(docx_path))}
$pdf = {json.dumps(os.path.abspath(pdf_path))}
$progIds = @('KWPS.Application', 'WPS.Application')
$wps = $null
$last = $null
foreach ($progId in $progIds) {{
  try {{
    $wps = New-Object -ComObject $progId
    break
  }} catch {{
    $last = $_.Exception.Message
  }}
}}
if ($wps -eq $null) {{
  throw "WPS COM is not available: $last"
}}
$wps.Visible = $false
$doc = $null
try {{
  $doc = $wps.Documents.Open($docx, $false, $true)
  try {{
    $doc.ExportAsFixedFormat($pdf, 17)
  }} catch {{
    $doc.SaveAs($pdf, 17)
  }}
  $doc.Close($false)
  $doc = $null
}} finally {{
  if ($doc -ne $null) {{
    try {{ $doc.Close($false) }} catch {{ }}
  }}
  if ($wps -ne $null) {{
    try {{ $wps.Quit() }} catch {{ }}
  }}
}}
"""
    with tempfile.TemporaryDirectory(prefix="wordchat_wps_ps_") as td:
        ps1 = os.path.join(td, "export_wps.ps1")
        with open(ps1, "w", encoding="utf-8-sig") as f:
            f.write(script)
        result = _run([powershell, "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", ps1], timeout=180)
    if result.returncode != 0:
        raise RuntimeError((result.stderr or result.stdout or "WPS COM export failed").strip()[:1000])
    if not os.path.exists(pdf_path):
        raise RuntimeError("WPS COM finished but PDF was not created.")
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


def _export_wps_pdf(docx_path: str, visual_dir: str) -> str:
    work = tempfile.mkdtemp(prefix="wordchat_wps_docx_")
    try:
        safe_docx = os.path.join(work, "input.docx")
        safe_pdf = os.path.join(work, "output.pdf")
        shutil.copy2(docx_path, safe_docx)
        _export_pdf_with_wps(safe_docx, safe_pdf)
        final_pdf = os.path.join(visual_dir, "rendered_wps.pdf")
        shutil.copy2(safe_pdf, final_pdf)
        return final_pdf
    finally:
        shutil.rmtree(work, ignore_errors=True)

