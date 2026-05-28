"""Generated-DOCX helper routines for regression cases."""
from __future__ import annotations

import json
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from typing import Any, Dict

from qa_checker import check_output
from script_generator import generate

from regression_suite_modules.harness import assert_true, base_format, fail, new_workdir, write_json


def run_generated_case(name: str, content: Dict[str, Any], fmt: Dict[str, Any] | None = None) -> Dict[str, Any]:
    work = new_workdir(name)
    fmt_path = work / "format.json"
    cnt_path = work / "content.json"
    out_docx = work / "out.docx"
    write_json(fmt_path, fmt or base_format())
    write_json(cnt_path, content)
    write_json(work / "workflow_mode.json", {"mode": "developer"})

    generate(str(fmt_path), str(cnt_path), str(work), "out.docx")
    build_py = work / "build_generated.py"
    subprocess.run([sys.executable, "-m", "py_compile", str(build_py)], check=True, timeout=120)
    result = subprocess.run(
        [sys.executable, str(build_py)],
        cwd=str(work),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=180,
    )
    if result.returncode != 0:
        fail(f"{name}: generated build failed: {result.stderr[:1000] or result.stdout[:1000]}")
    assert_true(out_docx.exists(), f"{name}: output docx was not created")

    with zipfile.ZipFile(out_docx) as zf:
        xml = zf.read("word/document.xml").decode("utf-8", errors="replace")
    manifest = json.loads((work / "build_manifest.json").read_text(encoding="utf-8"))
    report = check_output(str(work), mode="developer", output_docx_name="out.docx")
    return {"work": work, "xml": xml, "manifest": manifest, "report": report}


def omath_count(xml: str) -> int:
    return len(re.findall(r"<(?:[A-Za-z_][\w.-]*:)?oMath\b", xml))


def omath_para_count(xml: str) -> int:
    return len(re.findall(r"<(?:[A-Za-z_][\w.-]*:)?oMathPara\b", xml))


def make_vml_picture_docx(src_docx: Path, dst_docx: Path) -> None:
    tmp = Path(tempfile.mkdtemp(prefix="wordchat_vml_zip_"))
    try:
        with zipfile.ZipFile(src_docx) as zf:
            zf.extractall(tmp)
        document_xml = tmp / "word" / "document.xml"
        xml = document_xml.read_text(encoding="utf-8")
        m = re.search(r'r:embed="([^"]+)"', xml)
        if not m:
            fail("source docx did not contain a drawing relationship")
        rid = m.group(1)
        xml = re.sub(
            r"<w:drawing>.*?</w:drawing>",
            f'<w:pict><v:shape><v:imagedata r:id="{rid}"/></v:shape></w:pict>',
            xml,
            count=1,
            flags=re.S,
        )
        document_xml.write_text(xml, encoding="utf-8")
        with zipfile.ZipFile(dst_docx, "w", zipfile.ZIP_DEFLATED) as zf:
            for path in tmp.rglob("*"):
                if path.is_file():
                    zf.write(path, path.relative_to(tmp).as_posix())
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
