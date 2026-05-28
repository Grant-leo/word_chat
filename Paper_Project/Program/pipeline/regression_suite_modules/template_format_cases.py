"""Template profiling and format extraction regression cases."""
from __future__ import annotations

import json

from docx import Document
from format_extractor import extract as extract_docx_format
from regression_suite_modules.harness import assert_true, base_format, case, new_workdir
from template_profiler import profile_format


@case
def template_profile_sanitizes_private_source() -> None:
    fmt = base_format(source="private_school_template.docx")
    fmt["_meta"]["assets_dir"] = r"E:\private\Templates\private_school_template_assets"
    profile = profile_format(fmt, project_root=r"E:\private")
    text = json.dumps(profile, ensure_ascii=False)
    assert_true("private_school_template" not in text, "template profile leaked source filename")
    assert_true(profile["source"]["source_ext"] == ".docx", "template profile lost source extension")
    assert_true(profile["source"]["has_assets_dir"] is True, "template profile lost assets flag")


@case
def format_extractor_stops_cover_before_spaced_abstract_heading() -> None:
    work = new_workdir("format_cover_stop_abstract")
    docx = work / "cover_stop.docx"
    doc = Document()
    doc.add_paragraph("Cover title")
    doc.add_paragraph("摘  要")
    doc.add_paragraph("Template abstract sample paragraph should not be replayed as cover.")
    doc.save(docx)

    fmt, _ = extract_docx_format(str(docx))
    cover_text = "\n".join(
        "".join(run.get("t", "") for run in el.get("r", []))
        for el in fmt.get("cover") or []
        if isinstance(el, dict)
    )
    assert_true("Cover title" in cover_text, "cover text before abstract was not extracted")
    assert_true("摘" not in cover_text and "Template abstract sample" not in cover_text, "abstract page leaked into cover extraction")

