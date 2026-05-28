"""Optional PDF/render QA entry point for generated DOCX outputs."""
from __future__ import annotations

from typing import Any, Dict

try:
    from qa_visual_modules import checks as _checks
    from qa_visual_modules.exporters import _export_pdf
    from qa_visual_modules.pdf_tools import _pdf_pages_text, _pdfinfo, _render_samples, _sample_pages
    from qa_visual_modules.reports import report_to_markdown, write_reports
except ImportError:  # pragma: no cover - package-style imports
    from .qa_visual_modules import checks as _checks
    from .qa_visual_modules.exporters import _export_pdf
    from .qa_visual_modules.pdf_tools import _pdf_pages_text, _pdfinfo, _render_samples, _sample_pages
    from .qa_visual_modules.reports import report_to_markdown, write_reports

__all__ = [
    "_export_pdf",
    "_pdf_pages_text",
    "_pdfinfo",
    "_render_samples",
    "_sample_pages",
    "check_and_write",
    "check_visual",
    "report_to_markdown",
    "write_reports",
]


def check_visual(
    out_dir: str,
    output_docx_name: str = "最终论文.docx",
    project_root: str | None = None,
    render_all_pages: bool = True,
    require_wps: bool = False,
    golden_dir: str | None = None,
    update_golden: bool = False,
) -> Dict[str, Any]:
    """Run visual QA while preserving legacy monkeypatch hooks on qa_visual."""
    _checks._export_pdf = _export_pdf
    _checks._pdfinfo = _pdfinfo
    _checks._pdf_pages_text = _pdf_pages_text
    _checks._render_samples = _render_samples
    return _checks.check_visual(
        out_dir,
        output_docx_name=output_docx_name,
        project_root=project_root,
        render_all_pages=render_all_pages,
        require_wps=require_wps,
        golden_dir=golden_dir,
        update_golden=update_golden,
    )


def check_and_write(
    out_dir: str,
    output_docx_name: str = "最终论文.docx",
    project_root: str | None = None,
    render_all_pages: bool = True,
    require_wps: bool = False,
    golden_dir: str | None = None,
    update_golden: bool = False,
) -> Dict[str, Any]:
    report = check_visual(
        out_dir,
        output_docx_name=output_docx_name,
        project_root=project_root,
        render_all_pages=render_all_pages,
        require_wps=require_wps,
        golden_dir=golden_dir,
        update_golden=update_golden,
    )
    write_reports(report, out_dir)
    return report

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Run optional visual QA on a generated DOCX output directory.")
    parser.add_argument("out_dir")
    parser.add_argument("--docx", default="最终论文.docx")
    parser.add_argument("--sample-only", action="store_true", help="Render only sample pages instead of every page.")
    parser.add_argument("--require-wps", action="store_true", help="Fail visual QA if WPS export is unavailable.")
    parser.add_argument("--golden-dir", default=None, help="Directory containing golden baseline JSON files.")
    parser.add_argument("--update-golden", action="store_true", help="Create/update the golden baseline for this output.")
    args = parser.parse_args()
    result = check_and_write(
        args.out_dir,
        output_docx_name=args.docx,
        render_all_pages=not args.sample_only,
        require_wps=args.require_wps,
        golden_dir=args.golden_dir,
        update_golden=args.update_golden,
    )
    print(report_to_markdown(result))
    raise SystemExit(0 if result.get("passed") else 1)
