"""Strict DOCX conformance QA entry point."""
from __future__ import annotations

from typing import Any, Dict

try:
    from qa_conformance_modules.checks import check_conformance
    from qa_conformance_modules.registry import VALID_MODES
    from qa_conformance_modules.reports import report_to_markdown, write_reports
    from qa_conformance_modules.requirements import build_requirements, requirements_to_markdown, write_requirements
except ImportError:  # pragma: no cover - package-style imports
    from .qa_conformance_modules.checks import check_conformance
    from .qa_conformance_modules.registry import VALID_MODES
    from .qa_conformance_modules.reports import report_to_markdown, write_reports
    from .qa_conformance_modules.requirements import build_requirements, requirements_to_markdown, write_requirements

__all__ = [
    "build_requirements",
    "check_and_write",
    "check_conformance",
    "report_to_markdown",
    "requirements_to_markdown",
    "write_reports",
    "write_requirements",
]


def check_and_write(out_dir: str, mode: str = "user", output_docx_name: str = "最终论文.docx", project_root: str | None = None) -> Dict[str, Any]:
    report = check_conformance(out_dir, mode=mode, output_docx_name=output_docx_name, project_root=project_root)
    write_reports(report, out_dir)
    return report

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Run strict DOCX conformance QA on a generated output directory.")
    parser.add_argument("out_dir")
    parser.add_argument("--mode", choices=sorted(VALID_MODES), default="user")
    parser.add_argument("--docx", default="最终论文.docx")
    args = parser.parse_args()

    result = check_and_write(args.out_dir, mode=args.mode, output_docx_name=args.docx)
    print(report_to_markdown(result))
    raise SystemExit(0 if result.get("passed") else 1)
