"""
qa_checker.py - lightweight QA checks for generated Word pipeline outputs.

The checker does not fix files. It writes a structured report that tells the AI
which artifact should be edited in user mode and which core engine owns the same
class of issue in developer mode.
"""
from __future__ import annotations

from typing import Any, Dict

try:
    from qa_checker_modules.checks import check_output
    from qa_checker_modules.registry import VALID_MODES
    from qa_checker_modules.reports import report_to_markdown, write_reports
except ImportError:  # pragma: no cover - package-style imports
    from .qa_checker_modules.checks import check_output
    from .qa_checker_modules.registry import VALID_MODES
    from .qa_checker_modules.reports import report_to_markdown, write_reports

def check_and_write(out_dir: str, mode: str = "user", output_docx_name: str = "最终论文.docx") -> Dict[str, Any]:
    report = check_output(out_dir, mode=mode, output_docx_name=output_docx_name)
    write_reports(report, out_dir)
    return report

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Check generated pipeline output.")
    parser.add_argument("out_dir")
    parser.add_argument("--mode", choices=sorted(VALID_MODES), default="user")
    parser.add_argument("--docx", default="最终论文.docx")
    args = parser.parse_args()

    result = check_and_write(args.out_dir, mode=args.mode, output_docx_name=args.docx)
    print(report_to_markdown(result))
    raise SystemExit(0 if result.get("passed") else 1)
