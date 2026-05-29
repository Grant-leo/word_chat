"""Structural QA check orchestration for generated pipeline outputs."""
from __future__ import annotations

from typing import Any, Dict, List

try:
    from qa_checker_modules.artifact_phase import (
        build_output_paths,
        run_artifact_checks,
    )
    from qa_checker_modules.content_phase import run_content_checks
    from qa_checker_modules.docx_phase import run_docx_checks
    from qa_checker_modules.format_phase import run_format_checks
    from qa_checker_modules.report_phase import build_report
    from qa_checker_modules.registry import OWNER_BY_CODE, VALID_MODES
    from qa_checker_modules.repair_guides import REPAIR_GUIDES
except ImportError:  # pragma: no cover - package-style imports
    from .artifact_phase import (
        build_output_paths,
        run_artifact_checks,
    )
    from .content_phase import run_content_checks
    from .docx_phase import run_docx_checks
    from .format_phase import run_format_checks
    from .report_phase import build_report
    from .registry import OWNER_BY_CODE, VALID_MODES
    from .repair_guides import REPAIR_GUIDES


USER_INPUT_AUTO_LEVELS = {
    "needs_user_file",
    "needs_user_input",
    "needs_user_confirmation",
    "optional_user_input",
}
USER_INPUT_CODES = {
    "CONTENT_EMPTY",
    "CONTENT_IMAGE_MISSING",
    "IMAGE_EXTRACT_FAILED",
    "MISSING_CONTENT_JSON",
    "MISSING_FORMAT_JSON",
    "PDF_TEMPLATE_UNSUPPORTED",
    "EXTRACTION_VERIFICATION_FAILED",
}


def _user_owner_for_code(code: str) -> str:
    auto_level = str((REPAIR_GUIDES.get(code) or {}).get("auto_level") or "")
    if code in USER_INPUT_CODES or auto_level in USER_INPUT_AUTO_LEVELS:
        return "User input/template file"
    return "Outputs/<run>/build_generated.py"


def _issue(code: str, severity: str, message: str, mode: str, detail: str = "") -> Dict[str, Any]:
    owner_dev = OWNER_BY_CODE.get(code, "script_generator.py")
    owner_user = _user_owner_for_code(code)
    return {
        "code": code,
        "severity": severity,
        "message": message,
        "detail": detail,
        "owner_user": owner_user,
        "owner_developer": owner_dev,
        "active_owner": owner_user if mode == "user" else owner_dev,
    }


def check_output(out_dir: str, mode: str = "user", output_docx_name: str = "最终论文.docx") -> Dict[str, Any]:
    mode = mode if mode in VALID_MODES else "user"
    issues: List[Dict[str, Any]] = []
    counts: Dict[str, Any] = {}

    def add(code: str, severity: str, message: str, detail: str = "") -> None:
        issues.append(_issue(code, severity, message, mode, detail))

    paths = build_output_paths(out_dir, output_docx_name)
    manifest_counts = run_artifact_checks(out_dir, paths, counts, add)
    run_format_checks(paths, counts, add)
    content = run_content_checks(out_dir, paths, counts, manifest_counts, add)
    run_docx_checks(paths, counts, content, manifest_counts, add)
    return build_report(out_dir, mode, counts, issues)


