"""Aggregate QA/comparison evidence into an auditable delivery decision."""
from __future__ import annotations

import argparse
import json
import os
from datetime import datetime
from typing import Any, Dict, Iterable, List


DECISIONS = {
    "BLOCKED_UNJUDGEABLE",
    "FAILED_AUTOMATIC",
    "PASSED_WITH_REVIEW",
    "PASSED_MACHINE",
    "APPROVED_DELIVERY",
}
REPORT_FILES = (
    ("structural", "qa_report.json"),
    ("strict", "conformance_report.json"),
    ("visual", "visual_report.json"),
    ("reference_compare", "reference_compare_report.json"),
)
UNJUDGEABLE_CODES = {
    "STRUCTURAL_QA_UNAVAILABLE",
    "STRUCTURAL_QA_FAILED",
    "CONFORMANCE_QA_FAILED",
    "VISUAL_QA_FAILED",
    "PDFINFO_UNAVAILABLE",
    "PAGE_IMAGE_UNREADABLE",
}


def _load_json(path: str) -> Dict[str, Any] | None:
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else None
    except Exception:
        return None


def _issues(report: Dict[str, Any]) -> List[Dict[str, Any]]:
    issues = report.get("issues")
    if isinstance(issues, list):
        return [item for item in issues if isinstance(item, dict)]
    qa = report.get("qa")
    if isinstance(qa, dict) and isinstance(qa.get("issues"), list):
        return [item for item in qa.get("issues") if isinstance(item, dict)]
    return []


def _status(report: Dict[str, Any], issues: List[Dict[str, Any]]) -> str:
    value = str(report.get("status") or "").strip().lower()
    if value in {"passed", "passed_with_warnings", "failed"}:
        return value
    if report.get("passed") is False:
        return "failed"
    severities = {str(issue.get("severity") or "").lower() for issue in issues}
    if "error" in severities:
        return "failed"
    if "warning" in severities:
        return "passed_with_warnings"
    if report.get("passed") is True:
        return "passed"
    return "missing"


def _golden_evidence(report: Dict[str, Any]) -> Dict[str, Any] | None:
    artifacts = report.get("artifacts") or {}
    if isinstance(artifacts, dict) and isinstance(artifacts.get("golden_baseline"), dict):
        return artifacts.get("golden_baseline")
    golden = report.get("golden_baseline")
    return golden if isinstance(golden, dict) else None


def _collect_reports(out_dir: str) -> List[Dict[str, Any]]:
    collected: List[Dict[str, Any]] = []
    for phase, filename in REPORT_FILES:
        path = os.path.join(out_dir, filename)
        data = _load_json(path)
        if data is None:
            continue
        issues = _issues(data)
        collected.append(
            {
                "phase": phase,
                "path": filename,
                "status": _status(data, issues),
                "result_label": data.get("result_label"),
                "issues": issues,
                "golden_baseline": _golden_evidence(data),
            }
        )
    return collected


def _blocking_codes(reports: Iterable[Dict[str, Any]]) -> List[str]:
    codes: List[str] = []
    for report in reports:
        for issue in report.get("issues") or []:
            code = str(issue.get("code") or "").strip()
            if code and code not in codes:
                codes.append(code)
    return codes


def _review_pages(reports: Iterable[Dict[str, Any]]) -> List[int]:
    pages: List[int] = []
    for report in reports:
        for issue in report.get("issues") or []:
            for key in ("page", "pages", "review_pages"):
                value = issue.get(key)
                if isinstance(value, int):
                    pages.append(value)
                elif isinstance(value, list):
                    pages.extend(int(x) for x in value if isinstance(x, int) or str(x).isdigit())
    return sorted(set(page for page in pages if page > 0))


def _baseline_type(reports: Iterable[Dict[str, Any]]) -> str:
    has_visual = False
    has_golden = False
    has_reference = False
    for report in reports:
        if report.get("phase") == "visual":
            has_visual = True
        if report.get("phase") == "reference_compare":
            has_reference = True
        golden = report.get("golden_baseline") or {}
        if golden.get("enabled"):
            has_golden = True
    if has_reference:
        return "reference_compare"
    if has_golden:
        return "golden_visual"
    if has_visual:
        return "visual_render"
    return "machine_reports"


def _confidence(reports: List[Dict[str, Any]], decision: str) -> str:
    phases = {report.get("phase") for report in reports}
    if decision in {"FAILED_AUTOMATIC", "BLOCKED_UNJUDGEABLE"}:
        return "high" if phases else "low"
    if {"structural", "strict", "visual"} <= phases and decision == "PASSED_MACHINE":
        return "high"
    if {"structural", "strict"} <= phases:
        return "medium"
    return "low"


def assess_run(out_dir: str, approved_deviations: List[Dict[str, Any]] | None = None) -> Dict[str, Any]:
    """Assess one output directory using available QA/comparison reports."""
    out_dir = os.path.abspath(out_dir)
    reports = _collect_reports(out_dir)
    approved_deviations = approved_deviations or []
    codes = _blocking_codes(reports)
    statuses = {str(report.get("status")) for report in reports}
    warning_present = any(
        str(issue.get("severity") or "").lower() == "warning"
        for report in reports for issue in report.get("issues") or []
    )
    error_present = any(
        str(issue.get("severity") or "").lower() == "error"
        for report in reports for issue in report.get("issues") or []
    )
    if not reports:
        decision = "BLOCKED_UNJUDGEABLE"
        manual_review_required = True
        codes = ["QA_REPORTS_MISSING"]
    elif any(code in UNJUDGEABLE_CODES for code in codes):
        decision = "BLOCKED_UNJUDGEABLE"
        manual_review_required = True
    elif error_present or "failed" in statuses:
        decision = "FAILED_AUTOMATIC"
        manual_review_required = True
    elif warning_present or "passed_with_warnings" in statuses:
        decision = "PASSED_WITH_REVIEW"
        manual_review_required = True
    elif approved_deviations:
        decision = "APPROVED_DELIVERY"
        manual_review_required = False
    else:
        decision = "PASSED_MACHINE"
        manual_review_required = False

    assessment = {
        "schema_version": 1,
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "decision": decision,
        "confidence": _confidence(reports, decision),
        "baseline_type": _baseline_type(reports),
        "manual_review_required": manual_review_required,
        "review_pages": _review_pages(reports),
        "approved_deviations": approved_deviations,
        "blocking_issue_codes": codes,
        "reports": [
            {
                "phase": report.get("phase"),
                "path": report.get("path"),
                "status": report.get("status"),
                "result_label": report.get("result_label"),
                "issue_codes": [
                    str(issue.get("code")) for issue in report.get("issues") or []
                    if issue.get("code")
                ],
            }
            for report in reports
        ],
    }
    return assessment


def write_assessment(out_dir: str, approved_deviations: List[Dict[str, Any]] | None = None) -> Dict[str, str]:
    """Write comparison assessment JSON/Markdown reports for one run."""
    assessment = assess_run(out_dir, approved_deviations=approved_deviations)
    json_path = os.path.join(out_dir, "comparison_assessment.json")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(assessment, f, ensure_ascii=False, indent=2)
    md_path = os.path.join(out_dir, "comparison_assessment.md")
    with open(md_path, "w", encoding="utf-8") as f:
        f.write("# Comparison Assessment\n\n")
        f.write(f"- decision: `{assessment['decision']}`\n")
        f.write(f"- confidence: `{assessment['confidence']}`\n")
        f.write(f"- baseline_type: `{assessment['baseline_type']}`\n")
        f.write(f"- manual_review_required: `{assessment['manual_review_required']}`\n")
        f.write(f"- review_pages: `{assessment['review_pages']}`\n")
        f.write(f"- blocking_issue_codes: `{', '.join(assessment['blocking_issue_codes']) or '-'}`\n")
        f.write("\n## Reports\n")
        for report in assessment.get("reports") or []:
            codes = ", ".join(report.get("issue_codes") or []) or "-"
            f.write(f"- `{report.get('phase')}` `{report.get('status')}` codes: `{codes}`\n")
    return {"json_path": json_path, "md_path": md_path, "decision": assessment["decision"]}


def main(argv: List[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Summarize QA/comparison reports into a delivery decision.")
    parser.add_argument("out_dir", help="Pipeline output directory.")
    parser.add_argument("--save", action="store_true", help="Write comparison_assessment.json/md in the output directory.")
    args = parser.parse_args(argv)
    assessment = assess_run(args.out_dir)
    if args.save:
        write_assessment(args.out_dir)
    print(json.dumps(assessment, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
