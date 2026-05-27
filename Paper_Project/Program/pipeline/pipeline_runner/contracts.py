"""Lightweight data contracts for the Word pipeline JSON handoffs."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Dict, Iterable, List, Mapping, TypedDict


class FormatData(TypedDict, total=False):
    _meta: Dict[str, Any]
    sections: List[Dict[str, Any]]
    paragraphs: List[Dict[str, Any]]
    tables: List[Dict[str, Any]]
    cover: List[Dict[str, Any]]
    style_profiles: Dict[str, Dict[str, Any]]


class ContentSection(TypedDict, total=False):
    heading: str
    level: int
    role: str
    paragraphs: List[Any]
    images: List[str]


class ContentData(TypedDict, total=False):
    _meta: Dict[str, Any]
    title_info: Dict[str, Any]
    sections: List[ContentSection]
    references: List[Any]


class BuildManifest(TypedDict, total=False):
    schema_version: int
    counts: Dict[str, int]


class QaIssue(TypedDict, total=False):
    code: str
    severity: str
    message: str
    detail: str
    active_owner: str


class QaReport(TypedDict, total=False):
    schema_version: int
    mode: str
    passed: bool
    counts: Dict[str, Any]
    issues: List[QaIssue]
    next_action: str
    repair_plan: Dict[str, Any]


@dataclass(frozen=True)
class ContractIssue:
    code: str
    message: str
    path: str = ""
    severity: str = "error"

    def to_dict(self) -> Dict[str, str]:
        return {
            "code": self.code,
            "severity": self.severity,
            "path": self.path,
            "message": self.message,
        }


def _is_mapping(value: Any) -> bool:
    return isinstance(value, Mapping)


def _is_list(value: Any) -> bool:
    return isinstance(value, list)


def _add(
    issues: List[ContractIssue],
    code: str,
    message: str,
    path: str,
    severity: str = "error",
) -> None:
    issues.append(ContractIssue(code=code, message=message, path=path, severity=severity))


def _require_mapping(data: Any, issues: List[ContractIssue], name: str) -> bool:
    if not _is_mapping(data):
        _add(issues, f"{name.upper()}_NOT_OBJECT", f"{name} must be a JSON object.", "$")
        return False
    return True


def _require_list(data: Mapping[str, Any], key: str, issues: List[ContractIssue], name: str) -> bool:
    if key not in data:
        _add(issues, f"{name.upper()}_{key.upper()}_MISSING", f"{name}.{key} is required.", f"$.{key}")
        return False
    if not _is_list(data.get(key)):
        _add(issues, f"{name.upper()}_{key.upper()}_NOT_LIST", f"{name}.{key} must be a list.", f"$.{key}")
        return False
    return True


def validate_format_data(data: Any) -> List[ContractIssue]:
    issues: List[ContractIssue] = []
    if not _require_mapping(data, issues, "format"):
        return issues
    for key in ("sections", "paragraphs", "tables"):
        _require_list(data, key, issues, "format")
    meta = data.get("_meta")
    if meta is not None and not _is_mapping(meta):
        _add(issues, "FORMAT_META_NOT_OBJECT", "format._meta should be an object.", "$._meta", "warning")
    sections = data.get("sections") if _is_list(data.get("sections")) else []
    if not sections:
        _add(issues, "FORMAT_SECTIONS_EMPTY", "format.sections should contain at least one section.", "$.sections")
    for idx, section in enumerate(sections[:20]):
        if not _is_mapping(section):
            _add(issues, "FORMAT_SECTION_NOT_OBJECT", "each format section should be an object.", f"$.sections[{idx}]")
    paragraphs = data.get("paragraphs") if _is_list(data.get("paragraphs")) else []
    for idx, para in enumerate(paragraphs[:50]):
        if not _is_mapping(para):
            _add(issues, "FORMAT_PARAGRAPH_NOT_OBJECT", "each format paragraph should be an object.", f"$.paragraphs[{idx}]")
            continue
        runs = para.get("runs")
        if runs is not None and not _is_list(runs):
            _add(issues, "FORMAT_PARAGRAPH_RUNS_NOT_LIST", "paragraph.runs should be a list when present.", f"$.paragraphs[{idx}].runs")
    return issues


def validate_content_data(data: Any) -> List[ContractIssue]:
    issues: List[ContractIssue] = []
    if not _require_mapping(data, issues, "content"):
        return issues
    _require_list(data, "sections", issues, "content")
    if "references" in data and not _is_list(data.get("references")):
        _add(issues, "CONTENT_REFERENCES_NOT_LIST", "content.references should be a list when present.", "$.references")
    meta = data.get("_meta")
    if meta is not None and not _is_mapping(meta):
        _add(issues, "CONTENT_META_NOT_OBJECT", "content._meta should be an object.", "$._meta", "warning")
    title_info = data.get("title_info")
    if title_info is not None and not _is_mapping(title_info):
        _add(issues, "CONTENT_TITLE_INFO_NOT_OBJECT", "content.title_info should be an object when present.", "$.title_info", "warning")
    sections = data.get("sections") if _is_list(data.get("sections")) else []
    if not sections:
        _add(issues, "CONTENT_SECTIONS_EMPTY", "content.sections should contain at least one section.", "$.sections")
    for idx, section in enumerate(sections[:50]):
        if not _is_mapping(section):
            _add(issues, "CONTENT_SECTION_NOT_OBJECT", "each content section should be an object.", f"$.sections[{idx}]")
            continue
        if "paragraphs" in section and not _is_list(section.get("paragraphs")):
            _add(issues, "CONTENT_SECTION_PARAGRAPHS_NOT_LIST", "section.paragraphs should be a list.", f"$.sections[{idx}].paragraphs")
        if "images" in section and not _is_list(section.get("images")):
            _add(issues, "CONTENT_SECTION_IMAGES_NOT_LIST", "section.images should be a list.", f"$.sections[{idx}].images")
    return issues


def validate_build_manifest(data: Any) -> List[ContractIssue]:
    issues: List[ContractIssue] = []
    if not _require_mapping(data, issues, "manifest"):
        return issues
    counts = data.get("counts")
    if not _is_mapping(counts):
        _add(issues, "MANIFEST_COUNTS_NOT_OBJECT", "manifest.counts must be an object.", "$.counts")
        return issues
    for key, value in counts.items():
        if not isinstance(value, int) or value < 0:
            _add(issues, "MANIFEST_COUNT_INVALID", "manifest count values must be non-negative integers.", f"$.counts.{key}")
    return issues


def validate_qa_report(data: Any) -> List[ContractIssue]:
    issues: List[ContractIssue] = []
    if not _require_mapping(data, issues, "qa_report"):
        return issues
    if not isinstance(data.get("passed"), bool):
        _add(issues, "QA_REPORT_PASSED_NOT_BOOL", "qa_report.passed must be a boolean.", "$.passed")
    if not _is_mapping(data.get("counts")):
        _add(issues, "QA_REPORT_COUNTS_NOT_OBJECT", "qa_report.counts must be an object.", "$.counts")
    issues_value = data.get("issues")
    if not _is_list(issues_value):
        _add(issues, "QA_REPORT_ISSUES_NOT_LIST", "qa_report.issues must be a list.", "$.issues")
        return issues
    for idx, item in enumerate(issues_value[:100]):
        if not _is_mapping(item):
            _add(issues, "QA_REPORT_ISSUE_NOT_OBJECT", "each QA issue should be an object.", f"$.issues[{idx}]")
            continue
        for key in ("code", "severity", "message"):
            if not str(item.get(key) or "").strip():
                _add(issues, f"QA_REPORT_ISSUE_{key.upper()}_MISSING", f"QA issue {key} is required.", f"$.issues[{idx}].{key}")
    return issues


def has_contract_errors(issues: Iterable[ContractIssue]) -> bool:
    return any(issue.severity == "error" for issue in issues)


def format_contract_issues(issues: Iterable[ContractIssue], limit: int = 8) -> List[str]:
    lines: List[str] = []
    for issue in list(issues)[:limit]:
        location = f" {issue.path}" if issue.path else ""
        lines.append(f"{issue.severity} {issue.code}{location}: {issue.message}")
    return lines
