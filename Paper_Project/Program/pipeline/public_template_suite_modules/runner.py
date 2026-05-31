"""Scenario execution for the public template compatibility suite."""
from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path
from typing import Any, Dict, Optional

from public_template_suite_modules.paths import PIPELINE_DIR, ROOT, RUNS_DIR
from public_template_suite_modules.scenarios import SCENARIOS, ScenarioBuilder
from public_template_suite_modules.storage import download_template, sha256_file, write_json

if str(PIPELINE_DIR) not in sys.path:
    sys.path.insert(0, str(PIPELINE_DIR))

from format_extractor import extract as extract_format
from qa_checker import check_output, write_reports as write_qa_reports
from qa_conformance import check_and_write as conformance_check_and_write
from qa_conformance import write_requirements as write_template_requirements
from script_generator import generate
from template_profiler import write_profile

try:
    from qa_visual import check_and_write as visual_check_and_write
except Exception:
    visual_check_and_write = None


def run_build(build_py: Path, cwd: Path) -> Dict[str, Any]:
    subprocess.run([sys.executable, "-m", "py_compile", str(build_py)], check=True, timeout=120)
    completed = subprocess.run(
        [sys.executable, str(build_py)],
        cwd=str(cwd),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=240,
    )
    return {
        "returncode": completed.returncode,
        "stdout_tail": (completed.stdout or "")[-2000:],
        "stderr_tail": (completed.stderr or "")[-2000:],
    }


def run_scenario(
    template_id: str,
    fmt: Dict[str, Any],
    scenario: Dict[str, Any],
    template_run_dir: Path,
    visual: bool = False,
    golden_dir: Optional[Path] = None,
    update_golden: bool = False,
) -> Dict[str, Any]:
    scenario_id = str(scenario["id"])
    scenario_dir = template_run_dir / scenario_id
    if scenario_dir.exists():
        shutil.rmtree(scenario_dir)
    scenario_dir.mkdir(parents=True, exist_ok=True)

    result: Dict[str, Any] = {
        "id": scenario_id,
        "description": scenario.get("description"),
        "qa_passed": False,
        "run_dir": str(scenario_dir.relative_to(ROOT)).replace("\\", "/"),
    }

    try:
        write_json(scenario_dir / "format.json", fmt)
        builder: ScenarioBuilder = scenario["builder"]
        content = builder(scenario_dir / "figures_src", scenario_dir)
        write_json(scenario_dir / "content.json", content)
        write_json(scenario_dir / "workflow_mode.json", {"mode": "developer", "template_id": template_id, "scenario": scenario_id})
        write_template_requirements(fmt, content, str(scenario_dir))
        generate(str(scenario_dir / "format.json"), str(scenario_dir / "content.json"), str(scenario_dir), "synthetic_output.docx")

        build_result = run_build(scenario_dir / "build_generated.py", scenario_dir)
        result["build_returncode"] = build_result["returncode"]
        if build_result["returncode"] != 0:
            result["error"] = (build_result["stderr_tail"] or build_result["stdout_tail"] or "build failed")[:2000]
            return result

        report = check_output(str(scenario_dir), mode="developer", output_docx_name="synthetic_output.docx")
        write_qa_reports(report, str(scenario_dir))
        issues = report.get("issues") or []
        result.update(
            {
                "qa_passed": bool(report.get("passed")),
                "qa_errors": sum(1 for issue in issues if issue.get("severity") == "error"),
                "qa_warnings": sum(1 for issue in issues if issue.get("severity") == "warning"),
                "qa_issues": len(issues),
                "content": {
                    "sections": len(content.get("sections") or []),
                    "paragraphs": content.get("_meta", {}).get("paragraphs"),
                    "tables": content.get("_meta", {}).get("tables_count"),
                    "images": content.get("_meta", {}).get("images_extracted"),
                    "references": len(content.get("references") or []),
                },
            }
        )
        conformance_report = conformance_check_and_write(str(scenario_dir), mode="developer", output_docx_name="synthetic_output.docx")
        conformance_issues = conformance_report.get("issues") or []
        result.update(
            {
                "conformance_passed": bool(conformance_report.get("passed")),
                "conformance_errors": sum(1 for issue in conformance_issues if issue.get("severity") == "error"),
                "conformance_warnings": sum(1 for issue in conformance_issues if issue.get("severity") == "warning"),
                "conformance_issues": len(conformance_issues),
                "conformance_counts": conformance_report.get("counts") or {},
            }
        )
        if not conformance_report.get("passed"):
            result["qa_passed"] = False
        if visual:
            if visual_check_and_write is None:
                result.update({"visual_passed": False, "visual_errors": 1, "visual_warnings": 0, "visual_issues": 1, "visual_error": "qa_visual is not available"})
                result["qa_passed"] = False
            else:
                visual_report = visual_check_and_write(
                    str(scenario_dir),
                    output_docx_name="synthetic_output.docx",
                    project_root=str(ROOT),
                    golden_dir=str(golden_dir) if golden_dir else None,
                    update_golden=update_golden,
                )
                visual_issues = visual_report.get("issues") or []
                result.update(
                    {
                        "visual_passed": bool(visual_report.get("passed")),
                        "visual_errors": sum(1 for issue in visual_issues if issue.get("severity") == "error"),
                        "visual_warnings": sum(1 for issue in visual_issues if issue.get("severity") == "warning"),
                        "visual_issues": len(visual_issues),
                        "visual_counts": visual_report.get("counts") or {},
                    }
                )
                if not visual_report.get("passed"):
                    result["qa_passed"] = False
    except Exception as exc:
        result["error"] = str(exc)
    return result


def run_template(
    item: Dict[str, Any],
    force_download: bool = False,
    scenario_filter: Optional[str] = None,
    visual: bool = False,
    golden_dir: Optional[Path] = None,
    update_golden: bool = False,
) -> Dict[str, Any]:
    docx_path = download_template(item, force=force_download)
    run_id = str(item.get("id") or docx_path.stem)
    template_run_dir = RUNS_DIR / run_id
    if template_run_dir.exists():
        shutil.rmtree(template_run_dir)
    template_run_dir.mkdir(parents=True, exist_ok=True)

    result = {k: item.get(k) for k in ("id", "name", "source", "url", "download_url", "license_hint")}
    result.update(
        {
            "file": str(docx_path.relative_to(ROOT)).replace("\\", "/") if docx_path.is_relative_to(ROOT) else str(docx_path),
            "bytes": docx_path.stat().st_size,
            "sha256": sha256_file(docx_path),
            "run_dir": str(template_run_dir.relative_to(ROOT)).replace("\\", "/"),
        }
    )

    fmt, fmt_md = extract_format(str(docx_path), output_dir=str(template_run_dir))
    write_json(template_run_dir / "format.json", fmt)
    (template_run_dir / "format_report.md").write_text(fmt_md, encoding="utf-8")
    profile = write_profile(fmt, str(template_run_dir), project_root=str(ROOT))
    result.update(
        {
            "format": {
                "paragraphs": len(fmt.get("paragraphs") or []),
                "tables": len(fmt.get("tables") or []),
                "sections": len(fmt.get("sections") or []),
                "cover_elements": len(fmt.get("cover") or []),
                "style_profiles": len(fmt.get("style_profiles") or {}),
            },
            "profile_risks": [k for k, v in (profile.get("risk_flags") or {}).items() if v],
            "scenarios": [],
        }
    )

    selected = [scenario for scenario in SCENARIOS if scenario_filter in (None, "", scenario["id"])]
    if not selected:
        raise RuntimeError(f"unknown scenario: {scenario_filter}")

    for scenario in selected:
        scenario_result = run_scenario(
            run_id,
            fmt,
            scenario,
            template_run_dir,
            visual=visual,
            golden_dir=golden_dir,
            update_golden=update_golden,
        )
        result["scenarios"].append(scenario_result)
        status = "PASS" if scenario_result.get("qa_passed") else "FAIL"
        visual_text = ""
        if visual:
            visual_text = f" visual={scenario_result.get('visual_issues', 'n/a')}"
        print(
            f"  {status} {run_id}/{scenario_result.get('id')} "
            f"issues={scenario_result.get('qa_issues', 'n/a')} "
            f"conformance={scenario_result.get('conformance_issues', 'n/a')}{visual_text}"
        )

    scenarios = result["scenarios"]
    result.update(
        {
            "qa_passed": bool(scenarios) and all(bool(s.get("qa_passed")) for s in scenarios),
            "qa_errors": sum(int(s.get("qa_errors") or 0) for s in scenarios),
            "qa_warnings": sum(int(s.get("qa_warnings") or 0) for s in scenarios),
            "qa_issues": sum(int(s.get("qa_issues") or 0) for s in scenarios),
            "conformance_errors": sum(int(s.get("conformance_errors") or 0) for s in scenarios),
            "conformance_warnings": sum(int(s.get("conformance_warnings") or 0) for s in scenarios),
            "conformance_issues": sum(int(s.get("conformance_issues") or 0) for s in scenarios),
            "visual_errors": sum(int(s.get("visual_errors") or 0) for s in scenarios),
            "visual_warnings": sum(int(s.get("visual_warnings") or 0) for s in scenarios),
            "visual_issues": sum(int(s.get("visual_issues") or 0) for s in scenarios),
        }
    )
    return result

