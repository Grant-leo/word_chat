"""
public_template_suite.py - local public-template compatibility checks.

The corpus lives under TestData/PublicTemplates:
- manifest.json records public source URLs and hashes.
- files/ stores downloaded DOCX templates and is ignored by Git.
- runs/ stores generated outputs and is ignored by Git.

The suite uses only synthetic non-private content.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Any, Dict, List

from public_template_suite_modules.paths import (
    FILES_DIR,
    MANIFEST_PATH,
    PIPELINE_DIR,
    ROOT,
    RUNS_DIR,
    SELECTION_SUMMARY_PATH,
    TEST_ROOT,
)

if str(PIPELINE_DIR) not in sys.path:
    sys.path.insert(0, str(PIPELINE_DIR))

from format_extractor import extract as extract_format
from public_template_suite_modules.reporting import write_readme
from public_template_suite_modules.runner import run_template
from public_template_suite_modules.scenarios import DEFAULT_TEMPLATES, SCENARIOS
from public_template_suite_modules.storage import (
    read_manifest,
    resolve_existing_file,
    write_json,
)

def main() -> None:
    parser = argparse.ArgumentParser(description="Run public template compatibility checks.")
    parser.add_argument("--download", action="store_true", help="Download missing public templates before running.")
    parser.add_argument("--force-download", action="store_true", help="Re-download all templates.")
    parser.add_argument("--template", help="Run only one template id.")
    parser.add_argument("--scenario", help="Run only one scenario id.")
    parser.add_argument("--visual", action="store_true", help="Also run PDF/PNG visual QA for each generated DOCX.")
    parser.add_argument("--golden-dir", default=str(ROOT / "TestData" / "GoldenBaselines"), help="Directory for visual golden baseline JSON records.")
    parser.add_argument("--update-golden", action="store_true", help="Create or refresh visual golden baseline JSON records during visual QA.")
    args = parser.parse_args()

    TEST_ROOT.mkdir(parents=True, exist_ok=True)
    golden_dir = Path(args.golden_dir) if args.golden_dir else None
    manifest = read_manifest()
    templates = manifest.get("templates") or DEFAULT_TEMPLATES
    filtered_run = bool(args.template or args.scenario)
    if args.template:
        templates = [item for item in templates if str(item.get("id")) == args.template]
        if not templates:
            raise SystemExit(f"unknown template: {args.template}")

    results: List[Dict[str, Any]] = []
    failed = 0
    for item in templates:
        try:
            if args.download or args.force_download:
                result = run_template(
                    item,
                    force_download=args.force_download,
                    scenario_filter=args.scenario,
                    visual=args.visual,
                    golden_dir=golden_dir,
                    update_golden=args.update_golden,
                )
            else:
                expected = resolve_existing_file(item) or FILES_DIR / (item.get("name") or f"{item.get('id')}.docx")
                if not expected.exists():
                    raise RuntimeError(f"template file missing; rerun with --download: {expected}")
                result = run_template(
                    item,
                    force_download=False,
                    scenario_filter=args.scenario,
                    visual=args.visual,
                    golden_dir=golden_dir,
                    update_golden=args.update_golden,
                )
        except Exception as exc:
            result = {k: item.get(k) for k in ("id", "name", "source", "url", "download_url", "license_hint")}
            result.update({"qa_passed": False, "error": str(exc), "scenarios": []})
        if not result.get("qa_passed"):
            failed += 1
        status = "PASS" if result.get("qa_passed") else "FAIL"
        print(
            f"{status} {result.get('id')} scenarios={len(result.get('scenarios') or [])} "
            f"errors={result.get('qa_errors', 'n/a')} warnings={result.get('qa_warnings', 'n/a')}"
        )
        results.append(result)

    summary = {
        "schema_version": 1,
        "purpose": "Public non-private Word template regression corpus for local compatibility testing.",
        "files_ignored_by_git": True,
        "visual_qa_enabled": bool(args.visual),
        "golden_dir": str(golden_dir.relative_to(ROOT)).replace("\\", "/") if golden_dir and golden_dir.is_relative_to(ROOT) else (str(golden_dir) if golden_dir else None),
        "update_golden": bool(args.update_golden),
        "scenarios": [{k: scenario[k] for k in ("id", "description")} for scenario in SCENARIOS],
        "templates": results,
    }
    if filtered_run:
        write_json(SELECTION_SUMMARY_PATH, summary)
        print(f"Selection summary written to {SELECTION_SUMMARY_PATH.relative_to(ROOT)}")
    else:
        write_json(MANIFEST_PATH, summary)
        write_readme(results)
    print(f"RESULT passed={len(results) - failed} failed={failed} scenario_count={len(SCENARIOS)}")
    raise SystemExit(1 if failed else 0)

if __name__ == "__main__":
    main()
