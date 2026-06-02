# Release Validation Plan

This plan is the publish-before-you-ship gate for the Word paper pipeline. It uses only synthetic or local developer-owned inputs; do not commit `Inputs/`, `Templates/`, `Outputs/`, generated DOCX/PDF/PNG, visual renders, or customer files.

## Gate 1: Environment

Run:

```powershell
python --version
python -c "import docx; from PIL import Image; print('OK')"
```

Pass criteria:

- Python is 3.10+.
- `python-docx` and `Pillow` import successfully.

## Gate 2: Static Engine Hygiene

Run:

```powershell
python -m compileall -q run_pipeline.py Paper_Project\Program\pipeline
git diff --check
python scripts\project_memory.py validate
```

Pass criteria:

- No syntax errors.
- No whitespace errors from `git diff --check`.
- Memory validation is OK.

## Gate 3: Synthetic Regression Suite

Run:

```powershell
python Paper_Project\Program\pipeline\regression_suite.py
```

Pass criteria:

- All synthetic cases pass.
- The suite covers CLI routing, extraction verification, Markdown, DOCX/PDF template extraction, formulas, tables, images, QA routing, conformance QA, visual dependency failures, auto-repair loop behavior, and privacy/path sanitization.

## Gate 4: User-Experience Failure Routing

Check these targeted cases when QA or repair-loop code changes:

```powershell
python Paper_Project\Program\pipeline\regression_suite.py --filter pipeline_agent_auto_guides_missing_and_ambiguous_inputs
python Paper_Project\Program\pipeline\regression_suite.py --filter pipeline_agent_summary
python Paper_Project\Program\pipeline\regression_suite.py --filter pipeline_qa_points_to_conformance_report_when_strict_fails
python Paper_Project\Program\pipeline\regression_suite.py --filter pipeline_qa_writes_report_when_conformance_dependency_missing
python Paper_Project\Program\pipeline\regression_suite.py --filter pipeline_qa_writes_report_when_visual_dependency_missing
python Paper_Project\Program\pipeline\regression_suite.py --filter pipeline_auto_repair_blocks_unrepairable_visual_errors
python Paper_Project\Program\pipeline\regression_suite.py --filter run_pipeline_writes_repair_plan_for_build_failure
python Paper_Project\Program\pipeline\regression_suite.py --filter pipeline_agent_summary_surfaces_repair_loop_next_action
python Paper_Project\Program\pipeline\regression_suite.py --filter qa_repair_guides_cover_registered_issue_codes
```

Pass criteria:

- Every failed QA phase writes a matching Markdown/JSON report.
- Generated-script build failures write `qa_report`, `qa_repair_plan`, `qa_fix_prompt`, and `agent_summary` with a concrete `MISSING_DOCX` next action.
- Terminal output points to the correct report.
- `--agent-auto` auto-selects only single candidates, refuses ambiguous choices, and writes `agent_summary.md/json`.
- Registered structural QA issue codes have user-facing repair guides.
- Auto-repair stops with a clear reason when it needs user input, dependencies, or manual visual confirmation.
- Stopped auto-repair reports expose `next_action`, `resume_scope`, and `resume_command`, and `agent_summary` surfaces that next action.

## Gate 4.5: Fresh-Folder Agent Smoke

Create a disposable copy of the repository outside the git checkout, with only:

- one DOCX/PDF template in `Templates/`
- one DOCX/Markdown content file in `Inputs/`

Run the ordinary-user Agent entry:

```powershell
python run_pipeline.py --agent-auto
```

Pass criteria:

- The run does not ask the user to type Python commands.
- The single template/content pair is selected automatically.
- `workflow_mode.json` records `mode=user`, `auto_repair=true`, and `agent_auto=true`.
- `agent_summary.md/json` exists and is the first handoff artifact.
- Structural QA and strict conformance QA have 0 errors.
- The final DOCX exists and the summary tells the user to open it in Word/WPS for visual review.

## Gate 5: Matrix Compatibility

For release candidates, run a private local matrix without committing artifacts:

- Text domains: humanities, business, engineering, science, medicine.
- Template types: DOCX thesis template, Markdown instruction template, PDF instruction/visual template.
- QA levels: `strict` for all combinations; `visual` for representative DOCX/PDF cases when Word COM and Poppler are available.
- Repair mode: at least one `--mode user --auto-repair --qa-level strict` run and one `--mode user --auto-repair --qa-level visual` run.

Pass criteria per run:

- `qa_report.json`: 0 errors.
- `conformance_report.json`: 0 errors for strict/visual runs.
- `visual_report.json`: 0 errors for visual runs; warnings must be reviewed.
- `build_manifest.json` rendered counts meet or exceed expected content counts for images, tables, and formulas.
- Final DOCX opens in Word/WPS and visually preserves title, abstracts, TOC, headings, formulas, images, tables, references, appendices, and pagination.

## Gate 6: Bad-Input Boundary

Use synthetic or developer-owned samples for:

- Markdown front format-instruction block with a noisy or encoding-damaged heading, followed by a `---` delimiter and real content.
- Missing Markdown image.
- Header/footer content image.
- Table-cell image.
- Low-resolution image shard.
- Duplicate abstract/front-matter heading.
- Long paragraph truncation.
- Formula-heavy body with inline/display formulas.
- Scanned or text-poor PDF template.
- Corrupt or incomplete generated DOCX.

Pass criteria:

- Engine either produces a correct DOCX or fails closed with a QA report and actionable next step.
- User-file problems are routed to input/template repair, not blind generated-script editing.
- Developer engine problems name the owning module.
- Format-instruction paragraphs must not leak into `content.json`, `内容提取.md`, or strict conformance expected-content checks.
- DOCX template-only notes, source-TOC examples, cover field hints, and TOC page-number samples must not leak into the final DOCX.
- Common backmatter heading equivalents such as `Acknowledgements` / `Acknowledgment` / `致谢`, `References` / `参考文献`, and `Appendix` / `附录` must not create false `CONTENT_HEADING_MISSING` warnings.
- Public-template visual smoke should run without golden comparison by default; compare golden baselines only when `--golden-dir` or `--update-golden` is explicitly selected.

For parser/QA/release-candidate changes, also run the high-risk matrix that covers pure Markdown strict, missing Markdown images, header/footer image boundaries, user auto-repair, DOCX/PDF visual smoke, and dense media/math strict content. The expected release gate is that every case either passes or fails closed with the intended QA code and next-step guidance.

## Gate 7: Repository Hygiene

Before git:

```powershell
git status --short --ignored
```

Pass criteria:

- Only core scripts/docs intended for release are modified.
- No private `Inputs/`, `Templates/`, `Outputs/`, generated DOCX/PDF/PNG, visual QA renders, cache directories, or memory files are staged unless explicitly intended.
