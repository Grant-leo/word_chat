---
name: word-paper-pipeline
description: Use for the Word paper typesetting project when running or fixing the DOCX pipeline, validating template/content matrices, reviewing QA reports, preparing demo runs, cleaning Outputs, or making reusable engine changes under Paper_Project/Program/pipeline.
---

# Word Paper Pipeline

## Purpose

This skill keeps work on `E:\career\word_chat_v1` repeatable: load the project toolbox, keep private DOCX data local, route fixes to the right layer, and verify with structural and strict conformance QA before reporting.

This file is the repository source copy. A Codex runtime copy may also exist under `$CODEX_HOME/skills/word-paper-pipeline/SKILL.md` or `%USERPROFILE%\.codex\skills\word-paper-pipeline\SKILL.md`. If the runtime copy is missing, read this tracked source file and continue instead of treating the project workflow as unavailable.

## Required Startup

1. Run:
   ```powershell
   python --version
   python -c "import docx, lxml; from PIL import Image; print('OK')"
   ```
2. Read `AGENTS.md`.
3. Read `Paper_Project/基础操作.md`.
4. Read `memory/PROJECT_MEMORY.md` and `memory/active_context.md` when present.
5. Check `Templates/*.docx`, `Templates/*.pdf`, `Inputs/*.docx`, and `Inputs/*.md`.

## Mode Decision

- Ordinary user fine-tuning: edit only `Outputs/<latest>/build_generated.py`.
- Reusable engine or product behavior: edit only core files under `Paper_Project/Program/pipeline/`, then rerun the full pipeline.
- If unsure, inspect `Outputs/<latest>/workflow_mode.json`; otherwise default to user mode unless the request is clearly developer or maintainer work.

## Pipeline Verification

Use strict QA for engine confidence:

```powershell
python run_pipeline.py --mode developer --qa-level strict --template <template.docx> --content <content.docx>
```

For visual delivery gates, use:

```powershell
python run_pipeline.py --mode developer --qa-level visual --template <template.docx> --content <content.docx>
```

After each run, inspect:

- `agent_summary.md/json`
- `qa_report.md/json`
- `qa_repair_plan.md/json`
- `template_profile.md/json`
- `conformance_report.md/json` when strict or visual QA is used
- `visual_report.md/json` when visual QA is used
- `build_manifest.json`
- `最终论文.docx`

## Productization Rules

- Any preflight, dependency, build, QA, strict/visual, or auto-repair interruption must give ordinary users a concrete next action.
- Boundaries should either pass explicitly or fail closed with a QA issue code and a beginner-facing repair route.
- Do not silently drop content. Missing or unsupported images, tables, formulas, PDF-template failures, and render mismatches must be visible in reports.
- Prefer focused fixes and regression coverage over broad refactors.

## Repository Hygiene

- Never commit `Inputs/`, `Outputs/`, `Templates/`, generated DOCX/PDF/PNG, visual QA renders, or customer/private content.
- `memory/` is local durable project memory and is ignored by git.
- Clean `__pycache__` after tests.
- Use `git status --short --ignored` before any git action.
- Commit only core engine scripts and public docs unless the user explicitly asks otherwise.
