# Pipeline Engine Layout

This directory contains the reusable Word paper typesetting engine. User data,
templates, generated DOCX/PDF/PNG files, and local memory stay outside the
tracked engine.

## Stable Entrypoints

- `run_pipeline.py` at the repository root is the one-click CLI.
- `content_parser.py` extracts structured paper content.
- `format_extractor.py` extracts template style and layout.
- `template_profiler.py` builds template capability/risk profiles.
- `script_generator.py` writes `Outputs/<run>/build_generated.py`.
- `qa_checker.py`, `qa_conformance.py`, and `qa_visual.py` verify generated output.
- `regression_suite.py` is the synthetic engine regression gate.

## Runner Helpers

`pipeline_runner/` keeps `run_pipeline.py` as orchestration code while hiding
CLI, output, verification, and QA details in a focused package:

- `io.py`: input scanning, interactive choices, mode normalization.
- `cli.py`: CLI arguments, banner, interactive/non-interactive dispatch.
- `context.py`: path resolution, QA-level normalization, output folder creation, workflow metadata.
- `dependencies.py`: optional QA/template/Markdown imports and import-error details.
- `artifacts.py`: `format.json`, `content.json`, and markdown handoff reports.
- `verification.py`: repeated extraction verification, arbitration, and stable-content convergence.
- `template_phase.py`: template profile and template requirements report phase.
- `build_phase.py`: generated-script creation and DOCX build execution.
- `execution.py`: generated-script subprocess execution and UTF-8 output decoding.
- `contracts.py`: lightweight JSON handoff structure checks.
- `qa.py`: structural, strict, and visual QA orchestration.
- `reports.py`: terminal progress, contract warnings, and repair hints.
- `summary.py`: completion output inventory and repair workflow summary.

## Verification Baseline

Current baseline as of 2026-05-27:

- Synthetic regression: `112 passed, 0 failed`.
- End-to-end strict QA matrix: 5 complex content documents × 3 templates = `15/15` passed.
- Structural QA and conformance QA completed with no errors in that matrix.

## Parser Submodules

`content_parser_modules/` owns reusable content extraction rules behind
`content_parser.extract`: placeholders, style helpers, text cleanup, front
matter, captions, paragraph streams, source TOC filtering, images, tables,
formulas, headings, references, body dispatch, and section post-processing.

## Generator Submodules

`script_generator_modules/` owns reusable generation planning and runtime
fragments behind `script_generator.generate`: section planning, template
rules, style profiles, cover/front matter/body rendering, text formula
conversion, formula rendering, media/table/code blocks, references/backmatter,
TOC/page resolution, and build manifest orchestration.

## Git Hygiene

Commit core engine files and public docs only. Do not commit `Inputs/`,
`Outputs/`, `Templates/`, render artifacts, customer documents, or local
private memory. The private memory bank is intentionally ignored via
`.gitignore`.
