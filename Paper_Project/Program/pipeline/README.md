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

Current baseline as of 2026-05-28:

- Synthetic regression after the latest architecture split: `113 passed, 0 failed`.
- End-to-end strict QA matrix: 5 complex content documents × 3 templates = `15/15` passed.
- Structural QA and conformance QA completed with no errors in that matrix.

## Parser Submodules

`format_extractor_modules/` owns reusable template extraction rules behind
`format_extractor.extract`: OOXML scalar conversion, paragraph metrics, style
inheritance resolution, semantic style profiles, cover assets, and cover table
layout extraction. `extractor.py` owns the extraction orchestration, keeping
`format_extractor.py` as a thin stable entrypoint.

`content_parser_modules/` owns reusable content extraction rules behind
`content_parser.extract`: placeholders, style helpers, text cleanup, front
matter, captions, paragraph streams, source TOC filtering, images, tables,
formulas, headings, references, body dispatch, and section post-processing.
The formula path is split into label cleanup, source OMML extraction, text
formula item creation, and split-layout repair strategies. `extractor.py` owns
DOCX content extraction orchestration.

`md_parser_modules/` owns Markdown-specific helper rules behind `md_parser`:
YAML/natural-language format extraction, inline/display math tokenization,
Markdown image copying/missing-image metadata, table parsing, and Markdown text
cleanup. `content_extractor.py` owns Markdown content orchestration. The public
Markdown entrypoints stay `extract_format` and `extract_content`.

## Generator Submodules

`script_generator_modules/` owns reusable generation planning and runtime
fragments behind `script_generator.generate`: section planning, template
rules, style profiles, cover/front matter/body rendering, text formula
conversion, formula rendering, media/table/code blocks, references/backmatter,
TOC/page resolution, and build manifest orchestration. `runtime_template.py`
assembles generated-script fragments and `generator.py` owns the reusable
build-script generation workflow.

## QA Submodules

`qa_checker_modules/` owns structural QA helpers behind
`qa_checker.check_output`: issue ownership/repair-guide registry,
DOCX/content metrics and samples, repair-plan generation, and Markdown/JSON
report writers. Structural checks are organized by artifact, format, content,
DOCX XML, and report phases, with JSON/docx/content metric helpers split out
for targeted maintenance.

`qa_conformance_modules/` owns strict DOCX conformance helpers behind
`qa_conformance.check_conformance`: OOXML scalar/text/style readers,
content paragraph/style expectations, DOCX XML element checks,
template/content requirement generation, and conformance Markdown/JSON report
writers. `checks.py` owns validation orchestration while `qa_conformance.py`
stays a thin entrypoint.

`qa_visual_modules/` owns optional render QA helpers behind
`qa_visual.check_visual`: Word/WPS PDF export, Poppler text/page rendering,
sample page selection, rendered image statistics, golden-baseline comparison,
and visual QA report writers. `checks.py` owns render QA orchestration while
the entrypoint preserves legacy monkeypatch hooks used by regression tests.

## Public Template Suite Modules

`public_template_suite_modules/` owns reusable public-template test data behind
`public_template_suite.py`: shared paths, manifest/download/storage helpers,
execution runners, Markdown report writers, default public template metadata,
synthetic non-private scenarios, and generated PNG test assets. Downloads and
run outputs still stay under ignored `TestData/PublicTemplates/` paths.

## Formula Converter Modules

`latex_omath_modules/` owns reusable formula converter helpers behind
`latex_omath.py`: tokenizer, recursive parser, public API helpers, symbol
registries, Greek letters, arrows, n-ary operators, delimiters, matrix bracket
mappings, and low-level OOXML Math builders.
`script_generator.py` copies this dependency directory beside generated
`latex_omath.py` so output builds remain standalone.

## Comment Utility Modules

`comment_utils_modules/` owns OOXML comment injection behind
`comment_utils.py`: comment range/reference insertion, `word/comments.xml`
generation, relationship updates, and content-type updates. Generated scripts
keep the stable `from comment_utils import CommentCollector` import.

## Template Profiler Modules

`template_profiler_modules/` owns template profile construction and report
writing behind `template_profiler.py`. The public functions stay
`profile_format`, `report_to_markdown`, and `write_profile`.

## Regression Suite Modules

`regression_suite_modules/` owns reusable test harness helpers behind
`regression_suite.py`: case registration, assertions, temporary workspace
cleanup, base format/content fixtures, PNG fixtures, generated-DOCX smoke
helpers, and concrete case groups for pipeline orchestration, content parsing,
formula/OMML, Markdown, QA, script generation, template/format extraction, and
operational privacy/visual/CLI gates. The suite entrypoint is now a thin
registration runner.

## Git Hygiene

Commit core engine files and public docs only. Do not commit `Inputs/`,
`Outputs/`, `Templates/`, render artifacts, customer documents, or local
private memory. The private memory bank is intentionally ignored via
`.gitignore`.
