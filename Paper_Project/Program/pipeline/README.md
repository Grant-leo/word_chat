# Pipeline Engine Layout

This directory contains the reusable Word paper typesetting engine. User data,
templates, generated DOCX/PDF/PNG files, and local memory stay outside the
tracked engine.

## Stable Entrypoints

- `run_pipeline.py` at the repository root is the one-click CLI; `--agent-auto` is the Agent-first ordinary-user entry.
- `content_parser.py` extracts structured paper content.
- `format_extractor.py` extracts DOCX/PDF template style and layout.
- `template_profiler.py` builds template capability/risk profiles.
- `script_generator.py` writes `Outputs/<run>/build_generated.py`.
- `qa_checker.py`, `qa_conformance.py`, and `qa_visual.py` verify generated output.
- `regression_suite.py` is the synthetic engine regression gate.
- `RELEASE_VALIDATION.md` is the publish-before-release validation checklist.
- Standalone extractor CLIs write debug artifacts under `Outputs/_...` by default so `Inputs/` and `Templates/` remain source-only.
- Interactive cancellation/EOF and QA/dependency interruptions must always surface a concrete next step for ordinary users.

## Runner Helpers

`pipeline_runner/` keeps `run_pipeline.py` as orchestration code while hiding
CLI, output, verification, and QA details in a focused package:

- `io.py`: input scanning, interactive choices, mode normalization.
- `cli.py`: CLI arguments, banner, interactive/non-interactive dispatch, and Agent-first auto selection.
- `context.py`: path resolution, QA-level normalization, output folder creation, workflow metadata, and Agent-auto flags.
- `dependencies.py`: optional QA/template/Markdown imports and import-error details.
- `artifacts.py`: `format.json`, `content.json`, markdown handoff reports, QA-shaped build-failure reports, and early format-blocker handoffs.
- `verification.py`: repeated extraction verification, arbitration, and stable-content convergence.
- `template_phase.py`: template profile and template requirements report phase.
- `build_phase.py`: generated-script creation, DOCX build execution, and generated-script failure handoff.
- `execution.py`: generated-script subprocess execution and UTF-8 output decoding.
- `contracts.py`: lightweight JSON handoff structure checks.
- `qa.py`: structural, strict, and visual QA orchestration.
- `repair_loop.py`: bounded user-mode auto-repair loop; edits only `Outputs/<run>/build_generated.py`, reruns the enabled QA levels, and writes `repair_loop_report.md/json` with next-action resume fields.
- `reports.py`: terminal progress, contract warnings, and repair hints.
- `summary.py`: completion output inventory, repair workflow summary, and `agent_summary.md/json` handoff reports.

## Verification Baseline

Current baseline as of 2026-06-01:

- Synthetic regression after the Markdown reference-title continuation fix: `230 passed, 0 failed`.
- Agent-first flow: `--agent-auto` scans local inputs, auto-selects only single candidates, defaults to user auto-repair, and writes `agent_summary.md/json`.
- Novice interruption coverage: interactive cancellation/EOF, missing preflight inputs, generated-script build failures, QA dependency failures, and auto-repair blockers all route to a next action.
- Strict/visual report handoff coverage: `conformance_report.md/json` and `visual_report.md/json` top-level `next_action` values name the leading issue code before the beginner-facing repair step, so users can connect codes such as `PLACEHOLDER_TEXT_LEFT`, `PDF_PAGE_COUNT_INVALID`, and `GOLDEN_BASELINE_MISSING` to the next concrete action even without opening `agent_summary.md`.
- Workflow rerun command hygiene: absolute inputs outside this project's `Inputs/` / `Templates/`, including external same-named folders, do not collapse to misleading basename rerun commands; reports instead tell users to place the file in the correct source folder and rerun by file name.
- Markdown remote image handoff: remote `http://` / `https://` image URLs surface `CONTENT_IMAGE_REMOTE_UNSUPPORTED`, stop as user-file input blockers, and tell users to download the image locally and update the Markdown relative path before rerunning.
- Markdown local image path variants: local paths continue resolving `%20` spaces, `<path with spaces>` wrappers, balanced filename parentheses, optional image titles such as `![图](path "title")`, and local `?query` / `#fragment` suffixes copied from Markdown tools before checking the filesystem.
- Markdown unreadable image handoff: existing local image files that are corrupt, mislabeled, or unsupported surface `CONTENT_IMAGE_UNREADABLE`, stop as user-file blockers, and tell users to re-export a normal PNG/JPG before rerunning.
- Markdown reference-style images: `![alt][id]` plus `[id]: path` and shortcut reference images `![alt]` plus `[alt]: path` now copy local images into the content stream; optional title continuation lines after reference definitions are stripped with the definition; undefined image references become `CONTENT_IMAGE_MISSING` instead of staying as ordinary body text, and reference-definition-like lines inside fenced code blocks stay in code.
- Strict conformance body-start detection: default body paragraphs before the first explicit Markdown heading stay inside strict content checks instead of being skipped as TOC/front matter.
- Content-summary coverage: `内容提取.md` renders structured `role="image"` items, including table-cell images, as `[图片]` instead of opaque `[结构化内容]`.
- Output-boundary coverage: standalone/default `format_extractor`, `content_parser`, and `md_parser` outputs stay under `Outputs/_...` instead of beside private source files.
- Controlled auto-repair loop regression: repairable build-script error, no-improvement stop, rebuild-failure stop, needs-user-file stop, strict/visual dependency failure, WPS page-count/page-size/text-page/sample-image visual blockers, visual option preservation, summary next-action promotion, and sanitized report paths passed.
- PDF template end-to-end strict QA: synthetic instruction PDF template + DOCX content passed.
- Sparse PDF instruction handoff: incomplete text-rule PDFs now surface `PDF_TEMPLATE_INSTRUCTION_INCOMPLETE` warning guidance naming missing heading/caption/reference-style rule families, while still producing the DOCX for manual warning review.
- Visual sample PDF template handoff: visual sample PDFs now surface `PDF_TEMPLATE_VISUAL_APPROXIMATION` warning guidance and `pdf_template_visual_approximation` profile risk so users know the DOCX layout is estimated and must be reviewed in Word/WPS.
- Landscape PDF template handoff: landscape PDFs now surface `PDF_TEMPLATE_LANDSCAPE_PAGE` warning guidance and `pdf_template_landscape_page` profile risk so users know to review final DOCX orientation, margins, and compressed tables/body in Word/WPS.
- PDF template dependency handoff: missing `pdfinfo` / `pdftotext` stops after template profiling and before `build_generated.py`, while writing `PDF_TEMPLATE_DEPENDENCY_MISSING` QA/agent reports with `resume_scope=environment` and Poppler repair/rerun next steps.
- PDF template read-failure handoff: corrupt/unreadable PDFs stop after template profiling and before `build_generated.py`, while writing `PDF_TEMPLATE_READ_FAILED` QA/agent reports with re-export/openable-PDF or DOCX next steps.
- PDF template protection handoff: password-protected or copy-restricted PDFs stop after template profiling and before `build_generated.py`, while writing `PDF_TEMPLATE_PROTECTED` QA/agent reports with unlock/export-unprotected-PDF or DOCX next steps.
- Scanned/textless PDF template handoff: unsupported PDF templates stop after template profiling and before `build_generated.py`, while writing `PDF_TEMPLATE_UNSUPPORTED` QA/agent reports with DOCX/text-PDF/OCR next steps.
- PDF extreme stress gate: 9 cases covering uppercase extensions, visual samples, landscape pages, sparse instructions, scanned/corrupt/blank/too-short PDFs met expected outcomes.
- Public-template compatibility suite: 5 public templates × 5 synthetic scenarios = `25/25` passed.
- Local DOCX strict QA matrix: 5 DOCX templates × 5 DOCX contents = `25/25` passed.
- PDF boundary probe: parseable PDF templates passed strict QA; missing Poppler tools fail closed with `PDF_TEMPLATE_DEPENDENCY_MISSING`; protected/password or copy-restricted PDFs fail closed with `PDF_TEMPLATE_PROTECTED`; corrupt/unreadable PDFs fail closed with `PDF_TEMPLATE_READ_FAILED`; unsupported/scanned-style PDFs fail closed with `PDF_TEMPLATE_UNSUPPORTED` guidance.
- High-risk pipeline matrix: pure Markdown strict, missing Markdown image, header/footer image boundary, user auto-repair, DOCX/PDF visual smoke, and dense media/math strict checks all matched expectations (`7/7`).
- Fresh-folder novice smoke test: DOCX template + plain DOCX content + `--auto-repair --qa-level visual` converged with structural, strict, and visual QA all at zero errors.

## Parser Submodules

`format_extractor_modules/` owns reusable template extraction rules behind
`format_extractor.extract`: OOXML scalar conversion, paragraph metrics, style
inheritance resolution, PDF template parsing, semantic style profiles, cover
assets, and cover table layout extraction. `extractor.py` owns the extraction
orchestration, keeping `format_extractor.py` as a thin stable entrypoint.

PDF templates are handled as best-effort format sources. Instruction-style PDFs
are parsed as text rules; sparse instruction PDFs surface
`PDF_TEMPLATE_INSTRUCTION_INCOMPLETE` warnings that name missing rule families
such as headings, captions, and references for beginner-facing warning review.
Visual sample PDFs surface `PDF_TEMPLATE_VISUAL_APPROXIMATION` warning guidance
and a profile risk flag so users review estimated layout details in Word/WPS.
Landscape PDFs surface `PDF_TEMPLATE_LANDSCAPE_PAGE` warning guidance and a
profile risk flag so users review final DOCX page orientation, margins, and
compressed tables/body in Word/WPS. Visual sample PDFs estimate page geometry
and styles from Poppler text bounding boxes. Missing Poppler tools surface
`PDF_TEMPLATE_DEPENDENCY_MISSING` after template profiling with
`resume_scope=environment`, so users are told to repair `pdfinfo`/`pdftotext`
and rerun. Protected/password or copy-restricted PDFs surface
`PDF_TEMPLATE_PROTECTED` with a next step to remove the password/permission
restriction, export an unprotected copyable-text PDF, or use DOCX.
Corrupt/unreadable PDFs surface `PDF_TEMPLATE_READ_FAILED` with a next step to
re-export an openable text PDF or use DOCX. Scanned/textless PDFs surface
`PDF_TEMPLATE_UNSUPPORTED` before content extraction or script generation, so
users are routed to DOCX/text-PDF/OCR input repair instead of receiving a
misleading default-formatted DOCX.

`content_parser_modules/` owns reusable content extraction rules behind
`content_parser.extract`: placeholders, style helpers, text cleanup, front
matter, captions, paragraph streams, source TOC filtering, images, tables,
formulas, headings, references, body dispatch, and section post-processing.
The formula path is split into label cleanup, source OMML extraction, text
formula item creation, and split-layout repair strategies. `extractor.py` owns
DOCX content extraction orchestration.

Caption detection deliberately separates true captions such as `图 1 xxx 示意图`
from prose references such as `图 1 展示了...`, so body prose keeps body style
while captions keep caption style.

`md_parser_modules/` owns Markdown-specific helper rules behind `md_parser`:
YAML/natural-language format extraction, inline/display math tokenization,
front format-instruction stripping, Markdown image copying/missing-image
metadata, local image URI-suffix normalization, reference-style and shortcut
reference image definitions, remote-image blocker metadata, UTF-8 BOM-safe YAML/front-format
stripping, BOM/H1 and Setext `===` title detection, table parsing, and
Markdown text cleanup.
`content_extractor.py` owns Markdown content orchestration. The public
Markdown entrypoints stay `extract_format` and `extract_content`.

## Generator Submodules

`script_generator_modules/` owns reusable generation planning and runtime
fragments behind `script_generator.generate`: section planning, template
rules, style profiles, cover/front matter/body rendering, text formula
conversion, formula rendering, media/table/code blocks, references/backmatter,
TOC/page resolution, and build manifest orchestration. `runtime_template.py`
assembles generated-script fragments and `generator.py` owns the reusable
build-script generation workflow.

Generated scripts suppress Python bytecode cache creation to keep `Outputs/`
clean during user-mode rebuilds.

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
WPS PDF metadata/page-count/page-size/text-page validation, WPS sample-image comparison, separate Word/WPS rendered-text diagnostics, and visual QA report writers. `checks.py` owns render QA orchestration while
the entrypoint preserves legacy monkeypatch hooks used by regression tests.

## Public Template Suite Modules

`public_template_suite_modules/` owns reusable public-template test data behind
`public_template_suite.py`: shared paths, manifest/download/storage helpers,
execution runners, Markdown report writers, default public template metadata,
synthetic non-private scenarios, and generated PNG test assets. Downloads and
run outputs still stay under ignored `TestData/PublicTemplates/` paths.
Template downloads must use HTTPS, and manifest entries may pin `sha256` (or
`expected_sha256`) so local/downloaded DOCX files are verified before use.

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
helpers, synthetic PDF fixtures, and concrete case groups for pipeline orchestration, content parsing,
formula/OMML, Markdown, QA, script generation, template/format extraction, and
operational privacy/visual/CLI gates. The suite entrypoint is now a thin
registration runner.

## Git Hygiene

Commit core engine files and public docs only. Do not commit `Inputs/`,
`Outputs/`, `Templates/`, render artifacts, customer documents, or local
private memory. The private memory bank is intentionally ignored via
`.gitignore`.
