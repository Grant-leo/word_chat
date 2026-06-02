# Paper Project — AI Assistant Workflow

This file is the operating guide for an AI assistant helping a user run or repair
the Word paper typesetting pipeline. Treat it as the first file to read when a
user asks Claude/Codex/another AI to work inside this repository.

## Your Job

You are an AI assistant for the Word paper typesetting pipeline.

Your default workflow:

1. Read `Paper_Project/基础操作.md` into context. It is the OOXML toolbox for document work.
2. Run the pipeline when the user asks.
3. Verify the output through QA reports and, when layout matters, visual QA or Word/WPS.
4. Help ordinary users fine-tune the current generated document by editing `Outputs/<latest>/build_generated.py`.
5. Only edit reusable core engine scripts when the user is the developer/maintainer or explicitly asks for a product-level engine fix.

Do not guess low-level OOXML patterns. Look them up in `Paper_Project/基础操作.md`.

---

## Every Session Checklist

### 0. Environment Check

```bash
python --version
python -c "import docx, lxml; from PIL import Image; print('OK')"
```

If missing:

```bash
python -m pip install python-docx Pillow lxml
```

Dependency notes:

- Normal pipeline and strict QA require Python 3.10+, `python-docx`, `Pillow`, and `lxml`.
- Generated DOCX builds copy and use local engine modules; the public-template downloader uses stdlib `urllib`, so no `requests` package is required.
- PDF template parsing requires Poppler command-line tools on `PATH`: `pdfinfo` and `pdftotext`. Missing tools should surface `PDF_TEMPLATE_DEPENDENCY_MISSING`; password-protected or copy-restricted PDFs should surface `PDF_TEMPLATE_PROTECTED`; corrupt/unreadable PDFs should surface `PDF_TEMPLATE_READ_FAILED`; scanned/textless PDFs should become QA errors, not silent defaults.
- Automatic Word TOC/page-number updating is optional and uses Microsoft Word COM through `pywin32` (`python -m pip install pywin32`) when available; without it, static visible TOC lines remain.
- `--qa-level visual` requires Windows PowerShell plus Microsoft Word COM for PDF export, and Poppler command-line tools on `PATH`: `pdfinfo`, `pdftotext`, `pdftoppm`.
- Optional WPS cross-render QA requires WPS COM (`KWPS.Application` or `WPS.Application`); missing WPS is a warning unless `--require-wps` is used.

### 1. Load The Toolbox

Read:

```text
Paper_Project/基础操作.md
```

This file contains implementation references for tables, cross-references,
headers, pagination, formulas, comments, footnotes, images, and other OOXML work.

### 1.5 Optional Local Memory

If the repository owner has a private local memory bank and explicitly wants you
to use it, read:

```text
memory/PROJECT_MEMORY.md
memory/active_context.md
```

`memory/` is private local development state and is ignored by git. Do not require
it for ordinary users, do not create it for them unless asked, and do not commit it.

### 2. Check Inputs

Check that the user has files in the local folders:

```bash
python -c "from pathlib import Path; print('Templates:', [p.name for p in Path('Templates').glob('*.docx')] + [p.name for p in Path('Templates').glob('*.pdf')]); print('Inputs:', [p.name for p in Path('Inputs').glob('*.docx')] + [p.name for p in Path('Inputs').glob('*.md')])"
```

If no template or content exists, tell the user to put `.docx` or `.pdf`
template files under `Templates/` and `.docx` or `.md` content files under
`Inputs/`.

### 3. Select Workflow Mode

Use the mode recorded in `Outputs/<latest>/workflow_mode.json` when it exists.

- `user`: edit only the latest output directory's `build_generated.py`.
- `developer`: edit reusable core scripts under `Paper_Project/Program/pipeline/`, then rerun the full pipeline.

If no mode is known and the user has not said they are the developer, use `user`.

### 3.5 Agent-First Default

Ordinary users are expected to be guided by an Agent, not by terminal commands.
When the user asks you to start formatting, run the Agent entry first:

```bash
python run_pipeline.py --agent-auto
```

This non-interactive path scans `Templates/` and `Inputs/`, auto-selects only
when there is a single valid choice, defaults to `user` mode, enables the
bounded auto-repair loop, and writes `agent_summary.md/json`. If there are
multiple candidates, ask the user for the file name only.
The preflight report should also list direct reply sentences for each
candidate, such as `使用 Templates/<文件名> 作为模板` or
`使用 Inputs/<文件名> 作为内容`, so a beginner can simply choose one and
send it back to the Agent.

If interactive selection is cancelled or stdin closes, do not leave the user
waiting. Tell them to rerun with `python run_pipeline.py --agent-auto`, or with
explicit `--template` / `--content` file names when the choices are known.

If anything interrupts before or during the run, do not leave ordinary users
waiting without direction. Read or write the relevant `agent_preflight_report.md`,
`agent_summary.md`, or QA report, then state the next concrete action they
should take.

### 4. Run Pipeline

```bash
# Agent-first ordinary user workflow
python run_pipeline.py --agent-auto

# DOCX template + DOCX content
python run_pipeline.py --mode user --template <模板文件名> --content <内容文件名>

# DOCX template + Markdown content
python run_pipeline.py --mode user --template <模板文件名> --content <md文件名>

# PDF template + DOCX/Markdown content
python run_pipeline.py --mode user --template <PDF模板文件名> --content <内容文件名>

# Pure Markdown mode: format + content in one .md
python run_pipeline.py --mode user --md <md文件名>

# Developer engine-fix run
python run_pipeline.py --mode developer --template <模板文件名> --content <内容文件名>

# Product-grade verification: structure QA + strict XML/conformance + PDF/render QA
python run_pipeline.py --mode developer --qa-level visual --template <模板文件名> --content <内容文件名>

# Ordinary user controlled auto-repair loop
python run_pipeline.py --mode user --auto-repair --template <模板文件名> --content <内容文件名>
```

Interactive mode is also supported:

```bash
python run_pipeline.py
```

### 5. Verify Outputs

Open the newest `Outputs/<run>/` directory and inspect:

- `agent_summary.md` / `agent_summary.json`: user-facing handoff; read this first when present. Structural/strict/visual QA failures are summarized here with issue codes and beginner-facing next actions.
- `格式提取.md`: template format extraction summary.
- `内容提取.md`: content extraction summary.
- `template_profile.md`: template capability and risk flags.
- For PDF templates, check PDF type, confidence, warnings, and any `PDF_TEMPLATE_DEPENDENCY_MISSING`, `PDF_TEMPLATE_PROTECTED`, `PDF_TEMPLATE_READ_FAILED`, `PDF_TEMPLATE_UNSUPPORTED`, `PDF_TEMPLATE_INSTRUCTION_INCOMPLETE`, `PDF_TEMPLATE_VISUAL_APPROXIMATION`, or `PDF_TEMPLATE_LANDSCAPE_PAGE` issue in `template_profile.md`, `格式提取.md`, and `qa_report.md`.
- `template_requirements.md`: machine-checkable template/content requirements, when strict/visual QA is available.
- `qa_report.md`: first repair entry point; its top next action names the leading issue code and concrete beginner-facing fix. If `build_generated.py` fails before normal QA, the pipeline still writes this report with `MISSING_DOCX` guidance.
- `qa_repair_plan.md` / `qa_repair_plan.json`: step-by-step repair plan with top-level `next_action`, `resume_scope`, and `resume_command`, including generated-script build failures that should resume from the current DOCX build script.
- `qa_fix_prompt.txt`: user-copyable prompt for another AI repair round.
- `repair_loop_report.md` / `repair_loop_report.json`: bounded auto-repair audit when `--auto-repair` was used; stopped loops expose top-level `next_action`, `resume_scope`, and `resume_command`.
- `conformance_report.md`: strict DOCX/XML/template-content conformance report.
- `visual_report.md` and `visual_qa/samples/`: PDF/render QA outputs when `--qa-level visual` was used.
  Their top-level next action should name the leading issue code and be
  issue-code-specific for common blockers such as placeholders, Word field
  errors, invalid PDF page count, unreadable page PNGs, and missing render
  dependencies.
- `build_manifest.json`: rendered images/tables/formulas counts.
- `最终论文.docx`: final generated Word document.

Office Viewer is not enough for final delivery. Use Word/WPS or visual QA when
layout correctness matters.

### 6. Report Then Offer Fine-Tuning

After reporting results, ask:

```text
排版完成了，需要微调吗？
```

---

## Two Modification Modes

### Controlled Auto Repair

For ordinary users who want the Agent to continue after QA errors, prefer:

```bash
python run_pipeline.py --mode user --auto-repair --template <模板文件名> --content <内容文件名>
```

The loop is deliberately bounded. It edits only `Outputs/<run>/build_generated.py`, rebuilds the current DOCX, reruns every enabled QA level, and writes `repair_loop_report.md/json`. In `strict`/`visual` mode, missing conformance or visual dependencies are errors rather than skipped passes, and visual options such as explicit golden baseline path, update-golden mode, and WPS requirement must be preserved across rounds. It stops after repeated non-improvement, max rounds, build failure, or any `needs_user_file` / `needs_user_input` blocker, and the report plus `agent_summary.md/json` must tell the beginner the next concrete resume action. Treat convergence as "automatic QA has no error", not as a 100% correctness guarantee; still tell the user to inspect the final DOCX in Word/WPS.

For public-template visual checks, plain `python Paper_Project/Program/pipeline/public_template_suite.py --visual` is a render-QA run and does not compare golden baselines by default. Use `--golden-dir` when intentionally comparing against a baseline, and use `--update-golden` when intentionally refreshing baseline files.

`build_generated.py` is generated by `script_generator.py`. It is the normal
user-facing AI edit surface for one document, but not the long-term source of
truth for product behavior.

Use it to inspect or tune:

- helper functions such as `body()`, `heading1/2/3()`, `body_with_formula()`
- generated `DATA`, profile values, image paths, TOC entries, and section order
- exact error locations when a generated document fails to build

### User-Level Fine-Tuning

For ordinary users:

1. Read the latest output: `qa_report.md`, `qa_repair_plan.md`, `格式提取.md`, `内容提取.md`, and `build_generated.py`.
2. Edit only `Outputs/<latest>/build_generated.py`.
3. Re-run:

   ```bash
   python Outputs/<latest>/build_generated.py
   ```

4. Re-run QA if needed:

   ```bash
   python Paper_Project/Program/pipeline/qa_checker.py Outputs/<latest> --mode user
   ```

5. Verify the updated `最终论文.docx`.
6. Do not ask ordinary users to edit core engine scripts.

### Developer-Level Engine Fixes

For reusable fixes, product behavior, parser/generator changes, or maintainer
requests:

1. Read the latest output: `qa_report.md`, `qa_repair_plan.md`, `格式提取.md`, `内容提取.md`, `build_manifest.json`, and when useful `build_generated.py`.
2. Identify the owning core area.
3. Consult `Paper_Project/基础操作.md` for OOXML details.
4. Edit core scripts under `Paper_Project/Program/pipeline/`.
5. Keep the change generic and template-driven. Do not add school-name or file-name hardcoding.
6. Add or update focused regression tests in `regression_suite.py`.
7. Re-run the whole pipeline and the relevant regression tests.

Recommended developer verification:

```bash
python Paper_Project/Program/pipeline/regression_suite.py
python run_pipeline.py --mode developer --qa-level strict --template <模板文件名> --content <内容文件名>
```

Use visual QA for delivery gates:

```bash
python run_pipeline.py --mode developer --qa-level visual --template <模板文件名> --content <内容文件名>
```

---

## Core Ownership Map

- `content_parser.py`: stable public content extraction entrypoint.
- `content_parser_modules/`: reusable content extraction rules: extraction orchestration, placeholders, styles, text cleanup, front matter, captions, paragraph streams, body dispatch, source TOC, images, tables, formula label cleanup, source OMML extraction, text formula items, split-layout repair strategies, headings, references, and section building. Standalone/default extraction output stays under `Outputs/_content_parser_cli/` or `Outputs/_content_parser_extract/`, not under `Inputs/`.
- `format_extractor.py`: stable DOCX/PDF template format extraction entry point.
- `format_extractor_modules/`: PDF template parsing, OOXML metrics, style inheritance, semantic style profiles, cover assets, and cover table extraction. Standalone CLI output defaults to `Outputs/_format_extractor_cli/`, and template assets should not be written beside files in `Templates/`.
- `md_parser.py`: Markdown format/content parsing.
- `md_parser_modules/`: Markdown parser helpers for content extraction orchestration, format extraction, front format-instruction stripping, inline/display math tokens, image path resolution, table parsing, and text cleanup. Standalone CLI output defaults to `Outputs/_md_parser_cli/`.
- `template_profiler.py`: stable template capability/risk profile entry point.
- `template_profiler_modules/`: template profile construction and report writing.
- `script_generator.py`: stable public script generation entrypoint.
- `script_generator_modules/`: generator planning and generated-script runtime fragments: generator orchestration, runtime template assembly, sections, template rules, style profiles, base runtime, cover/front matter/body rendering, formula text conversion, formula rendering, media/tables/code, references/backmatter, TOC, and build manifest orchestration.
- `latex_omath.py`: LaTeX/text formula to native Word OOXML Math conversion.
- `latex_omath_modules/`: formula converter tokenizer, parser, API helpers, symbol registries, and OOXML builders copied with generated build scripts.
- `qa_checker.py`: structural QA report, issue routing, repair-plan output.
- `qa_checker_modules/`: structural QA phase checks, issue registry, JSON/DOCX/content metrics and samples, repair-plan generation, and Markdown/JSON report writers.
- `qa_conformance.py`: strict DOCX/XML/template-content conformance checks and template requirements.
- `qa_conformance_modules/`: strict QA orchestration, OOXML helpers, content/style checks, DOCX XML checks, template requirement generation, and conformance report writers.
- `qa_visual.py`: Word COM/PDF/render visual QA.
- `qa_visual_modules/`: visual QA orchestration, export, Poppler/render, image stats, opt-in golden baseline, and report helpers.
- `privacy.py`: report path sanitization helpers.
- `comment_utils.py`: Word comment injection system.
- `comment_utils_modules/`: comment_utils implementation for OOXML comments, relationships, and content types.
- `public_template_suite.py`: public-template compatibility suite; downloads/runs stay local.
- `public_template_suite_modules/`: shared paths, storage/download helpers, execution runners, Markdown reports, default public template metadata, synthetic non-private scenarios, and generated test image assets.

Current engine rules: template-only instructions, source-TOC examples, cover field hints, and TOC page-number samples are not paper content and must be stripped before final rendering. Structural QA should also treat common backmatter heading equivalents such as `Acknowledgements` / `Acknowledgment` / `致谢`, `References` / `参考文献`, and `Appendix` / `附录` as matching rather than forcing the user to resolve false `CONTENT_HEADING_MISSING` warnings.
- `regression_suite.py`: synthetic engine regression suite.
- `regression_suite_modules/`: regression harness, assertions, temp workspace cleanup, base fixtures, generated-DOCX helpers, and grouped pipeline/content/formula/Markdown/QA/generator/template/operational cases.

### Runner Helper Modules

`run_pipeline.py` is now mostly orchestration. `pipeline_runner/` holds the details:

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
- `repair_loop.py`: controlled user-mode auto-repair loop and `repair_loop_report.md/json` writer.
- `reports.py`: terminal progress, contract warnings, and repair hints.
- `summary.py`: completion output inventory and repair workflow summary.

---

## Common Edits

Look up implementation details in `Paper_Project/基础操作.md`.

| User says | Where to look |
|-----------|---------------|
| 改页边距/纸张 | 页面设置 |
| 表格加三线 | 三线表 |
| 加参考文献跳转 | 交叉引用 |
| 加页码 | 页眉与页码 |
| 图片居中 | 图片居中 |
| 加公式/矩阵 | 公式（OOXML Math） |
| 加脚注 | Word 脚注 |
| 分页不对 | A4 自动分页 |
| 加批注 | 批注（Comment） |

### Adding Features Not In The Template

Template may not include certain features such as cross-references, formulas,
or comments, so generated scripts may not expose dedicated helpers yet.

When a user wants a document-specific feature:

1. Read the relevant section in `基础操作.md`.
2. Add the implementation to the current `Outputs/<latest>/build_generated.py`.
3. Adapt parameters to the current document's extracted format values.
4. Re-run the generated script.
5. If the feature should become reusable product behavior, move it into the appropriate core engine module later.

---

## Technical Rules

- All format values should come from `P{}` template data or `D{}` derived data. Avoid hardcoding.
- Always honor workflow mode: user mode changes only `build_generated.py`; developer mode changes reusable core scripts and reruns the whole pipeline.
- `template_profile.json/md` is the reusable template decision layer. Use profile capabilities/risk flags instead of school-name logic.
- QA reports are routing-focused and block on `error`; they do not replace Word/WPS visual verification for final delivery.
- QA also writes `qa_repair_plan.md/json` and `qa_fix_prompt.txt`; generated-script build failures should get the same QA-shaped handoff before asking the user to repair `build_generated.py`. `qa_report.md/json` and the repair plan should name the leading issue code and concrete next action first when repairing. The repair plan JSON should also expose `next_action`, `resume_scope`, and `resume_command` so the handoff says whether to fix input files, rebuild the current DOCX, rerun the full pipeline, or do final Word/WPS review. Strict/visual reports should also name the leading issue code in the top-level `next_action` and avoid generic "inspect the report" guidance when the issue code can provide a concrete next action.
- QA/user-facing reports should prefer run-relative paths and avoid leaking absolute local paths.
- `workflow_mode.json` should only create copyable rerun commands for inputs that can be safely expressed under this project's `Inputs/` or `Templates/`. If an absolute source file is outside those folders, including another same-named external `Inputs` or `Templates` directory, do not collapse it to a misleading basename; omit the fake command and tell the user to move/copy the file into the project source folder, then rerun by file name.
- PDF templates are best-effort format sources: instruction-style PDFs provide text rules, sparse instruction PDFs must surface missing-rule warnings, visual sample PDFs provide estimated geometry/styles and must surface `PDF_TEMPLATE_VISUAL_APPROXIMATION` for Word/WPS layout review, landscape PDFs must surface `PDF_TEMPLATE_LANDSCAPE_PAGE`, missing Poppler tools must surface `PDF_TEMPLATE_DEPENDENCY_MISSING`, protected/password or copy-restricted PDFs must surface `PDF_TEMPLATE_PROTECTED`, corrupt/unreadable PDFs must surface `PDF_TEMPLATE_READ_FAILED`, and scanned/textless PDFs must surface `PDF_TEMPLATE_UNSUPPORTED`.
- `--qa-level visual` is the preferred delivery gate for developer/product checks. It requires Word COM for PDF export and Poppler tools (`pdfinfo`, `pdftotext`, `pdftoppm`) for page/text/sample checks. Missing required render tools fail visual QA.
- Missing local Markdown images, including images embedded in Markdown table cells, must surface as `CONTENT_IMAGE_MISSING`; unreadable or unsupported local Markdown image files must surface as `CONTENT_IMAGE_UNREADABLE` with a next step to re-export a normal PNG/JPG. Stable local Markdown image formats are `.png`, `.jpg`, and `.jpeg`; GIF/WebP/SVG/no-extension files, extension/actual-format mismatches, and data URI MIME/actual-format mismatches must fail closed as `CONTENT_IMAGE_UNREADABLE` instead of reaching Word generation. Remote Markdown image URLs must surface as `CONTENT_IMAGE_REMOTE_UNSUPPORTED` with a next step to download the image locally and change the Markdown link to a local relative path. DOCX image extraction failures, corrupt relationship image bytes, and unsupported DOCX relationship image formats must surface as `IMAGE_EXTRACT_FAILED` QA errors with a next step to re-export/reinsert the source image as a normal PNG/JPG rather than disappearing from `content.json` or flowing into `figures/`.
- DOCX table-cell images must surface in the content image stream.
- Header/footer images from the content source are non-body content and should surface as `NON_BODY_IMAGE_UNSUPPORTED` unless product behavior explicitly changes.
- Chinese text needs `w:eastAsia` set; the generator handles this automatically.
- A4 pagination uses separate Latin and CJK characters-per-line estimates.
- Each pipeline run creates an independent `Outputs/<date>_<content>/` directory; same-day duplicates get `_2`, `_3`, etc.
- Formulas should render as native OOXML Math. Check the generated docx XML for `<m:oMathPara>` / `<m:oMath>` when formula correctness matters.
- Markdown `$...$` / `$$...$$` formulas in abstracts and body sections should render as native OOXML Math.
- Markdown image paths resolve relative to the `.md` file first, including `%20` spaces, `<path with spaces>` wrappers, balanced parentheses in filenames, optional image titles such as `![图](path "title")`, local `?query` / `#fragment` suffixes copied from Markdown tools, reference-style images such as `![图][id]` plus `[id]: path` with optional title lines, shortcut reference images such as `![图]` plus `[图]: path`, HTML image tags such as `<img src="path" alt="图">`, HTML lazy attributes such as `data-src` / `data-original`, the first candidate in `srcset`, PNG/JPG data URI images, and Markdown table-cell images, then copy readable `.png` / `.jpg` / `.jpeg` local images into the current output `figures/` folder. Markdown table-cell images are attached to `table_cell_items`, keep `location="markdown_table_cell"`, render inside generated Word table cells, and must not be silently dropped or duplicated after the table. Undefined image references must surface as `CONTENT_IMAGE_MISSING`; unreadable/corrupt or unsupported local image files, GIF/WebP/SVG/no-extension local files, extension/actual-format mismatches, malformed/unsupported data URI images, and data URI MIME/actual-format mismatches must surface as `CONTENT_IMAGE_UNREADABLE`; reference-definition-like lines inside fenced code blocks must remain code. Remote `http://` / `https://` image URLs are not downloaded automatically; users should save them locally and update the Markdown path before rerunning.
- Markdown front format-instruction sections are format-only. They must be stripped from the content stream, including noisy/encoding-damaged headings followed by obvious format rules and a `---` delimiter.
- `内容提取.md` should summarize images, tables, and formulas by their real roles, mention table-cell image counts in table summaries when present, avoid duplicate image listings, and never display non-formula structured content as `[公式]`.
- Caption detection should distinguish true captions from prose references: `图 1 xxx 示意图` can be a caption, while `图 1 展示了...` remains body text.
- Generated scripts should suppress Python bytecode cache creation so output folders stay clean.
- OOXML math runs need `m:rPr` for WPS compatibility.
- TOC defaults to static visible TOC lines; Word COM can resolve heading page numbers automatically when available.
- Formula numbering supports `\tag{1.1}`, `\begin{equation}`, and `\begin{align}`. Appendix-style labels such as `A.1` should be preserved.

---

## Repository Hygiene

Never commit or upload private or generated artifacts:

- real files under `Inputs/`
- real files under `Templates/`
- generated files under `Outputs/`
- generated DOCX/PDF/PNG/visual QA renders
- customer/private content
- API keys or credentials
- local private memory under `memory/`

Commit only reusable engine scripts and public docs when the maintainer asks.

Before git:

```bash
git status --short --ignored
git diff --check
```

Do not use `git add .` blindly. Prefer explicit paths.

---

## Project Files

```text
run_pipeline.py              <- one-click entry
Paper_Project/Program/pipeline/
    format_extractor.py
    format_extractor_modules/
    content_parser.py
    content_parser_modules/
    md_parser.py
    md_parser_modules/
    template_profiler.py
    template_profiler_modules/
    script_generator.py
    script_generator_modules/
    latex_omath.py
    qa_checker.py
    qa_checker_modules/
    qa_conformance.py
    qa_conformance_modules/
    qa_visual.py
    qa_visual_modules/
    latex_omath_modules/
    privacy.py
    public_template_suite.py
    public_template_suite_modules/
    regression_suite.py
    regression_suite_modules/
    comment_utils.py
    comment_utils_modules/
    pipeline_runner/
Paper_Project/Program/pipeline/README.md <- engine layout map
Paper_Project/基础操作.md                <- OOXML toolbox
Inputs/                                  <- local content files, ignored except placeholder
Templates/                               <- local template files, ignored except placeholder
Outputs/                                 <- generated runs, ignored except placeholder
```

## Workflow Diagram

```text
Templates/模板.docx/.pdf or .md
        |
        v
format_extractor / md_parser
        |
        +--> Outputs/<run>/format.json
        +--> Outputs/<run>/格式提取.md
        |
        v
template_profiler
        |
        +--> Outputs/<run>/template_profile.json/md

Inputs/内容.docx/.md
        |
        v
content_parser / md_parser
        |
        +--> Outputs/<run>/content.json
        +--> Outputs/<run>/内容提取.md

format.json + content.json
        |
        v
script_generator
        |
        +--> Outputs/<run>/build_generated.py
        |
        v
python build_generated.py
        |
        +--> Outputs/<run>/最终论文.docx
        +--> Outputs/<run>/build_manifest.json
        |
        v
qa_checker / qa_conformance / qa_visual
        |
        +--> qa_report.md
        +--> qa_repair_plan.md
        +--> conformance_report.md
        +--> visual_report.md
```
