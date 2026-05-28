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
- PDF template parsing requires Poppler command-line tools on `PATH`: `pdfinfo` and `pdftotext`. Scanned/textless PDFs should become QA errors, not silent defaults.
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

### 4. Run Pipeline

```bash
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
```

Interactive mode is also supported:

```bash
python run_pipeline.py
```

### 5. Verify Outputs

Open the newest `Outputs/<run>/` directory and inspect:

- `格式提取.md`: template format extraction summary.
- `内容提取.md`: content extraction summary.
- `template_profile.md`: template capability and risk flags.
- For PDF templates, check PDF type, confidence, warnings, and any `PDF_TEMPLATE_UNSUPPORTED` issue in `template_profile.md`, `格式提取.md`, and `qa_report.md`.
- `template_requirements.md`: machine-checkable template/content requirements, when strict/visual QA is available.
- `qa_report.md`: first repair entry point; it names the active fix target.
- `qa_repair_plan.md` / `qa_repair_plan.json`: step-by-step repair plan.
- `qa_fix_prompt.txt`: user-copyable prompt for another AI repair round.
- `conformance_report.md`: strict DOCX/XML/template-content conformance report.
- `visual_report.md` and `visual_qa/samples/`: PDF/render QA outputs when `--qa-level visual` was used.
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
- `content_parser_modules/`: reusable content extraction rules: extraction orchestration, placeholders, styles, text cleanup, front matter, captions, paragraph streams, body dispatch, source TOC, images, tables, formula label cleanup, source OMML extraction, text formula items, split-layout repair strategies, headings, references, and section building.
- `format_extractor.py`: stable DOCX/PDF template format extraction entry point.
- `format_extractor_modules/`: PDF template parsing, OOXML metrics, style inheritance, semantic style profiles, cover assets, and cover table extraction.
- `md_parser.py`: Markdown format/content parsing.
- `md_parser_modules/`: Markdown parser helpers for content extraction orchestration, format extraction, inline/display math tokens, image path resolution, table parsing, and text cleanup.
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
- `qa_visual_modules/`: visual QA orchestration, export, Poppler/render, image stats, golden baseline, and report helpers.
- `privacy.py`: report path sanitization helpers.
- `comment_utils.py`: Word comment injection system.
- `comment_utils_modules/`: comment_utils implementation for OOXML comments, relationships, and content types.
- `public_template_suite.py`: public-template compatibility suite; downloads/runs stay local.
- `public_template_suite_modules/`: shared paths, storage/download helpers, execution runners, Markdown reports, default public template metadata, synthetic non-private scenarios, and generated test image assets.
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
- QA also writes `qa_repair_plan.md/json` and `qa_fix_prompt.txt`; use these files first when repairing.
- PDF templates are best-effort format sources: instruction-style PDFs provide text rules, visual sample PDFs provide estimated geometry/styles, and scanned/textless PDFs must surface `PDF_TEMPLATE_UNSUPPORTED`.
- `--qa-level visual` is the preferred delivery gate for developer/product checks. It requires Word COM for PDF export and Poppler tools (`pdfinfo`, `pdftotext`, `pdftoppm`) for page/text/sample checks. Missing required render tools fail visual QA.
- Missing or remote Markdown images, and DOCX image extraction failures, must surface as QA errors rather than disappearing from `content.json`.
- DOCX table-cell images must surface in the content image stream.
- Header/footer images from the content source are non-body content and should surface as `NON_BODY_IMAGE_UNSUPPORTED` unless product behavior explicitly changes.
- Chinese text needs `w:eastAsia` set; the generator handles this automatically.
- A4 pagination uses separate Latin and CJK characters-per-line estimates.
- Each pipeline run creates an independent `Outputs/<date>_<content>/` directory; same-day duplicates get `_2`, `_3`, etc.
- Formulas should render as native OOXML Math. Check the generated docx XML for `<m:oMathPara>` / `<m:oMath>` when formula correctness matters.
- Markdown `$...$` / `$$...$$` formulas in abstracts and body sections should render as native OOXML Math.
- Markdown image paths resolve relative to the `.md` file first, then copy into the current output `figures/` folder.
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
