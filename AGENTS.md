# Paper Project — AI Assistant Workflow

## Your Job

You are an AI assistant for the Word paper typesetting pipeline. Your core workflow:

1. **Read `Paper_Project/基础操作.md` into context** — this is your toolbox for ALL document manipulations
2. Run the pipeline when the user asks
3. Verify outputs
4. Help fine-tune the current document by editing `build_generated.py`, unless the user is asking for a reusable engine fix

---

## Every Session: Must-Do Checklist

### 0. Environment Check
```bash
python --version
python -c "import docx, lxml; from PIL import Image; print('OK')"
```
If missing: `python -m pip install python-docx Pillow lxml`.

Dependency notes:
- Normal pipeline and strict QA require Python 3.10+, `python-docx`, `Pillow`, and `lxml`.
- Generated DOCX builds copy and use local engine modules; no extra Python package is needed for the public-template downloader because it uses stdlib `urllib`.
- PDF template parsing requires Poppler command-line tools on `PATH`: `pdfinfo` and `pdftotext`. Missing tools must become `PDF_TEMPLATE_DEPENDENCY_MISSING` with `resume_scope=environment`; password-protected or copy-restricted PDFs must become `PDF_TEMPLATE_PROTECTED` with unlock/export-unprotected-PDF guidance; unreadable/corrupt PDFs must become `PDF_TEMPLATE_READ_FAILED` with re-export/openable-PDF guidance; scanned/textless PDFs should become `PDF_TEMPLATE_UNSUPPORTED`, not silent defaults. Sparse instruction PDFs should continue only with `PDF_TEMPLATE_INSTRUCTION_INCOMPLETE` / `PDF_TEMPLATE_LIMITED_CONFIDENCE` warnings that name the missing rule families and the next review step. Visual sample PDFs should continue with `PDF_TEMPLATE_VISUAL_APPROXIMATION` warning guidance for Word/WPS layout review. Landscape PDF templates should continue with `PDF_TEMPLATE_LANDSCAPE_PAGE` warning guidance for Word/WPS page-orientation review.
- Automatic Word TOC/page-number updating is optional and uses Microsoft Word COM via `pywin32` (`python -m pip install pywin32`) when available; without it the pipeline keeps static visible TOC lines.
- `--qa-level visual` requires Windows PowerShell plus Microsoft Word COM for PDF export, and Poppler command-line tools on `PATH`: `pdfinfo`, `pdftotext`, `pdftoppm`.
- Optional WPS cross-render QA requires WPS COM (`KWPS.Application` or `WPS.Application`); missing WPS is a warning unless `--require-wps` is used.

### 1. Load Your Toolbox
**Read `Paper_Project/基础操作.md`.** This file contains all OOXML code snippets: tables, cross-references, headers, pagination, formulas, footnotes. You will consult it for every modification. Do NOT guess OOXML — look it up here.

### 1.5 Load Long-Term Memory
Read `memory/PROJECT_MEMORY.md` and `memory/active_context.md` when they exist. These files are the disk-backed project memory and should guide long-running architecture work.

### 2. Check Files
```bash
ls Templates/*.docx Templates/*.pdf Inputs/*.docx Inputs/*.md 2>/dev/null
```
If no template/content files exist, tell the user to put `.docx` or `.pdf` templates in `Templates/`, and `.docx` or `.md` content in `Inputs/`.

### 3. Select Workflow Mode

Use the mode recorded in `Outputs/<latest>/workflow_mode.json` when it exists.

- `user`: edit only `Outputs/<latest>/build_generated.py`
- `developer`: edit only core scripts under `Paper_Project/Program/pipeline/`, then rerun the full pipeline

If no mode is known and the user has not said they are the developer, use `user`.

### 3.5 Agent-First Default

普通用户主要通过 Agent 引导使用项目，不应被要求自己打开终端或拼命令。When the user says "开始排版", "帮我跑论文", "自动排版", or otherwise asks for ordinary paper formatting:

1. Use the project Agent entry: `python run_pipeline.py --agent-auto`
2. Let it scan `Templates/` and `Inputs/`
3. If there is exactly one valid template/content pair, run it directly
4. If there are multiple candidates, ask the user to choose only the file name; `agent_preflight_report.md/json` should also list direct reply sentences such as `使用 Templates/<文件名> 作为模板` or `使用 Inputs/<文件名> 作为内容`, and include a stable `文件应该放哪里` / `source_folders` handoff that names `Templates/` (`.docx`/`.pdf`) and `Inputs/` (`.docx`/`.md`)
5. After the run, read `Outputs/<latest>/agent_summary.md` first, then the detailed QA reports; structural/strict/visual QA failures are summarized there with issue codes and beginner-facing next actions
6. If anything interrupts before or during the run, do not leave ordinary users waiting: read or write the relevant `agent_preflight_report.md`, `agent_summary.md`, or QA report, and state the next concrete action they should take
7. If interactive selection is cancelled or stdin closes, tell the user to rerun through `python run_pipeline.py --agent-auto`, or rerun with explicit `--template` / `--content` file names

Only use explicit `--template` / `--content` commands when the user or the file situation makes the choice unambiguous.

### 4. Run Pipeline
```bash
# Agent-first ordinary user workflow
python run_pipeline.py --agent-auto

# DOCX template + DOCX content
python run_pipeline.py --mode user --template <模板文件名> --content <内容文件名>

# PDF template + DOCX/MD content
python run_pipeline.py --mode user --template <PDF模板文件名> --content <内容文件名>

# DOCX template + MD content
python run_pipeline.py --mode user --template <模板文件名> --content <md文件名>

# Pure MD mode (format + content in one .md)
python run_pipeline.py --mode user --md <md文件名>

# Developer engine-fix run
python run_pipeline.py --mode developer --template <模板文件名> --content <内容文件名>

# Product-grade verification: structure QA + PDF/render QA
python run_pipeline.py --mode developer --qa-level visual --template <模板文件名> --content <内容文件名>

# Ordinary user controlled auto-repair loop
python run_pipeline.py --mode user --auto-repair --template <模板文件名> --content <内容文件名>
```
Or interactive: `python run_pipeline.py`

### 5. Verify Outputs
- Read `Outputs/<latest>/格式提取.md` — check paragraph counts, fonts, sizes
- Read `Outputs/<latest>/内容提取.md` — check all sections present
- Read `Outputs/<latest>/template_profile.md` — check template capabilities and risk flags
- For PDF templates, check `template_profile.md` and `格式提取.md` for PDF type, confidence, warnings, and possible `PDF_TEMPLATE_DEPENDENCY_MISSING` / `PDF_TEMPLATE_PROTECTED` / `PDF_TEMPLATE_READ_FAILED` / `PDF_TEMPLATE_UNSUPPORTED` / `PDF_TEMPLATE_INSTRUCTION_INCOMPLETE` / `PDF_TEMPLATE_VISUAL_APPROXIMATION` / `PDF_TEMPLATE_LANDSCAPE_PAGE`; missing Poppler tools should fail closed after template profiling with `resume_scope=environment` and a Poppler repair/rerun step, password-protected or copy-restricted PDFs should ask the user to remove the password/permission restriction or export an unprotected copyable-text PDF, unreadable/corrupt PDFs should ask the user to re-export a normal openable text PDF or switch to DOCX, scanned/textless PDF templates should fail closed before content extraction or `build_generated.py` creation with DOCX/text-instruction-PDF/OCR guidance, sparse instruction PDFs should tell users which missing rules (such as headings, captions, or references) need supplementation or manual review, visual sample PDFs should tell users to review estimated layout details in Word/WPS, and landscape PDFs should tell users to review final DOCX page orientation, margins, and compressed tables/body in Word/WPS.
- Read `Outputs/<latest>/qa_report.md` first; it names the active fix target, the first issue-code-specific next action for the current mode, and the repair-plan `先打开这些文件` review paths when available, including warning-only structural QA runs. Warning-only structural repair plans should use `resume_scope=warning_review`, keep a rerun command when the workflow can be inferred, and make clear whether the user can accept the warning manually or should update inputs and rerun. If `build_generated.py` failed before QA, the pipeline still writes `qa_report.md/json`, `qa_repair_plan.md/json`, and `qa_fix_prompt.txt` with `MISSING_DOCX` guidance. If `qa_checker.py` itself is unavailable, the pipeline writes the same structural handoff with `STRUCTURAL_QA_UNAVAILABLE`, `resume_scope=full_pipeline`, a workflow-derived rerun command when available, and a next step to repair `qa_checker.py` / `qa_checker_modules` before rerunning. If structural QA starts but crashes, the handoff uses `STRUCTURAL_QA_FAILED`; strict/visual runtime crashes write `CONFORMANCE_QA_FAILED` or `VISUAL_QA_FAILED` reports instead of leaving only a traceback.
- If `--auto-repair` was used, read `Outputs/<latest>/repair_loop_report.md/json`; it records every repair round, stop reason, top-level `next_action`, `resume_scope`, `resume_command`, and remaining manual checks. If the repair loop's internal structural/strict/visual QA dependency is missing or crashes, it must still write the corresponding QA report plus `repair_loop_report.md/json` with `STRUCTURAL_QA_UNAVAILABLE`, `STRUCTURAL_QA_FAILED`, `CONFORMANCE_QA_FAILED`, or `VISUAL_QA_FAILED`, sanitized details, and a workflow-derived rerun command when available.
- Read `Outputs/<latest>/agent_summary.md/json` first when present; it is the user-facing handoff with final DOCX path, QA status, per-report result labels, repair-loop result, structural/strict/visual QA issue-code next actions, and manual checks
- Read `conformance_report.md` and `visual_report.md` when strict/visual fails or has warnings; their top-level next action should name the leading issue code and be issue-code-specific for common blockers and warnings such as placeholders, Word field errors, strict style/page/content warnings, invalid PDF page count, unreadable page PNGs, missing render dependencies, missing golden baselines, blank-page warnings, TOC text warnings, optional WPS export warnings, WPS PDF metadata/page-count invalid results, WPS/Word page-count mismatches, WPS/Word page-size or orientation mismatches, WPS/Word extractable-text page mismatches, and WPS/Word sample-image mismatches that must tell users to inspect WPS/Word PDFs or PNG samples, fix, and rerun visual QA. `conformance_report.md` includes a `核对入口` section with the final DOCX, content summary, structured content, build manifest, template requirements, and format data paths for first-pass review. `visual_report.md` includes a `诊断产物` section with the rendered PDFs, text diagnostics, and bounded Word/WPS sample PNG paths. For WPS text mismatches, also inspect `visual_qa/rendered_word.txt` and `visual_qa/rendered_wps.txt` when present. For WPS sample-image mismatches, compare `visual_qa/samples/` with `visual_qa/wps/samples/`.
- If `--qa-level visual` was used, read `Outputs/<latest>/visual_report.md` and inspect sample PNGs under `visual_qa/samples/`
- Confirm `Outputs/<latest>/最终论文.docx` exists
- Render/check with Word/WPS when layout matters; Office Viewer alone is not enough
- For ordinary user fine-tuning, edit `Outputs/<latest>/build_generated.py` and run it.
- For developer/maintainer requests or reusable bug fixes, update the core pipeline scripts and run the full pipeline.

### 6. Report Then Offer Fine-Tuning
After reporting results, ask: "排版完成了，需要微调吗？"

---

## Two Modification Modes

`build_generated.py` is generated by `script_generator.py`. It is the normal user-facing AI edit surface for one document, but not the long-term source of truth for product behavior.

Use it to inspect or tune:
- helper functions such as `body()`, `heading1/2/3()`, `body_with_formula()`
- generated DATA, profile values, image paths, TOC entries, and section order
- the exact error location when a generated document fails to build

### User-Level Fine-Tuning

For ordinary users:

1. Read the latest output (`qa_report.md`, `格式提取.md`, `内容提取.md`, and `build_generated.py`) to locate the mismatch
2. Edit `Outputs/<latest>/build_generated.py`
3. Re-run `python Outputs/<latest>/build_generated.py`
4. Re-run QA if needed: `python Paper_Project/Program/pipeline/qa_checker.py Outputs/<latest> --mode user`
5. Verify the updated `最终论文.docx`
6. Do not ask users to edit core engine scripts

### User-Level Auto Repair

When the user wants the project to keep fixing obvious QA failures automatically, use:

```bash
python run_pipeline.py --mode user --auto-repair --template <模板文件名> --content <内容文件名>
```

Rules:
- The loop may edit only `Outputs/<latest>/build_generated.py`.
- It writes `repair_loop_report.md/json` with each round's actions, QA error/warning changes, new/resolved issue codes, stop reason, top-level next action, resume scope, and resume command.
- It reruns every enabled QA level. In `strict`/`visual` mode, missing conformance or visual dependencies must be reported as errors instead of being treated as convergence.
- It preserves visual QA options such as golden baseline path, update-golden mode, and WPS requirement across repair rounds.
- It stops after repeated non-improvement, after the max round limit, or when QA says a user file/input is required.
- A converged loop means automatic QA has no error; it is not a 100% correctness guarantee. Still ask the user to open the final DOCX in Word/WPS for visual review.

### Developer-Level Engine Fixes

For reusable fixes, product behavior, parser changes, or when the requester is the developer/maintainer:

1. **Read the latest output** (`qa_report.md`, `格式提取.md`, `内容提取.md`, and when useful `build_generated.py`) to locate the mismatch
2. **Identify the owning core script**
   - `script_generator.py`: stable script-generation entry point
   - `script_generator_modules/`: generator orchestration, runtime template assembly, DOCX rendering fragments, cover, TOC, styles, tables, images, formulas, references
   - `latex_omath.py`: LaTeX/text formula to native OOXML Math conversion
   - `latex_omath_modules/`: formula converter tokenizer, parser, API helpers, symbol registries, and OOXML builders copied with generated scripts
   - `format_extractor.py`: stable template format extraction entry point
   - `format_extractor_modules/`: extraction orchestration, PDF template parsing, OOXML metrics, style inheritance, semantic style profiles, cover assets, and cover table extraction. Standalone CLI output defaults to `Outputs/_format_extractor_cli/`; template assets should not be written beside files in `Templates/`.
   - `content_parser.py`: content sections, figures, references, metadata extraction
   - `content_parser_modules/`: reusable content extraction helpers for extraction orchestration, placeholders, text cleanup, front matter, captions, OOXML streams, body dispatch, source TOC, images, tables, formula labels/OMML/text items/repair strategies, headings, references, and section building. Standalone/default extraction output stays under `Outputs/_content_parser_cli/` or `Outputs/_content_parser_extract/`, never under `Inputs/`.
   - `md_parser.py`: Markdown input parsing
   - `md_parser_modules/`: Markdown parser helpers for content orchestration, format extraction, math tokens, image path resolution, tables, and text cleanup. Standalone CLI output defaults to `Outputs/_md_parser_cli/`.
   - `template_profiler.py`: stable template profile entry point
   - `template_profiler_modules/`: template capability/risk profile construction and report writing
   - `qa_visual.py`: optional PDF export/render QA
   - `qa_checker_modules/`: structural QA phase orchestration, issue registry, JSON/DOCX/content metrics, sample detectors, repair-plan generation, and report writers
   - `qa_conformance_modules/`: strict QA orchestration, OOXML helpers, content/style checks, DOCX XML checks, template requirements, and report writers
   - `qa_visual_modules/`: visual QA orchestration, export, Poppler/render, image stats, golden baseline, and report helpers
   - `public_template_suite.py`: public-template compatibility suite entry point
   - `public_template_suite_modules/`: paths, storage, runner, reporting, default public template metadata, synthetic scenarios, and generated test image assets
   - `privacy.py`: report path sanitization helpers
   - `regression_suite.py`: synthetic engine regression suite
   - `regression_suite_modules/`: regression harness, assertions, temp workspace cleanup, base fixtures, generated-DOCX helpers, and concrete case groups
   - `run_pipeline.py`: one-click entry and high-level orchestration
   - `pipeline_runner/`: CLI, IO, run context, dependency loading, artifact writing, extraction verification, template phases, build execution, JSON contracts, QA orchestration, controlled auto-repair loops, terminal reports, and completion summaries
3. **Consult `基础操作.md`** — find the correct OOXML implementation
4. **Edit the core script** — keep the change generic and template-driven
5. **Re-run the whole pipeline** — `python run_pipeline.py --mode developer --template <模板文件名> --content <内容文件名>`
6. **Verify the new output directory**

### Common Edits (look up details in 基础操作.md)

| User says | Where to look in 基础操作.md |
|-----------|------|
| "改页边距/纸张" | 页面设置 |
| "表格加三线" | 三线表 |
| "加参考文献跳转" | 交叉引用 |
| "加页码" | 页眉与页码 |
| "图片居中" | 图片居中 |
| "加公式/矩阵" | 公式（OOXML Math） |
| "加脚注" | Word 脚注 |
| "分页不对" | A4 自动分页 |

### Adding Features NOT in the Template

Template may not include certain features (cross-references, formulas, etc.) so `script_generator.py` won't generate them. When user wants to add them:

1. Read the relevant section in `基础操作.md`
2. Add the implementation to the current `build_generated.py` for the user's document
3. Adapt parameters to match the current paper's format
4. Re-run `python Outputs/<latest>/build_generated.py`
5. If the feature should become reusable product behavior, move the implementation into the appropriate core pipeline script

Example: template has no cross-references → generated script has no `B_ref()`. User says "加参考文献引用". Read 基础操作.md → 交叉引用, add the `add_ref_hyperlink` + `bookmarkStart/End` pattern to the current `build_generated.py`, then re-run it. A developer can later move the pattern into `script_generator.py`.

---

## Key Technical Rules

- All format values from `P{}` (template) or `D{}` (derived). Never hardcode.
- Do not commit or upload private test data: real files under `Inputs/`, `Outputs/`, `Templates/`, generated DOCX/PDF/PNG, QA renders, and template assets are local-only.
- Always honor the workflow mode: user mode changes only `build_generated.py`; developer mode changes only reusable core scripts and reruns the whole pipeline.
- Generated QA reports are routing-focused and block the pipeline on `error`; they still do not replace WPS/Word visual verification for final delivery.
- QA also writes `qa_repair_plan.md/json` and `qa_fix_prompt.txt`; generated-script build failures and structural QA dependency failures must also get the same QA-shaped handoff before the user is sent back to code. If `qa_checker.py` cannot import, `qa_report.md/json`, `qa_repair_plan.md/json`, and `qa_fix_prompt.txt` must use `STRUCTURAL_QA_UNAVAILABLE`, sanitize local paths, point to `qa_checker.py` / `qa_checker_modules`, infer a full-pipeline rerun command from `workflow_mode.json` when possible, and tell the user to repair the dependency then rerun the full pipeline. If structural QA crashes after starting, write `STRUCTURAL_QA_FAILED` with the same structural handoff and point to `qa_checker.py` / `qa_checker_modules`; if strict or visual QA crashes after starting, write `CONFORMANCE_QA_FAILED` or `VISUAL_QA_FAILED` into the corresponding report with sanitized detail and a concrete rerun route. The same rule applies inside `--auto-repair`: internal QA crashes must not bubble out as raw tracebacks, and `repair_loop_report.md/json` must stop with a beginner-facing next action plus `resume_scope`/`resume_command`. `qa_report.md/json` and the repair plan should name the leading issue code and concrete next action before editing user-level or developer-level code, even when the structural QA issue is warning-only and non-blocking. The repair plan JSON must expose `next_action`, `resume_scope`, and `resume_command` so users know whether to fix input files, rebuild the current DOCX, rerun the full pipeline, review warning-only issues with `warning_review`, or do final Word/WPS review. Warning-only structural plans must not summarize as a plain pass; they should say there is no blocker but there are warnings to confirm, and preserve a rerun command when available. Strict/visual reports should also name the leading issue code in the top-level `next_action` and avoid generic "go inspect the report" guidance when an issue code can name the next concrete action; `WPS_PDFINFO_UNAVAILABLE`, `WPS_PDFINFO_FAILED`, and `WPS_PAGE_COUNT_INVALID` must explicitly say to inspect WPS PDF/DOCX output, repair PDF/Poppler or blank-output causes, and rerun visual QA; `WPS_PAGE_COUNT_MISMATCH` must explicitly say to compare Word/WPS PDFs, repair the pagination cause, and rerun visual QA; `WPS_PAGE_SIZE_MISMATCH` must explicitly say to compare Word/WPS PDF paper size, page dimensions, and orientation, repair the page-setting or WPS compatibility cause, and rerun visual QA; `WPS_TEXT_PAGE_MISMATCH` must explicitly say to compare Word/WPS PDF body, TOC, formula, and image content, repair WPS export/font compatibility or generation causes, and rerun visual QA, with `visual_qa/rendered_word.txt` and `visual_qa/rendered_wps.txt` preserved when text extraction succeeds; `WPS_SAMPLE_RENDER_FAILED` and `WPS_SAMPLE_IMAGE_MISMATCH` must explicitly say to compare Word samples in `visual_qa/samples/` with WPS samples in `visual_qa/wps/samples/`, repair WPS rendering compatibility or generation causes, and rerun visual QA. Strict and visual warning-only runs still need warning-specific next actions, Markdown result labels such as "通过但有警告", and `agent_summary.md/json` must not hide them behind a plain "QA passed" status. Runtime terminal QA summaries must use the same "通过但有警告" label, show warning counts, and print the report `next_action` for warning-only runs so users are not left thinking the run is fully clean. In `agent_summary`, each structural/strict/visual report entry should expose `result_label` and render "通过但有警告" when that report passed with warnings.
- QA/user-facing reports should avoid leaking absolute local paths; use run-relative paths whenever possible.
- `workflow_mode.json` may generate copyable rerun commands only when inputs can be represented safely under this project's `Inputs/` or `Templates/`. If a user/developer passed an absolute file outside those source folders, even under another directory that happens to be named `Inputs` or `Templates`, reports must not collapse it to a fake basename command; instead they should tell the user to move/copy the file into the project source folder and rerun by file name.
- `template_profile.json/md` is the reusable template decision layer. Do not add school-name logic when a profile capability/risk flag can describe the same need.
- PDF templates are best-effort format sources: instruction-style PDFs provide text rules, sparse instruction PDFs must surface `PDF_TEMPLATE_INSTRUCTION_INCOMPLETE` warning guidance for missing rule families, visual sample PDFs must surface `PDF_TEMPLATE_VISUAL_APPROXIMATION` warning guidance for Word/WPS layout review, landscape PDFs must surface `PDF_TEMPLATE_LANDSCAPE_PAGE` warning guidance for page-orientation review, missing Poppler tools must surface `PDF_TEMPLATE_DEPENDENCY_MISSING` with environment repair guidance, password-protected or copy-restricted PDFs must surface `PDF_TEMPLATE_PROTECTED` with unlock/export-unprotected-PDF guidance, corrupt/unreadable PDFs must surface `PDF_TEMPLATE_READ_FAILED` with re-export/openable-PDF guidance, and scanned/textless PDFs must surface `PDF_TEMPLATE_UNSUPPORTED` before generation so users do not receive a misleading DOCX built from default formatting.
- When the user says "更新记忆" or asks to save durable progress, update the disk memory under `memory/` and validate with `python scripts/project_memory.py validate`.
- Do not write private test data, generated DOCX/PDF/PNG, customer content, API keys, or raw chat logs into memory.
- Standalone extractor/debug outputs must stay under `Outputs/_...`; never create derived JSON, Markdown reports, copied figures, or template assets beside private source files in `Inputs/` or `Templates/`.
- `--qa-level visual` is the preferred delivery gate for developer/product checks. It requires Word COM for PDF export and Poppler tools (`pdfinfo`, `pdftotext`, `pdftoppm`) for page/text/sample checks; missing required render tools fail visual QA and make the pipeline exit nonzero. Visual sample pages should stay bounded but prioritize cover/TOC/body anchors plus detectable figure, table, and formula risk pages so WPS/Word sample-image comparison does not only inspect early pages. `visual_report.md` must expose the key diagnostic artifact paths so ordinary users can open the rendered PDFs, text diagnostics, and sample PNG folders without digging through JSON.
- Missing local Markdown images, including images embedded in Markdown table cells, must surface as `CONTENT_IMAGE_MISSING`; unreadable or unsupported local Markdown image files must surface as `CONTENT_IMAGE_UNREADABLE` with a next step to re-export a normal PNG/JPG; remote Markdown image URLs must surface as `CONTENT_IMAGE_REMOTE_UNSUPPORTED` with a next step to download the image locally and change the Markdown link to a local relative path. DOCX image extraction failures must also surface as QA errors rather than disappearing from `content.json`.
- DOCX table-cell images must be surfaced in the content image stream; header/footer images from the content source are non-body content and must surface as `NON_BODY_IMAGE_UNSUPPORTED` unless product behavior explicitly changes.
- Chinese text needs `w:eastAsia` set (handled automatically by generator).
- A4 pagination uses dual cpl: Latin and CJK separately.
- Each pipeline run = independent Outputs directory; same-day duplicate names get `_2`, `_3`, etc.
- Office Viewer ≠ WPS/Word. Final verification MUST use WPS/Word.
- Formulas: use `latex_to_omath(r"\frac{a}{b}")` — LaTeX math string → native Word OOXML equation. Write formulas in LaTeX syntax, they auto-convert. Covers fractions, roots, sums, integrals, matrices, cases, Greek letters, arrows, accents, limits, braces, boxed, and more. See `latex_omath.py` for full reference.
- Plain-text formulas extracted from content docx must become formula items (`role="formula"`, `source="text"`) and render as native `m:oMathPara`; verify by checking the docx XML for `<m:oMathPara>` and by rendering in Word/WPS.
- Markdown `$...$` / `$$...$$` formulas in abstracts and body sections must also render as native OOXML Math; cleanup code must preserve math-only paragraphs.
- Markdown image paths must resolve relative to the `.md` file first, including common local forms such as `%20` percent-encoded spaces, `<path with spaces>` wrappers, balanced parentheses in filenames, optional image titles such as `![图](path "title")`, local `?query` / `#fragment` suffixes copied from Markdown tools, reference-style images such as `![图][id]` plus `[id]: path` with optional title lines, shortcut reference images such as `![图]` plus `[图]: path`, and images inside Markdown table cells, then copy readable local images into the current output `figures/` folder. Markdown table-cell images are promoted into the content stream with `location="markdown_table_cell"` so they are not silently dropped. Undefined image references must surface as `CONTENT_IMAGE_MISSING`; unreadable/corrupt or unsupported local image files must surface as `CONTENT_IMAGE_UNREADABLE`; reference-definition-like lines inside fenced code blocks must remain code. Remote `http://` / `https://` image URLs are not downloaded automatically; users should save them locally and update the Markdown path before rerunning.
- Markdown files may start with a UTF-8 BOM before YAML/front matter `---` or before an H1 title; strip the BOM before YAML/front-format skipping and heading detection. H1 titles may also use Setext `Title` + `===`; store CJK titles as `title_cn` and non-CJK titles as `title_en`, and avoid false `TITLE_MISSING` warnings for valid H1 titles. Treat Setext `---` conservatively because it conflicts with delimiters and horizontal rules.
- Markdown front format-instruction sections are format-only. They must be stripped from content parsing, including noisy or encoding-damaged headings followed by obvious format rules and a `---` delimiter.
- `内容提取.md` should display structured images, including `role="image"` table-cell/body images, tables, and formulas by role; avoid duplicate image listings; and never label non-formula structured content as `[公式]` or opaque `[结构化内容]`.
- Caption detection must not treat prose references such as `图 1 展示了...` or `表 1 显示...` as captions.
- Generated scripts should suppress Python bytecode cache creation so `Outputs/` stays clean.
- Matrix short-cut: `formula_build_matrix(cells, cols, brackets)` as alternative.
- All legacy formula tools (`formula_text/remove/replace`) remain available.
- OOXML math: every `m:r` needs `m:rPr` (even empty) for WPS compatibility.
- Comments: use `CommentCollector` from `comment_utils.py`. Add `comment="导师: ..."` parameter to `body()`/`heading()` calls.
- TOC: the pipeline renders static visible TOC lines; Word COM can resolve heading page numbers automatically when available.
- Formula numbering: `\tag{1.1}`, `\begin{equation}`, `\begin{align}` for auto-numbered equations.

---

## Project Files

```
run_pipeline.py              ← One-click entry
memory/                      ← Disk-backed project memory: summaries + JSONL audit streams
scripts/project_memory.py    ← Memory append/validation helper
Paper_Project/Program/pipeline/
    format_extractor.py       ← Phase 1: template → format JSON
    format_extractor_modules/ ← format_extractor submodules: PDF template parsing, OOXML metrics, style inheritance, style profiles, cover assets/tables
    content_parser.py         ← Phase 2: content → structured JSON
    content_parser_modules/   ← content_parser submodules: extraction orchestration, placeholders, styles, text cleanup, front matter, caption flow, paragraph streams, body dispatch, source TOC, images, tables, formula labels/OMML/text items/repair strategies, headings, references, section building
    md_parser.py              ← MD parser (format + content from .md)
    md_parser_modules/        ← md_parser submodules: content extraction orchestration, format extraction, math tokens, image paths, tables, text cleanup
    template_profiler.py      ← template capability profile entry point
    template_profiler_modules/← template profile construction and report writers
    script_generator.py       ← Phase 3: JSON → build_generated.py
    script_generator_modules/ ← script_generator submodules: generator orchestration, runtime assembly, sections/front matter, template rules, style profiles, and runtime template fragments
    pipeline_runner/          ← run_pipeline orchestration helpers
    qa_checker_modules/       ← qa_checker submodules: phase checks, issue registry, metrics, repair plans, report writers
    qa_conformance_modules/   ← qa_conformance submodules: OOXML helpers, content/style checks, DOCX XML checks, requirements, reports
    qa_visual_modules/        ← qa_visual submodules: export, PDF/render, image stats, golden, reports
    latex_omath.py            ← LaTeX→OOXML formula converter
    latex_omath_modules/      ← latex_omath submodules: tokenizer, parser, API helpers, symbol registries and OOXML builders
    qa_checker.py             ← Generated-output QA report and fix-target routing
    qa_visual.py              ← optional PDF/render QA
    privacy.py                ← path sanitization helpers
    public_template_suite.py  ← public-template compatibility suite; downloads/runs stay local
    public_template_suite_modules/ ← public-template paths, storage, runner, reporting, scenarios, and asset helpers
    regression_suite.py       ← synthetic regression tests
    regression_suite_modules/ ← regression harness, fixtures, and grouped case modules
    comment_utils.py          ← Word comment injection system
    comment_utils_modules/    ← comment_utils implementation: OOXML comments, rels, content types
Paper_Project/基础操作.md     ← ★ YOUR TOOLBOX: all OOXML code snippets
build_acta_manuscript.py      ← Reference: Acta journal paper
build_comprehensive_doc.py    ← Reference: all features demo
```

## Workflow Diagram

```
Templates/模版.docx/.pdf ─→ format_extractor ─→ Outputs/format.json
    or .md (# 格式说明)     or md_parser         Outputs/格式提取.md

format.json ─────────→ template_profiler ─→ Outputs/template_profile.json

Inputs/内容.docx/.md ──→ content_parser ──→ Outputs/content.json
                        or md_parser          Outputs/内容提取.md

format.json ──┬──→ script_generator ──→ build_generated.py
content.json ─┘

build_generated.py ──→ python run ──→ 最终论文.docx
                           ↓
                      qa_checker / qa_visual
                           ↓
                   Codex + 基础操作.md
                   (user fine-tuning; core fixes for maintainers)
```
