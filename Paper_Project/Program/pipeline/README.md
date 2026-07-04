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
- `private_corpus_audit.py` scans a local private corpus, classifies files by
  structural features only, and writes ignored reports under
  `Outputs/_private_realdata_audit/`.
- `comparison_assessment.py` aggregates structural/strict/visual/golden/reference
  evidence into an auditable decision for each run.
- `regression_suite.py` is the synthetic engine regression gate.
- `RELEASE_VALIDATION.md` is the publish-before-release validation checklist.
- Standalone extractor CLIs write debug artifacts under `Outputs/_...` by default so `Inputs/` and `Templates/` remain source-only.
- Interactive cancellation/EOF and QA/dependency interruptions must always surface a concrete next step for ordinary users.
- Explicit unsupported input formats (`.doc`, `.wps`, archives, shortcuts,
  spreadsheets, and similar non-paper files) must stop in preflight with an
  issue code and a next action instead of reaching extraction or throwing a raw
  traceback.

## Private Real-Data Audit

Real documents used for product hardening stay local and ignored. The first
stage is inventory, not formatting:

```powershell
python Paper_Project/Program/pipeline/private_corpus_audit.py Templates/<private-corpus>
```

The audit writes:

- `Outputs/_private_realdata_audit/inventory.json`
- `Outputs/_private_realdata_audit/inventory.md`
- `Outputs/_private_realdata_audit/review_queue.json`

The classifier uses only structural features such as extension, DOCX/PDF
readability, paragraph/table/image/formula counts, sections, headers/footers,
fields, comments/revisions, footnotes/endnotes, textboxes, content controls,
embedded objects, landscape pages, PDF page counts, and PDF text density. It
must not store body text. Classifications are fixed to:

- `template_candidate`
- `content_candidate`
- `reference_candidate`
- `attachment_or_nonpaper`
- `unsupported_or_conversion_needed`

`Templates/<private-corpus>/` is treated as a private real-data pool and is not part of the
ordinary recursive `--agent-auto` file choice path. `.doc` / `.wps` files are
blocked by default and should be manually saved as DOCX before joining an E2E
matrix.

## Comparison Assessment

Every QA-enabled pipeline run writes `comparison_assessment.json/md` after QA
or after an auto-repair stop. The assessment reads available
`qa_report.json`, `conformance_report.json`, `visual_report.json`,
`reference_compare_report.json`, and golden-baseline artifacts, then emits:

- `decision`: `BLOCKED_UNJUDGEABLE`, `FAILED_AUTOMATIC`,
  `PASSED_WITH_REVIEW`, `PASSED_MACHINE`, or `APPROVED_DELIVERY`
- `confidence`
- `baseline_type`
- `manual_review_required`
- `review_pages`
- `approved_deviations`
- `blocking_issue_codes`

Warning-only runs are classified as `PASSED_WITH_REVIEW`; they are never counted
as perfect machine passes until the manual review/approval trail is recorded.
Golden baselines are compare-only unless `--update-golden` is explicitly used
after human approval.

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

Current baseline as of 2026-07-04:

- Synthetic regression after private real-data audit/source-audit/comparison-assessment hardening, boxed-content recovery, native note rendering, merged-table preservation, rectangular merged-cell coalescing, legacy `hMerge` horizontal-merge preservation/source-audit/private-inventory counting, legacy `hMerge` + `vMerge` rectangular coalescing, nonrectangular legacy `hMerge` + `vMerge` fail-open/source-audit handling, mixed `gridSpan` + legacy `hMerge` duplicate-encoding de-duplication/source-audit guidance, mixed `gridSpan` + legacy `hMerge` + `vMerge` 2D rectangle de-duplication/source-audit guidance, visible mixed `gridSpan` + legacy `hMerge` + `vMerge` continuation-text/image fail-open handling, rich `vMerge continue` fail-open handling/source-audit guidance, visible irregular `vMerge continue` source-audit guidance, `gridBefore` vertical-merge round-trip preservation with generated leading row-omission restore, `gridAfter` trailing row-omission restore, row-omission media guard for `gridBefore` / `gridAfter` cells with images, inline formulas, nested tables, direct formula items, direct note refs, or `math=[...]` list payloads, row-omission text-guard manifest evidence for non-empty omitted-zone text, `gridAfter` trailing row-omission source-audit wide-table counting, revision-wrapped table row source-audit width counting, deleted revision table-risk suppression, table-column-width preservation, partial/zero table-column-width repair, table layout-detail preservation, explicit table/cell border preservation, border/layout property-order hardening, source-order six-level nested table preservation, too-deep nested-table visible-text/media/formula/note flattening, too-deep multi-cell note/image/caption source-order preservation, DOCX table-cell and nested-table image in-cell preservation, table-cell image/text run-order preservation, table-cell block/inline/nested-inline content-control preservation and de-duplication, content-control simple-field visible-value preservation, content-control transparent-container visible-value preservation, body transparent-container paragraph/table preservation, body content-control paragraph/table in-place preservation, body content-control inline rich media/formula/note preservation, DOCX revision final-view extraction for body/table/table-row/table-cell/heading/textbox-recovery content, content-control hyperlink media/formula/note preservation, body content-control/table-text overlap/exact-match dedupe protection, table-cell inline OMML formula preservation, table-cell inline LaTeX formula preservation, mixed table-cell image/LaTeX/OMML ordering preservation, inline table-cell image/formula/footnote ordering preservation, nested table-cell inline image/LaTeX/OMML/footnote ordering preservation, table-cell footnote reference preservation, table-cell image-before-note-only reference preservation, irregular merge-grid audit, orphan/mismatched `vMerge` visible-text repair, `gridSpan` overflow width repair, tall table body-row split pagination, structured table-caption landscape-section grouping, consecutive compatible landscape-table no-note grouping, adjacent landscape-table short-note/short-sentence-note grouping, numbered-heading/trailing-dot bridge table-caption promotion guard, bounded rich bridge-note landscape-table grouping, long explanatory bridge split-back-to-portrait handling, display-formula run/item bridge split-back-to-portrait handling, legacy direct rich-text formula-item bridge split-back-to-portrait handling, rich-text image-run/item plus table/code/caption-item bridge rendering and split-back-to-portrait handling with mixed block/media source-order, inline item/display run math-shape preservation, and run-item nested block/media source-order preservation, list-like, roman-number, bracketed-number, and Chinese-numbered bridge split-back-to-portrait handling, plain overwide-table auto-landscape protection, source-audit DOCX fixture rewrite hygiene, section-scoped/de-duplicated wide-table risk counts, DOCX landscape-section wide-table page setup plus source-width preservation, Markdown GB18030/UTF-8 BOM safe reading, generated Unicode-literal preservation, unicode-escape decode guard, generated-script unsafe-unicode-decode QA blocking including general `codecs.decode(...)` text re-decoding routes, `codecs` import aliases, dynamic `getattr(codecs, "decode")` direct/assignment aliases, `functools.partial(codecs.decode, ...)` bound decode aliases, `codecs.decode` list/tuple/object-attribute/object-alias/SimpleNamespace/class-attribute/`__init__` instance-attribute/class-alias/factory-instance/temporary-instance/staticmethod-factory/classmethod-factory/direct-child-local-function handoffs, no-required-argument functions returning `codecs.decode` / `getattr(codecs, "decode")` / `functools.partial(codecs.decode, ...)` or those attributes, method `.encode(...).decode(...)`, bytes `.decode` list/tuple/dict container handoffs and no-required-argument function returns, `memoryview(...).tobytes()` re-decoding, plus constructor-based, simple constructor-alias, and `builtins` import/getattr constructor-alias `str(bytes/bytearray, encoding)` / `bytes(text, encoding).decode(...)` / `bytearray(text, encoding).decode(...)` mismatched text-byte re-decoding routes, direct `codecs.decode` assignment aliases, wrapper-function forwarding, higher-order decode helpers including no-required-argument functions returning direct-child higher-order wrappers and higher-order calls that pass no-required-argument decode-function results or `codecs` module objects/results/attributes before internal `module.decode` / `getattr(module, "decode")` calls, including local aliases assigned from those dynamic module-parameter attributes, normalized codec spellings, decoder factory routes including non-unicode-escape charset re-decoders, decoder-factory list/tuple/object-attribute/object-alias/SimpleNamespace/class-attribute/`__init__` instance-attribute/class-alias/factory-instance/temporary-instance/staticmethod-factory/classmethod-factory/direct-child-local-function handoffs and no-required-argument function returns, higher-order decoder-factory parameters receiving `codecs.getdecoder` / `codecs.lookup`, no-required-argument helpers returning those factory functions, and `codecs` module-object parameters/attributes that call `module.getdecoder` / `module.lookup` / dynamic `getattr(module, ...)`, including local aliases assigned from `getattr(module, "getdecoder")` / `getattr(module, "lookup")`, same-named global/local safe factory protection, and static codec-name expressions, and recursive structural/strict QA counting for nested rich images, formulas, body tables, notes, and deep text: `419 passed, 0 failed`.
- DOCX textbox/content-control recovery: visible text inside `w:txbxContent`
  and `w:sdtContent` is now recovered into the body stream, deduplicated, and
  reported in content metadata. The fallback uses the same Word final-view text
  rules as normal paragraphs, so `w:moveTo`/`w:ins` stays visible while
  `w:moveFrom`/`w:del` stays out; textbox position fidelity remains a warning
  requiring Word/WPS review.
- DOCX footnote/endnote coverage: paragraph-level note anchors and visible note
  text are extracted, rendered as native Word footnotes/endnotes, and protected
  by manifest count mismatch QA.
- DOCX body table fidelity: basic merged cells, column widths, row heights,
  repeated header rows, default cell margins, cell-level margins, and vertical
  alignment now round-trip through `content.json` and generated DOCX XML.
  Rectangular cells that combine `gridSpan` and `vMerge` are coalesced into
  one merge record, so generated scripts do not repeat overlapping merge
  operations.
  Legacy `w:hMerge restart/continue` horizontal merges are normalized into
  `table_merges` and generated as standard `gridSpan`, while source audit and
  private inventory count them as merged-cell review risk.
  Mixed modern `gridSpan` plus legacy `hMerge` duplicate encodings use the
  `gridSpan` width once, skip empty duplicate continuation cells, and surface
  `irregular_hmerge_count` in source-audit details for review.
  When the same duplicated encoding also appears inside a `vMerge` 2D block,
  empty redundant continuation cells are skipped and same-width horizontal
  merge records are folded into the single rectangular merge.
  Non-empty duplicated continuations fail open as visible cells so user content
  is not cleared as an invisible merge placeholder. Source audit counts
  non-empty duplicate continuations as `visible_hmerge_continuation_count`
  without leaking their text. Structural QA repair guidance now names
  `visible_hmerge_continuations` directly and tells users to compare the source
  table with the final DOCX in Word/WPS when those visible continuation cells
  appear.
  Legacy `hMerge` rectangles combined with `vMerge` are also coalesced into
  one merge record.
  Nonrectangular legacy `hMerge` / `vMerge` combinations fail open by
  preserving continuation text as visible content and keeping only the safe
  horizontal merge; source audit marks the table as irregular for review.
  Orphaned or span-mismatched `vMerge continue` cells keep their visible text
  as normal cells instead of being cleared as invisible merge placeholders;
  source audit still marks the table for Word/WPS visual review. If those
  irregular continuation cells carry visible content, including a `gridAfter`
  trailing-omission mismatch in a mixed `gridSpan`/`hMerge`/`vMerge` table,
  audit detail also records `visible_vmerge_continuations` without leaking
  the cell text.
  Valid `vMerge continue` cells still follow Word merge semantics for plain
  hidden continuation text, but rich continuation payloads such as images,
  formulas, nested tables, or note anchors fail open as visible cells. Source
  audit records `visible_vmerge_continuation_count`, and QA guidance names
  `visible_vmerge_continuations` so users know to compare the source and final
  tables in Word/WPS.
  If a `gridSpan` extends beyond an incomplete `tblGrid`, missing generated
  grid widths are repaired from source cell-width evidence instead of producing
  zero-width columns. Generated table-width data with partial or explicit zero
  column widths is also repaired before `tblGrid`/`tcW` are written, preserving
  known positive widths and filling missing columns with a visible fallback.
  Repeated header rows keep `tblHeader`/`cantSplit`; unusually tall non-header
  rows omit `cantSplit` so Word/WPS can split them across pages instead of
  forcing large blank areas or row overflow.
- DOCX landscape and overwide tables: tables extracted from landscape sections
  carry source page setup into `content.json`; generated DOCX creates a
  landscape section around the table and uses the source landscape text width
  before scaling columns. Plain top-level tables that are clearly too wide for
  the current portrait text area are also wrapped in a generated landscape
  section, then the following body returns to the template's portrait section.
  Structured `table_caption` items immediately before a landscape or
  auto-landscaped table are rendered inside the same landscape section so
  captions do not stay behind on the portrait page. Consecutive compatible
  landscape table groups with no bridge note also share one landscape section
  instead of creating an unnecessary portrait/landscape flip between tables.
  Adjacent landscape tables separated only by a short note, including a very short English sentence such
  as `Short bridge note.`, share the same landscape section, so the note stays
  with the table group and the document avoids unnecessary
  portrait/landscape page flips. Longer explanatory bridge prose that exceeds
  the compact-note budget is treated as body flow: the generator closes the
  first landscape section, restores portrait flow for that prose, and opens a
  separate landscape section for the next wide table. Bridge content containing
  display/block math
  is not grouped: the generator closes the first landscape section, restores
  portrait flow for the equation paragraph, and opens a separate landscape
  section for the next wide table. Legacy rich-text formula child items that
  carry `latex` or `xml` directly are treated the same way instead of being
  mistaken for compact table notes. Bridge content carrying rich-text image runs
  is also treated as body flow and the image run is rendered in the portrait
  bridge paragraph, preventing both landscape-section misgrouping and image
  loss. Bridge content carrying a block table under `rich_text.items` is also
  treated as body flow; the nested table is rendered as a portrait table before
  the next wide table opens its own landscape section, preventing silent table
  loss. Bridge content carrying a code block under `rich_text.items` follows
  the same body-flow route, rendering the code block in portrait before the
  next wide table opens its own landscape section. Bridge content carrying
  figure or table captions under `rich_text.items` also follows the body-flow
  route, rendering the caption in portrait before the next wide table opens
  its own landscape section. When one `rich_text.items` bridge carries several
  media and block children such as code, images, inline formulas, captions,
  and small tables, those children are rendered in source order instead of
  being appended by the media collector or batched by type; explicit
  `type=inline` item formulas keep inline OMML shape instead of being promoted
  to display formulas. List-like bridge paragraphs starting with
  bullets, ASCII numbering, bracketed numbering, roman numbering, full-width
  Chinese numbering, Chinese numeral markers, or circled numerals are also
  treated as body flow rather than table notes, so they close the preceding
  landscape section and render in portrait before the next wide table opens its
  own landscape section. Numbered heading-style bridge paragraphs such as
  `2.2 Experimental setup` or `2.2. Experimental setup` are preserved as
  body flow and are not promoted to a generated table caption when the
  following wide table has no explicit caption. Complex or extreme wide tables still
  require Word/WPS visual review before delivery.
- Private real-data inventory smoke: local-only private corpus scans stay under ignored output folders, report structural classifications only, and must not publish source file contents, local paths, or corpus statistics.
- Agent-first flow: `--agent-auto` scans local inputs, auto-selects only single candidates, defaults to user auto-repair, and writes `agent_summary.md/json`.
- Novice interruption coverage: interactive cancellation/EOF, missing preflight inputs, generated-script build failures, QA dependency failures, and auto-repair blockers all route to a next action.
- Strict/visual report handoff coverage: `conformance_report.md/json` and `visual_report.md/json` top-level `next_action` values name the leading issue code before the beginner-facing repair step, so users can connect codes such as `PLACEHOLDER_TEXT_LEFT`, `PDF_PAGE_COUNT_INVALID`, and `WPS_SAMPLE_IMAGE_MISMATCH` to the next concrete action even without opening `agent_summary.md`.
- QA JSON status contract: structural `qa_report.json`, strict `conformance_report.json`, and visual `visual_report.json` now expose `status` (`passed`, `passed_with_warnings`, or `failed`) plus `result_label`; dependency-missing, QA-crash, build-failure, and extraction-failure fallback reports use the same fields. `run_qa_phases()` prints non-blocking contract warnings for structural, strict, and visual QA reports, including dependency-missing and QA-crash fallback reports, when fields are missing or inconsistent, and `agent_summary.json` preserves the same per-report status so UI/Agent consumers do not have to infer warning-only state from `passed`.
- Workflow rerun command hygiene: absolute inputs outside this project's `Inputs/` / `Templates/`, including external same-named folders, do not collapse to misleading basename rerun commands; reports instead tell users to place the file in the correct source folder and rerun by file name.
- Markdown remote image handoff: remote `http://` / `https://` image URLs surface `CONTENT_IMAGE_REMOTE_UNSUPPORTED`, stop as user-file input blockers, and tell users to download the image locally and update the Markdown relative path before rerunning.
- Markdown local image path variants: local paths continue resolving `%20` spaces, `<path with spaces>` wrappers, balanced filename parentheses, optional image titles such as `![图](path "title")`, and local `?query` / `#fragment` suffixes copied from Markdown tools before checking the filesystem.
- Markdown unreadable image handoff: existing local image files that are corrupt, mislabeled, unsupported, or outside stable Word generation formats (`.png`, `.jpg`, `.jpeg`) surface `CONTENT_IMAGE_UNREADABLE`, stop as user-file blockers, and tell users to re-export a normal PNG/JPG before rerunning. GIF/WebP/SVG/no-extension local files should not flow into generated scripts as late render surprises.
- Markdown reference-style images: `![alt][id]` plus `[id]: path` and shortcut reference images `![alt]` plus `[alt]: path` now copy local images into the content stream; optional title continuation lines after reference definitions are stripped with the definition; undefined image references become `CONTENT_IMAGE_MISSING` instead of staying as ordinary body text, and reference-definition-like lines inside fenced code blocks stay in code.
- Markdown HTML images: `<img src="path" alt="...">`, lazy `data-src` / `data-original`, the first `srcset` candidate, and PNG/JPG data URI images use the same image routing as Markdown image syntax; malformed/unsupported data URI images and data URI MIME/actual-format mismatches surface as `CONTENT_IMAGE_UNREADABLE`, and raw tags or inline payloads must not leak into generated DOCX text or QA metadata.
- Markdown table-cell images: images embedded inside GitHub-style Markdown table cells are attached to `table_cell_items`, keep `location="markdown_table_cell"`, and render inside generated Word table cells; missing table-cell images remain QA-visible `CONTENT_IMAGE_MISSING` blockers instead of being stripped by table text cleanup.
- Markdown text encoding: Markdown format/content readers decode file bytes as UTF-8/UTF-8 BOM first and GB18030 fallback second, covering common Windows Chinese `.md` files without applying `unicode_escape` or `codecs.decode()` to already-Unicode text. Generated build scripts preserve both real Chinese characters and literal strings such as `\u4e2d\u6587`; structural QA blocks generated scripts that contain dangerous text decode calls with `GENERATED_SCRIPT_UNSAFE_UNICODE_DECODE`, including general `codecs.decode(...)` re-decoding routes such as UTF-8 bytes decoded again as GBK, common `codecs` module/function aliases, dynamic `getattr(codecs, "decode")` direct/assignment aliases, `functools.partial(codecs.decode, ...)` bound decode aliases, `codecs.decode` / `getattr(codecs, "decode")` stored in literal `list` / `tuple` / `dict` containers, object attributes, object-alias attributes, `SimpleNamespace(...)` attributes, class attributes, simple `__init__` instance attributes reached through direct class instantiation, class aliases, temporary class-alias instances, no-required-argument factory functions, `@staticmethod` / `@classmethod` instance factories, direct-child local functions returned through an outer no-required-argument function, or no-required-argument functions that return those attributes, no-required-argument functions that return `codecs.decode` / `getattr(codecs, "decode")` / `functools.partial(codecs.decode, ...)`, method `.encode(...).decode(...)`, `memoryview(...).tobytes()` re-decoding, plus constructor-based, simple constructor-alias, and `builtins` import/getattr constructor-alias `str(bytes/bytearray, encoding)` / `bytes(text, encoding).decode(...)` / `bytearray(text, encoding).decode(...)` mismatched text-byte re-decoding routes, direct `decode_text = codecs.decode` aliases, wrapper functions that forward codec parameters into `codecs.decode(...)`, higher-order helpers that receive `codecs.decode` as a function object, a no-required-argument helper result such as `get_decoder()`, or a `codecs` module object/result/attribute before calling `module.decode(...)`, `getattr(module, "decode")(...)`, or a local alias assigned from `getattr(module, "decode")`, no-required-argument functions returning direct-child higher-order wrappers before calls such as `build_apply()(codecs.decode, payload, "gbk")`, `build_apply()(get_decoder(), payload, "gbk")`, or `build_apply()(codecs, payload, "gbk")`, normalized codec spellings such as `unicode-escape`, `Unicode Escape`, and `raw unicode escape`, decoder-factory routes such as `codecs.getdecoder(...)` / `codecs.lookup(...).decode(...)` including non-unicode-escape charset re-decoding, decoder-factory results stored in literal `list` / `tuple` containers, object attributes, object-alias attributes, `SimpleNamespace(...)` attributes, class attributes, simple `__init__` instance attributes reached through direct class instantiation, class aliases, temporary class-alias instances, no-required-argument factory functions, `@staticmethod` / `@classmethod` instance factories, direct-child local functions returned through an outer no-required-argument function, no-required-argument functions that return `codecs.getdecoder(...)` / `codecs.lookup(...)` results or those attributes, higher-order helpers that receive `codecs.getdecoder` / `codecs.lookup` factory functions, no-required-argument helpers returning those factory functions, or `codecs` module-object parameters/attributes before constructing decoder objects internally through `module.getdecoder` / `module.lookup`, dynamic `getattr(module, ...)`, or local aliases assigned from those dynamic module-parameter attributes, and statically provable codec-name expressions such as `codec = "unicode" + "_escape"`.
- Generated-script text re-decoding guard: the same QA route treats `bytes.__new__(bytes, text, encoding)` and simple constructor-alias `.__new__` calls as text-derived byte boundaries, recognizes `__builtins__.bytes`, `getattr(__builtins__, "bytes")`, and `__builtins__["bytes"]` constructor aliases before mismatched `.decode(...)` / `str(payload, encoding)` calls, and follows direct-child local-function result handoffs such as `def build_decoder(): def local(): return codecs.decode; return local()` before `build_decoder()(payload, "gbk")`. Regression: `qa_flags_generated_script_method_decode_text_reencoding` / `qa_flags_generated_script_general_codecs_decode_text_reencoding`.
- Generated-script type-derived constructor guard: the same encoded-byte boundary logic recognizes static `type(b"")`, `b"".__class__`, `type(bytearray())`, and `bytearray().__class__` aliases before mismatched `.decode(...)` / `str(payload, encoding)` calls, while leaving unrelated `type("...")` factories alone.
- Generated-script dynamic `codecs.decode` guard: structural QA now also recognizes statically provable dynamic module/function routes such as `__import__("codecs").decode(...)`, `importlib.import_module("codecs").decode(...)`, `codecs.__dict__["decode"](...)`, `vars(codecs)["decode"](...)`, `globals()["codecs"].decode(...)` / `locals()["codecs"].decode(...)`, literal dictionaries that store either `codecs`, `codecs.decode`, or their simple aliases, literal `list` / `tuple` containers holding `codecs.decode` / `getattr(codecs, "decode")`, object attributes and object-alias attributes such as `box.decode = codecs.decode; alias.decode(...)`, `SimpleNamespace(decode=codecs.decode)`, class attributes such as `class Holder: decode = codecs.decode`, simple `__init__` instance attributes such as `self.decode = codecs.decode; holder.decode(...)`, class-alias named and temporary instances such as `Alias = Holder; holder = Alias()` and `Alias().decode(...)`, no-required-argument instance factories such as `def make_holder(): return Holder(); holder = make_holder()`, `@staticmethod` factories such as `holder = Holder.make()`, `@classmethod` temporary factories such as `Holder.make().decode(...)`, direct-child local function return handoffs such as `def build_decoder(): def local(): return codecs.decode; return local()`, no-required-argument functions that return direct-child higher-order wrappers such as `def build_apply(): def apply_decoder(decoder, value, encoding): return decoder(value, encoding); return apply_decoder`, higher-order calls that pass a no-required-argument decode-function result such as `apply_decoder(get_decoder(), payload, "gbk")`, higher-order helpers that receive `codecs`, `get_module()`, or a module-object attribute such as `box.module` before calling `module.decode(payload, encoding)`, `getattr(module, "decode")(payload, encoding)`, or a local alias such as `decode = getattr(module, "decode"); decode(payload, encoding)`, and no-required-argument functions that return `codecs.decode`, `getattr(codecs, "decode")`, `functools.partial(codecs.decode, ...)`, or those attributes before re-decoding UTF-8 Chinese bytes as GBK/Latin-1-style text. Regression: `qa_flags_generated_script_general_codecs_decode_text_reencoding` / `qa_flags_generated_script_partial_codecs_decode_text_reencoding`.
- Generated-script decoder-factory handoff guard: structural QA now follows decoder objects returned by `codecs.getdecoder(...)` and codec-info objects returned by `codecs.lookup(...)` after they are stored in literal `list` / `tuple` containers, object attributes, object-alias attributes, `SimpleNamespace(...)` attributes, class attributes, simple `__init__` instance attributes, class-alias named or temporary instances, no-required-argument instance factories, `@staticmethod` / `@classmethod` instance factories, direct-child local functions returned by an outer no-required-argument function, or returned by no-required-argument functions, before calls such as `routes[0](payload)`, `box.decoder(payload)`, `alias.codec.decode(payload)`, `Holder.decoder(payload)`, `holder.decoder(payload)`, `holder.codec.decode(payload)` after `holder = Alias()`, `holder = make_holder()`, or `holder = Holder.make()`, `Alias().codec.decode(payload)`, `Holder.make().codec.decode(payload)`, `build_decoder()(payload)`, or `build_codec().decode(payload)`. It also follows higher-order helpers that receive `codecs.getdecoder` / `codecs.lookup` factory functions, no-required-argument helpers returning those factory functions, or `codecs` module objects/attributes before creating the decoder inside the helper, including returned wrapper forms such as `build_apply()(codecs.getdecoder, payload, "gbk")`, `build_apply()(get_factory(), payload, "gbk")`, `build_apply()(codecs.lookup, payload, "gbk")`, and `apply_module(codecs, payload, "gbk")` / `apply_module(box.module, payload, "gbk")` where the helper calls `module.getdecoder(...)`, `module.lookup(...).decode(...)`, `getattr(module, "getdecoder")(...)`, `getattr(module, "lookup")(...).decode(...)`, or local aliases such as `factory = getattr(module, "getdecoder")` / `lookup = getattr(module, "lookup")`. Regression: `qa_flags_generated_script_general_codecs_decoder_factories_text_reencoding`.
- Generated-script dynamic bytes `.decode` guard: structural QA now follows bytes decode methods obtained via `getattr(payload, "decode")`, `payload.decode` bound-method aliases, `payload.__getattribute__("decode")`, `operator.methodcaller("decode", ...)`, and `operator.attrgetter("decode")(payload)` when `payload` is text-derived bytes. It also follows those decode callables through simple literal `list` / `tuple` / `dict` containers such as `routes = [payload.decode]` or `routes = {"decode": operator.methodcaller("decode", "gbk")}` and no-required-argument functions such as `def get_decoder(): return payload.decode` or `return operator.methodcaller("decode", "gbk")` before a fixed-index/key/immediate call. Same-codec routes such as UTF-8 bytes decoded as UTF-8 remain allowed; mismatched or dynamic re-decodes remain `GENERATED_SCRIPT_UNSAFE_UNICODE_DECODE`.
- QA recursive content counts: structural QA and strict `template_requirements` now recurse through `rich_text.items`, `runs[].items`, `runs[].table_cell_items`, and table-cell item trees for images, formulas, body tables, footnote/endnote anchors, and deep nested text. Code-block `table_rows` remain excluded from body-table counts even when `_meta.tables_count` is nonzero.
- DOCX table-cell images: images inside DOCX body table cells are attached to the source cell's `table_cell_items` instead of being appended after the table; same-paragraph image-before-text and text-before-image ordering is preserved from OOXML run order; images inside supported six-level nested tables stay inside the nested cell, and structural/strict image counts recurse through nested `table_cell_items`.
- DOCX table-cell inline OMML formulas: formula-bearing cell paragraphs are attached as replaceable `rich_text` cell items and rendered back as native inline Word math inside the generated table cell instead of being flattened into plain text.
- DOCX table-cell inline LaTeX formulas: plain cell text such as `Energy $E=mc^2$ model` is split into text/math/text rich runs and rendered as native inline Word math without leaking `$...$` delimiters into the final DOCX.
- DOCX mixed table-cell media/formulas: when a cell contains preceding paragraphs, an image paragraph, and a later paragraph mixing `$...$` LaTeX with inline OMML, the image remains before the formula paragraph and both formula sources render as native Word math.
- DOCX inline table-cell media/formula/notes: when the same source cell paragraph contains text, an inline image, later LaTeX/OMML formulas, and a footnote reference, extraction splits the cell at the image boundary so the generated table keeps text, image, formulas, and note in source order.
- DOCX nested table-cell inline media/formula/notes: supported six-level nested tables preserve the same source order when a nested cell paragraph mixes text, an inline image, LaTeX, OMML, and a footnote reference.
- DOCX table-cell content controls: block-level, inline, and nested inline `w:sdtContent` inside a source table cell are consumed in cell order, rendered back in the same generated table cell, and skipped by the body-level content-control fallback so they are not duplicated after the table. Simple-field containers (`w:fldSimple`) and transparent containers (`w:customXml` / `w:smartTag`) inside those controls keep their visible result text in source order. Hyperlink containers inside those controls keep mixed images, LaTeX/OMML formulas, and note anchors in source order. Body-level content controls that partially overlap or exactly match table-cell text are still recovered as body content.
- DOCX body content controls: body-level `w:sdtContent` wrappers around paragraphs and tables are recursively dispatched in source order. Wrapped paragraphs count toward `_meta.recovered_content_control_paragraphs`, while wrapped tables remain table items and table-cell text is not leaked as loose body paragraphs. Inline `w:sdt`, `w:fldSimple`, hyperlink, `w:customXml`, and `w:smartTag` containers inside body paragraphs are recursively consumed so images, OMML/LaTeX formulas, and footnote/endnote anchors keep paragraph order.
- DOCX body transparent containers: body-level `w:customXml` / `w:smartTag` wrappers around paragraphs and tables are recursively expanded by the body dispatcher, so wrapped content keeps source order and following direct paragraphs do not get displaced by python-docx paragraph/table indexing differences.
- DOCX table-cell footnote references: note references in table-cell paragraphs are attached to the same `rich_text` cell item, carry extracted note text, render as native Word footnote references inside the generated table cell, and participate in recursive note-count QA.
- DOCX table-cell note-only anchors: an image followed only by a footnote anchor, with no visible text after the image, still renders as a native footnote reference after the image instead of being dropped.
- DOCX relationship-image fail-closed handoff: corrupt or unsupported image bytes in source DOCX relationships are validated before copying into `figures/`; invalid images are not counted as extracted, do not leak into the content stream, and surface `IMAGE_EXTRACT_FAILED` with a beginner-facing step to re-export/reinsert the source image as a normal PNG/JPG before rerunning.
- Strict conformance body-start detection: default body paragraphs before the first explicit Markdown heading stay inside strict content checks instead of being skipped as TOC/front matter.
- Content-summary coverage: `内容提取.md` renders structured `role="image"` items, including table-cell images, as `[图片]` and mentions table-cell image counts in table summaries instead of opaque `[结构化内容]`.
- Template-instruction cleanup coverage: DOCX template format notes, source-TOC examples, cover field hints, and TOC page-number samples are stripped or ignored before final rendering, so template instructions do not appear in `最终论文.docx`.
- Semantic heading coverage: structural QA treats common Chinese/English backmatter equivalents such as `Acknowledgements` / `Acknowledgment` / `致谢`, `References` / `参考文献`, and `Appendix` / `附录` as matching headings for `CONTENT_HEADING_MISSING`.
- Public-template visual baseline routing: `public_template_suite.py --visual` runs render QA without golden comparison by default. Golden comparison is opt-in through `--golden-dir`; `--update-golden` writes/refreshes the default baseline directory when intentionally maintaining golden data.
- Output-boundary coverage: standalone/default `format_extractor`, `content_parser`, and `md_parser` outputs stay under `Outputs/_...` instead of beside private source files.
- Controlled auto-repair loop regression: repairable build-script error, no-improvement stop, rebuild-failure stop, needs-user-file stop, strict/visual dependency failure, WPS page-count/page-size/text-page/sample-image visual blockers, visual option preservation, summary next-action promotion, and sanitized report paths passed.
- PDF template end-to-end strict QA: synthetic instruction PDF template + DOCX content passed.
- Sparse PDF instruction handoff: incomplete text-rule PDFs now surface `PDF_TEMPLATE_INSTRUCTION_INCOMPLETE` warning guidance naming missing heading/caption/reference-style rule families, while still producing the DOCX for manual warning review.
- Visual sample PDF template handoff: visual sample PDFs now surface `PDF_TEMPLATE_VISUAL_APPROXIMATION` warning guidance and `pdf_template_visual_approximation` profile risk so users know the DOCX layout is estimated and must be reviewed in Word/WPS.
- Landscape PDF template handoff: landscape PDFs now surface `PDF_TEMPLATE_LANDSCAPE_PAGE` warning guidance and `pdf_template_landscape_page` profile risk so users know to review final DOCX orientation, margins, and compressed tables/body in Word/WPS.
- PDF template dependency handoff: missing `pdfinfo` / `pdftotext` stops after template profiling and before `build_generated.py`, while writing `PDF_TEMPLATE_DEPENDENCY_MISSING` QA/agent reports with `resume_scope=environment` and Poppler repair/rerun next steps.
- Poppler discovery hardening: visual QA scans `PATH` / `PATHEXT` directly before using `where.exe` / `which -a`, so a broken early `pdfinfo.cmd` shim does not hide a later working `pdfinfo.exe` and misreport page metadata as unreadable.
- PDF template read-failure handoff: corrupt/unreadable PDFs stop after template profiling and before `build_generated.py`, while writing `PDF_TEMPLATE_READ_FAILED` QA/agent reports with re-export/openable-PDF or DOCX next steps.
- PDF template protection handoff: password-protected or copy-restricted PDFs stop after template profiling and before `build_generated.py`, while writing `PDF_TEMPLATE_PROTECTED` QA/agent reports with unlock/export-unprotected-PDF or DOCX next steps.
- Scanned/textless PDF template handoff: unsupported PDF templates stop after template profiling and before `build_generated.py`, while writing `PDF_TEMPLATE_UNSUPPORTED` QA/agent reports with DOCX/text-PDF/OCR next steps.
- LOF/LOT (list of figures/tables) coverage: `collect_figure_entries` / `collect_table_entries` scan `DATA['sections']` for figure/table captions; `add_figure_list` / `add_table_list` render static entry lines with right-aligned tab stops and dot leaders; `_infer_caption_pages_from_word_com` resolves page numbers in the two-pass build using the same `v - first_heading_page + 1` normalization as TOC entries. Each list gets its own page via `doc.add_page_break()` between TOC and body. Lists are only rendered when content has captions (no empty pages), table-list legacy fallback requires a real table object, and strict conformance skips LOF/LOT tabbed listing lines when matching expected body content. Regression: 8 LOF/LOT-focused cases pass.
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

`content_parser_modules/source_audit.py` performs a privacy-safe DOCX
pre-extraction structure audit and stores counts/issue codes in
`content.json` `_meta.source_audit`. Structural QA promotes these issues into
normal reports. P0 fail-closed / review codes include:

- `SOURCE_FORMAT_UNSUPPORTED`
- `LEGACY_DOC_UNSUPPORTED`
- `SOURCE_TEXTBOX_UNSUPPORTED`
- `SOURCE_FOOTNOTE_UNSUPPORTED`
- `SOURCE_ENDNOTE_UNSUPPORTED`
- `TRACKED_CHANGES_PRESENT`
- `COMMENTS_PRESENT`
- `CONTENT_CONTROL_UNSUPPORTED`
- `SOURCE_EMBEDDED_OBJECT_UNSUPPORTED`
- `SOURCE_LANDSCAPE_SECTION_UNSUPPORTED`
- `CONTENT_IMAGE_FORMAT_UNSUPPORTED`
- `COMPLEX_TABLE_UNSUPPORTED`
- `TABLE_MERGE_UNSUPPORTED`

`content_parser_modules/boxed_content.py` provides the first P1 fallback for
real DOCX structures that python-docx skips: visible text inside textboxes and
content controls is recovered into the normal body stream, deduplicated against
already extracted content, including table rows and nested table-cell items,
and counted in `_meta.recovered_textbox_paragraphs` and
`_meta.recovered_content_control_paragraphs`. Textbox QA remains a warning
because the floating anchor/position still needs Word/WPS visual review, but
ordinary users should no longer be asked to copy visible textbox text into
regular paragraphs before rerunning.
Containment-based table dedupe is limited to recovered content-control records
tagged as table-cell context, so a body-level content control is not dropped
just because its text is a substring of an existing table cell.
Body-level `w:sdtContent` that wraps ordinary paragraphs or tables is handled
by `content_parser_modules/body_dispatcher.py` first; the boxed-content fallback
only contributes remaining recoverable text and duplicate counts, so original
table structure and source order are preserved.

`content_parser_modules/notes.py` extracts visible `footnotes.xml` /
`endnotes.xml` note bodies and preserves `footnoteReference` /
`endnoteReference` anchors as ordered `note_ref` runs in `rich_text` items.
`script_generator_modules/runtime_notes.py` renders those runs back as native
Word footnotes/endnotes and writes the required OOXML package parts after DOCX
save. Structural QA compares extracted note references with
`build_manifest.json` rendered counts and blocks
`FOOTNOTE_RENDER_COUNT_MISMATCH` / `ENDNOTE_RENDER_COUNT_MISMATCH`; source
footnote/endnote issue codes remain warnings so users know to review final
numbering and placement in Word/WPS.

`content_parser_modules/table_extractor.py` now expands DOCX tables onto a
stable cell grid and preserves basic horizontal/vertical merged cells as
`table_merges` (`gridSpan` / legacy `hMerge` / `vMerge`) plus source grid/cell widths as
`table_col_widths_twips`, common row/header/margin/alignment layout details,
explicit table/cell borders as `table_borders` / cell override `borders`, and
up to six-level nested tables as `table_cell_items` with
`location="nested_table_cell"` and `after_paragraph_index` when the nested
table appears between visible paragraphs inside the same source cell. When a
source table nests deeper than that structured limit, the engine keeps the
deep table's final-view visible text, images, inline formulas, and note
anchors by flattening them into the parent cell instead of silently dropping
the content. In multi-cell over-depth tables, note-only anchors no longer
replace a later flattened caption paragraph at the same insertion index, so
endnotes/footnotes, images, and captions keep source order; source audit still
keeps seven-level-and-deeper nesting as a Word/WPS review risk. DOCX
images inside table-cell paragraphs are also attached as cell media with
`location="table_cell"`, including inside supported nested tables, so they do
not become ordinary body images after the table. When an image shares a source
paragraph with text, the extractor records whether the image appeared before
or after the visible text from OOXML run order. Block-level, inline, and nested
inline content controls inside source cells are consumed through their
`w:sdtContent` children in cell order, including hyperlink containers that mix
images, LaTeX/OMML formulas, and footnote/endnote anchors. Inline OMML formulas
and footnote/endnote references inside table-cell paragraphs are recorded as
replaceable `rich_text` cell items so the cell keeps native formula/note markup
instead of only the formula's plain text.
`script_generator_modules/runtime_media_tables.py` applies those merges back
into the generated DOCX, restores fixed table grids/cell widths, layout
details, explicit border OOXML, nested tables, table-cell images, and
table-cell rich text inside parent cells in source order. Note-only table-cell
rich text that has no visible replacement text is inserted at its anchor rather
than replacing a real cell paragraph, preventing later captions from being
dropped. The generator records the rendered table-fidelity counts in
`build_manifest.json`. Structural and strict image counters recurse through
nested `table_cell_items`, so mixed body images plus nested-table images are
not undercounted.
The
historical `TABLE_MERGE_UNSUPPORTED` and `COMPLEX_TABLE_UNSUPPORTED` audit
codes remain warning/review signals for seven-level-and-deeper nested /
overwide / irregular
tables. Source audit details now include irregular merge-grid counts and
section-scoped landscape wide-table risk counts; nested wide tables are counted
once, and valid `gridBefore` vertical-merge continuations are not treated as
orphaned merges or dropped during table extraction. Generated DOCX now restores
`w:gridBefore` and `w:gridAfter` row omission metadata instead of rendering the
omitted leading or trailing grid columns as visible empty cells. If an
internally inconsistent row-omission zone carries `table_cell_items` such as
images, inline formulas, nested tables, or direct note refs, the generator
skips the omission restore for that row and keeps the visible cell content
instead of deleting it as a blank
placeholder. Direct `role="formula"` / `latex` / `xml` items and `math=[...]`
list items in those preserved table-cell media paths render as native inline
OMML, and direct `role="note_ref"` / `type="note_ref"` items render as native
Word note references, including inside cells carried by the omitted zone. If
the omitted zone carries non-empty plain text, the generator also skips the
restore and records `content_table_grid_before_text_guard_rows_skipped` or
`content_table_grid_after_text_guard_rows_skipped` so the manifest explains why
the row stayed fully visible. Source audit
also counts omitted `w:gridAfter` trailing grid columns when deriving
`max_table_columns` and `wide_table_count`, so wide-table QA guidance is not
lost when a source row has only a few visible cells. Source audit also walks
final-view-visible table rows and cells inside `w:ins` / `w:moveTo`,
`w:sdtContent`, `w:customXml`, and `w:smartTag` wrappers while ignoring
deleted revision wrappers, so wide-table and irregular-table guidance is not
lost when Word wraps a row. Fully deleted revision tables are excluded from
final-view table counts, merge counts, nesting depth, wide-table counts, and
irregular-table risk metrics, while tracked-change warnings still tell users
to confirm revisions before delivery. Legacy `hMerge` rows combined with
`vMerge` columns are coalesced into one rectangular merge when the spans form
the same safe 2D block in every affected row, including mixed
`gridSpan`/`hMerge` duplicate encodings whose empty continuations are only
compatibility artifacts. Nonrectangular old-style
`hMerge` / `vMerge` conflicts preserve the continuation text and raise
irregular-table source-audit guidance instead of emitting overlapping merge
records. Orphaned or
span-mismatched `vMerge continue` cells keep their visible text as normal
cells instead of creating a fake vertical merge. Incomplete `tblGrid` data for
overflowing `gridSpan` cells is repaired with nonzero inferred column widths
so generated tables do not silently collapse extra columns. Plain top-level
overwide tables can be auto-landscaped during generation; warnings remain for
visual review when the table is extreme, irregular, or part of a complex
multi-section layout. Repeated header
rows stay together, while tall body rows are allowed to split across pages by
omitting row-level `cantSplit`. QA can therefore point users to the specific tables that need
Word/WPS visual review instead of asking them to flatten simple merged,
bordered, or six-level nested tables.

Caption detection deliberately separates true captions such as `图 1 xxx 示意图`
from prose references such as `图 1 展示了...`, so body prose keeps body style
while captions keep caption style.

`md_parser_modules/` owns Markdown-specific helper rules behind `md_parser`:
YAML/natural-language format extraction, inline/display math tokenization,
front format-instruction stripping, UTF-8 BOM/GB18030-safe Markdown file
reading, Markdown image copying/missing-image
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
`qa_visual.check_visual`: Word/WPS PDF export, PATH/PATHEXT-aware Poppler tool
discovery, Poppler text/page rendering,
sample page selection, rendered image statistics, opt-in golden-baseline comparison,
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
`Outputs/`, `Templates/`, render artifacts, customer documents, local
private memory, or user-local Codex skill files. The private memory bank and
`docs/skills/` are intentionally ignored via `.gitignore`.
