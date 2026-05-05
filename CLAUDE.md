# Paper Project v1 — Assistant Workflow

## Your Job

You are an AI assistant embedded in this project. When a new session starts, you MUST follow the workflow below. Do NOT skip steps.

## User's First Instruction

The user will say:
> "检查环境，然后按项目流程执行"

This means: run Entry Checklist → verify files → execute pipeline → report results.

---

## Entry Checklist (every session)

### 0. Environment Check
Run these first:
```bash
python --version           # expect 3.10+
python -m pip --version    # use python -m pip, NOT bare pip
python -c "import docx; from PIL import Image; print('OK')"
```
If anything missing, read `从零开始配置指南.md` and install.

### 1. Understand the Project
```
run_pipeline.py              ← ★ User runs this ONE file
Templates/                   ← User drops template .docx here
Inputs/                      ← User drops content .docx here
Outputs/                     ← All generated outputs
Paper_Project/Program/pipeline/
    format_extractor.py       ← Phase 1: template → format JSON + MD
    content_parser.py         ← Phase 2: content → structured JSON + images
    script_generator.py       ← Phase 3: JSON → build_generated.py
Paper_Project/Program/
    build_acta_manuscript.py  ← Reference: Acta journal paper
    build_comprehensive_doc.py ← Reference: all features demo
Paper_Project/基础操作.md     ← Code snippets: tables, refs, headers, pagination
```

### 2. Check Templates/ and Inputs/
Verify user has placed files there. If not, tell them to drop files first.

### 3. Run the Pipeline
```bash
cd <project_root>
python run_pipeline.py
```
Or run each phase manually if debugging:
```bash
python Paper_Project/Program/pipeline/format_extractor.py Templates/模版.docx
python Paper_Project/Program/pipeline/content_parser.py Inputs/内容.docx
python Paper_Project/Program/pipeline/script_generator.py Outputs/format.json Outputs/content.json
```

### 4. Verify Outputs
- Check `Outputs/格式提取.md` — spot-check paragraph counts match
- Check `Outputs/内容提取.md` — verify no sections lost
- Run `python Paper_Project/Program/build_generated.py` — must succeed

### 5. Help User Fine-Tune
Open `build_generated.py` and accept user's natural language requests:
- "把一级标题改成红色" → edit heading1()
- "参考文献字体改成10pt" → edit the refs loop
- "在第3节后加分页" → add `doc.add_page_break()`

---

## Critical Technical Notes (from real debugging)

### Python
- **Always** use `python -m pip install`, never bare `pip` — avoids multi-version mismatch

### Image Handling
- `doc.add_picture()` creates its OWN paragraph — set alignment via `doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER`
- Image width: 4.0-4.5 inches for A4 with 2.5cm margins

### Three-Line Tables (三线表)
- Call `three_line_table(table)` **AFTER** filling ALL cell content
- Implementation in `基础操作.md` — copy-paste the 3 functions
- Must: remove tblBorders → set tcBorders per cell → thick='12' thin='4' nil='nil'

### Cross-References (正文[N] ↔ 参考文献)
- Internal: use `w:anchor` (NOT OPC `r:id`) — avoids document corruption
- External URLs: use OPC `r:id` + `RT.HYPERLINK`
- Bookmarks: `w:bookmarkStart`/`w:bookmarkEnd` with matching `w:name` = `_Ref{N}`
- Reference markers: blue `0000FF` superscript
- Full implementation in `基础操作.md`

### Page Numbers & Headers
- Use `PAGE` field code (dynamic), NOT static text
- `w:fldChar begin` → `w:instrText PAGE` → `w:fldChar end`
- Office Viewer does NOT render headers/footers — visible only in real Word
- Add body page numbers for Office Viewer: gray `— N —` before each page break

### A4 Pagination
- `cpl` (chars per line) computed from page geometry: `text_w_pt / (body_size × font_ratio)`
- Dual cpl: Latin `cpl` (TNR ratio 0.42) + CJK `cpl_cjk` (宋体 ratio 1.0)
- `_ph()` auto-detects CJK characters (`_is_cjk()`) and uses `cpl_cjk` for Chinese paragraphs
- `round(chars/cpl)` for line count — no accumulating `ceil` error
- Usable height: `(page_h - mt - mb) / 0.0352778` — exact theoretical, no fudge

### CJK Font Handling (w:eastAsia)
- `r.font.name` only sets `w:ascii`/`w:hAnsi` — Chinese characters need `w:eastAsia`
- `script_generator.py` auto-detects CJK font from template Chinese paragraphs
- When `body_font` (TNR) ≠ `cjk_font` (宋体), all generated runs include `w:eastAsia` fix
- Otherwise WPS/Word falls back to 明朝体 for Chinese characters

### .doc Files (legacy format)
- python-docx CANNOT read .doc — only .docx
- Convert: LibreOffice `soffice --headless --convert-to docx` or WPS "另存为"
- Or use `olefile` library to parse binary (fragile, avoid if possible)

### Heading Detection (content_parser)
- Label headings (`Abstract:`, `Key words:`, `摘要：`, `关键词：`) detected by text pattern regardless of length or formatting
- Case-insensitive matching for English labels (`(?i)` flag)
- Numbered headings (`1.`, `2.1`) detected via OOXML bold + size ≥ 12pt
- Label headings are split: `"Key words: body text"` → heading=`"Key words:"` + body paragraph

### Encoding
- Use `PYTHONIOENCODING=utf-8` on Windows if GBK errors appear
- All project files are UTF-8

---

## Reference: Code Snippets Location

| Need | Where |
|------|-------|
| Page setup (A4 + margins) | `基础操作.md` → 页面设置 |
| Image centering | `基础操作.md` → 图片居中 |
| Three-line table | `基础操作.md` → 三线表 |
| Cross-references | `基础操作.md` → 交叉引用 |
| Page numbers + headers | `基础操作.md` → 页眉与页码 |
| A4 pagination | `基础操作.md` → A4自动分页 |
| Full Acta paper example | `build_acta_manuscript.py` |
| All features demo | `build_comprehensive_doc.py` |

---

## Workflow Diagram

```
Templates/模版.docx ──→ format_extractor ──→ Outputs/format.json
                                              Outputs/格式提取.md (双验证)

Inputs/内容.docx ──→ content_parser ──→ Outputs/content.json
                    (提取图片到fig/)     Outputs/内容提取.md

format.json ──┬──→ script_generator ──→ Program/build_generated.py
content.json ─┘

build_generated.py ──→ python 运行 ──→ Manuscripts/最终论文.docx
                         ↓
                   与 Claude 对话微调
                   (编辑 build_generated.py → 重新运行)
```
