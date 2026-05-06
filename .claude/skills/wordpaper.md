---
name: wordpaper
description: Run the Word paper typesetting pipeline, or fine-tune formatting of generated papers.
---

# Word Paper Typesetting Pipeline

When user says "排版" "排论文" "typeset" "wordpaper" or "run pipeline", execute the full workflow. When they say "微调" "fine-tune" or "adjust formatting", edit build_generated.py.

## Full Pipeline

### 1. Check Environment

```bash
python --version
python -c "import docx; from PIL import Image; print('OK')"
```

If dependencies are missing: `python -m pip install python-docx Pillow`.

### 2. Scan Files

```bash
ls Templates/*.docx 2>/dev/null
ls Inputs/*.docx 2>/dev/null
```

- No .docx found: tell user "Please put your template .docx in Templates/ and your content .docx in Inputs/, then re-run."
- Files found: list them, let user choose, or just run `python run_pipeline.py` which handles interactive selection.
- Format: `.docx` only (Office 2007+), no `.doc` files.

### 3. Run

```bash
python run_pipeline.py
```

The script auto-scans files, lets user pick from a numbered list, and creates `Outputs/{date}_{name}/` for each run.

### 4. Verify Results

- Read `Outputs/<latest>/格式提取.md` — check paragraph count, font sizes, fonts
- Read `Outputs/<latest>/内容提取.md` — confirm all sections present
- Confirm the final .docx was generated

### 5. Report

Brief summary: X paragraphs / X sections / X references, N pages, output directory.

## Fine-Tuning

Find the latest `Outputs/*/build_generated.py`, edit per user's instructions, then:

```bash
python Outputs/<dir>/build_generated.py
```

### Quick Reference

| Request | Location |
|------|------|
| Body font/size/line-spacing/alignment/indent | `body()` function |
| Heading size/alignment | `heading1/2/3()` functions |
| Reference font size | `D['ref_size']` |
| Image width | `D['img_width']` |
| Add page break before section | `doc.add_page_break()` before that section |
| Table cell font size | `C()` function `size` parameter |

## Important Rules

- All format values come from `P{}` (template extraction) or `D{}` (derived), never hardcoded
- Chinese text must set `w:eastAsia`, otherwise WPS falls back to Mincho font
- A4 pagination uses dual cpl: Latin cpl and CJK cpl computed separately
- Never run `run_pipeline.py` after editing `build_generated.py` — it will overwrite your edits
- Each pipeline run creates a separate output directory, no overwrites
