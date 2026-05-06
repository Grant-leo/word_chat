---
name: wordpaper
description: Word论文排版：从模板docx+内容docx一键生成格式规范的论文，支持对话微调。
---

# Word Paper Typesetting Pipeline

When the user says "排版", "排论文", "typeset", "wordpaper", or "run the pipeline", follow this workflow.

## Step 1: Check Environment

```bash
python --version
python -c "import docx; from PIL import Image; print('OK')"
```

If `python-docx` or `Pillow` are missing, tell the user and offer to install:
```bash
python -m pip install python-docx Pillow
```

## Step 2: Scan and Choose Files

```bash
ls Templates/*.docx 2>/dev/null
ls Inputs/*.docx 2>/dev/null
```

Report what you found. Ask the user which template and which content file to use. Present them as numbered lists:

```
Templates 里有:
  [1] 本科毕业论文模版.docx
  [2] Acta_Journal_模版.docx

Inputs 里有:
  [1] 张三论文.docx

用哪个模版和哪个内容？
```

If only one file in a directory, auto-select it and tell the user.

If no files found in a directory, tell the user to put files there first.

## Step 3: Run Pipeline

Once the user confirms the two files, run in non-interactive mode:

```bash
python run_pipeline.py --template "<文件名>" --content "<文件名>"
```

This runs all 4 phases without pausing for input.

## Step 4: Report Results

After the pipeline completes, read the verification outputs:

```bash
ls Outputs/<latest>/
```

Read `Outputs/<latest>/格式提取.md` and `Outputs/<latest>/内容提取.md`. Summarize:

- Template used, content used
- Number of paragraphs / sections / references extracted
- Number of pages in the final document
- Output directory path
- Any warnings or anomalies

## Step 5: Offer Fine-Tuning

After reporting, ask:

> 排版完成了。需要微调什么吗？比如正文字号、行距、标题格式、参考文献大小等。

## Fine-Tuning Workflow

When the user asks for adjustments:

1. Find the latest `Outputs/*/build_generated.py`
2. Read it, understand what the user wants to change
3. Edit the relevant function or parameter
4. Run the script to regenerate:
   ```bash
   python Outputs/<dir>/build_generated.py
   ```
5. Report what changed

### Quick Reference for Common Edits

| User says | Edit this |
|-----------|-----------|
| "正文改成 XX 号" / "body font size" | `body()` → `Pt(...)` |
| "行距改成 X 倍" | `body()` → `pf.line_spacing = X` |
| "标题居中/加粗/左对齐" | `heading1/2/3()` → `p.alignment` |
| "参考文献字号改成 X pt" | `D['ref_size'] = X` |
| "图片太小/太大" | `D['img_width'] = ...` |
| "在 XX 节前加分页" | Insert `doc.add_page_break()` before that section |
| "表格字号" | `C()` → `size=...` |

## Important Rules

- Always use `--template` and `--content` flags when calling `run_pipeline.py` from Claude
- Never call `run_pipeline.py` without arguments, or it will block on `input()`
- Do not run `run_pipeline.py` again after editing `build_generated.py` — it would overwrite
- Chinese text rendering requires `w:eastAsia` to be set (the generator handles this automatically)
- A4 pagination uses dual character-per-line metrics (Latin vs CJK)
