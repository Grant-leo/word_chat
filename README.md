# Word 论文自动排版流水线

<p align="center">
  <a href="#chinese"><strong>中文</strong></a> &nbsp;|&nbsp;
  <a href="#english"><strong>English</strong></a>
</p>

---

<a id="chinese"></a>

从模版 docx + 文本资料 docx（或 Markdown）一键生成格式规范的论文 docx。

## 核心思路

**格式与内容分离**。模版提供字体/字号/行距/边距，文本资料提供章节/段落/图片/参考文献。流水线自动提取两者，生成 python-docx 构建脚本，最终输出排好版的 docx。

## 推荐配套插件

### [DOCX Live Preview](https://github.com/Grant-leo/docx-livepreview) — VSCode 中像素级一致的 DOCX 预览

在 VSCode 中编辑 docx 时，**其他预览插件显示的格式与 WPS 实际渲染不一致**——字体偏移、表格错位、公式变形。你不得不在 VSCode 和 WPS 之间来回切换，严重打断排版效率。

DOCX Live Preview **直接调用 WPS 引擎渲染**，所见即所得——与 WPS 显示完全一致。专为本项目打造，强烈建议安装。

**安装方式：** VSCode 扩展商店搜索 `DOCX Live Preview` 直接安装；或从 [Releases](https://github.com/Grant-leo/docx-livepreview/releases) 下载 `.vsix`，`Ctrl+Shift+P` → `Extensions: Install from VSIX...`

## 快速开始

```bash
# 1. 安装依赖
python -m pip install python-docx Pillow

# 2. 放入文件
#    模版 docx → Templates/
#    内容 docx 或 .md → Inputs/

# 3. 交互模式（终端用户）
python run_pipeline.py
# → 自动扫描文件，编号列表供选择

# 4. 参数模式 — DOCX 模板 + DOCX 内容
python run_pipeline.py --template 模版.docx --content 论文.docx

# 5. 参数模式 — DOCX 模板 + Markdown 内容
python run_pipeline.py --template 模版.docx --content 论文.md

# 6. 纯 MD 模式（格式说明 + 内容都在一个 .md 文件里）
python run_pipeline.py --md 论文.md
```

每次运行生成独立目录 `Outputs/{日期}_{内容名}/`，互不覆盖。

### Markdown 文件格式

纯 MD 模式下，一个 `.md` 文件同时承载格式说明和论文内容：

```markdown
# 格式说明

一级标题：黑体，小三号(15pt)，加粗，居中，段前12pt。
正文：Times New Roman，小四号(12pt)，两端对齐，1.5倍行距。
页面：A4，上2.5cm，下2.4cm，左2.8cm，右2.2cm。

---

# 论文标题

## Abstract

正文内容...

## 1. Introduction

内联公式：$E = mc^2$

显示公式：
$$
\int_0^\infty e^{-x^2} dx = \frac{\sqrt{\pi}}{2}
$$

图片：![fig1](./images/chart.png)

## References

[1] Author. Title [J]. Journal, Year.
```

也可用 YAML frontmatter 替代自然语言格式说明：

```yaml
---
body_font: Times New Roman
body_size: 12
body_cjk_font: 宋体
heading1_size: 15
heading1_font: 黑体
page_width_cm: 21.0
---
```

如果完全不写格式说明，流水线使用默认值：A4、Times New Roman 12pt、1.5 倍行距。

## 架构

四个固定引擎 + 一个动态脚本 + 两个知识库：

```
┌──────────────────┐        ┌──────────────────┐
│                  │        │                  │
│  格式提取器       │──┐     │  构建脚本生成     │
│  内容解析器       │  │     │                  │
│  脚本生成器       │  │传参  │  零硬编码        │──→ 最终论文
│                  │──┘     │                  │
│  （固定不改）     │        │  AI 对话微调      │
│                  │        │                  │
└──────────────────┘        └──────────────────┘

 知识库
 ├── CLAUDE.md      ← AI 工作流
 └── 基础操作.md     ← AI 工具箱（所有 OOXML 代码片段）
```

- **四个固定引擎**：只需维护，不改动。通过 template/content JSON 传参
- **md_parser.py**：Markdown 解析器，从 `.md` 文件同时提取格式说明和论文内容
- **latex_omath.py**：独立 LaTeX→OOXML 公式转换器，像写 `.tex` 一样写公式
- **comment_utils.py**：Word 批注系统，`comment="导师: ..."` 即可添加批注
- **build_generated.py**：引擎生成的零硬编码脚本，AI 可在此基础上对话微调
- **基础操作.md**：持续维护的代码片段库——测试越多越完善，AI 能改的功能越多
- **模板缺的功能**（如交叉引用）→ 引擎不会生成 → 用户提出 → AI 查基础操作.md → 加代码 → 重跑

## 输出结构

每次运行生成独立目录，互不覆盖：

```
Outputs/
├── 2026-05-06_我的论文/
│   ├── 格式提取.md          ← 核对模版格式
│   ├── 内容提取.md          ← 核对文本内容
│   ├── format.json
│   ├── content.json
│   ├── build_generated.py   ← 生成脚本（可对话微调）
│   └── 最终论文.docx
├── 2026-05-07_另一篇/
│   └── ...
```

## 环境

- Python 3.10+
- `python -m pip install python-docx Pillow`

## 流水线四阶段

```
Templates/模版.docx ──→ [Phase 1] format_extractor ──→ format.json
    或 .md (# 格式说明)      或 md_parser              格式提取.md

Inputs/内容.docx/.md ──→ [Phase 2] content_parser ──→ content.json
                       或 md_parser                 内容提取.md

format.json ──┬──→ [Phase 3] script_generator ──→ build_generated.py
content.json ─┘

build_generated.py ──→ [Phase 4] python 运行 ──→ 最终论文.docx
```

每个阶段内建双验证：提取器独立运行两次，比对段落/表格/run 数，不一致时第三轮仲裁。

## 微调排版

```bash
# 打开生成脚本，改函数/参数，重跑
python Outputs/<目录>/build_generated.py
```

| 意图 | 位置 |
|------|------|
| 改正文字号/字体/行距 | `body()` 函数 |
| 改标题字号/居中 | `heading1/2/3()` 函数 |
| 参考文献字号 | `D['ref_size']` |
| 图片宽度 | `D['img_width']` |
| 编辑公式（LaTeX） | `body_with_formula("", [latex_to_omath(r"...")])` |
| 加批注 | `body("text", comment="导师: ...")` |
| 加目录 | 取消 `# insert_toc(doc)` 的注释 |

## 项目结构

```
├── run_pipeline.py              ← 一键入口
├── CLAUDE.md                    ← AI 工作流（Claude Code 自动加载）
├── .claude/settings.json        ← 项目权限配置（可选）
├── .gitignore
├── Templates/模版放这里.txt
├── Inputs/文本资料放这里.txt
├── Outputs/                     ← 每次运行生成独立子目录
└── Paper_Project/
    ├── 基础操作.md               ← AI 工具箱（所有 OOXML 代码片段）
    └── Program/
        ├── pipeline/
        │   ├── format_extractor.py   ← Phase 1: 模版 → 格式 JSON
        │   ├── content_parser.py     ← Phase 2: 内容 → 结构化 JSON
        │   ├── md_parser.py          ← MD 解析（格式 + 内容）
        │   ├── script_generator.py   ← Phase 3: JSON → 生成脚本
        │   ├── latex_omath.py        ← LaTeX→OOXML 公式转换器
        │   └── comment_utils.py      ← Word 批注系统
        ├── build_acta_manuscript.py  ← 参考：Acta Materialia 期刊格式
        ├── build_comprehensive_doc.py ← 参考：全功能演示
        └── master.py                 ← 参考：编排器骨架
```

## 功能

- **Markdown 内容支持**：`.md` 文件可直接作为内容输入，内置 `# 格式说明` 或 YAML frontmatter 描述排版参数
- 页面设置 (A4 + 四边距)
- 封面（字体/空行/表格全从模版提取，零硬编码）
- 三级标题自动检测（正文说明 > OOXML 直读）
- 三线表（OOXML 直写，顶粗线 / 表头细线 / 底粗线）
- 交叉引用（正文 [N] 蓝色上标 → 参考文献 w:anchor 跳转）
- 参考文献 [N] 格式，悬挂缩进，自动去重前缀
- 中文字体：自动检测 CJK 字体，设置 w:eastAsia 防止回退
- A4 自动分页（双 cpl：拉丁文 + CJK 各自度量）
- 页眉页脚（PAGE 域代码动态页码）
- 图片居中 + Fig. 图注
- **LaTeX 公式转换**：`latex_to_omath(r"\frac{a}{b}")` — 像写 .tex 一样直写 LaTeX → 原生 Word 方程
  支持分式/根式/求和/积分/矩阵/cases/希腊字母/符号/箭头/重音/函数/极限/定界符/括号/框（42+ 构造）
- **公式编号**：`\tag{1.1}`, `\begin{equation}`, `\begin{align}` 自动编号
- **双线体/手写体**：`\mathbb{R}`, `\mathcal{F}` 等数学字体
- **Word 批注**：`body("text", comment="导师: 请确认")` → 生成原生 Word 批注
- **目录 (TOC)**：`insert_toc(doc)` 生成 Word 域代码，打开后右键更新
- 传统公式工具：`formula_build_matrix()` 传参构建，`formula_text/remove/replace` 对话修改
- 双验证提取（独立运行两次交叉比对，不一致第三轮仲裁）

---

<a id="english"></a>

Automated academic paper formatting pipeline — from template docx + content docx to a beautifully formatted paper.

## Core Idea

**Separate formatting from content.** Templates define fonts, sizes, line spacing, and margins. Content documents provide chapters, paragraphs, images, and references. The pipeline extracts both, generates a python-docx build script, and outputs a fully formatted docx.

## Recommended Companion Plugin

### [DOCX Live Preview](https://github.com/Grant-leo/docx-livepreview) — Pixel-Perfect DOCX Preview in VSCode

When editing docx in VSCode, **other preview extensions render differently from WPS** — shifted fonts, misaligned tables, broken equations. You end up switching back and forth between VSCode and WPS, killing your flow.

DOCX Live Preview **uses WPS as its rendering engine** — what you see is exactly what WPS outputs. Built specifically for this project. Strongly recommended.

**Install:** Search `DOCX Live Preview` in the VSCode Extensions panel (`Ctrl+Shift+X`), or download `.vsix` from [Releases](https://github.com/Grant-leo/docx-livepreview/releases) → `Ctrl+Shift+P` → `Extensions: Install from VSIX...`

## Quick Start

```bash
# 1. Install dependencies
python -m pip install python-docx Pillow

# 2. Place your files
#    Template docx → Templates/
#    Content docx or .md → Inputs/

# 3. Interactive mode
python run_pipeline.py
# → Auto-scan files, numbered list for selection

# 4. CLI — DOCX template + DOCX content
python run_pipeline.py --template template.docx --content paper.docx

# 5. CLI — DOCX template + Markdown content
python run_pipeline.py --template template.docx --content paper.md

# 6. Pure MD mode (format + content in one .md file)
python run_pipeline.py --md paper.md
```

Each run produces an independent directory `Outputs/{date}_{content_name}/`, never overwriting previous results.

## Architecture

Four fixed engines + one dynamic script + two knowledge bases:

```
+---------------------------+       +---------------------------+
| format_extractor.py       |       |                           |
| content_parser.py         |       |   build_generated.py      |
| script_generator.py       |params |                           |
|                           |------>|   zero hardcoding         |---> final .docx
| (fixed engines)           |       |                           |
+---------------------------+       |   fine-tuned by Claude    |
                                    |   + 基础操作.md            |
                                    +---------------------------+

 Knowledge Base
 +-- CLAUDE.md       <- AI workflow instructions
 +-- 基础操作.md      <- AI toolbox (all OOXML code snippets)
```

- **Four fixed engines**: maintain only, never modify. Parameterized via template/content JSON
- **md_parser.py**: Markdown parser — extracts both format specs and content from `.md` files
- **latex_omath.py**: standalone LaTeX→OOXML formula converter — write formulas like `.tex`
- **comment_utils.py**: Word comment system — `comment="advisor: ..."` adds native Word comments
- **build_generated.py**: zero-hardcoding generated script, AI can fine-tune iteratively
- **基础操作.md**: continuously maintained code snippet library — the more it's tested, the more AI can do
- **Missing template features** (e.g. cross-references) → engine won't generate → user requests → AI checks 基础操作.md → adds code → re-run

## Output Structure

Each run produces an independent directory:

```
Outputs/
├── 2026-05-06_my_paper/
│   ├── 格式提取.md          ← verify template formats
│   ├── 内容提取.md          ← verify content
│   ├── format.json
│   ├── content.json
│   ├── build_generated.py   ← generated script (fine-tunable)
│   └── final_paper.docx
├── 2026-05-07_another_paper/
│   └── ...
```

## Environment

- Python 3.10+
- `python -m pip install python-docx Pillow`

## Pipeline Stages

```
Templates/template.docx → [Phase 1] format_extractor → format.json
                                                         format_report.md

Inputs/content.docx ──→ [Phase 2] content_parser ──→ content.json
                    (images → Inputs/xxx/figures/)    content_report.md

format.json ──┬──→ [Phase 3] script_generator ──→ build_generated.py
content.json ─┘

build_generated.py ──→ [Phase 4] python execution ──→ final_paper.docx
```

Each stage has built-in dual verification: the extractor runs independently twice, comparing paragraph/table/run counts; a third arbitration run resolves mismatches.

## Fine-Tuning

```bash
# Open the generated script, modify functions/parameters, re-run
python Outputs/<directory>/build_generated.py
```

| Intent | Location |
|--------|----------|
| Fix font size / typeface / line spacing | `body()` function |
| Change heading size / centering | `heading1/2/3()` functions |
| Reference font size | `D['ref_size']` |
| Image width | `D['img_width']` |
| Edit formulas (LaTeX) | `body_with_formula("", [latex_to_omath(r"...")])` |
| Add comment | `body("text", comment="advisor: ...")` |
| Add TOC | Uncomment `# insert_toc(doc)` |

## Project Structure

```
├── run_pipeline.py              ← one-click entry point
├── CLAUDE.md                    ← AI workflow (auto-loaded by Claude Code)
├── .claude/settings.json        ← project permissions config (optional)
├── .gitignore
├── Templates/                   ← place template docx here
├── Inputs/                      ← place content docx here
├── Outputs/                     ← independent sub-directory per run
└── Paper_Project/
    ├── 基础操作.md               ← AI toolbox (all OOXML code snippets)
    └── Program/
        ├── pipeline/
        │   ├── format_extractor.py   ← Phase 1: template → format JSON
        │   ├── content_parser.py     ← Phase 2: content → structured JSON
        │   ├── md_parser.py          ← MD parser (format + content)
        │   ├── script_generator.py   ← Phase 3: JSON → build script
        │   ├── latex_omath.py        ← LaTeX→OOXML formula converter
        │   └── comment_utils.py      ← Word comment system
        ├── build_acta_manuscript.py  ← reference: Acta Materialia journal format
        ├── build_comprehensive_doc.py ← reference: full feature demo
        └── master.py                 ← reference: orchestrator skeleton
```

## Features

- **Markdown content support**: `.md` files as content input, with built-in `# 格式说明` section or YAML frontmatter for format specification
- Page setup (A4 + four margins)
- Cover page (fonts/spacing/tables extracted from template, zero hardcoding)
- Three-level heading auto-detection (content description > OOXML direct read)
- Three-line tables (OOXML direct write: top thick / header thin / bottom thick)
- Cross-references (body [N] blue superscript → reference w:anchor jump)
- Reference [N] format, hanging indent, auto dedup prefix
- CJK fonts: auto-detect and set w:eastAsia to prevent fallback
- A4 auto pagination (dual cpl: Latin + CJK measured separately)
- Headers & footers (PAGE field code for dynamic page numbers)
- Image centering + Fig. captions
- **LaTeX formula conversion**: `latex_to_omath(r"\frac{a}{b}")` — write LaTeX directly → native Word equations
  Supports fractions/radicals/sums/integrals/matrices/cases/Greek/symbols/arrows/accents/functions/limits/delimiters/brackets/boxes (42+ constructs)
- **Formula numbering**: `\tag{1.1}`, `\begin{equation}`, `\begin{align}` auto-numbering
- **Math fonts**: `\mathbb{R}`, `\mathcal{F}` etc.
- **Word comments**: `body("text", comment="advisor: please confirm")` → native Word comments
- **TOC**: `insert_toc(doc)` generates Word field codes, right-click to update
- Traditional formula tools: `formula_build_matrix()` parameterized construction
- Dual verification extraction (two independent runs cross-compared, third arbitration on mismatch)

---

## 许可 / License

Copyright © 2025 Youwei Zhang

本软件仅供个人学习和研究使用。**未经作者明确书面授权，禁止将本软件用于任何商业目的**，包括但不限于：将本软件或衍生作品作为商业产品、付费服务、SaaS 平台的一部分进行销售、出租、许可或分发。如需商业使用授权，请联系作者。

This software is provided for personal learning and research purposes only. **Commercial use is prohibited without explicit written authorization from the author.**
