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

# 7. 指定工作模式
#    user: 只修本次输出目录的 build_generated.py
#    developer: 只修核心引擎脚本并重跑完整流水线
python run_pipeline.py --mode user
python run_pipeline.py --mode developer --template 模版.docx --content 论文.docx

# 8. 产品级检查：结构 QA + Word COM 导出 PDF + 抽样渲染
python run_pipeline.py --mode developer --qa-level visual --template 模版.docx --content 论文.docx
```

每次运行生成独立目录 `Outputs/{日期}_{内容名}/`，互不覆盖；同一天同名内容会自动追加 `_2`、`_3` 等后缀。

## 当前实现要点

- `Inputs/`、`Outputs/`、`Templates/` 中的实际论文、模板和生成产物默认视为本地隐私数据，已通过 `.gitignore` 排除；不要提交论文原文、模板文件、生成的 docx/PDF/PNG。
- `build_generated.py` 是用户侧的对话微调层：普通用户只需要和 AI 对话修改当前输出目录里的生成脚本，再运行它得到本次最终稿。
- `Paper_Project/Program/pipeline/` 是开发者维护层：通用能力、Bug 修复、格式规则升级应改核心引擎脚本，再重新运行流水线。
- `run_pipeline.py` 支持 `--mode user|developer|auto`：交互模式会询问身份，参数模式默认普通用户；每次输出都会写入 `workflow_mode.json`。
- 流水线生成后自动运行 `qa_checker.py`，输出 `qa_report.json` / `qa_report.md`，并生成 `qa_repair_plan.md/json` 与 `qa_fix_prompt.txt`。QA 不自动改代码，但会把每个问题路由到用户侧或开发者侧的下一步修复动作；检测到 `error` 时会让流水线以失败退出，避免把未通过结果当成成功。
- 流水线会生成 `template_profile.json` / `template_profile.md`，把模板能力归纳为封面、目录、标题、图表题、参考文献、公式、页眉页脚和风险标记，不依赖学校名或私有文件名。
- `--qa-level visual` 会在结构 QA 后运行 `qa_visual.py`：调用 Word COM 导出 PDF，使用 Poppler 工具抽取页数/纸张/文本并渲染抽样 PNG，输出 `visual_report.json` / `visual_report.md`；缺少必要渲染工具时 visual QA 会判定未通过并阻断流水线。
- 内容文档中的段落图片和表格单元格图片会提取到本次输出目录的 `figures/`，避免污染 `Inputs/` 或覆盖其他任务的图片；Markdown 图片路径按 `.md` 文件所在目录解析，支持相对路径。缺失、远程 Markdown 图片、DOCX 图片抽取失败都会进入 QA 错误报告，不会静默丢失。内容源里的页眉/页脚图片不会被当作正文迁移，会以 `NON_BODY_IMAGE_UNSUPPORTED` 明确提示用户移动到正文或删除装饰图。
- 内容文档中的普通文本公式会由 `content_parser.py` 识别为公式项，Markdown 的 inline/display math 会由 `md_parser.py` 提取为公式项，再由 `script_generator.py` 调用 `latex_omath.py` 生成原生 OOXML Math；行内公式保留在当前段落，显示公式使用 WPS/Word 可编辑的 `m:oMathPara`，不是图片或纯文本。单美元 `$...$` 只有在内容确实像公式时才会切分，`$100$` 这类金额/纯数字不会误转，`$$...$$` 会按显示公式处理。对源 DOCX 中被拆行的求和、分式求和、max/min 目标函数会做保守重组；最终 DOCX 若残留原始 `$...$` / `$$...$$` 分隔符，会由 QA 作为 `LATEX_DELIMITER_TEXT` 报错。
- 对已经排版过的源 DOCX，`content_parser.py` 会前瞻识别并跳过源文档静态目录，包括无页码目录项和混入重复公式/页码碎片的目录块；真实使用 Heading/标题样式的“目录”章节会保留，即使它后面跟着分页。未填写的 `[报名序号]` 等占位符、重复公式编号 `(1)(1)`、公式碎片段落（如 `E` / `rgreen` / `RE`）、疑似低分辨率文字/公式图片碎片会由 `qa_checker.py` 报为 `CONTENT_TOC_POLLUTION`、`UNFILLED_PLACEHOLDER_TEXT`、`FORMULA_NUMBER_CONFLICT`、`FORMULA_TEXT_FRAGMENTED` 或 `LOW_RES_IMAGE_FRAGMENT`，避免把坏输入静默当成可交付结果。
- 目录默认生成静态目录行；在 Windows + Word COM 可用时，会自动读取正文标题页码并写入目录页码，不再依赖手动取消 TOC 注释。
- 正文图片按正文文本宽度适配；封面图片按模板提取的宽高插入，避免校徽、Logo 被段落行高裁切。

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

核心引擎 + 一个生成脚本 + 两个知识库：

```
┌──────────────────┐        ┌──────────────────┐
│                  │        │                  │
│  格式提取器       │──┐     │  构建脚本生成     │
│  内容解析器       │  │     │                  │
│  脚本生成器       │  │传参  │  零硬编码        │──→ 最终论文
│                  │──┘     │                  │
│  （开发维护层）   │        │  用户微调层       │
│                  │        │                  │
└──────────────────┘        └──────────────────┘

 知识库
 ├── CLAUDE.md      ← AI 工作流
 └── 基础操作.md     ← AI 工具箱（所有 OOXML 代码片段）
```

- **核心引擎**：开发者长期维护入口。所有可复用修复都应落在 `pipeline/` 脚本中，通过 template/content JSON 传参
- **md_parser.py**：Markdown 解析器，从 `.md` 文件同时提取格式说明和论文内容
- **template_profiler.py**：模板画像器，把低层 `format.json` 归纳为可决策的模板能力和风险标记
- **latex_omath.py**：独立 LaTeX→OOXML 公式转换器；既支持手写 LaTeX，也承接 `content_parser.py` 从文本公式归一化出的 LaTeX
- **formula_semantics.py**：公式语义判别层，区分正文数量/单位、行内公式、显示公式和公式碎片问题，避免把正常中文段落整段转成公式
- **comment_utils.py**：Word 批注系统，`comment="导师: ..."` 即可添加批注
- **qa_visual.py**：可选 PDF 渲染 QA，检查页数、纸张、目录文本并抽样渲染页面 PNG
- **public_template_suite.py**：公共模板兼容性套件，下载公开模板到本地忽略目录，用合成内容跑结构/视觉 QA；只提交套件代码，不提交下载模板和运行输出
- **privacy.py**：报告脱敏辅助，避免对外报告暴露绝对路径
- **regression_suite.py**：开发者合成回归测试，覆盖公式语义、行内/显示公式、混合段落、MD 表格/代码、manifest 计数、表格单元格图片、页眉页脚图片告警、原始 LaTeX 分隔符残留和 QA 失败路径
- **build_generated.py**：每次运行生成的用户微调脚本。用户通过 AI 改它即可完成当前文档的排版调整；长期修复由开发者回写核心引擎
- **基础操作.md**：持续维护的代码片段库——测试越多越完善，AI 能改的功能越多
- **模板缺的功能**（如交叉引用）→ 引擎不会生成 → 用户提出 → AI 查基础操作.md → 先改当前 `build_generated.py` 完成本次文档；需要产品化时再由开发者沉淀到核心引擎

## 输出结构

每次运行生成独立目录，互不覆盖：

```
Outputs/
├── 2026-05-06_我的论文/
│   ├── 格式提取.md          ← 核对模版格式
│   ├── 内容提取.md          ← 核对文本内容
│   ├── format.json
│   ├── content.json
│   ├── template_profile.md  ← 模板能力画像
│   ├── workflow_mode.json  ← 本轮用户/开发者模式
│   ├── build_manifest.json ← 本轮正文图片/表格/公式渲染数量
│   ├── qa_report.md        ← 自动 QA 检测报告
│   ├── visual_report.md     ← PDF 渲染 QA（--qa-level visual）
│   ├── visual_qa/           ← 本地 PDF/抽样 PNG（不提交）
│   ├── figures/             ← 本次内容文档提取的图片
│   ├── assets/              ← 本次模板提取的封面/Logo 资源
│   ├── build_generated.py   ← 生成脚本（用户可通过 AI 微调）
│   └── 最终论文.docx
├── 2026-05-06_我的论文_2/
│   └── ...
```

## 环境

- Python 3.10+
- `python -m pip install python-docx Pillow`

## 流水线六阶段

```
Templates/模版.docx ──→ [Phase 1] format_extractor ──→ format.json
    或 .md (# 格式说明)      或 md_parser              格式提取.md

format.json ─────────→ [Phase 2] template_profiler ─→ template_profile.json

Inputs/内容.docx/.md ──→ [Phase 3] content_parser ──→ content.json
                       或 md_parser                 内容提取.md

format.json ──┬──→ [Phase 4] script_generator ──→ build_generated.py
content.json ─┘

build_generated.py ──→ [Phase 5] python 运行 ──→ 最终论文.docx
最终论文.docx ──→ [Phase 6] qa_checker ──→ qa_report.json / qa_report.md
                    可选 qa_visual ───────→ visual_report.json / visual_report.md
```

每个阶段内建双验证：提取器独立运行两次，比对段落/表格/run 数，不一致时第三轮仲裁。

## 微调排版

```bash
# 用户侧：让 AI 修改当前输出目录的生成脚本，然后重跑生成脚本
python run_pipeline.py --mode user
python Outputs/<目录>/build_generated.py
```

`Outputs/<目录>/build_generated.py` 是用户侧微调入口，适合处理当前这篇文档的一次性排版要求。下一次重新运行完整流水线会覆盖它。

开发者维护通用规则时，修改核心引擎后重跑流水线：

```bash
python Paper_Project/Program/pipeline/regression_suite.py
python run_pipeline.py --mode developer --template 模版.docx --content 论文.docx
```

| 意图 | 用户侧微调 | 开发者侧沉淀 |
|------|------------|--------------|
| 当前文档的正文/标题/参考文献样式 | `build_generated.py` | `script_generator.py` |
| 当前文档的图片、表格、目录细节 | `build_generated.py` | `script_generator.py` / `content_parser.py` |
| 内容识别、章节、图片抽取规则 | 不建议用户处理 | `content_parser.py` |
| 模板格式、封面结构、页眉页脚抽取 | 不建议用户处理 | `format_extractor.py` |
| Markdown 输入规则 | 不建议用户处理 | `md_parser.py` |
| 输出目录、文件选择、双验证流程 | 不建议用户处理 | `run_pipeline.py` |
| 生成后 QA 检测规则 | 不建议用户处理 | `qa_checker.py` |
| PDF 渲染 QA 和视觉抽样 | 不建议用户处理 | `qa_visual.py` |
| 模板能力画像 | 不建议用户处理 | `template_profiler.py` |

## 项目结构

```
├── run_pipeline.py              ← 一键入口
├── CLAUDE.md                    ← AI 工作流（Claude Code 自动加载）
├── .claude/settings.local.json  ← 本地 Claude 权限配置（忽略提交）
├── .gitignore
├── Templates/模版放这里.txt
├── Inputs/文本资料放这里.txt
├── Outputs/                     ← 每次运行生成独立子目录（忽略提交）
└── Paper_Project/
    ├── 基础操作.md               ← AI 工具箱（所有 OOXML 代码片段）
    └── Program/
        ├── pipeline/
        │   ├── format_extractor.py   ← Phase 1: 模版 → 格式 JSON
        │   ├── content_parser.py     ← Phase 2: 内容 → 结构化 JSON
        │   ├── md_parser.py          ← MD 解析（格式 + 内容）
        │   ├── script_generator.py   ← Phase 3: JSON → 生成脚本
        │   ├── latex_omath.py        ← LaTeX/文本公式→OOXML 公式转换器
        │   ├── qa_checker.py         ← 输出结构 QA 检测与修复目标报告
        │   ├── qa_visual.py          ← 可选 PDF 渲染 QA
        │   ├── public_template_suite.py ← 公共模板兼容性测试套件
        │   ├── privacy.py            ← 报告脱敏辅助
        │   ├── comment_utils.py      ← Word 批注系统
        │   └── regression_suite.py   ← 开发者合成回归测试
        ├── build_acta_manuscript.py  ← 参考：Acta Materialia 期刊格式
        ├── build_comprehensive_doc.py ← 参考：全功能演示
        └── master.py                 ← 参考：编排器骨架
```

## 功能

- **Markdown 内容支持**：`.md` 文件可直接作为内容输入，内置 `# 格式说明` 或 YAML frontmatter 描述排版参数
- **Markdown 相对图片**：`![图](media/a.png)` 按 `.md` 文件目录解析并复制到本次输出目录，适合把图片和正文放在同一资料文件夹
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
- **文本公式重建**：内容 docx 中的普通文本工程公式会被识别、转换为 LaTeX，再以 `m:oMathPara` 插入，WPS/Word 中可继续编辑公式。
- **Markdown 公式保真**：摘要和正文里的 `$...$` / `$$...$$` 会转为原生 OOXML Math；生成前的空段落清理会保留纯公式段落，避免公式被误删。
- **WPS 原生公式兼容**：`latex_omath.py` 为每个 `m:r` 保留 `m:rPr`，并合并 `\text{}` / `\mathrm{}` 的连续文本 run，避免公式显示成 `M|b|p|s` 一类分隔伪影。
- **公式编号**：`\tag{1.1}`, `\begin{equation}`, `\begin{align}` 自动编号
- **双线体/手写体**：`\mathbb{R}`, `\mathcal{F}` 等数学字体
- **Word 批注**：`body("text", comment="导师: 请确认")` → 生成原生 Word 批注
- **目录 (TOC)**：默认生成静态目录；Word COM 可用时自动解析标题页码并写入目录
- **QA 检测报告 + 修复计划**：生成后自动检查关键输出、样式 profile、目录、图片、公式和 LaTeX 错误，并按用户/开发者模式标注修复目标，额外输出可复制给 AI 的 `qa_fix_prompt.txt`。
- **DOCX 图片边界检测**：正文段落图片和表格单元格图片进入正文图片流；页眉页脚图片明确报 `NON_BODY_IMAGE_UNSUPPORTED`，防止用户以为图片已迁移。
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

# 7. Workflow mode
#    user: edit only the generated build_generated.py for this run
#    developer: edit only core pipeline engines and rerun the full pipeline
python run_pipeline.py --mode user
python run_pipeline.py --mode developer --template template.docx --content paper.docx

# 8. Product-grade check: structure QA + PDF export + sample render
python run_pipeline.py --mode developer --qa-level visual --template template.docx --content paper.docx
```

Each run produces an independent directory `Outputs/{date}_{content_name}/`, never overwriting previous results. If the same content name is run again on the same day, the folder gets `_2`, `_3`, etc.

## Current Behavior

- Real paper content, templates, and generated artifacts under `Inputs/`, `Outputs/`, and `Templates/` are treated as local private data and ignored by Git. Do not commit paper content, templates, generated docx/PDF/PNG files, or QA renders.
- `build_generated.py` is the user-facing AI fine-tuning layer: regular users can ask AI to edit the generated script in the current output folder, then run it for the current document.
- `Paper_Project/Program/pipeline/` is the developer maintenance layer: reusable features, bug fixes, and rule upgrades belong in the core engine scripts, followed by a fresh pipeline run.
- `run_pipeline.py` supports `--mode user|developer|auto`: interactive runs ask for the identity, non-interactive runs default to user mode, and each output folder records `workflow_mode.json`.
- After generation, `qa_checker.py` writes `qa_report.json` and `qa_report.md`, plus `qa_repair_plan.md/json` and `qa_fix_prompt.txt`. QA does not edit code automatically, but it routes every issue to the user-side or developer-side next repair action; any `error` makes the pipeline fail so unverified output is not treated as successful.
- Each run writes `template_profile.json` / `template_profile.md`, a template capability profile covering cover, TOC, headings, captions, references, formulas, headers/footers, and risk flags without school-name or private-filename rules.
- `--qa-level visual` runs `qa_visual.py` after structural QA: it exports PDF through Word COM, uses Poppler tools to inspect page count/size/text, renders sample PNG pages, and writes `visual_report.json` / `visual_report.md`; missing required render tools fail visual QA and block the pipeline.
- Paragraph images and table-cell images from the content document are extracted into the current output directory's `figures/` folder, so `Inputs/` is not polluted and runs do not overwrite each other's images. Markdown image paths are resolved relative to the `.md` file, so colocated `media/` folders work; missing or remote Markdown images and DOCX image extraction failures are reported as QA errors instead of disappearing silently. Header/footer images in the content source are not migrated as body content; QA reports `NON_BODY_IMAGE_UNSUPPORTED` so users can move real content images into the body or remove decorative source images.
- Plain-text formulas in content documents are recognized by `content_parser.py`; Markdown inline/display math is extracted by `md_parser.py`; both paths are normalized by `script_generator.py` and rendered through `latex_omath.py` as native OOXML Math. Inline formulas stay in the paragraph, and display formulas use editable WPS/Word `m:oMathPara`, not screenshots or plain text. Single-dollar `$...$` spans are accepted only when the content is semantically math-like, so `$100$` is not misclassified, and `$$...$$` remains display math. Split source-DOCX sum, fraction-plus-sum, and max/min objective layouts are conservatively repaired. If raw `$...$` / `$$...$$` delimiters remain in the final DOCX, QA reports `LATEX_DELIMITER_TEXT`.
- For already-typeset source DOCX files, `content_parser.py` looks ahead to skip static source TOC blocks, including unpaged TOC entries and TOC blocks polluted by repeated formula/page fragments, while preserving real "Contents" sections that use explicit Heading styles even when followed by a page break. Unfilled placeholders such as `[报名序号]`, duplicate formula labels like `(1)(1)`, formula-fragment paragraphs such as `E` / `rgreen` / `RE`, and likely low-resolution text/formula image shards are surfaced by `qa_checker.py` as `CONTENT_TOC_POLLUTION`, `UNFILLED_PLACEHOLDER_TEXT`, `FORMULA_NUMBER_CONFLICT`, `FORMULA_TEXT_FRAGMENTED`, or `LOW_RES_IMAGE_FRAGMENT` instead of being silently treated as deliverable output.
- TOC output is static by default; when Windows Word COM is available, the pipeline reads heading page numbers from Word and writes them into the visible TOC.
- Body images fit the full text width. Cover images use the template-extracted width and height, preventing logos and seals from being clipped by paragraph line height.

## Architecture

Core engines + one generated script + two knowledge bases:

```
+---------------------------+       +---------------------------+
| format_extractor.py       |       |                           |
| content_parser.py         |       |   build_generated.py      |
| script_generator.py       |params |                           |
|                           |------>|   zero hardcoding         |---> final .docx
| (developer layer)         |       |                           |
+---------------------------+       |   user fine-tuning layer  |
                                    |   + 基础操作.md            |
                                    +---------------------------+

 Knowledge Base
 +-- CLAUDE.md       <- AI workflow instructions
 +-- 基础操作.md      <- AI toolbox (all OOXML code snippets)
```

- **Core engines**: the developer maintenance surface. Reusable fixes belong in `pipeline/` scripts and stay parameterized via template/content JSON
- **md_parser.py**: Markdown parser — extracts both format specs and content from `.md` files
- **template_profiler.py**: template profiler — derives reusable capabilities and risk flags from `format.json`
- **latex_omath.py**: standalone LaTeX→OOXML formula converter; it supports handwritten LaTeX and LaTeX normalized from plain-text formulas extracted by `content_parser.py`
- **formula_semantics.py**: deterministic formula semantics layer for quantities/units, inline math, display math, and fragmented formula problems
- **comment_utils.py**: Word comment system — `comment="advisor: ..."` adds native Word comments
- **qa_visual.py**: optional PDF/render QA for page count, paper size, TOC text, and sample PNG rendering
- **public_template_suite.py**: public-template compatibility suite; downloads public templates into ignored local folders and runs synthetic structure/visual QA without committing templates or outputs
- **privacy.py**: report-sanitization helper to avoid leaking absolute machine paths
- **regression_suite.py**: developer synthetic regression suite covering formula semantics, inline/display formulas, mixed paragraphs, MD tables/code, manifest counts, table-cell images, non-body image warnings, raw LaTeX delimiter leftovers, and QA failure paths
- **build_generated.py**: generated user fine-tuning script. Users can ask AI to edit it for the current document; developers move reusable fixes back into the core engines
- **基础操作.md**: continuously maintained code snippet library — the more it's tested, the more AI can do
- **Missing template features** (e.g. cross-references) → engine won't generate → user requests → AI checks 基础操作.md → edits the current `build_generated.py`; developers can later productize the pattern in the core engine

## Output Structure

Each run produces an independent directory:

```
Outputs/
├── 2026-05-06_my_paper/
│   ├── 格式提取.md          ← verify template formats
│   ├── 内容提取.md          ← verify content
│   ├── format.json
│   ├── content.json
│   ├── template_profile.md  ← template capability profile
│   ├── workflow_mode.json  ← user/developer mode for this run
│   ├── build_manifest.json ← rendered body image/table/formula counts
│   ├── qa_report.md        ← generated QA report
│   ├── visual_report.md     ← PDF/render QA (--qa-level visual)
│   ├── visual_qa/           ← local PDF and sample PNGs (ignored)
│   ├── figures/             ← images extracted for this run
│   ├── assets/              ← template assets for this run
│   ├── build_generated.py   ← generated script, AI-tunable for this document
│   └── final_paper.docx
├── 2026-05-06_my_paper_2/
│   └── ...
```

## Environment

- Python 3.10+
- `python -m pip install python-docx Pillow`

## Pipeline Stages

```
Templates/template.docx → [Phase 1] format_extractor → format.json
                                                         format_report.md

format.json ─────────→ [Phase 2] template_profiler ─→ template_profile.json

Inputs/content.docx ──→ [Phase 3] content_parser ──→ content.json
                    (images → Outputs/<run>/figures/) content_report.md

format.json ──┬──→ [Phase 4] script_generator ──→ build_generated.py
content.json ─┘

build_generated.py ──→ [Phase 5] python execution ──→ final_paper.docx
final_paper.docx ──→ [Phase 6] qa_checker ──→ qa_report.json / qa_report.md
                    optional qa_visual ─────→ visual_report.json / visual_report.md
```

Each stage has built-in dual verification: the extractor runs independently twice, comparing paragraph/table/run counts; a third arbitration run resolves mismatches.

## Fine-Tuning

```bash
# User-level: ask AI to edit the generated script, then re-run it
python run_pipeline.py --mode user
python Outputs/<directory>/build_generated.py
```

`Outputs/<directory>/build_generated.py` is the user-facing fine-tuning entry point for the current document. A full pipeline run regenerates it.

For reusable behavior, developers update the core engine and run the whole pipeline again:

```bash
python Paper_Project/Program/pipeline/regression_suite.py
python run_pipeline.py --mode developer --template template.docx --content paper.docx
```

| Intent | User-level tuning | Developer productization |
|--------|-------------------|--------------------------|
| Current document body, heading, reference styles | `build_generated.py` | `script_generator.py` |
| Current document images, tables, TOC details | `build_generated.py` | `script_generator.py` / `content_parser.py` |
| Content recognition, sections, image extraction rules | not expected from users | `content_parser.py` |
| Template formatting, cover, headers/footers extraction | not expected from users | `format_extractor.py` |
| Markdown input rules | not expected from users | `md_parser.py` |
| Output folders, file selection, verification flow | not expected from users | `run_pipeline.py` |
| Generated-output QA rules | not expected from users | `qa_checker.py` |
| PDF/render QA | not expected from users | `qa_visual.py` |
| Template capability profiling | not expected from users | `template_profiler.py` |

## Project Structure

```
├── run_pipeline.py              ← one-click entry point
├── CLAUDE.md                    ← AI workflow (auto-loaded by Claude Code)
├── .claude/settings.local.json  ← local Claude permissions (ignored)
├── .gitignore
├── Templates/                   ← place template docx here
├── Inputs/                      ← place content docx here
├── Outputs/                     ← independent sub-directory per run (ignored)
└── Paper_Project/
    ├── 基础操作.md               ← AI toolbox (all OOXML code snippets)
    └── Program/
        ├── pipeline/
        │   ├── format_extractor.py   ← Phase 1: template → format JSON
        │   ├── content_parser.py     ← Phase 2: content → structured JSON
        │   ├── md_parser.py          ← MD parser (format + content)
        │   ├── template_profiler.py  ← template capability profile
        │   ├── script_generator.py   ← Phase 3: JSON → build script
        │   ├── latex_omath.py        ← LaTeX/plain-text formula → OOXML converter
        │   ├── qa_checker.py         ← output QA report and fix-target routing
        │   ├── qa_visual.py          ← optional PDF/render QA
        │   ├── public_template_suite.py ← public-template compatibility suite
        │   ├── privacy.py            ← report path sanitization
        │   ├── comment_utils.py      ← Word comment system
        │   └── regression_suite.py   ← developer synthetic regression tests
        ├── build_acta_manuscript.py  ← reference: Acta Materialia journal format
        ├── build_comprehensive_doc.py ← reference: full feature demo
        └── master.py                 ← reference: orchestrator skeleton
```

## Features

- **Markdown content support**: `.md` files as content input, with built-in `# 格式说明` section or YAML frontmatter for format specification
- **Markdown relative images**: `![figure](media/a.png)` resolves from the `.md` file directory and is copied into the current output folder
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
- **Plain-text formula reconstruction**: engineering formulas in source docx paragraphs are detected, normalized to LaTeX, and inserted as editable `m:oMathPara` equations.
- **Markdown equation preservation**: `$...$` and `$$...$$` in abstracts and body sections render as native OOXML Math; cleanup preserves math-only paragraphs instead of deleting them as empty text.
- **WPS-native equation compatibility**: `latex_omath.py` keeps `m:rPr` on every math run and merges contiguous `\text{}` / `\mathrm{}` text runs to avoid delimiter artifacts such as `M|b|p|s`.
- **Formula numbering**: `\tag{1.1}`, `\begin{equation}`, `\begin{align}` auto-numbering
- **Math fonts**: `\mathbb{R}`, `\mathcal{F}` etc.
- **Word comments**: `body("text", comment="advisor: please confirm")` → native Word comments
- **TOC**: static visible TOC by default; Word COM can resolve and write heading page numbers automatically
- **QA reports + repair plans**: after generation, checks key artifacts, style profiles, TOC, images, formulas, and LaTeX errors, then routes fixes by user/developer mode and writes `qa_fix_prompt.txt` for follow-up repair.
- **DOCX image boundary detection**: paragraph images and table-cell images enter the body image stream; header/footer images report `NON_BODY_IMAGE_UNSUPPORTED` so users know they were not migrated.
- Traditional formula tools: `formula_build_matrix()` parameterized construction
- Dual verification extraction (two independent runs cross-compared, third arbitration on mismatch)

---

## 许可 / License

Copyright © 2025 Youwei Zhang

本软件仅供个人学习和研究使用。**未经作者明确书面授权，禁止将本软件用于任何商业目的**，包括但不限于：将本软件或衍生作品作为商业产品、付费服务、SaaS 平台的一部分进行销售、出租、许可或分发。如需商业使用授权，请联系作者。

This software is provided for personal learning and research purposes only. **Commercial use is prohibited without explicit written authorization from the author.**
