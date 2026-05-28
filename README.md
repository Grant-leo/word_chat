# Word 论文自动排版流水线

把「模板 DOCX」和「内容 DOCX / Markdown」分开处理：模板提供格式，内容提供章节、段落、图片、表格、公式和参考文献，流水线自动生成排版后的 `最终论文.docx`。

## 适合做什么

- 用学校、期刊或自定义模板快速排版论文。
- 从无格式/弱格式内容中提取正文、图表、公式和参考文献。
- 生成可检查、可微调、可追踪的 Word 输出。
- 开发者可以把通用修复沉淀到核心引擎，而不是反复改一次性脚本。

## 快速开始

默认你已经在 VSCode 中安装了 Agent 插件。普通用户不需要手动记命令，只需要把模板和内容放好，然后把下面的提示词发给 Agent。

推荐安装配套预览插件：在 VSCode 扩展商店搜索 `DOCX Live Preview` 并安装，用来预览生成的 Word 文档。

先放文件：

- 模板 DOCX 放入 `Templates/`
- 内容 DOCX 或 Markdown 放入 `Inputs/`

然后复制这段给 Agent：

```text
请读取本项目的 AGENTS.md、README.md 和 Paper_Project/基础操作.md。
我已经把论文模板放在 Templates/，把论文内容放在 Inputs/。
请先检查 Python 依赖是否齐全：python-docx、Pillow、lxml；缺失时请帮我安装。
请帮我检查文件是否齐全，选择合适的模板和内容，以普通用户模式运行论文排版流水线。
运行完成后，请读取最新 Outputs 目录里的 qa_report.md、qa_repair_plan.md、template_profile.md、格式提取.md、内容提取.md，并告诉我：
1. 最终论文 docx 在哪里；
2. QA 是否通过；
3. 如果没有通过，下一步应该怎么修。
```

如果 QA 报错，继续把这段发给 Agent：

```text
请读取最新 Outputs 目录里的 qa_report.md 和 qa_repair_plan.md。
如果这是一次性排版问题，请只修改该输出目录里的 build_generated.py，然后重新生成最终论文并再次运行 QA。
如果这是可复用的引擎问题，请先说明原因，不要直接改核心脚本。
```

开发者做核心引擎验证时，可以把这段发给 Agent：

```text
请按开发者模式验证当前流水线。
先检查基础依赖 python-docx、Pillow、lxml。
先运行合成回归，再选择 Templates/ 和 Inputs/ 中合适的测试文件运行 strict QA。
如果要跑 visual QA，请确认 Word/WPS COM 和 Poppler 工具 pdfinfo、pdftotext、pdftoppm 可用。
只修改 Paper_Project/Program/pipeline/ 下的核心引擎脚本和必要文档。
不要提交 Inputs、Outputs、Templates、TestData、memory 或任何生成的 DOCX/PDF/PNG。
```

## 输出在哪里

每次运行都会生成独立目录：

```text
Outputs/日期_内容名/
├── 最终论文.docx
├── build_generated.py
├── qa_report.md / qa_report.json
├── qa_repair_plan.md / qa_repair_plan.json
├── template_profile.md / template_profile.json
├── 格式提取.md
├── 内容提取.md
├── build_manifest.json
├── figures/
└── assets/
```

最常看的文件：

- `最终论文.docx`：生成结果。
- `qa_report.md`：是否有图片、公式、表格、占位符、内容缺失等问题。
- `qa_repair_plan.md`：下一步该修哪里，适合直接交给 AI 继续处理。
- `build_generated.py`：本次文档的用户级微调脚本。
- `template_profile.md`：模板能力和风险画像。

## 两种工作模式

| 模式 | 适合谁 | 修改哪里 |
|---|---|---|
| `user` | 普通用户微调当前文档 | `Outputs/<本次输出>/build_generated.py` |
| `developer` | 修复可复用引擎能力 | `Paper_Project/Program/pipeline/` |

原则很简单：一次性排版问题改 `build_generated.py`；所有可复用能力、解析规则、QA 规则都沉淀到核心引擎。

## 流水线概览

```text
Templates/模板.docx ──→ format_extractor ──→ format.json
                         template_profiler ─→ template_profile.json

Inputs/内容.docx/.md ─→ content_parser/md_parser ─→ content.json

format.json + content.json ─→ script_generator ─→ build_generated.py
build_generated.py ─────────→ 最终论文.docx
最终论文.docx ─────────────→ qa_checker / qa_conformance / qa_visual
```

核心入口：

- `run_pipeline.py`：一键运行入口。
- `format_extractor.py`：提取模板格式。
- `content_parser.py`：提取 DOCX 内容。
- `md_parser.py`：解析 Markdown 内容和格式说明。
- `script_generator.py`：生成构建脚本。
- `latex_omath.py`：LaTeX / 文本公式转原生 Word 公式。
- `qa_checker.py`：结构 QA。
- `qa_conformance.py`：严格 DOCX/XML 合规检查。
- `qa_visual.py`：可选 PDF/PNG 视觉检查。

更细的模块布局见 [pipeline README](Paper_Project/Program/pipeline/README.md)。

## 公式、图片和 QA

- Markdown 的 `$...$` / `$$...$$` 会转成 Word 原生 OOXML Math。
- DOCX 中的文本公式会尽量识别并重建为可编辑公式。
- 图片会复制到本次输出目录的 `figures/`，不会污染 `Inputs/`。
- 缺图、图片抽取失败、公式丢失、表格数量不匹配、占位符残留会进入 QA 报告。
- `--qa-level visual` 会尝试用 Word/WPS 导出 PDF，并做页数、纸张、文本和抽样 PNG 检查。

基础依赖：Python 3.10+、`python-docx`、`Pillow`、`lxml`。自动目录页码可选依赖 `pywin32`；视觉 QA 还需要 Word/WPS COM 和 Poppler 工具 `pdfinfo`、`pdftotext`、`pdftoppm`。

最终交付前仍建议用 WPS/Word 打开 `最终论文.docx` 做视觉核对。

## 仓库规则

这些是本地隐私数据，不提交：

- `Inputs/`
- `Outputs/`
- `Templates/`
- `TestData/`
- `memory/`
- 生成的 DOCX / PDF / PNG / QA 渲染图

提交时只提交核心脚本和公共文档。

## 当前验证基线

截至 2026-05-28：

- 合成回归：`113 passed, 0 failed`
- 端到端 strict QA：5 个复杂测试文本 × 3 个模板，`15/15 passed`
- 矩阵结果：`0` QA error，`0` QA warning，`0` conformance error，`0` conformance warning

## 进一步文档

- [AGENTS.md](AGENTS.md)：AI 助手工作流。
- [CLAUDE.md](CLAUDE.md)：Claude/Codex 操作说明。
- [Paper_Project/基础操作.md](Paper_Project/基础操作.md)：OOXML 和 python-docx 操作速查。
- [Paper_Project/Program/pipeline/README.md](Paper_Project/Program/pipeline/README.md)：核心引擎模块结构。

## English Quick Start

This project generates a formatted Word paper from a DOCX template plus DOCX/Markdown content. The recommended workflow is to use a VSCode Agent instead of typing commands manually.

Put the template in `Templates/`, put the content file in `Inputs/`, then send this to your Agent:

```text
Please read AGENTS.md, README.md, and Paper_Project/基础操作.md.
I have placed the Word template in Templates/ and the paper content in Inputs/.
Please check the files, run the paper formatting pipeline in user mode, then inspect the latest Outputs folder.
Report where the final DOCX is, whether QA passed, and what to fix next if QA failed.
```

Outputs are written to `Outputs/<date_content>/`. Use `qa_report.md` and `qa_repair_plan.md` to inspect issues.

## License

See the repository license file if present.
