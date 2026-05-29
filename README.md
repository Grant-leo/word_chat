# Word 论文自动排版流水线

把「模板 DOCX/PDF」和「内容 DOCX / Markdown」分开处理：模板提供格式，内容提供章节、段落、图片、表格、公式和参考文献，流水线自动生成排版后的 `最终论文.docx`。

## 适合做什么

- 用学校、期刊或自定义模板快速排版论文。
- 从无格式/弱格式内容中提取正文、图表、公式和参考文献。
- 生成可检查、可微调、可追踪的 Word 输出。
- 开发者可以把通用修复沉淀到核心引擎，而不是反复改一次性脚本。

## 快速开始

默认你已经在 VSCode 中安装了 Agent 插件。普通用户不需要打开终端，也不需要手动记命令；把模板和内容放好，然后把下面的提示词发给 Agent。

推荐安装配套预览插件：在 VSCode 扩展商店搜索 `DOCX Live Preview` 并安装，用来预览生成的 Word 文档。

先放文件：

- 模板 DOCX 或 PDF 放入 `Templates/`
- 内容 DOCX 或 Markdown 放入 `Inputs/`

然后复制这段给 Agent：

```text
请读取本项目的 AGENTS.md、README.md 和 Paper_Project/基础操作.md。
我已经把论文模板放在 Templates/，把论文内容放在 Inputs/。模板可能是 DOCX，也可能是 PDF。
请先检查 Python 依赖是否齐全：python-docx、Pillow、lxml；缺失时请帮我安装。
请使用项目的 Agent 自动入口完成排版：自动扫描 Templates/ 和 Inputs/，能唯一确定文件时直接运行；有多个候选时只问我选择哪一个。
请以普通用户模式运行，开启自动修复闭环，能自动修的 QA 问题请直接修复并重跑。若模板是 PDF，请读取 template_profile.md 和 qa_report.md 说明 PDF 提取置信度与风险。
运行完成后，请优先读取最新 Outputs 目录里的 agent_summary.md，再读取 repair_loop_report.md、qa_report.md、qa_repair_plan.md、template_profile.md、格式提取.md、内容提取.md，并告诉我：
1. 最终论文 docx 在哪里；
2. QA 是否通过；
3. 自动修复了什么，是否还有需要我补文件或人工检查的事项。
```

如果自动修复闭环停止后仍有问题，继续把这段发给 Agent：

```text
请读取最新 Outputs 目录里的 repair_loop_report.md、qa_report.md 和 qa_repair_plan.md。
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

## 自动修复闭环

普通用户推荐让 Agent 直接开启受控自动修复：

```text
请以 user 模式运行流水线并开启 --auto-repair。
如果 QA 报告存在可自动修复的问题，请只修改最新 Outputs/<run>/build_generated.py，
重建最终 DOCX，重新运行 QA，并循环到没有 QA error 或修复循环停止。
不要修改 Paper_Project/Program/pipeline/，除非我明确要求做可复用引擎修复。
停止后请读取 repair_loop_report.md、qa_report.md、qa_repair_plan.md，
告诉我最终 DOCX 位置、自动修复了什么、还剩什么问题、需要我在 Word/WPS 里人工检查什么。
```

Agent 内部优先使用项目的自动入口；普通用户不用自己输入 Python 命令。开发者需要手动复现时，可在高级场景下使用 `run_pipeline.py --agent-auto` 或显式传入模板/内容参数。

如果流程在预检、QA、依赖或自动修复阶段中断，Agent 必须明确告诉用户下一步该做什么；预检阶段会写入 `Outputs/_agent_preflight_latest/agent_preflight_report.md`，正式运行后优先看 `agent_summary.md`。

自动修复闭环会生成 `repair_loop_report.md/json`，最多运行有限轮次，只允许修改本次输出目录的 `build_generated.py`。如果连续修复没有减少 error，或遇到缺图、扫描 PDF、内容缺失等必须由用户补文件的问题，会停止并说明原因。即使自动 QA 已无 error，也不代表 100% 正确，最终仍建议用 Word/WPS 做视觉检查。

`strict` / `visual` 等级下，自动修复闭环会把结构 QA、DOCX/XML 合规 QA、PDF/PNG 视觉 QA 一起作为收敛条件；缺少必要检查工具时不会假装通过。报告中的重建命令使用相对路径，便于直接复制给 Agent 继续处理。

## 输出在哪里

每次运行都会生成独立目录：

```text
Outputs/日期_内容名/
├── 最终论文.docx
├── build_generated.py
├── agent_summary.md / agent_summary.json
├── qa_report.md / qa_report.json
├── qa_repair_plan.md / qa_repair_plan.json
├── repair_loop_report.md / repair_loop_report.json   <- --auto-repair 时生成
├── template_profile.md / template_profile.json
├── 格式提取.md
├── 内容提取.md
├── build_manifest.json
├── figures/
└── assets/
```

最常看的文件：

- `最终论文.docx`：生成结果。
- `agent_summary.md`：面向用户和 Agent 的最终摘要，先看它。
- `qa_report.md`：是否有图片、公式、表格、占位符、内容缺失等问题。
- `qa_repair_plan.md`：下一步该修哪里，适合直接交给 AI 继续处理。
- `build_generated.py`：本次文档的用户级微调脚本。
- `template_profile.md`：模板能力和风险画像。
- `内容提取.md`：按上下文展示正文、图片、表格和公式；不会把表格/图片误写成公式，也会避免重复列出同一张正文图片。

## 两种工作模式

| 模式 | 适合谁 | 修改哪里 |
|---|---|---|
| `user` | 普通用户微调当前文档 | `Outputs/<本次输出>/build_generated.py` |
| `developer` | 修复可复用引擎能力 | `Paper_Project/Program/pipeline/` |

原则很简单：一次性排版问题改 `build_generated.py`；所有可复用能力、解析规则、QA 规则都沉淀到核心引擎。

## 流水线概览

```text
Templates/模板.docx/.pdf ─→ format_extractor ─→ format.json
                         template_profiler ─→ template_profile.json

Inputs/内容.docx/.md ─→ content_parser/md_parser ─→ content.json

format.json + content.json ─→ script_generator ─→ build_generated.py
build_generated.py ─────────→ 最终论文.docx
最终论文.docx ─────────────→ qa_checker / qa_conformance / qa_visual
```

核心入口：

- `run_pipeline.py`：一键运行入口。
- `run_pipeline.py --agent-auto`：Agent-first 自动入口；自动扫描、唯一选择、普通用户模式、自动修复、写出 `agent_summary.md/json`。
- `format_extractor.py`：提取 DOCX/PDF 模板格式；PDF 会区分文字说明模板、精排样张模板、扫描/不可解析模板。
- `content_parser.py`：提取 DOCX 内容。
- `md_parser.py`：解析 Markdown 内容和格式说明。
- `script_generator.py`：生成构建脚本。
- `latex_omath.py`：LaTeX / 文本公式转原生 Word 公式。
- `qa_checker.py`：结构 QA。
- `qa_conformance.py`：严格 DOCX/XML 合规检查。
- `qa_visual.py`：可选 PDF/PNG 视觉检查。
- `pipeline_runner/repair_loop.py`：`--auto-repair` 的受控自动修复闭环。

更细的模块布局见 [pipeline README](Paper_Project/Program/pipeline/README.md)。

## 公式、图片和 QA

- Markdown 的 `$...$` / `$$...$$` 会转成 Word 原生 OOXML Math。
- DOCX 中的文本公式会尽量识别并重建为可编辑公式。
- 图片会复制到本次输出目录的 `figures/`，不会污染 `Inputs/`。
- “图 1 展示了……”这类正文引用句会保持正文样式；“图 1 xxx 示意图”这类真实图注才按图注排版。
- 缺图、图片抽取失败、公式丢失、表格数量不匹配、占位符残留会进入 QA 报告。
- PDF 模板需要 Poppler 的 `pdfinfo`、`pdftotext`；扫描件或不可复制文字会进入 QA error，并提示用户提供 DOCX、文字说明 PDF 或 OCR 后重跑。
- `--qa-level visual` 会尝试用 Word/WPS 导出 PDF，并做页数、纸张、文本和抽样 PNG 检查。

基础依赖：Python 3.10+、`python-docx`、`Pillow`、`lxml`。PDF 模板解析需要 Poppler 的 `pdfinfo`、`pdftotext`；自动目录页码可选依赖 `pywin32`；视觉 QA 还需要 Word/WPS COM 和 Poppler 工具 `pdfinfo`、`pdftotext`、`pdftoppm`。

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

截至 2026-05-29：

- 合成回归：`150 passed, 0 failed`
- 自动修复闭环回归：可修复 QA error、连续无改善停止、needs_user_file 停止、strict/visual QA 依赖缺失、visual 参数保持、报告路径脱敏均已覆盖
- Agent-first 自动入口：`--agent-auto` 可自动扫描单候选模板/内容，默认普通用户自动修复，并写出 `agent_summary.md/json`
- PDF 模板端到端 strict QA：合成文字说明 PDF 模板 + DOCX 内容，`passed`
- PDF 极端压力测试：9 个场景覆盖大写扩展名、精排样张、横向页面、稀疏说明、扫描/损坏/空白/过短 PDF，`9/9` 符合预期
- 端到端 strict QA：5 个复杂测试文本 × 3 个模板，`15/15 passed`
- 矩阵结果：`0` QA error，`0` QA warning，`0` conformance error，`0` conformance warning
- fresh-folder 小白用户 visual 冒烟：DOCX 模板 + 无格式机器学习内容 + `--auto-repair --qa-level visual`，结构 QA / strict conformance / visual QA 均为 `0` error，自动修复闭环 `converged`

## 进一步文档

- [AGENTS.md](AGENTS.md)：AI 助手工作流。
- [CLAUDE.md](CLAUDE.md)：Claude/Codex 操作说明。
- [Paper_Project/基础操作.md](Paper_Project/基础操作.md)：OOXML 和 python-docx 操作速查。
- [Paper_Project/Program/pipeline/README.md](Paper_Project/Program/pipeline/README.md)：核心引擎模块结构。

## English Quick Start

This project generates a formatted Word paper from a DOCX/PDF template plus DOCX/Markdown content. The recommended workflow is to use a VSCode Agent instead of typing commands manually.

Put the template in `Templates/`, put the content file in `Inputs/`, then send this to your Agent:

```text
Please read AGENTS.md, README.md, and Paper_Project/基础操作.md.
I have placed the template in Templates/ and the paper content in Inputs/. The template may be DOCX or PDF.
Please use the Agent-first entry to scan the files, run the paper formatting pipeline in user mode with automatic repair, then inspect agent_summary.md in the latest Outputs folder first.
Report where the final DOCX is, whether automatic QA converged, what was repaired, and what I should still check manually.
```

Outputs are written to `Outputs/<date_content>/`. Read `agent_summary.md` first, then use `repair_loop_report.md`, `qa_report.md`, and `qa_repair_plan.md` to inspect details.

## License

See the repository license file if present.
