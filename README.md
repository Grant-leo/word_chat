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

如果流程在预检、交互选择、构建最终 DOCX、QA、依赖或自动修复阶段中断，Agent 必须明确告诉用户下一步该做什么；预检阶段会写入 `Outputs/_agent_preflight_latest/agent_preflight_report.md`，多候选时会列出“可以直接回复：使用 Templates/... 作为模板 / 使用 Inputs/... 作为内容”的候选句，并固定列出“文件应该放哪里”：模板放 `Templates/`（`.docx`/`.pdf`），内容放 `Inputs/`（`.docx`/`.md`）。正式运行后优先看 `agent_summary.md`。如果 `build_generated.py` 执行失败，流水线也会写出 `qa_report.md/json`、`qa_repair_plan.md/json` 和 `qa_fix_prompt.txt`，并把下一步指向本次 `build_generated.py`。结构/strict/visual QA 失败时，`agent_summary.md/json` 会直接列出前几个问题码和“小白用户下一步”；`conformance_report.md` 和 `visual_report.md` 顶部下一步也会点名 leading issue code，并针对占位符、Word 域、PDF 页数、不可读页面等常见阻断给出具体下一步。交互模式被取消或输入流中断时，下一步默认是改用 `python run_pipeline.py --agent-auto`。

如果高级用户用绝对路径传入了不在本项目 `Inputs/` / `Templates/` 下的文件，即使外部目录也恰好叫 `Inputs` 或 `Templates`，后续报告也不会生成可能失效的“只剩文件名”的重跑命令；下一步会提示先把文件放入本项目对应目录，再用文件名重跑。

自动修复闭环会生成 `repair_loop_report.md/json`，最多运行有限轮次，只允许修改本次输出目录的 `build_generated.py`。如果连续修复没有减少 error、重建失败、达到轮次上限，或遇到缺图、扫描 PDF、内容缺失等必须由用户补文件的问题，会停止并说明原因。报告顶部会写明 `next_action`、`resume_scope` 和 `resume_command`，`agent_summary.md/json` 也会汇总这条下一步。即使自动 QA 已无 error，也不代表 100% 正确，最终仍建议用 Word/WPS 做视觉检查。

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
├── conformance_report.md / conformance_report.json <- strict/visual 时生成
├── visual_report.md / visual_report.json           <- --qa-level visual 时生成
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
- `agent_summary.md`：面向用户和 Agent 的最终摘要，先看它；如果结构/strict/visual QA 失败，这里会直接汇总问题码和小白下一步。
- `qa_report.md`：是否有图片、公式、表格、占位符、内容缺失等问题；顶部“下一步”会点名优先处理的问题码和具体动作。
- `qa_repair_plan.md/json`：下一步该修哪里，顶部会写明 `next_action`、`resume_scope` 和 `resume_command`，适合直接交给 AI 继续处理。
- `conformance_report.md` / `visual_report.md`：strict/visual QA 的中文报告；顶部“下一步”会点名 leading issue code，缺依赖、渲染失败或常见阻断时给出具体下一步。
- `build_generated.py`：本次文档的用户级微调脚本。
- `template_profile.md`：模板能力和风险画像。
- `内容提取.md`：按上下文展示正文、图片、表格和公式；表格单元格图这类 `role="image"` 项也会显示为 `[图片]`，表格摘要会标出 `单元格图片 N`，不会变成 `[结构化内容]`，也不会把表格/图片误写成公式或重复列出同一张正文图片。

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
- `format_extractor.py`：提取 DOCX/PDF 模板格式；PDF 会区分文字说明模板、稀疏/不完整文字说明模板、精排样张模板、损坏/不可读取模板、扫描/不可解析模板；独立运行时默认写入 `Outputs/_format_extractor_cli/`，模板 assets 不写回 `Templates/`。
- `content_parser.py`：提取 DOCX 内容；独立运行时默认写入 `Outputs/_content_parser_cli/`，图片不写回 `Inputs/`。
- `md_parser.py`：解析 Markdown 内容和格式说明；会剥离开头的 YAML/自然语言格式块，识别 Windows 常见 UTF-8 BOM 开头的 YAML/front matter、H1 题名和 Setext 一级题名，避免格式规则进入正文；独立运行时默认写入 `Outputs/_md_parser_cli/`。
- `script_generator.py`：生成构建脚本。
- `latex_omath.py`：LaTeX / 文本公式转原生 Word 公式。
- `qa_checker.py`：结构 QA。
- `qa_conformance.py`：严格 DOCX/XML 合规检查。
- `qa_visual.py`：可选 PDF/PNG 视觉检查。
- `pipeline_runner/repair_loop.py`：`--auto-repair` 的受控自动修复闭环。

更细的模块布局见 [pipeline README](Paper_Project/Program/pipeline/README.md)。

## 公式、图片和 QA

- Markdown 的 `$...$` / `$$...$$` 会转成 Word 原生 OOXML Math。
- DOCX 中的文本公式会尽量识别并重建为可编辑公式；DOCX 表格单元格里的 block-level/inline/嵌套 inline 内容控件文字、`w:fldSimple` 可见字段结果、`w:customXml` / `w:smartTag` 透明容器里的绑定/下拉显示值、内容控件内 hyperlink 包裹的图片、inline OMML/LaTeX 公式和脚注引用会留在原表格单元格内，不会压成普通字符串、丢到表格外或重复追加到正文末尾。
- DOCX 正文级 `w:sdtContent` 内容控件包住段落和表格时，会按源顺序进入正文流；表格仍保持为表格，单元格文字不会散落成普通正文。
- DOCX 正文级 `w:customXml` / `w:smartTag` 包住的段落和表格会按源顺序展开，后续正常段落不会因为包装节点而错位或丢失。
- 图片会复制到本次输出目录的 `figures/`，不会污染 `Inputs/`；Markdown 本地图片路径支持 `%20` 空格编码、`<带空格路径>` 包裹写法、文件名括号、可选图片 title（如 `![图](path "title")`）、复制链接时常见的 `?query` / `#fragment` 后缀，以及引用式图片 `![图][id]` + `[id]: path`、引用定义下一行 title、shortcut 引用式图片 `![图]` + `[图]: path`、HTML 图片标签 `<img src="path" alt="图">`、懒加载 `data-src` / `data-original`、`srcset` 首候选和 PNG/JPG data URI。本地图片只把 `.png` / `.jpg` / `.jpeg` 作为稳定可生成格式；GIF、WebP、SVG、无扩展名图片，或扩展名和真实格式不一致的图片，会用 `CONTENT_IMAGE_UNREADABLE` 提示重新导出普通 PNG/JPG/JPEG 后更新 Markdown 链接并重跑。内联 data URI 也会核对 MIME 声明和真实图片格式，`data:image/png` 里实际是 JPEG 这类标错内容会作为 `CONTENT_IMAGE_UNREADABLE` 阻断，不会按错误扩展名写入 `figures/`。Markdown 表格单元格里的图片会记录到 `table_cell_items`、标记 `markdown_table_cell` 并渲染在生成的 Word 表格单元格内，缺失时同样进入 QA；损坏、打不开、不支持的本地图片或坏 data URI 会用 `CONTENT_IMAGE_UNREADABLE` 提示重新导出普通 PNG/JPG；远程 `http://` / `https://` 图片不会自动下载，QA 会用 `CONTENT_IMAGE_REMOTE_UNSUPPORTED` 明确提示先下载到本地并改成相对路径后重跑。
- DOCX 表格单元格里的正文图片会挂到对应 `table_cell_items` 并渲染在原 Word 表格单元格内；同一单元格段落里的“图片在文字前/后”顺序会按 OOXML run 顺序保留。四层嵌套表中的单元格图片也会保留在嵌套表内，QA 与 strict 预期计数会递归统计这些图片，避免“渲染成功但图片跑到表格外”或“QA 少数图片”的静默误排。DOCX 图片关系如果是损坏字节、特殊嵌入对象、扩展名和真实格式不一致或不支持的图片格式，会以 `IMAGE_EXTRACT_FAILED` 阻断并提示用户把源图重新导出/插入为普通 PNG/JPG 后重跑，不会把坏图片写进 `figures/`；Markdown 表格单元格图片也会保留在原表格单元格中渲染，缺失时同样进入 QA；源文件页眉/页脚图片属于 non-body 内容，会以 `NON_BODY_IMAGE_UNSUPPORTED` 提示用户移到正文或确认忽略。
- Markdown 开头的 YAML/front matter、第一行 H1 题名支持 UTF-8 BOM，题名也支持 Setext `Title` + `===`；中文题名写入 `title_cn`，英文/非 CJK 题名写入 `title_en`，避免有效标题触发 `TITLE_MISSING`。
- “图 1 展示了……”这类正文引用句会保持正文样式；“图 1 xxx 示意图”这类真实图注才按图注排版。
- 缺图、损坏/不支持的本地 Markdown 图片、远程 Markdown 图片 URL、DOCX 图片关系读取失败、公式丢失、表格数量不匹配、占位符残留会进入 QA 报告；本地坏图、GIF/WebP/SVG 等不支持格式、扩展名不匹配或 data URI MIME/真实格式不一致会点名 `CONTENT_IMAGE_UNREADABLE` 并提示重新导出 PNG/JPG，DOCX 坏图片关系会点名 `IMAGE_EXTRACT_FAILED` 并提示重新导出/插入普通 PNG/JPG，远程图片会点名 `CONTENT_IMAGE_REMOTE_UNSUPPORTED`，不会误导用户等待自动下载。
- PDF 模板需要 Poppler 的 `pdfinfo`、`pdftotext`；如果工具缺失，会在生成脚本前进入 `PDF_TEMPLATE_DEPENDENCY_MISSING` 并提示修复 Poppler 后重跑；如果 PDF 需要打开密码或禁止复制/提取文字，会进入 `PDF_TEMPLATE_PROTECTED`，提示解除密码/权限或重新导出无密码、可复制文字的 PDF；如果 Poppler 已运行但 PDF 损坏/不可读取，会进入 `PDF_TEMPLATE_READ_FAILED`，提示重新导出可正常打开、可复制文字的 PDF 或改用 DOCX；扫描件或不可复制文字会进入 `PDF_TEMPLATE_UNSUPPORTED`，提示用户提供 DOCX、文字说明 PDF 或 OCR 后重跑；文字说明 PDF 如果缺少标题、图表题注、参考文献等关键规则，会以 `PDF_TEMPLATE_INSTRUCTION_INCOMPLETE` warning 告诉用户补哪些规则或做重点人工核对；视觉样张 PDF 会以 `PDF_TEMPLATE_VISUAL_APPROXIMATION` warning 提醒用户用 Word/WPS 核对估算版式；横向 PDF 模板会以 `PDF_TEMPLATE_LANDSCAPE_PAGE` warning 提醒用户核对最终 DOCX 的页面方向、页边距和正文/表格压缩情况。
- `--qa-level visual` 会尝试用 Word/WPS 导出 PDF，并做页数、纸张、文本和抽样 PNG 检查；抽样页会优先覆盖封面、目录/正文起点，以及能识别到的图片、表格、公式风险页。

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
- Codex 本地 skill 属于用户运行环境，不放进仓库，也不恢复 `docs/skills/` 作为项目源码。

提交时只提交核心脚本和公共文档。

## 当前验证基线

截至 2026-06-23：

- 合成回归：`333 passed, 0 failed`
- DOCX 表格/嵌套表注释边界：四层嵌套表单元格会保持同段文字、图片、LaTeX、OMML 和脚注的源顺序；表格单元格中“图片后只有脚注锚点、没有可见文字”的情况也会在图片后原位渲染为 Word 原生脚注引用；表格单元格里的 block-level、inline 和嵌套 inline 内容控件文字会原位进入该单元格，内容控件内 `w:fldSimple` 的可见字段结果、`w:customXml` / `w:smartTag` 透明容器里的显示值也会按源顺序保留，内容控件内 hyperlink 包住图片、LaTeX、OMML 和脚注时也会保留原顺序，并在正文级内容控件兜底恢复时去重，避免重复正文；表格外正文级内容控件即使与表格单元格文本部分重叠或完全相同，也不会被误判为表格重复项；正文级 `w:sdtContent` 同时包住段落和表格时也会原位展开，元数据只统计正文段落，表格单元格文本不会被当成散落正文；正文级内容控件段落里的 inline `w:sdt` / `w:fldSimple` / hyperlink / `w:customXml` / `w:smartTag` 会递归保留图片、OMML/LaTeX 公式和脚注/尾注锚点顺序；带 `w:ins` / `w:moveTo` 的修订插入内容会按 Word 最终视图进入正文、表格、修订包裹的整行表格/单元格、标题路由和文本框恢复通道，`w:del` / `w:moveFrom` 删除内容与批注正文不会混入最终论文。
- DOCX 正文透明容器边界：正文级 `w:customXml` / `w:smartTag` 包住段落和表格时，内容会原位进入正文流，后续普通段落不会被包装节点造成的索引差异替换或丢失。
- 自动修复闭环回归：可修复 QA error、连续无改善停止、重建失败停止、needs_user_file 停止、strict/visual QA 依赖缺失、visual 参数保持、报告路径脱敏、停止后 `agent_summary` 汇总下一步均已覆盖
- Agent-first 自动入口：`--agent-auto` 可自动扫描单候选模板/内容；多候选时预检报告会把每个候选转成可直接回复给 Agent 的句子，并在 Markdown/JSON 中列出 `Templates/` 与 `Inputs/` 的放置位置和支持格式；默认普通用户自动修复，并写出 `agent_summary.md/json`
- 小白中断体验：交互取消、EOF、预检失败、生成脚本构建失败、QA/依赖失败都会给出下一步，`agent_summary.md/json` 会聚合结构/strict/visual QA 的问题码和具体修复动作，构建失败也会生成 `qa_report.md/json`、`qa_repair_plan.md/json` 和 `qa_fix_prompt.txt`；`qa_report.md/json` 顶部会点名首个结构 QA 问题码和动作；strict/visual 报告顶部下一步也会点名 leading issue code，并针对占位符、Word 域、PDF 页数无效、页面图片不可读等问题给出更具体的下一步；外部绝对路径输入不会生成失效的 basename 重跑命令，即使外部路径中也有同名 `Inputs` / `Templates` 目录，而会提示放入本项目 `Inputs/` / `Templates/` 后按文件名重跑；Markdown 图片路径已覆盖 `%20` 空格编码、`<带空格路径>` 本地写法、文件名括号、可选图片 title、本地图片 `?query` / `#fragment` 后缀、引用式图片 `![图][id]` + `[id]: path`、引用定义下一行 title、shortcut 引用式图片 `![图]` + `[图]: path`、HTML `<img src>`、HTML 懒加载 `data-src`、`srcset` 首候选、PNG/JPG data URI 图片和 Markdown 表格单元格图片，未定义图片引用会以 `CONTENT_IMAGE_MISSING` 阻断，损坏图片、GIF/WebP/SVG 等不支持本地格式、扩展名不匹配、坏 data URI 或 data URI MIME/真实格式不一致会以 `CONTENT_IMAGE_UNREADABLE` 提示重新导出 PNG/JPG，远程图片 URL 会以 `CONTENT_IMAGE_REMOTE_UNSUPPORTED` 提示下载到本地并改相对路径，UTF-8 BOM 开头的 YAML/front matter、Markdown H1、Setext 一级英文题名，以及“格式块 + 公式 + 缺图”的组合边界已覆盖；DOCX 表格单元格图和四层嵌套表内图片原位渲染、同段落图文 run 顺序保留、DOCX 表格单元格 inline OMML/LaTeX 行内公式原生渲染、DOCX 表格单元格图片 + LaTeX + OMML 混排顺序保真、DOCX 表格单元格同段落图片 + 公式 + 脚注顺序保真、DOCX 表格单元格脚注引用原位渲染、QA/strict 图片计数递归覆盖、DOCX 损坏/不支持图片关系的 `IMAGE_EXTRACT_FAILED` 阻断、Markdown 表格单元格图原位渲染、页眉/页脚 non-body 图、正文表格合并/列宽/行高/重复表头/单元格边距/垂直对齐/显式边框/四层嵌套表保真、普通超宽正文表自动横向页保护、嵌套表前后段落顺序保真、五层及以上嵌套风险审计、异常合并网格审计、横向宽表风险计数、嵌套宽表去重、`gridBefore` 纵向合并误报防护、`gridBefore`/`gridAfter` 行省略解析与生成端省略单元格保真、`内容提取.md` 图片摘要也有回归覆盖，strict QA 已覆盖默认正文段落出现在第一个显式标题前的场景，visual/WPS 样张对比会优先抽封面、目录/正文锚点和图表公式风险页
- DOCX 行省略宽表审计：源审计现在会把 `w:gridAfter` 省略的行尾网格列计入 `max_table_columns` / `wide_table_count`，因此“1 个可见单元格 + 多个尾部省略列”的宽表会触发 `COMPLEX_TABLE_UNSUPPORTED` 复核提示，不会因为只数可见单元格而漏警告。
- DOCX 修订包装表格审计：源审计现在按 Word 最终视图穿透 `w:ins` / `w:moveTo`、内容控件和透明容器中的表格行/单元格，宽表和异常表格风险不会因为行被包装而漏掉 `COMPLEX_TABLE_UNSUPPORTED`。
- 模板说明清理：DOCX 模板里的“格式说明”、封面字段说明、源目录样例和 TOC 页码样例不会再进入最终论文；本地脱敏真实样例已通过 developer visual 端到端验证，结构 QA、strict conformance、visual QA 均为 `0` error / `0` warning。
- 后置章节等价：结构 QA 现在把 `Acknowledgements` / `Acknowledgment` / `致谢`、`References` / `参考文献`、`Appendix` / `附录` 这类语义等价标题视为已覆盖，避免让用户为中英文模板标题差异处理误报的 `CONTENT_HEADING_MISSING`。
- QA JSON 契约：结构 `qa_report.json`、strict `conformance_report.json`、visual `visual_report.json` 都显式写入 `status`（`passed` / `passed_with_warnings` / `failed`）和 `result_label`（例如 `通过但有警告`）；依赖缺失、QA 崩溃、构建失败、提取验证失败等 fallback 报告也必须写同样字段。流水线会对结构、strict、visual 三类 QA 报告都运行契约检查，包括依赖缺失和 QA 崩溃 fallback 报告，并在终端报告缺失字段或与 `passed` / warning 状态不一致的字段，`agent_summary.json` 的每个报告条目也同步暴露该状态，避免界面或 Agent 只靠 `passed` 猜测。
- 输出边界：独立 `format_extractor.py` / `content_parser.py` / `md_parser.py` 默认写入 `Outputs/_...`，不污染 `Inputs/` 或 `Templates/`
- PDF 模板端到端 strict QA：合成文字说明 PDF 模板 + DOCX 内容，`passed`
- 稀疏 PDF 文字说明：缺少标题、图表题注、参考文献等关键规则时，结构 QA 以 `PDF_TEMPLATE_INSTRUCTION_INCOMPLETE` warning 暴露具体缺失项，`qa_report` / `qa_repair_plan` / `agent_summary` 都给出补规则或人工核对下一步
- 视觉样张 PDF 模板：结构 QA 以 `PDF_TEMPLATE_VISUAL_APPROXIMATION` warning 暴露估算版式风险，`template_profile` 标记 `pdf_template_visual_approximation`，并提示用 Word/WPS 核对页边距、标题层级、封面、页眉页脚、图表题注和参考文献版式
- 横向 PDF 模板：结构 QA 以 `PDF_TEMPLATE_LANDSCAPE_PAGE` warning 暴露页面方向风险，`template_profile` 标记 `pdf_template_landscape_page`，并提示用 Word/WPS 核对横向页面、页边距和正文/表格是否被压缩
- PDF 模板依赖缺失：在模板画像后、内容解析和 `build_generated.py` 创建前 fail closed，生成 `PDF_TEMPLATE_DEPENDENCY_MISSING` 的 QA/agent 交接，`resume_scope=environment`，提示修复 Poppler 后重跑
- PDF 模板读取失败：损坏/不可读取 PDF 在模板画像后、内容解析和 `build_generated.py` 创建前 fail closed，生成 `PDF_TEMPLATE_READ_FAILED` 的 QA/agent 交接，提示重新导出可正常打开、可复制文字的 PDF 或改用 DOCX
- PDF 模板受保护：需要密码或复制/提取权限受限的 PDF 在模板画像后、内容解析和 `build_generated.py` 创建前 fail closed，生成 `PDF_TEMPLATE_PROTECTED` 的 QA/agent 交接，提示解除密码/权限或导出无密码、可复制文字的 PDF 后重跑
- 扫描/不可复制 PDF 模板：在模板画像后、内容解析和 `build_generated.py` 创建前 fail closed，生成 `PDF_TEMPLATE_UNSUPPORTED` 的 QA/agent 交接
- PDF 极端压力测试：9 个场景覆盖大写扩展名、精排样张、横向页面、稀疏说明、扫描/损坏/空白/过短 PDF，`9/9` 符合预期
- 端到端 strict QA：5 个 DOCX 模板 × 5 个 DOCX 内容，`25/25 passed`
- 公共模板兼容套件：5 个公开模板 × 5 个合成场景，`25/25 passed`；`public_template_suite.py --visual` 默认只跑渲染 QA，不再自动比较 stale golden baseline。需要维护基线时显式使用 `--golden-dir`，需要刷新基线时使用 `--update-golden`。
- PDF 边界测试：可解析 PDF 模板继续通过 strict QA；缺 Poppler、受密码/权限保护、损坏/不可读取、不可用/扫描类模板按预期在生成前失败并给出下一步
- 高风险引擎流水线：纯 Markdown strict、缺图/远程图 Markdown、页眉图片边界、user auto-repair、DOCX/PDF visual smoke、密集媒体公式 strict 共 `7/7` 符合预期
- 矩阵结果：通过项均为 `0` QA error，`0` conformance error；visual smoke 为 `0` visual error
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

## Private Real-Data Hardening

Private corpora stay local and ignored. To inventory a private pool such as
`Templates/<private-corpus>/`, run:

```powershell
python Paper_Project/Program/pipeline/private_corpus_audit.py Templates/<private-corpus>
```

This writes only local reports under `Outputs/_private_realdata_audit/`:
`inventory.json`, `inventory.md`, and `review_queue.json`. The audit records
structural features and issue codes, not body text. Do not commit these reports
or the source documents.

QA-enabled runs now also write `comparison_assessment.json/md`, which records
whether the run is `FAILED_AUTOMATIC`, `PASSED_WITH_REVIEW`,
`PASSED_MACHINE`, and so on. Warning-only runs still require review; they are
not treated as perfect delivery.

Unsupported explicit inputs such as `.doc`, `.wps`, `.rar`, `.lnk`, and
`.xlsx` stop at preflight with a concrete next step. Save legacy Word files as
DOCX manually before adding them to a test matrix. Golden visual baselines are
not created during compare-only runs; use `--update-golden` only after manual
approval.

## License

See the repository license file if present.
