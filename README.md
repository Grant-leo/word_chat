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
- DOCX 表格单元格里的正文图片会挂到对应 `table_cell_items` 并渲染在原 Word 表格单元格内；同一单元格段落里的“图片在文字前/后”顺序会按 OOXML run 顺序保留。六层以内嵌套表中的单元格图片也会保留在嵌套表内，QA 与 strict 预期计数会递归统计这些图片，避免“渲染成功但图片跑到表格外”或“QA 少数图片”的静默误排。DOCX 图片关系如果是损坏字节、特殊嵌入对象、扩展名和真实格式不一致或不支持的图片格式，会以 `IMAGE_EXTRACT_FAILED` 阻断并提示用户把源图重新导出/插入为普通 PNG/JPG 后重跑，不会把坏图片写进 `figures/`；Markdown 表格单元格图片也会保留在原表格单元格中渲染，缺失时同样进入 QA；源文件页眉/页脚图片属于 non-body 内容，会以 `NON_BODY_IMAGE_UNSUPPORTED` 提示用户移到正文或确认忽略。
- Markdown 开头的 YAML/front matter、第一行 H1 题名支持 UTF-8 BOM，题名也支持 Setext `Title` + `===`；中文题名写入 `title_cn`，英文/非 CJK 题名写入 `title_en`，避免有效标题触发 `TITLE_MISSING`。
- “图 1 展示了……”这类正文引用句会保持正文样式；“图 1 xxx 示意图”这类真实图注才按图注排版。
- 缺图、损坏/不支持的本地 Markdown 图片、远程 Markdown 图片 URL、DOCX 图片关系读取失败、公式丢失、表格数量不匹配、占位符残留会进入 QA 报告；本地坏图、GIF/WebP/SVG 等不支持格式、扩展名不匹配或 data URI MIME/真实格式不一致会点名 `CONTENT_IMAGE_UNREADABLE` 并提示重新导出 PNG/JPG，DOCX 坏图片关系会点名 `IMAGE_EXTRACT_FAILED` 并提示重新导出/插入普通 PNG/JPG，远程图片会点名 `CONTENT_IMAGE_REMOTE_UNSUPPORTED`，不会误导用户等待自动下载。
- PDF 模板需要 Poppler 的 `pdfinfo`、`pdftotext`；如果工具缺失，会在生成脚本前进入 `PDF_TEMPLATE_DEPENDENCY_MISSING` 并提示修复 Poppler 后重跑；如果 PDF 需要打开密码或禁止复制/提取文字，会进入 `PDF_TEMPLATE_PROTECTED`，提示解除密码/权限或重新导出无密码、可复制文字的 PDF；如果 Poppler 已运行但 PDF 损坏/不可读取，会进入 `PDF_TEMPLATE_READ_FAILED`，提示重新导出可正常打开、可复制文字的 PDF 或改用 DOCX；扫描件或不可复制文字会进入 `PDF_TEMPLATE_UNSUPPORTED`，提示用户提供 DOCX、文字说明 PDF 或 OCR 后重跑；文字说明 PDF 如果缺少标题、图表题注、参考文献等关键规则，会以 `PDF_TEMPLATE_INSTRUCTION_INCOMPLETE` warning 告诉用户补哪些规则或做重点人工核对；视觉样张 PDF 会以 `PDF_TEMPLATE_VISUAL_APPROXIMATION` warning 提醒用户用 Word/WPS 核对估算版式；横向 PDF 模板会以 `PDF_TEMPLATE_LANDSCAPE_PAGE` warning 提醒用户核对最终 DOCX 的页面方向、页边距和正文/表格压缩情况。
- `--qa-level visual` 会尝试用 Word/WPS 导出 PDF，并做页数、纸张、文本和抽样 PNG 检查；抽样页会优先覆盖封面、目录/正文起点，以及能识别到的图片、表格、公式风险页；复杂长文的有界样张预算为最多 8 页，长表续页会在这些主风险页之后用剩余额度补充，并参与 Word/WPS 同页样张对比。

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

截至 2026-07-06：

- 合成回归：`448 passed, 0 failed`
- visual QA 多 section 宽表证据链：有界样张仍保持最多 8 页，但会优先保留后续 section 中宽表的延续页，同时不挤掉图/公式主风险页，便于在 `visual_qa/samples/` 和 `visual_qa/wps/samples/` 里核对复杂长文的跨页表头、续页和 Word/WPS 同页差异。
- DOCX 横向宽表短说明：相邻横向宽表之间的 `rich_text` 桥接段即使没有顶层 `text`、只在 `runs` 中保存普通文字和 `note_ref` 脚注/尾注锚点，也会按可见 run 文字识别为短说明并保留在同一个 landscape section；脚注/尾注继续以 Word 原生 note reference 渲染，不会要求用户手动处理。
- DOCX 富文本 run 内嵌块级内容：`rich_text.runs[].items` 里的代码、图片、题注和小表会按源顺序插入在前后文字 run 之间；不会被提前塞进当前段落，也不会被整体推迟到整段文字之后。
- DOCX 富文本 run 单元格来源块级内容：`rich_text.runs[].table_cell_items` 里的代码、图片、题注和小表也会按源顺序插入在前后文字 run 之间；不会因为被挂在 table-cell 兼容结构下而静默丢失。
- DOCX 表格富文本图片 run：`table_cell_items` 里的 `role="rich_text"` 如果直接携带 `runs[].type="image"`，图片会在生成的 Word 表格单元格内按表格图片尺寸渲染，并与文字、行内公式、脚注/尾注锚点保持同一单元格源顺序；`gridAfter` 省略区遇到这种富内容时会保留完整行，不会静默删掉图片。
- DOCX 表格富文本嵌套脚注：`table_cell_items` 里的 `role="rich_text"` 如果在 `items[]` 中携带 `note_ref`，脚注/尾注会在生成的 Word 表格单元格内原生渲染；即使该富文本位于 `gridAfter` 省略区，省略保护也不会只保留行却丢掉注释锚点。
- DOCX 表格/嵌套表注释边界：六层以内嵌套表单元格会保持同段文字、图片、LaTeX、OMML 和脚注的源顺序；表格单元格中“图片后只有脚注锚点、没有可见文字”的情况也会在图片后原位渲染为 Word 原生脚注引用；表格单元格里的 block-level、inline 和嵌套 inline 内容控件文字会原位进入该单元格，内容控件内 `w:fldSimple` 的可见字段结果、`w:customXml` / `w:smartTag` 透明容器里的显示值也会按源顺序保留，内容控件内 hyperlink 包住图片、LaTeX、OMML 和脚注时也会保留原顺序，并在正文级内容控件兜底恢复时去重，避免重复正文；表格外正文级内容控件即使与表格单元格文本部分重叠或完全相同，也不会被误判为表格重复项；正文级 `w:sdtContent` 同时包住段落和表格时也会原位展开，元数据只统计正文段落，表格单元格文本不会被当成散落正文；正文级内容控件段落里的 inline `w:sdt` / `w:fldSimple` / hyperlink / `w:customXml` / `w:smartTag` 会递归保留图片、OMML/LaTeX 公式和脚注/尾注锚点顺序；带 `w:ins` / `w:moveTo` 的修订插入内容会按 Word 最终视图进入正文、表格、修订包裹的整行表格/单元格、标题路由和文本框恢复通道，`w:del` / `w:moveFrom` 删除内容与批注正文不会混入最终论文。
- DOCX 正文透明容器边界：正文级 `w:customXml` / `w:smartTag` 包住段落和表格时，内容会原位进入正文流，后续普通段落不会被包装节点造成的索引差异替换或丢失。
- 自动修复闭环回归：可修复 QA error、连续无改善停止、重建失败停止、needs_user_file 停止、strict/visual QA 依赖缺失、visual 参数保持、报告路径脱敏、停止后 `agent_summary` 汇总下一步均已覆盖
- Agent-first 自动入口：`--agent-auto` 可自动扫描单候选模板/内容；多候选时预检报告会把每个候选转成可直接回复给 Agent 的句子，并在 Markdown/JSON 中列出 `Templates/` 与 `Inputs/` 的放置位置和支持格式；默认普通用户自动修复，并写出 `agent_summary.md/json`
- 小白中断体验：交互取消、EOF、预检失败、生成脚本构建失败、QA/依赖失败都会给出下一步，`agent_summary.md/json` 会聚合结构/strict/visual QA 的问题码和具体修复动作，构建失败也会生成 `qa_report.md/json`、`qa_repair_plan.md/json` 和 `qa_fix_prompt.txt`；`qa_report.md/json` 顶部会点名首个结构 QA 问题码和动作；strict/visual 报告顶部下一步也会点名 leading issue code，并针对占位符、Word 域、PDF 页数无效、页面图片不可读等问题给出更具体的下一步；外部绝对路径输入不会生成失效的 basename 重跑命令，即使外部路径中也有同名 `Inputs` / `Templates` 目录，而会提示放入本项目 `Inputs/` / `Templates/` 后按文件名重跑；Markdown 图片路径已覆盖 `%20` 空格编码、`<带空格路径>` 本地写法、文件名括号、可选图片 title、本地图片 `?query` / `#fragment` 后缀、引用式图片 `![图][id]` + `[id]: path`、引用定义下一行 title、shortcut 引用式图片 `![图]` + `[图]: path`、HTML `<img src>`、HTML 懒加载 `data-src`、`srcset` 首候选、PNG/JPG data URI 图片和 Markdown 表格单元格图片，未定义图片引用会以 `CONTENT_IMAGE_MISSING` 阻断，损坏图片、GIF/WebP/SVG 等不支持本地格式、扩展名不匹配、坏 data URI 或 data URI MIME/真实格式不一致会以 `CONTENT_IMAGE_UNREADABLE` 提示重新导出 PNG/JPG，远程图片 URL 会以 `CONTENT_IMAGE_REMOTE_UNSUPPORTED` 提示下载到本地并改相对路径，UTF-8 BOM 开头的 YAML/front matter、Markdown H1、Setext 一级英文题名，以及“格式块 + 公式 + 缺图”的组合边界已覆盖；生成脚本 QA 会阻断危险 unicode / `codecs.decode`（含 `setattr` 属性别名）/ `.encode(...).decode(...)` 等中文二次解码风险；DOCX 表格单元格图和六层以内嵌套表内图片原位渲染、同段落图文 run 顺序保留、DOCX 表格单元格 inline OMML/LaTeX 行内公式原生渲染、DOCX 表格单元格图片 + LaTeX + OMML 混排顺序保真、DOCX 表格单元格同段落图片 + 公式 + 脚注顺序保真、DOCX 表格单元格脚注引用原位渲染、QA/strict 图片计数递归覆盖、DOCX 损坏/不支持图片关系的 `IMAGE_EXTRACT_FAILED` 阻断、Markdown 表格单元格图原位渲染、页眉/页脚 non-body 图、正文表格合并/列宽/行高/重复表头/单元格边距/垂直对齐/显式边框/六层以内嵌套表保真、普通超宽正文表自动横向页保护、嵌套表前后段落顺序保真、七层及以上嵌套风险审计、异常合并网格审计、横向宽表风险计数、嵌套宽表去重、`gridBefore` 纵向合并误报防护、`gridBefore`/`gridAfter` 行省略解析与生成端省略单元格保真、`内容提取.md` 图片摘要也有回归覆盖，strict QA 已覆盖默认正文段落出现在第一个显式标题前的场景，visual/WPS 样张对比会优先抽封面、目录/正文锚点和图表公式风险页
- 中文编码防护：生成脚本 QA 还会阻断 `builtins.__import__("codecs")`、`importlib.import_module` / `builtins.__import__` 的赋值别名、`getattr(importlib, "import_module")` / `getattr(builtins, "__import__")`、`codecs.__dict__.get("decode")` / `vars(codecs).get("decode")` / 普通容器 `.get(...)` 取出的真实 `codecs.decode`、`operator.attrgetter("decode")(codecs)`、函数默认参数里的 `decoder=codecs.decode` / `module=codecs` / `factory=codecs.getdecoder` / `lookup=codecs.lookup`，以及 `from codecs import *` 后裸调用真实 `decode` / `getdecoder` / `lookup` 等路径，避免 UTF-8 中文 bytes 被 GBK 等错误编码二次解码；没有真实 `import builtins` 的本地安全 `builtins` 对象、星号导入后的同层安全覆盖、或 `import codecs` 后同层覆盖为安全自定义模块对象，不会误报；覆盖之前已经发生的真实 `codecs.decode` / `getdecoder` / `lookup` 调用仍会阻断。
- 中文编码防护补充：生成脚本 QA 现在也会识别 helper 中 `decoder(value, **kwargs)` / `module.decode(value, **kwargs)` 这类高阶转发；当外层把真实 `codecs.decode` / `codecs` 模块和 `encoding="gbk"` 等错误编码一起传入时，会以 `GENERATED_SCRIPT_UNSAFE_UNICODE_DECODE` 在执行前阻断，避免用户拿到看似通过但中文已经乱码的 DOCX。
- 中文编码防护补充 2：生成脚本 QA 现在也会识别批量回调形式的 `map(codecs.decode, payloads, encodings)` 和 `itertools.starmap(codecs.decode, rows)`；这些写法没有直接的 `codecs.decode(...)` 调用节点，但仍会在执行前按 `GENERATED_SCRIPT_UNSAFE_UNICODE_DECODE` 阻断，避免批量中文文本被错误编码静默二次解码。若用户脚本在调用前自定义了安全 `map`，QA 不会把它误判成内置 `map`。
- 中文编码防护补充 3：生成脚本 QA 现在也会识别自定义批处理包装器中的固定索引编码参数，例如 `decoder(values[0], encodings[0])`。当外层传入真实 `codecs.decode` 和 `encodings = ["gbk"]` 这类错误编码容器时，会在执行前阻断；同形状的自定义安全 decoder 不会误报。
- 中文编码防护补充 4：生成脚本 QA 现在会追踪自定义批处理包装器里的 `zip(values, encodings)` 循环变量，例如 `for value, encoding in zip(values, encodings): decoder(value, encoding)`；当外层传入真实 `codecs.decode` 和错误编码容器时会阻断，同形状的自定义安全 decoder 不会误报。
- 中文编码防护补充 5：生成脚本 QA 现在会追踪自定义批处理包装器里的 `enumerate(values)` 索引变量，例如 `for idx, value in enumerate(values): decoder(value, encodings[idx])`；当外层传入真实 `codecs.decode` 和 `encodings = ["gbk"]` 等错误编码容器时会在执行前阻断，避免中文 UTF-8 bytes 被静默二次解码成乱码。
- 中文编码防护补充 6：生成脚本 QA 现在会追踪自定义批处理包装器里的 `range(len(values))` 索引变量，例如 `for idx in range(len(values)): decoder(values[idx], encodings[idx])`；当外层传入真实 `codecs.decode` 和 `encodings = ["gbk"]` 等错误编码容器时会在执行前阻断，同形状的自定义安全 decoder 不会误报。
- 中文编码防护补充 7：生成脚本 QA 现在会追踪自定义批处理包装器里的列表推导式/生成器表达式局部变量，例如 `[decoder(value, encoding) for value, encoding in zip(values, encodings)]` 和 `next(decoder(value, encoding) for value, encoding in zip(values, encodings))`；当外层传入真实 `codecs.decode` 和 `encodings = ["gbk"]` 等错误编码容器时会在执行前阻断，同形状的自定义安全 decoder 不会误报。
- 中文编码防护补充 8：生成脚本 QA 现在会追踪自定义批处理包装器里已经配好对的 `rows` 元组拆包，例如 `for value, encoding in rows: decoder(value, encoding)` 和 `[decoder(value, encoding) for value, encoding in rows]`；当外层传入真实 `codecs.decode` 和 `rows = [(text.encode("utf-8"), "gbk")]` 等错误编码行时会在执行前阻断，同形状的自定义安全 decoder 不会误报。
- 中文编码防护补充 9：生成脚本 QA 现在会追踪高阶 helper 里先把 decoder 放进局部 list/dict 再调用的写法，例如 `callbacks = [decoder]; callbacks[0](value, encoding)` 和 `callbacks = {"decode": decoder}; callbacks["decode"](value, encoding)`；当外层传入真实 `codecs.decode`、中文 UTF-8 bytes 和 `"gbk"` 等错误编码时会在执行前阻断，同形状的自定义安全 decoder 或重新赋值后的安全 callback 不会误报。
- 中文编码防护补充 10：生成脚本 QA 现在会识别返回闭包捕获 decoder 参数的高阶写法，例如 `def make_decoder(decoder): ... return call` 后 `make_decoder(codecs.decode)(value, "gbk")`，也会追踪 `decode_text = make_decoder(codecs.decode)` 后再调用 `decode_text(value, "gbk")` 的赋值闭包，包括 `if`/循环等语句块内的赋值后调用；当闭包直接调用 `decoder(...)` 或调用外层局部 `callbacks=[decoder]` 容器里的 callback 时，会在执行前阻断真实 `codecs.decode` + UTF-8 中文 bytes + 错误编码，同形安全 decoder 或后续重赋值为安全函数不会误报。
- 中文编码防护补充 11：生成脚本 QA 现在会识别动态容器写入后的真实 `codecs.decode` 路由，例如 `routes=[]; routes.append(codecs.decode); routes[0](...)`、`routes={}; routes["decode"]=codecs.decode; routes["decode"](...)`，以及 `modules.append(codecs); modules[0].decode(...)`；这些写法以前没有直接调用节点或字面量容器节点，可能让中文 UTF-8 bytes 被 `"gbk"` 等错误编码静默二次解码，现在会在执行前以 `GENERATED_SCRIPT_UNSAFE_UNICODE_DECODE` 阻断。
- 中文编码防护补充 12：动态容器写入防护继续覆盖 `extend([codecs.decode])`、`insert(0, codecs.decode)`、`dict.update({"decode": codecs.decode})`、`dict.update(decode=codecs.decode)`、`setdefault("decode", codecs.decode)`，以及同形的 `codecs` 模块容器；已有安全 decoder 的 `setdefault` 不会被误报。
- 中文编码防护补充 13：生成脚本 QA 现在也会识别容器方法返回值被立刻调用的路由，例如 `routes.setdefault("decode", codecs.decode)(...)`、`routes.get("decode", codecs.decode)(...)`、`routes.pop("decode")(...)`，以及 `modules.get("m", codecs).decode(...)` 这类同形 `codecs` 模块返回值；已有安全 decoder 的 `setdefault/get/pop` 不会被误报。
- 中文编码防护补充 14：容器方法返回值先赋值或传入高阶 helper 后再调用也会被识别，例如 `decode_text = routes.get("decode", codecs.decode); decode_text(...)`、`text_codecs = modules.get("m", codecs); text_codecs.decode(...)`、`apply_decoder(routes.get("decode", codecs.decode), ...)`；同形安全 decoder 仍不误报。
- 中文编码防护补充 15：容器挂在对象/类属性上时也会进入同一条守卫链，例如 `SimpleNamespace(routes={})`、`class Holder: routes = {}` 后的 `holder.routes.get("decode", codecs.decode)` / `Holder.routes.get("decode", codecs.decode)`，以及 `box.decoder = routes.get("decode", codecs.decode)`、`Box.text_codecs = modules.get("m", codecs)` 这类方法返回值再挂到属性上的路径；同形安全 decoder 不误报。
- 中文编码防护补充 16：生成脚本 QA 现在会把 `codecs.decode.__call__(...)`、`getattr(codecs.decode, "__call__")(...)` 和 `operator.call(codecs.decode, ...)` 视为同一个真实 `codecs.decode` 路由；这类元编程写法以前没有普通 `codecs.decode(...)` 调用节点，可能把 UTF-8 中文 bytes 按 `"gbk"` 等错误编码静默二次解码，现在会在执行前阻断；同形安全 decoder 的 `.__call__` / `operator.call` 不误报。
- DOCX 行省略宽表审计：源审计现在会把 `w:gridAfter` 省略的行尾网格列计入 `max_table_columns` / `wide_table_count`，因此“1 个可见单元格 + 多个尾部省略列”的宽表会触发 `COMPLEX_TABLE_UNSUPPORTED` 复核提示，不会因为只数可见单元格而漏警告；生成端遇到省略区残留可见文本时会保留完整行并记录 text-guard manifest 计数，方便审计解释为什么没有恢复行省略。
- DOCX 修订包装表格审计：源审计现在按 Word 最终视图穿透 `w:ins` / `w:moveTo`、内容控件和透明容器中的表格行/单元格，宽表和异常表格风险不会因为行被包装而漏掉 `COMPLEX_TABLE_UNSUPPORTED`。
- DOCX 包装横向分节审计：源审计现在按 Word 最终视图穿透正文级 `w:sdt`、`w:customXml` / `w:smartTag` 和接受修订容器中的 `sectPr`，横向 section 内宽表会继续计入 `landscape_wide_table_risk_count`，不会因为 section break 被包装而漏掉人工复核提示。
- DOCX 删除修订表格审计：源审计的表格数量、合并计数、嵌套深度、宽表和异常合并风险都按 Word 最终视图可见表格计算；`w:del` / `w:moveFrom` 中已经删除的隐藏表格不会制造 `COMPLEX_TABLE_UNSUPPORTED` 假警告，但仍会保留 `TRACKED_CHANGES_PRESENT` 提醒用户确认修订。
- DOCX 矩形合并单元格：同一单元格同时跨多行、多列时，内容解析会归一成一个矩形 `table_merges`，避免把 `gridSpan` 和 `vMerge` 拆成重叠合并记录后让生成端重复 merge。
- DOCX 分裂 vMerge 延续修复：当上方单元格用 `gridSpan + vMerge restart` 表示二维合并、下方转换器却拆成多个空的 `vMerge continue` 单元格时，内容解析会把它们折叠为一个安全矩形 `table_merges`；如果任一延续单元格带有可见文字、图片、公式、嵌套表或注释锚点，则继续 fail-open 保留为可见单元格并进入复核链路。
- DOCX 旧式 hMerge 表格合并：兼容旧版/兼容模式 Word 的 `w:hMerge restart/continue`，内容解析会转为 `table_merges`，生成端用标准 `gridSpan` 输出；源审计和私有资料清分会把 `hMerge` 计入合并单元格风险并给出 `TABLE_MERGE_UNSUPPORTED` 复核提示。
- DOCX 旧式 hMerge + vMerge 矩形合并：兼容旧式横向合并和纵向合并组合成 2D 合并块时，内容解析会归一成一个矩形 `table_merges`，避免生成端重复执行多条重叠 merge。
- DOCX 非矩形旧式 hMerge + vMerge 冲突：如果旧式横向合并和纵向合并不是同宽矩形，引擎会 fail-open 保留 continuation 可见文本，只保留安全横向合并，并由源审计标记 `COMPLEX_TABLE_UNSUPPORTED` 复核。
- DOCX 混合 gridSpan + hMerge 重复编码：兼容转换工具同时写入现代 `gridSpan` 和旧式 `hMerge` 的表格，解析端不会把同一横向合并双计为更宽合并；当同一个 2D 合并块还叠加 `vMerge` 时，也会跳过空的重复 continuation 单元格并折叠成一个矩形 `table_merges`。如果重复 continuation 单元格里有可见文本、图片或嵌套内容，引擎会 fail-open 保留为可见单元格，不会当合并占位清空；源审计会用 `irregular_hmerge_count` 和 `visible_hmerge_continuation_count` 暴露复核提示，但不把这种兼容冗余误报为 `gridSpan` 越界。
- DOCX 省略列 + 半坏二维合并：当 `gridBefore` 省略了下一行的合并起点列，但源文件仍残留 `gridSpan + hMerge + vMerge` 组合时，封面字段探测不会因 `python-docx` 行枚举异常中断，内容解析会保留可见 continuation 文本，并只输出一个安全横向 `table_merges`，避免生成端执行重复或跨省略列的假 merge。
- DOCX vMerge 富内容延续单元格：普通 `vMerge continue` 文本继续按 Word 合并语义隐藏；但如果 continuation 单元格里有图片、公式、嵌套表或注释锚点，引擎会 fail-open 保留为可见单元格，源审计记录 `visible_vmerge_continuation_count` 并要求复核，避免富内容被静默吞掉；混合 `gridSpan`/`hMerge`/`vMerge` 叠加 `gridAfter` 尾部省略这类半坏合并中，异常 continuation 只要有可见内容也会计入 `visible_vmerge_continuations`，并且不泄漏单元格原文。
- DOCX 复杂表格小白指引：当 `COMPLEX_TABLE_UNSUPPORTED` 的 detail 出现 `visible_hmerge_continuations=N` 或 `visible_vmerge_continuations=N` 时，`qa_report` / `qa_repair_plan` 的下一步会明确提示用 Word/WPS 对照原文和最终 DOCX，重点核对这些带可见内容的合并延续单元格没有被吞掉、挪位或重复。
- DOCX 横向表格 section 保护：连续相邻且源页面设置签名一致的横向宽表，即使中间没有桥接说明，也会共用同一个 landscape section；带短说明时也可合并，纸张尺寸或页边距不同则保留独立横向 section，避免第二个表套用错误版心；当 section heading 本身是表题且第一项就是横向宽表时，表题会跟随首表进入同一个 landscape section，不会留在前一个纵向页；如果横向 section break 被正文级 `w:sdt`、`w:customXml` 或接受修订 `w:ins` 包装，解析端会按最终正文流展开后再把页面设置挂到对应宽表。
- DOCX 横向表格公式桥接保护：相邻横向宽表之间如果出现 display/block 公式（包括富文本段中的显示公式），引擎会结束前一个横向 section，先恢复纵向正文，再为后续宽表单独建立横向 section，避免公式段被静默套进横向页。
- DOCX 横向表格图片桥接保护：相邻横向宽表之间如果旧版/兼容结构里出现 `rich_text` 图片 run，引擎会先恢复纵向正文并渲染该图片，再为后续宽表单独建立横向 section，避免图片桥接段被静默留在横向页或图片丢失。
- DOCX 横向表格富文本内嵌表格保护：相邻横向宽表之间如果 `rich_text.items` 里携带块级小表，引擎会把该桥接内容放回纵向正文流并渲染小表，再为后续宽表单独建立横向 section，避免小表被当成短说明吞进横向页或静默丢失。
- DOCX 横向表格富文本内嵌代码保护：相邻横向宽表之间如果 `rich_text.items` 里携带块级代码，引擎会把桥接说明和代码块放回纵向正文流，再为后续宽表单独建立横向 section，避免代码被当成短说明留在横向页或静默丢失。
- DOCX 横向表格富文本内嵌题注保护：相邻横向宽表之间如果 `rich_text.items` 里携带图题或表题，引擎会把桥接说明和题注放回纵向正文流并渲染题注，再为后续宽表单独建立横向 section，避免题注被当成短说明吞进横向页或静默丢失。
- DOCX 横向表格富文本媒体/块级源序保护：当同一个 `rich_text.items` 桥接段里同时携带代码、图片、行内公式、题注、小表等内容时，引擎会按源顺序渲染这些媒体和块级子项，显式 `type=inline` 公式保持行内 OMML 形态，再为后续宽表单独建立横向 section，避免图片/公式收集器、display 公式误提升或按类型批量输出造成内容错位。
- DOCX 横向表格列表桥接保护：相邻横向宽表之间如果出现 `-` / 编号 / `[1]` / `（1）` / `一、` / `①` 等列表式正文，引擎会结束前一个横向 section，把列表段落放回纵向正文流，再为后续宽表单独建立横向 section，避免正文清单或参考文献式编号被误当成表格短注释。
- DOCX 长宽表跨页表头：自动横向保护的长宽表会默认给短文本第一行写入 Word `tblHeader`，跨页时表头可重复；首行含图片、公式、脚注/尾注、富文本或嵌套表时不会被自动猜成重复表头，短小表格和显式 `table_repeat_header_rows=0` 也不会被误加重复表头。
- visual QA 长表样张覆盖：图表清单、目录和点线页码清单不会被误当成真实图表风险页；复杂长文样张最多 8 页，会保留封面/目录/正文锚点、图/表/公式主风险页和长表续页，便于在 `visual_qa/samples/` 与 `visual_qa/wps/samples/` 中核对跨页表头、延续页和 Word/WPS 同页画面差异。`visual_report.md` 会直接列出这 8 个有界 Word/WPS 样张路径；PDF 模板解析和 visual QA 的 Poppler 工具查找都会跳过 PATH 中损坏或无法执行的 `pdfinfo` / `pdftotext` / `pdftoppm` shim，继续尝试后续可用候选。
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
