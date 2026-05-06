# Word 论文自动排版流水线

从模版 docx + 文本资料 docx 一键生成格式规范的论文 docx。

## 核心思路

**格式与内容分离**。模版提供字体/字号/行距/边距，文本资料提供章节/段落/图片/参考文献。流水线自动提取两者，生成 python-docx 构建脚本，最终输出排好版的 docx。

## 快速开始

```bash
# 1. 安装依赖
python -m pip install python-docx Pillow

# 2. 放入文件
#    模版 docx → Templates/
#    内容 docx → Inputs/

# 3. 交互模式（终端用户）
python run_pipeline.py
# → 自动扫描文件，编号列表供选择

# 4. 参数模式（脚本调用）
python run_pipeline.py --template 模版.docx --content 论文.docx
# → 指定文件名直接运行，无需交互
```

每次运行生成独立目录 `Outputs/{日期}_{内容名}/`，互不覆盖。

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
                                                         格式提取.md

Inputs/内容.docx ──→ [Phase 2] content_parser ──→ content.json
                    (图片 → Inputs/xxx/figures/)    内容提取.md

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

## 项目结构

```
├── run_pipeline.py              ← 一键入口
├── .claude/settings.json        ← 项目权限配置（可选）
├── .gitignore
├── Templates/模版放这里.txt
├── Inputs/文本资料放这里.txt
├── Outputs/                     ← 每次运行生成独立子目录
└── Paper_Project/
    └── Program/
        ├── pipeline/
        │   ├── format_extractor.py   ← Phase 1: 模版 → 格式 JSON
        │   ├── content_parser.py     ← Phase 2: 内容 → 结构化 JSON
        │   └── script_generator.py   ← Phase 3: JSON → 生成脚本
        ├── build_acta_manuscript.py  ← 参考：Acta Materialia 期刊格式
        ├── build_comprehensive_doc.py ← 参考：全功能演示
        └── master.py                 ← 参考：编排器骨架
```

## 功能

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

## 许可

Copyright © 2025 Youwei Zhang

本软件仅供个人学习和研究使用。**未经作者明确书面授权，禁止将本软件用于任何商业目的**，包括但不限于：将本软件或衍生作品作为商业产品、付费服务、SaaS 平台的一部分进行销售、出租、许可或分发。

如需商业使用授权，请联系作者。

This software is provided for personal learning and research purposes only. **Commercial use is prohibited without explicit written authorization from the author.**
