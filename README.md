# Word 论文自动排版流水线

从模版 docx + 文本资料 docx 一键生成格式规范的论文 docx。

## 核心思路

**格式与内容分离**。模版提供字体/字号/行距/边距，文本资料提供章节/段落/图片/参考文献。流水线自动提取两者，生成 python-docx 构建脚本，最终输出排好版的 docx。

## 快速开始

1. 把模版 docx 放入 `Templates/`，内容 docx 放入 `Inputs/`
2. 修改 `run_pipeline.py` 头部的 `TEMPLATE_DOCX` 和 `CONTENT_DOCX`
3. 运行：`python run_pipeline.py`
4. 查看 `Outputs/` 下的 MD 核对提取结果
5. 最终论文在 `Paper_Project/Manuscripts/最终论文.docx`

## 环境

- Python 3.10+
- `python -m pip install python-docx Pillow`

## 流水线四阶段

```
Templates/模版.docx ──→ format_extractor ──→ Outputs/format.json
                                              Outputs/格式提取.md

Inputs/内容.docx ──→ content_parser ──→ Outputs/content.json
                    (图片 → Inputs/xxx/figures/)    Outputs/内容提取.md

format.json ──┬──→ script_generator ──→ Program/build_generated.py
content.json ─┘

build_generated.py ──→ python 运行 ──→ Manuscripts/最终论文.docx
```

每个阶段内建双验证：提取器独立运行两次交叉比对。

## 项目结构

```
├── run_pipeline.py              ← 一键入口
├── Templates/模版放这里.txt
├── Inputs/文本资料放这里.txt
├── Outputs/输出放这里.txt
└── Paper_Project/
    └── Program/
        ├── pipeline/
        │   ├── format_extractor.py   ← 模版 → 格式 JSON
        │   ├── content_parser.py     ← 内容 → 结构化 JSON
        │   └── script_generator.py   ← JSON → build_generated.py
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
