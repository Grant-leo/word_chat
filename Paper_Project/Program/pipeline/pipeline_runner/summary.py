"""Completion-summary helpers for the one-click pipeline runner."""
from __future__ import annotations


def build_completion_summary(folder_name, output_docx, mode):
    active_mode = "普通用户" if mode == "user" else "开发者"
    return f"""
  输出目录: Outputs/{folder_name}/
    ├── 格式提取.md          <- 核对模版格式
    ├── 内容提取.md          <- 核对文本内容
    ├── format.json
    ├── content.json
    ├── template_profile.json <- 模板能力画像
    ├── template_requirements.json <- 机器可核查模板要求
    ├── workflow_mode.json <- 用户/开发者模式
    ├── build_manifest.json <- 正文元素渲染数量
    ├── qa_report.md       <- 自动检测报告
    ├── conformance_report.md <- strict DOCX/XML 合规报告
    ├── repair_loop_report.md <- 自动修复闭环报告（--auto-repair 时生成）
    ├── visual_report.md   <- PDF 渲染 QA（--qa-level visual 时生成）
    ├── build_generated.py   <- 生成脚本
    └── {output_docx}        <- 最终文件

  修复工作流:
    当前模式: {active_mode}
    普通用户模式: 修改本次输出目录中的 build_generated.py，然后重跑该脚本
    自动修复模式: 使用 --auto-repair 后读取 repair_loop_report.md 查看每轮修改和停止原因
    开发者模式: 修改 Paper_Project/Program/pipeline/ 下的核心脚本后重跑完整流水线
    目录: 生成脚本会优先用 Word COM 解析正文标题页码；不可用时仍保留静态目录行
"""


def print_completion_summary(folder_name, output_docx, mode):
    print(build_completion_summary(folder_name, output_docx, mode))
